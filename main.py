import customtkinter as ctk
from tkinter import messagebox
import serial
import serial.tools.list_ports
import threading
import json
import os
import time
from datetime import datetime
from configparser import ConfigParser

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

CONFIG_FILE = "config.ini"
DATA_FILE = "sensor_data.json"
LOG_FILE = "app.log"

class SensorMonitorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Sensor Monitor - 传感器数据监测")
        self.geometry("700x550")
        
        self.serial_port = None
        self.reading_thread = None
        self.running = False
        self.buffer = ""
        
        self.current_data = {
            "temperature": 0.0,
            "humidity": 0.0,
            "co2": 0,
            "frequency": 0,
            "wind_speed": 0.0,
            "dht_ok": False,
            "sgp_ok": False
        }
        
        self.data_buffer = {
            "temperature": [],
            "humidity": [],
            "co2": [],
            "frequency": [],
            "wind_speed": []
        }
        
        self.last_save_time = time.time()
        self.save_interval = 30
        
        self.last_receive_time = time.time()
        self.receive_timeout = 10
        
        self.load_config()
        self.create_widgets()
        self.refresh_ports()
        self.after(500, self.update_receive_status)
        
    def load_config(self):
        self.config = ConfigParser()
        if os.path.exists(CONFIG_FILE):
            self.config.read(CONFIG_FILE)
            self.baudrate = self.config.getint("serial", "baudrate", fallback=9600)
            self.selected_port = self.config.get("serial", "port", fallback="")
            self.save_interval = self.config.getint("data", "interval", fallback=30)
            self.frequency_ratio = self.config.getfloat("wind_speed", "frequency_ratio", fallback=75.0)
            if self.config.has_section("wind_speed") and self.config.has_option("wind_speed", "poly_coeffs"):
                self.poly_coeffs = [float(x) for x in self.config.get("wind_speed", "poly_coeffs").split(",")]
            else:
                self.poly_coeffs = [0.0, 1.0]
        else:
            self.baudrate = 9600
            self.selected_port = ""
            self.save_interval = 30
            self.frequency_ratio = 75.0
            self.poly_coeffs = [0.0, 1.0]
            
    def save_config(self):
        if not self.config.has_section("serial"):
            self.config.add_section("serial")
        if not self.config.has_section("data"):
            self.config.add_section("data")
        if not self.config.has_section("wind_speed"):
            self.config.add_section("wind_speed")
            
        self.config.set("serial", "port", self.selected_port)
        self.config.set("serial", "baudrate", str(self.baudrate))
        self.config.set("data", "interval", str(self.save_interval))
        self.config.set("wind_speed", "frequency_ratio", str(self.frequency_ratio))
        self.config.set("wind_speed", "poly_coeffs", ",".join(str(x) for x in self.poly_coeffs))
        
        with open(CONFIG_FILE, "w") as f:
            self.config.write(f)
            
    def calculate_wind_speed(self, frequency):
        rpm = (frequency * 1000.0 / self.frequency_ratio) * 60.0
        wind_speed = sum(coef * (rpm ** i) for i, coef in enumerate(self.poly_coeffs))
        return wind_speed
            
    def create_widgets(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        self.logo_label = ctk.CTkLabel(
            self.sidebar, 
            text="Sensor\nMonitor",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        self.port_label = ctk.CTkLabel(self.sidebar, text="串口:")
        self.port_label.grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")
        
        self.port_combo = ctk.CTkComboBox(self.sidebar, values=[""])
        self.port_combo.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        
        self.baudrate_label = ctk.CTkLabel(self.sidebar, text="波特率:")
        self.baudrate_label.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="w")
        
        self.baudrate_combo = ctk.CTkComboBox(
            self.sidebar, 
            values=["9600", "19200", "38400", "57600", "115200"],
            state="readonly"
        )
        self.baudrate_combo.set(str(self.baudrate))
        self.baudrate_combo.grid(row=4, column=0, padx=20, pady=5, sticky="ew")
        
        self.refresh_btn = ctk.CTkButton(
            self.sidebar, 
            text="刷新串口",
            command=self.refresh_ports,
            fg_color="gray"
        )
        self.refresh_btn.grid(row=5, column=0, padx=20, pady=5)
        
        self.connect_btn = ctk.CTkButton(
            self.sidebar, 
            text="连接",
            command=self.toggle_connection,
            fg_color="#2CC985"
        )
        self.connect_btn.grid(row=6, column=0, padx=20, pady=10)
        
        self.settings_btn = ctk.CTkButton(
            self.sidebar,
            text="设置",
            command=self.open_settings,
            fg_color="gray"
        )
        self.settings_btn.grid(row=7, column=0, padx=20, pady=5)
        
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="实时传感器数据",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.title_label.pack(pady=(10, 20))
        
        self.status_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.status_frame.pack(fill="x", padx=10)
        
        self.receive_indicator_frame = ctk.CTkFrame(self.status_frame, fg_color="transparent")
        self.receive_indicator_frame.pack(side="left")
        
        self.receive_dot = ctk.CTkCanvas(self.receive_indicator_frame, width=12, height=12, bg="#2A2A2A", highlightthickness=0)
        self.receive_dot.pack(side="left", padx=(0, 5))
        self.receive_circle = self.receive_dot.create_oval(1, 1, 12, 12, fill="#666666", outline="")
        self.receive_status_label = ctk.CTkLabel(
            self.receive_indicator_frame,
            text="等待数据",
            font=ctk.CTkFont(size=12)
        )
        self.receive_status_label.pack(side="left", padx=(0, 20))
        
        self.dht_status_label = ctk.CTkLabel(
            self.status_frame,
            text="DHT22: 未连接",
            font=ctk.CTkFont(size=14)
        )
        self.dht_status_label.pack(side="left", padx=10)
        
        self.sgp_status_label = ctk.CTkLabel(
            self.status_frame,
            text="SGP30: 未连接",
            font=ctk.CTkFont(size=14)
        )
        self.sgp_status_label.pack(side="right", padx=10)
        
        self.data_frame = ctk.CTkFrame(self.main_frame)
        self.data_frame.pack(pady=20, padx=10, fill="both", expand=True)
        
        self.temp_card = self.create_data_card(
            self.data_frame, "温度", "°C", "0.0", "#FF6B6B"
        )
        self.temp_card.pack(side="left", padx=10, pady=10, expand=True, fill="both")
        
        self.humidity_card = self.create_data_card(
            self.data_frame, "湿度", "%", "0.0", "#4ECDC4"
        )
        self.humidity_card.pack(side="left", padx=10, pady=10, expand=True, fill="both")
        
        self.co2_card = self.create_data_card(
            self.data_frame, "CO2", "ppm", "0", "#45B7D1"
        )
        self.co2_card.pack(side="left", padx=10, pady=10, expand=True, fill="both")
        
        self.wind_speed_card = self.create_data_card(
            self.data_frame, "风速", "m/s", "0.0", "#F9CA24"
        )
        self.wind_speed_card.pack(side="left", padx=10, pady=10, expand=True, fill="both")
        
        self.log_frame = ctk.CTkFrame(self.main_frame)
        self.log_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.log_label = ctk.CTkLabel(
            self.log_frame,
            text="数据日志",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.log_label.pack(pady=(5, 0))
        
        self.log_text = ctk.CTkTextbox(self.log_frame, height=100, font=("Consolas", 10))
        self.log_text.pack(padx=5, pady=5, fill="both", expand=True)
        
        self.info_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.info_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.save_info_label = ctk.CTkLabel(
            self.info_frame,
            text=f"数据保存间隔: {self.save_interval}秒",
            font=ctk.CTkFont(size=12)
        )
        self.save_info_label.pack(side="left")
        
        self.last_save_label = ctk.CTkLabel(
            self.info_frame,
            text="上次保存: --",
            font=ctk.CTkFont(size=12)
        )
        self.last_save_label.pack(side="right")
        
        if self.selected_port:
            self.port_combo.set(self.selected_port)
            
    def create_data_card(self, parent, title, unit, default_value, color):
        card = ctk.CTkFrame(parent, fg_color=("gray85", "gray17"))
        
        title_label = ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=14)
        )
        title_label.pack(pady=(10, 5))
        
        value_label = ctk.CTkLabel(
            card,
            text=default_value,
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color=color
        )
        value_label.pack(pady=5)
        
        unit_label = ctk.CTkLabel(
            card,
            text=unit,
            font=ctk.CTkFont(size=14)
        )
        unit_label.pack(pady=(0, 10))
        
        card.value_label = value_label
        return card
        
    def refresh_ports(self):
        ports = serial.tools.list_ports.comports()
        port_list = [p.device for p in ports]
        if not port_list:
            port_list = ["无可用串口"]
        self.port_combo.configure(values=port_list)
        if port_list:
            self.port_combo.set(port_list[0])
            
    def toggle_connection(self):
        if self.running:
            self.disconnect()
        else:
            self.connect()
            
    def connect(self):
        port = self.port_combo.get()
        if not port or port == "无可用串口":
            messagebox.showerror("错误", "请选择有效的串口")
            return
            
        try:
            self.baudrate = int(self.baudrate_combo.get())
            self.serial_port = serial.Serial(
                port=port,
                baudrate=self.baudrate,
                timeout=1
            )
            self.running = True
            self.last_receive_time = time.time()
            self.selected_port = port
            self.save_config()
            
            self.reading_thread = threading.Thread(target=self.read_from_serial, daemon=True)
            self.reading_thread.start()
            
            self.connect_btn.configure(text="断开", fg_color="#E84646")
            self.port_combo.configure(state="disabled")
            self.baudrate_combo.configure(state="disabled")
            self.log_message(f"已连接到 {port} @ {self.baudrate} bps")
            
        except Exception as e:
            messagebox.showerror("连接错误", f"无法打开串口: {str(e)}")
            
    def disconnect(self):
        self.running = False
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
            
        self.connect_btn.configure(text="连接", fg_color="#2CC985")
        self.port_combo.configure(state="normal")
        self.baudrate_combo.configure(state="normal")
        self.receive_dot.itemconfig(self.receive_circle, fill="#666666")
        self.receive_status_label.configure(text="未连接", text_color="gray")
        self.log_message("已断开连接")
        
    def read_from_serial(self):
        while self.running:
            try:
                if self.serial_port and self.serial_port.in_waiting:
                    data = self.serial_port.read(self.serial_port.in_waiting).decode('utf-8', errors='ignore')
                    self.buffer += data
                    
                    while '\n' in self.buffer:
                        line, self.buffer = self.buffer.split('\n', 1)
                        self.parse_data(line)
                        
                self.update_display()
                self.check_save_interval()
                time.sleep(0.1)
                
            except Exception as e:
                self.log_message(f"读取错误: {str(e)}")
                break
                
    def parse_data(self, line):
        line = line.strip()
        
        self.last_receive_time = time.time()
        
        if "Temperature:" in line:
            try:
                value = float(line.split(":")[1].strip().split()[0])
                self.current_data["temperature"] = value
                self.data_buffer["temperature"].append(value)
            except:
                pass
            
        elif "Humidity:" in line:
            try:
                value = float(line.split(":")[1].strip().split()[0])
                self.current_data["humidity"] = value
                self.data_buffer["humidity"].append(value)
            except:
                pass
                
        elif "CO2:" in line:
            try:
                value = int(line.split(":")[1].strip().split()[0])
                self.current_data["co2"] = value
                self.data_buffer["co2"].append(value)
            except:
                pass
                
        elif "AC Frequency:" in line:
            try:
                value = int(line.split(":")[1].strip().split()[0])
                self.current_data["frequency"] = value
                self.data_buffer["frequency"].append(value)
                wind_speed = self.calculate_wind_speed(value)
                self.current_data["wind_speed"] = wind_speed
                self.data_buffer["wind_speed"].append(wind_speed)
            except:
                pass
                
        elif "DHT22: OK" in line:
            self.current_data["dht_ok"] = True
        elif "DHT22: ERROR" in line:
            self.current_data["dht_ok"] = False
            
        elif "SGP30: OK" in line:
            self.current_data["sgp_ok"] = True
        elif "SGP30: ERROR" in line:
            self.current_data["sgp_ok"] = False
            
    def update_display(self):
        self.after(0, self._update_display_safe)
        
    def _update_display_safe(self):
        self.temp_card.value_label.configure(
            text=f"{self.current_data['temperature']:.1f}"
        )
        self.humidity_card.value_label.configure(
            text=f"{self.current_data['humidity']:.1f}"
        )
        self.co2_card.value_label.configure(
            text=str(self.current_data['co2'])
        )
        self.wind_speed_card.value_label.configure(
            text=f"{self.current_data['wind_speed']:.2f}"
        )
        
        if self.current_data["dht_ok"]:
            self.dht_status_label.configure(text="DHT22: 正常", text_color="#2CC985")
        else:
            self.dht_status_label.configure(text="DHT22: 异常", text_color="#E84646")
            
        if self.current_data["sgp_ok"]:
            self.sgp_status_label.configure(text="SGP30: 正常", text_color="#2CC985")
        else:
            self.sgp_status_label.configure(text="SGP30: 异常", text_color="#E84646")
            
    def check_save_interval(self):
        current_time = time.time()
        if current_time - self.last_save_time >= self.save_interval:
            self.save_average_data()
            self.last_save_time = current_time
            
    def save_average_data(self):
        if not any(len(v) > 0 for v in self.data_buffer.values()):
            return
            
        avg_temp = sum(self.data_buffer["temperature"]) / len(self.data_buffer["temperature"]) if self.data_buffer["temperature"] else 0
        avg_humidity = sum(self.data_buffer["humidity"]) / len(self.data_buffer["humidity"]) if self.data_buffer["humidity"] else 0
        avg_co2 = int(sum(self.data_buffer["co2"]) / len(self.data_buffer["co2"])) if self.data_buffer["co2"] else 0
        avg_frequency = int(sum(self.data_buffer["frequency"]) / len(self.data_buffer["frequency"])) if self.data_buffer["frequency"] else 0
        avg_wind_speed = sum(self.data_buffer["wind_speed"]) / len(self.data_buffer["wind_speed"]) if self.data_buffer["wind_speed"] else 0
        
        record = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "temperature": round(avg_temp, 2),
            "humidity": round(avg_humidity, 2),
            "co2": avg_co2,
            "frequency": avg_frequency,
            "wind_speed": round(avg_wind_speed, 2)
        }
        
        try:
            data_list = []
            if os.path.exists(DATA_FILE):
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    try:
                        data_list = json.load(f)
                    except:
                        data_list = []
                        
            data_list.append(record)
            
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(data_list, f, ensure_ascii=False, indent=2)
                
            self.last_save_label.configure(
                text=f"上次保存: {record['timestamp']}"
            )
            self.log_message(f"已保存平均值: 温度{avg_temp:.1f}°C, 湿度{avg_humidity:.1f}%, CO2{avg_co2}ppm, 频率{avg_frequency}Hz, 风速{avg_wind_speed:.2f}m/s")
            
        except Exception as e:
            self.log_message(f"保存错误: {str(e)}")
            
        self.data_buffer = {"temperature": [], "humidity": [], "co2": [], "frequency": [], "wind_speed": []}
        
    def open_settings(self):
        settings_win = ctk.CTkToplevel(self)
        settings_win.title("设置")
        settings_win.geometry("400x320")
        
        tabview = ctk.CTkTabview(settings_win)
        tabview.pack(pady=20, padx=20, fill="both", expand=True)
        
        tab_general = tabview.add("常规")
        tab_wind = tabview.add("风速计算")
        
        ctk.CTkLabel(tab_general, text="数据保存间隔 (秒):").pack(pady=(20, 5))
        interval_var = ctk.StringVar(value=str(self.save_interval))
        ctk.CTkEntry(tab_general, textvariable=interval_var, width=150).pack(pady=5)
        ctk.CTkLabel(
            tab_general,
            text="注: 每隔指定秒数计算缓冲区数据的平均值并保存",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        ).pack(pady=5)
        
        ctk.CTkLabel(tab_wind, text="频率-转速转换系数:").pack(pady=(20, 5))
        ctk.CTkLabel(
            tab_wind,
            text="注: 周期75ms对应100rpm，系数=75",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        ).pack()
        freq_ratio_var = ctk.StringVar(value=str(self.frequency_ratio))
        ctk.CTkEntry(tab_wind, textvariable=freq_ratio_var, width=150).pack(pady=5)
        
        ctk.CTkLabel(tab_wind, text="风速拟合多项式系数 (逗号分隔):").pack(pady=(20, 5))
        ctk.CTkLabel(
            tab_wind,
            text="格式: a0,a1,a2,... (按升幂排列)\n例如: 0,0.1 表示 wind_speed = 0.1 * rpm",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        ).pack()
        poly_var = ctk.StringVar(value=",".join(str(x) for x in self.poly_coeffs))
        ctk.CTkEntry(tab_wind, textvariable=poly_var, width=300).pack(pady=5)
        
        def save_settings():
            try:
                new_interval = int(interval_var.get())
                if new_interval < 5:
                    messagebox.showerror("错误", "间隔至少5秒")
                    return
                self.save_interval = new_interval
                
                new_freq_ratio = float(freq_ratio_var.get())
                self.frequency_ratio = new_freq_ratio
                
                new_poly = [float(x.strip()) for x in poly_var.get().split(",")]
                self.poly_coeffs = new_poly
                
                self.save_config()
                self.save_info_label.configure(text=f"数据保存间隔: {self.save_interval}秒")
                settings_win.destroy()
                messagebox.showinfo("成功", "设置已保存")
            except ValueError:
                messagebox.showerror("错误", "请输入有效数字")
                
        ctk.CTkButton(
            settings_win,
            text="保存",
            command=save_settings,
            fg_color="#2CC985"
        ).pack(pady=20)
        
    def update_receive_status(self):
        if not self.running:
            self.after(500, self.update_receive_status)
            return
            
        elapsed = time.time() - self.last_receive_time
        
        if elapsed < self.receive_timeout:
            self.receive_dot.itemconfig(self.receive_circle, fill="#2CC985")
            self.receive_status_label.configure(text="接收正常", text_color="#2CC985")
        else:
            self.receive_dot.itemconfig(self.receive_circle, fill="#E84646")
            self.receive_status_label.configure(text=f"无数据({int(elapsed)}s)", text_color="#E84646")
        
        self.after(500, self.update_receive_status)
        
    def log_message(self, message):
        self.after(0, self._log_message_safe, message)
        
    def _log_message_safe(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        
    def on_closing(self):
        self.running = False
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.destroy()

if __name__ == "__main__":
    app = SensorMonitorApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
