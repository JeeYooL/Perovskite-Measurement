
import serial
import tkinter as tk
from tkinter import messagebox
import time

# --- 설정부 ---
PORT = 'COM3'  # 박사님의 아두이노 포트 번호로 수정하세요
BAUD = 9600
# --------------

class RelayControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Solar Cell 16-CH Selector (Hyoungwoo Kwon)")
        self.root.geometry("400x550")
        
        self.current_ch = None # 현재 켜져 있는 채널 저장
        
        try:
            self.ser = serial.Serial(PORT, BAUD, timeout=1)
            time.sleep(2) # 연결 안정화
        except Exception as e:
            messagebox.showerror("연결 오류", f"아두이노를 찾을 수 없습니다: {e}")
            self.root.destroy()
            return

        self.create_widgets()

    def create_widgets(self):
        # 상단 타이틀
        title_label = tk.Label(self.root, text="Perovskite Solar Cell Measurement", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)

        # 버튼 그리드 프레임
        frame = tk.Frame(self.root)
        frame.pack(pady=10)

        self.buttons = []
        for i in range(1, 17):
            # 4x4 그리드 배치 (4개 소자 x 4개 픽셀 구조 반영)
            btn = tk.Button(frame, text=f"CH {i:02d}", width=8, height=3,
                           command=lambda idx=i: self.toggle_relay(idx))
            row = (i-1) // 4
            col = (i-1) % 4
            btn.grid(row=row, column=col, padx=5, pady=5)
            self.buttons.append(btn)

        # 전체 끄기 버튼
        off_btn = tk.Button(self.root, text="ALL OFF", bg="red", fg="white", 
                           command=self.all_off, width=20, height=2)
        off_btn.pack(pady=20)

        self.status_var = tk.StringVar(value="Status: All Channels OFF")
        status_label = tk.Label(self.root, textvariable=self.status_var, fg="blue")
        status_label.pack()

    def toggle_relay(self, ch_idx):
        # 만약 이미 켜져 있는 채널을 다시 누르면 끔
        if self.current_ch == ch_idx:
            self.all_off()
        else:
            # 새로운 채널을 켜기 (아두이노 코드가 숫자를 받으면 해당 채널만 켜도록 설계됨)
            try:
                self.ser.write(str(ch_idx).encode())
                self.update_button_colors(ch_idx)
                self.current_ch = ch_idx
                self.status_var.set(f"Status: Channel {ch_idx} is ON")
            except Exception as e:
                print(f"전송 에러: {e}")

    def all_off(self):
        try:
            self.ser.write(b'0') # 0은 모두 끄기 신호
            self.update_button_colors(None)
            self.current_ch = None
            self.status_var.set("Status: All Channels OFF")
        except Exception as e:
            print(f"전송 에러: {e}")

    def update_button_colors(self, active_ch):
        for i, btn in enumerate(self.buttons, 1):
            if i == active_ch:
                btn.config(bg="yellow") # 켜진 버튼은 노란색
            else:
                btn.config(bg="lightgray") # 꺼진 버튼은 회색

if __name__ == "__main__":
    root = tk.Tk()
    app = RelayControlApp(root)
    root.mainloop()