import cv2
import numpy as np
import tkinter as tk
from tkinter import messagebox, scrolledtext
from PIL import Image, ImageTk
import threading
import pyodbc  # MSSQL 연결을 위한 라이브러리
import time


class ColorAnalyzeSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Color Analyze System")
        self.root.geometry("1000x600")

        # 버튼 영역
        self.button_frame = tk.Frame(self.root)
        self.button_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nw")

        self.start_button = tk.Button(self.button_frame, text="시작", font=("Arial", 12), command=self.start_detection)
        self.start_button.pack(side="left", padx=5)

        self.stop_button = tk.Button(self.button_frame, text="중지", font=("Arial", 12), command=self.stop_detection, state="disabled")
        self.stop_button.pack(side="left", padx=5)

        self.exit_button = tk.Button(self.button_frame, text="종료", font=("Arial", 12), command=self.close)
        self.exit_button.pack(side="left", padx=5)

        # 웹캠 출력 영역
        self.video_frame = tk.Frame(self.root, width=650, height=500, bg="white", relief="solid", bd=2)
        self.video_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nw")

        self.video_label = tk.Label(self.video_frame, text="화면 준비 중...", font=("Arial", 14), bg="white")
        self.video_label.place(relx=0.5, rely=0.5, anchor="center")

        # 마우스 이벤트 바인딩
        self.video_label.bind("<Button-1>", self.on_mouse_down)
        self.video_label.bind("<B1-Motion>", self.on_mouse_move)
        self.video_label.bind("<ButtonRelease-1>", self.on_mouse_up)

        # 로그 출력 영역
        self.right_frame = tk.Frame(self.root, width=300, height=500)
        self.right_frame.grid(row=1, column=1, padx=10, pady=10, sticky="se")

        self.log_label = tk.Label(self.right_frame, text="로그 출력", font=("Arial", 12))
        self.log_label.pack(pady=5)

        self.log_box = scrolledtext.ScrolledText(self.right_frame, width=40, height=10, state="disabled")
        self.log_box.pack()

        # 변수 초기화
        self.cap = None
        self.running = False
        self.ROI_SIZE = 200
        self.roi_x, self.roi_y = 200, 200
        self.dragging = False
        self.avg_color = (0, 0, 0)  # 초기 RGB 값

        # MSSQL 연결 초기화
        self.conn = None
        self.connect_to_database()

    # database 연결
    def connect_to_database(self):
        try:
            self.conn = pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=118.39.27.73;'
                'DATABASE=SF_JFFAB;'
                'UID=sa;'
                'PWD=hntadmin;'
            )
            self.log_message("MSSQL 데이터베이스 연결 성공")
        except Exception as e:
            messagebox.showerror("Database Error", f"MSSQL 연결 실패: {e}")

    # 마우스 콜백 함수
    def on_mouse_down(self, event):
        if self.roi_x <= event.x <= self.roi_x + self.ROI_SIZE and self.roi_y <= event.y <= self.roi_y + self.ROI_SIZE:
            self.dragging = True

    def on_mouse_move(self, event):
        if self.dragging:
            self.roi_x, self.roi_y = event.x - self.ROI_SIZE // 2, event.y - self.ROI_SIZE // 2

    def on_mouse_up(self, event):
        self.dragging = False

    def start_detection(self):
        self.running = True
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")

        self.cap = cv2.VideoCapture(0)
        self.thread = threading.Thread(target=self.update_frame)
        self.thread.start()

        self.proc_thread = threading.Thread(target=self.run_procedure)
        self.proc_thread.daemon = True
        self.proc_thread.start()

    def stop_detection(self):
        self.running = False
        if self.cap:
            self.cap.release()
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")

    def update_frame(self):
        while self.running and self.cap.isOpened():
            ret, frame = self.cap.read()
            if not ret:
                break

            # ROI 설정
            roi_x = max(0, min(self.roi_x, frame.shape[1] - self.ROI_SIZE))
            roi_y = max(0, min(self.roi_y, frame.shape[0] - self.ROI_SIZE))
            roi = frame[roi_y:roi_y + self.ROI_SIZE, roi_x:roi_x + self.ROI_SIZE]
            avg_color = cv2.mean(roi)[:3]
            self.avg_color = (int(avg_color[2]), int(avg_color[1]), int(avg_color[0]))

            r, g, b = int(avg_color[2]), int(avg_color[1]), int(avg_color[0])
            hex_color = "#{:02X}{:02X}{:02X}".format(r, g, b)

            # ROI 표시
            cv2.rectangle(frame, (roi_x, roi_y), (roi_x + self.ROI_SIZE, roi_y + self.ROI_SIZE), (0, 255, 0), 2)
            text = f"RGB: {self.avg_color} HEX: {hex_color}"
            cv2.putText(frame, text, (roi_x, roi_y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)

            # Tkinter로 프레임 표시
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame_rgb)
            imgtk = ImageTk.PhotoImage(image=img)

            self.video_label.config(image=imgtk)
            self.video_label.image = imgtk

            # 로그 출력
            self.log_message(text)

        self.video_label.config(image="", text="화면 준비 중...")

    def run_procedure(self):
        time.sleep(1) # 초기에 프로시저가 바로 실행되면서 RGB( 0, 0, 0)이 들어감을 방지.
        while self.running:
            self.execute_procedure()
            time.sleep(10)

    def execute_procedure(self):
        if self.conn:
            try:
                r, g, b = int(self.avg_color[2]), int(self.avg_color[1]), int(self.avg_color[0])
                cursor = self.conn.cursor()
                cursor.execute("EXEC 색상추출_등록 ?, ?, ?", r, g, b)
                self.conn.commit()
                self.log_message(f"프로시저 실행: RGB({r}, {g}, {b})")
            except Exception as e:
                self.log_message(f"프로시저 실행 오류: {e}")

    def log_message(self, message):
        self.log_box.config(state="normal")
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.yview(tk.END)
        self.log_box.config(state="disabled")

    def close(self):
        self.stop_detection()
        self.root.quit()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = ColorAnalyzeSystem(root)
    root.mainloop()
