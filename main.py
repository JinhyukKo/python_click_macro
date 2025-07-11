import ctypes
import threading
import time
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import pyautogui  # 현재 마우스 위치 가져올 때만 사용
import random

# Windows API 상수
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_ABSOLUTE = 0x8000
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004

DELAY = 0.05
registered_delay = 0
# 화면 해상도 구하기
user32 = ctypes.windll.user32
screen_width = user32.GetSystemMetrics(0)
screen_height = user32.GetSystemMetrics(1)

# 마우스 입력 구조체 정의
class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = [("mi", MOUSEINPUT)]
    _anonymous_ = ("_input",)
    _fields_ = [("type", ctypes.c_ulong), ("_input", _INPUT)]

def send_mouse_click_lowlevel(x, y):
    # 좌표 0~65535 정규화
    normalized_x = int(x * 65535 / screen_width)
    normalized_y = int(y * 65535 / screen_height)

    # 마우스 이동 (절대 좌표)
    mi_move = MOUSEINPUT(dx=normalized_x, dy=normalized_y, mouseData=0,
                         dwFlags=MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, time=0, dwExtraInfo=None)
    inp_move = INPUT(type=0, mi=mi_move)
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp_move), ctypes.sizeof(inp_move))

    # 왼쪽 버튼 다운
    mi_down = MOUSEINPUT(dx=normalized_x, dy=normalized_y, mouseData=0,
                         dwFlags=MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_ABSOLUTE, time=0, dwExtraInfo=None)
    inp_down = INPUT(type=0, mi=mi_down)
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp_down), ctypes.sizeof(inp_down))

    # 왼쪽 버튼 업
    mi_up = MOUSEINPUT(dx=normalized_x, dy=normalized_y, mouseData=0,
                       dwFlags=MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE, time=0, dwExtraInfo=None)
    inp_up = INPUT(type=0, mi=mi_up)
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp_up), ctypes.sizeof(inp_up))


# 클릭 예약 리스트
click_queue = []

# 예약된 클릭 실행
def process_click_queue():
    for  target_time, x, y, delay_range in click_queue:
        adjusted_time = target_time + datetime.timedelta(milliseconds=delay_range)
        while True:
            now = datetime.datetime.now()
            if now >= adjusted_time:
                send_mouse_click_lowlevel(x, y)
                break
            time.sleep(DELAY)  # 너무 짧은 sleep으로 CPU 과다 사용 방지
    messagebox.showinfo("완료", "모든 클릭 예약을 완료했습니다.")

def add_to_queue():
    try:
        hour = int(hour_cb.get())
        minute = int(minute_cb.get())
        second = int(second_cb.get())
        millisecond = int(millisecond_cb.get())

        now = datetime.datetime.now()
        target_time = datetime.datetime(year=now.year, month=now.month, day=now.day,
                                        hour=hour, minute=minute, second=second, microsecond=millisecond * 1000)
        if target_time < now:
            target_time += datetime.timedelta(days=1)

        x = int(entry_x.get())
        y = int(entry_y.get())

        click_queue.append((target_time, x, y, registered_delay))
        update_queue_display()
        messagebox.showinfo("추가됨", f"{target_time.strftime('%H:%M:%S.%f')[:-3]} ±{registered_delay}ms에 ({x},{y}) 클릭 예약됨.")
    except Exception as e:
        messagebox.showerror("오류", f"입력 오류: {e}")


# 예약 리스트 화면 갱신
def update_queue_display():
    text_queue.delete("1.0", tk.END)
    for i, (t, x, y,delay) in enumerate(click_queue, 1):
        text_queue.insert(tk.END, f"{i}. {t.strftime('%H:%M:%S.%f')[:-3]} ±{delay}ms - ({x}, {y})\n")


# 예약 실행 시작
def start_queue():
    if not click_queue:
        messagebox.showwarning("경고", "예약된 클릭이 없습니다.")
        return
    threading.Thread(target=process_click_queue, daemon=True).start()
    messagebox.showinfo("시작됨", "예약된 클릭 실행을 시작합니다.")
    root.configure(bg='lightblue')

# 현재 마우스 위치 가져오기 (pyautogui 사용)
def get_mouse_position():
    pos = pyautogui.position()
    entry_x.delete(0, tk.END)
    entry_x.insert(0, str(pos.x))
    entry_y.delete(0, tk.END)
    entry_y.insert(0, str(pos.y))
# 삭제 함수
def delete_from_queue():
    try:
        idx = int(entry_delete.get())
        if 1 <= idx <= len(click_queue):
            del click_queue[idx - 1]
            update_queue_display()
            messagebox.showinfo("삭제됨", f"{idx}번 예약이 삭제되었습니다.")
            entry_delete.delete(0, tk.END)
        else:
            messagebox.showwarning("경고", "유효한 번호를 입력하세요.")
    except ValueError:
        messagebox.showwarning("경고", "숫자를 입력하세요.")

def clear_queue_and_restart():
    global click_queue
    if messagebox.askyesno("확인", "모든 예약을 삭제하고 다시 시작하시겠습니까?"):
        click_queue = []
        update_queue_display()
        messagebox.showinfo("초기화됨", "모든 예약이 삭제되었습니다.")
        root.configure(bg='SystemButtonFace')  # 배경색 원래대로
        
def register_delay():
    global registered_delay
    try:
        val = int(entry_delay.get())
        if val < 0:
            raise ValueError("0 이상의 숫자만 입력하세요.")
        registered_delay = val
        messagebox.showinfo("딜레이 등록", f"딜레이 {registered_delay}ms 가 등록되었습니다.")
        label_registered_delay.config(text=f"현재 등록된 딜레이: {registered_delay} ms")
    except Exception as e:
        messagebox.showerror("오류", f"딜레이 입력 오류: {e}")


# GUI 생성
root = tk.Tk()
root.title("상준 매크로")
root.geometry("370x800")

frame_time = tk.Frame(root)
frame_time.pack(pady=5)

hour_cb = ttk.Combobox(frame_time, values=[f"{i:02d}" for i in range(24)], width=3)
hour_cb.set("12")
hour_cb.pack(side=tk.LEFT)
tk.Label(frame_time, text=":").pack(side=tk.LEFT)

minute_cb = ttk.Combobox(frame_time, values=[f"{i:02d}" for i in range(60)], width=3)
minute_cb.set("00")
minute_cb.pack(side=tk.LEFT)
tk.Label(frame_time, text=":").pack(side=tk.LEFT)

second_cb = ttk.Combobox(frame_time, values=[f"{i:02d}" for i in range(60)], width=3)
second_cb.set("00")
second_cb.pack(side=tk.LEFT)

# 시간 선택 콤보박스 아래에 추가
tk.Label(frame_time, text=".").pack(side=tk.LEFT)

millisecond_cb = ttk.Combobox(frame_time, values=[f"{i:03d}" for i in range(1000)], width=4)
millisecond_cb.set("000")
millisecond_cb.pack(side=tk.LEFT)


tk.Label(root, text="X 좌표:").pack()
entry_x = tk.Entry(root)
entry_x.pack()



tk.Label(root, text="Y 좌표:").pack()
entry_y = tk.Entry(root)
entry_y.pack()

tk.Label(root, text="딜레이 범위 (ms):").pack()
entry_delay = tk.Entry(root)
entry_delay.insert(0, "0")  # 기본값 0
entry_delay.pack()


btn_register_delay = tk.Button(root, text="딜레이 등록", command=register_delay)
btn_register_delay.pack(pady=5)

label_registered_delay = tk.Label(root, text=f"현재 등록된 딜레이: {registered_delay} ms")
label_registered_delay.pack()


# 삭제 입력 라벨과 입력창
tk.Label(root, text="삭제할 예약 번호 입력:").pack()
entry_delete = tk.Entry(root)
entry_delete.pack()
tk.Button(root, text="예약 삭제", command=delete_from_queue).pack(pady=5)

tk.Button(root, text="현재 마우스 위치 불러오기 (ctrl)", command=get_mouse_position).pack(pady=5)
tk.Button(root, text="예약 추가 (Alt)", command=add_to_queue).pack(pady=5)
tk.Button(root, text="START", command=start_queue,bg="red", fg="white").pack(pady=10)

tk.Button(root, text="전체 삭제 후 다시 시작", command=clear_queue_and_restart).pack(pady=5)



tk.Label(root, text="예약된 클릭 목록:").pack()

text_queue = tk.Text(root, height=50, width=45)
text_queue.pack()

root.bind("<Control_L>", lambda event: get_mouse_position())
root.bind("<Control_L>", lambda event: get_mouse_position())
root.bind('<Alt_L>', lambda event: add_to_queue())



root.mainloop()
