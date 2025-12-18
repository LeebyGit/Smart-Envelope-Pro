import customtkinter as ctk
from tkinter import messagebox
import win32com.client as win32
import os
import requests
import re
import webbrowser  # 인터넷 창을 열어주는 도구 (기본 내장)
import urllib.parse # 한글 주소를 인터넷 주소로 바꿔주는 도구

# ==========================================
# 0. 디자인 및 폰트 설정
# ==========================================
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("dark-blue")

FONT_FAMILY = "Malgun Gothic"
APP_FONT = (FONT_FAMILY, 14)
LABEL_FONT = (FONT_FAMILY, 16, "bold")
SUBTITLE_FONT = (FONT_FAMILY, 12)

# ==========================================
# 1. API 설정
# ==========================================
JUSO_API_KEY = "YOUR_API_KEY_HERE="

def get_juso_data(keyword):
    if not keyword:
        return None, None
    try:
        url = "https://www.juso.go.kr/addrlink/addrLinkApi.do"
        params = {
            "confmKey": JUSO_API_KEY,
            "currentPage": 1,
            "countPerPage": 10,
            "keyword": keyword,
            "resultType": "json"
        }
        response = requests.get(url, params=params).json()
        if response['results']['common']['errorCode'] == "0":
            juso_list = response['results']['juso']
            if len(juso_list) > 0:
                return juso_list[0]['zipNo'], juso_list[0]['roadAddr']
            else:
                return None, None
        else:
            return None, None
    except Exception as e:
        print(f"Juso API Error: {e}")
        return None, None

# ==========================================
# 2. 전화번호 포맷팅
# ==========================================
def format_phone_number(number):
    if not number:
        return ""
    clean_num = re.sub(r'[^0-9]', '', number)
    if len(clean_num) == 11:
        return f"{clean_num[:3]}-{clean_num[3:7]}-{clean_num[7:]}"
    elif len(clean_num) == 10:
        if clean_num.startswith('02'):
            return f"{clean_num[:2]}-{clean_num[2:6]}-{clean_num[6:]}"
        else:
            return f"{clean_num[:3]}-{clean_num[3:6]}-{clean_num[6:]}"
    elif len(clean_num) == 9 and clean_num.startswith('02'):
        return f"{clean_num[:2]}-{clean_num[2:5]}-{clean_num[5:]}"
    elif len(clean_num) == 8:
        return f"{clean_num[:4]}-{clean_num[4:]}"
    return number

# ==========================================
# 3. 한글(HWP) 자동화 로직
# ==========================================
def fill_hwp_envelope():
    s_name = entry_s_name.get()
    s_addr = entry_s_addr.get()
    s_tel = format_phone_number(entry_s_tel.get())
    s_zip = entry_s_zip.get()

    r_name = entry_r_name.get()
    r_addr = entry_r_addr.get()
    r_tel = format_phone_number(entry_r_tel.get())
    r_zip = entry_r_zip.get()

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible = True
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        messagebox.showerror("오류", "한글(HWP)을 실행할 수 없습니다.")
        return

    hwp_path = os.path.join(os.getcwd(), "서류봉투(A4) 주소.hwp")
    if not os.path.exists(hwp_path):
        messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{hwp_path}")
        return

    hwp.Open(hwp_path)

    hwp.PutFieldText("send_name", s_name)
    hwp.PutFieldText("send_addr", s_addr)
    hwp.PutFieldText("send_tel", s_tel)
    if len(s_zip) == 5:
        for i in range(5):
            hwp.PutFieldText(f"s_zip{i+1}", s_zip[i])

    hwp.PutFieldText("recv_name", r_name)
    hwp.PutFieldText("recv_addr", r_addr)
    hwp.PutFieldText("recv_tel", r_tel)
    if len(r_zip) == 5:
        for i in range(5):
            hwp.PutFieldText(f"r_zip{i+1}", r_zip[i])

    messagebox.showinfo("완료", "봉투 출력이 준비되었습니다! 🖨️")

# ==========================================
# 4. 지도 열기 함수 (네이버/카카오)
# ==========================================
def open_map(address, service="naver"):
    """
    주소를 받아서 브라우저로 네이버/카카오 지도를 띄워주는 함수
    """
    if not address:
        messagebox.showwarning("경고", "주소가 입력되지 않았습니다.")
        return

    # 주소에 괄호나 특수문자가 있으면 검색이 잘 안될 수 있어서 제거
    clean_addr = address.split("(")[0].strip()
    
    # 인터넷 주소용으로 한글 변환
    encoded_addr = urllib.parse.quote(clean_addr)
    
    if service == "naver":
        url = f"https://map.naver.com/v5/search/{encoded_addr}"
    else: # kakao
        url = f"https://map.kakao.com/link/search/{encoded_addr}"
        
    webbrowser.open(url)

# 버튼 연결 함수들
def check_s_naver():
    open_map(entry_s_addr.get(), "naver")

def check_s_kakao():
    open_map(entry_s_addr.get(), "kakao")

def check_r_naver():
    open_map(entry_r_addr.get(), "naver")

def check_r_kakao():
    open_map(entry_r_addr.get(), "kakao")


def search_s():
    keyword = entry_s_search.get()
    zip_code, full_addr = get_juso_data(keyword)
    if zip_code:
        entry_s_zip.delete(0, ctk.END)
        entry_s_zip.insert(0, zip_code)
        entry_s_addr.delete(0, ctk.END)
        entry_s_addr.insert(0, full_addr)
    else:
        messagebox.showinfo("알림", "주소를 찾지 못했습니다.")

def search_r():
    keyword = entry_r_search.get()
    zip_code, full_addr = get_juso_data(keyword)
    if zip_code:
        entry_r_zip.delete(0, ctk.END)
        entry_r_zip.insert(0, zip_code)
        entry_r_addr.delete(0, ctk.END)
        entry_r_addr.insert(0, full_addr)
    else:
        messagebox.showinfo("알림", "주소를 찾지 못했습니다.")

# ==========================================
# 5. GUI 구성
# ==========================================
app = ctk.CTk()
app.title("Smart Envelope Pro")
app.geometry("580x750") # 지도 공간 뺐으므로 길이 조정
app.configure(fg_color="#F2F4F8")

ctk.CTkLabel(app, text="", height=10).pack()

# --- [보내는 사람] ---
card_s = ctk.CTkFrame(app, fg_color="white", corner_radius=20, border_width=0)
card_s.pack(padx=20, pady=10, fill="x")

header_s = ctk.CTkFrame(card_s, fg_color="transparent")
header_s.pack(fill="x", padx=20, pady=(20, 10))
ctk.CTkLabel(header_s, text="보내는 사람", font=LABEL_FONT, text_color="#3B8ED0").pack(side="left")
ctk.CTkLabel(header_s, text="Sender", font=SUBTITLE_FONT, text_color="#999999").pack(side="left", padx=5, pady=(5,0))

# 검색창
search_box_s = ctk.CTkFrame(card_s, fg_color="#F7F9FC", corner_radius=10)
search_box_s.pack(fill="x", padx=20, pady=5)
entry_s_search = ctk.CTkEntry(search_box_s, placeholder_text="동 이름 (예: 별양동)", font=APP_FONT, border_width=0, fg_color="transparent", height=40)
entry_s_search.pack(side="left", fill="x", expand=True, padx=10)
btn_search_s = ctk.CTkButton(search_box_s, text="검색", width=60, height=30, font=(FONT_FAMILY, 12, "bold"), command=search_s, fg_color="#E1E5EB", text_color="#333", hover_color="#D1D5DB")
btn_search_s.pack(side="right", padx=10)

# 입력폼
form_box_s = ctk.CTkFrame(card_s, fg_color="transparent")
form_box_s.pack(fill="x", padx=20, pady=(10, 5))
row_s1 = ctk.CTkFrame(form_box_s, fg_color="transparent")
row_s1.pack(fill="x", pady=5)
entry_s_name = ctk.CTkEntry(row_s1, placeholder_text="이름/직급", width=110, font=APP_FONT, height=35)
entry_s_name.pack(side="left", padx=(0, 5))
entry_s_tel = ctk.CTkEntry(row_s1, placeholder_text="연락처", width=140, font=APP_FONT, height=35)
entry_s_tel.pack(side="left", padx=(0, 5))
entry_s_zip = ctk.CTkEntry(row_s1, placeholder_text="우편번호", width=80, font=APP_FONT, height=35, fg_color="#EEF2FF", border_color="#C7D2FE")
entry_s_zip.pack(side="left")

ctk.CTkLabel(form_box_s, text="▼ 실제 봉투에 인쇄될 주소", font=(FONT_FAMILY, 11), text_color="#888888").pack(anchor="w", pady=(5,0))
entry_s_addr = ctk.CTkEntry(form_box_s, placeholder_text="주소 자동 입력", font=APP_FONT, height=40)
entry_s_addr.pack(fill="x", pady=2)

# [지도 버튼 영역]
map_box_s = ctk.CTkFrame(card_s, fg_color="transparent")
map_box_s.pack(fill="x", padx=20, pady=(5, 20))
ctk.CTkButton(map_box_s, text="N 네이버 지도 확인", width=120, height=30, fg_color="#03C75A", hover_color="#029f48", command=check_s_naver, font=(FONT_FAMILY, 12, "bold")).pack(side="left", padx=(0,5))
ctk.CTkButton(map_box_s, text="K 카카오맵 확인", width=120, height=30, fg_color="#FEE500", text_color="#000000", hover_color="#e6cf00", command=check_s_kakao, font=(FONT_FAMILY, 12, "bold")).pack(side="left")


# --- [받는 사람] ---
card_r = ctk.CTkFrame(app, fg_color="white", corner_radius=20, border_width=0)
card_r.pack(padx=20, pady=10, fill="x")

header_r = ctk.CTkFrame(card_r, fg_color="transparent")
header_r.pack(fill="x", padx=20, pady=(20, 10))
ctk.CTkLabel(header_r, text="받는 사람", font=LABEL_FONT, text_color="#E04F5F").pack(side="left")
ctk.CTkLabel(header_r, text="Receiver", font=SUBTITLE_FONT, text_color="#999999").pack(side="left", padx=5, pady=(5,0))

# 검색창
search_box_r = ctk.CTkFrame(card_r, fg_color="#FFF5F5", corner_radius=10)
search_box_r.pack(fill="x", padx=20, pady=5)
entry_r_search = ctk.CTkEntry(search_box_r, placeholder_text="동 이름 (예: 첨단로 39)", font=APP_FONT, border_width=0, fg_color="transparent", height=40)
entry_r_search.pack(side="left", fill="x", expand=True, padx=10)
btn_search_r = ctk.CTkButton(search_box_r, text="검색", width=60, height=30, font=(FONT_FAMILY, 12, "bold"), command=search_r, fg_color="#FEE2E2", text_color="#991B1B", hover_color="#FECACA")
btn_search_r.pack(side="right", padx=10)

# 입력폼
form_box_r = ctk.CTkFrame(card_r, fg_color="transparent")
form_box_r.pack(fill="x", padx=20, pady=(10, 5))
row_r1 = ctk.CTkFrame(form_box_r, fg_color="transparent")
row_r1.pack(fill="x", pady=5)
entry_r_name = ctk.CTkEntry(row_r1, placeholder_text="이름/직급", width=110, font=APP_FONT, height=35)
entry_r_name.pack(side="left", padx=(0, 5))
entry_r_tel = ctk.CTkEntry(row_r1, placeholder_text="연락처", width=140, font=APP_FONT, height=35)
entry_r_tel.pack(side="left", padx=(0, 5))
entry_r_zip = ctk.CTkEntry(row_r1, placeholder_text="우편번호", width=80, font=APP_FONT, height=35, fg_color="#FFF1F2", border_color="#FECDD3")
entry_r_zip.pack(side="left")

ctk.CTkLabel(form_box_r, text="▼ 실제 봉투에 인쇄될 주소", font=(FONT_FAMILY, 11), text_color="#888888").pack(anchor="w", pady=(5,0))
entry_r_addr = ctk.CTkEntry(form_box_r, placeholder_text="주소 자동 입력", font=APP_FONT, height=40)
entry_r_addr.pack(fill="x", pady=2)

# [지도 버튼 영역]
map_box_r = ctk.CTkFrame(card_r, fg_color="transparent")
map_box_r.pack(fill="x", padx=20, pady=(5, 20))
ctk.CTkButton(map_box_r, text="N 네이버 지도 확인", width=120, height=30, fg_color="#03C75A", hover_color="#029f48", command=check_r_naver, font=(FONT_FAMILY, 12, "bold")).pack(side="left", padx=(0,5))
ctk.CTkButton(map_box_r, text="K 카카오맵 확인", width=120, height=30, fg_color="#FEE500", text_color="#000000", hover_color="#e6cf00", command=check_r_kakao, font=(FONT_FAMILY, 12, "bold")).pack(side="left")

# --- 실행 버튼 ---
btn_run = ctk.CTkButton(app, text="✉️ 서류봉투 생성하기", font=(FONT_FAMILY, 18, "bold"), height=55, corner_radius=28, fg_color="#2563EB", hover_color="#1D4ED8", command=fill_hwp_envelope)
btn_run.pack(padx=20, pady=(10, 30), fill="x")

app.mainloop()
