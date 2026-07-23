# -*- coding: utf-8 -*-
"""
네이버 금융(finance.naver.com) 종목 시세 자동 수집기 (업그레이드 버전)

stock_price_monitor.py 를 기반으로 다음 기능을 추가했다.

1. 이전에 입력했던 종목코드/알림가격/설정값을 JSON 설정 파일에 저장하고,
   프로그램을 다시 실행하면 초기 화면에 자동으로 불러와 채워준다.
2. 종목별로 "하한가"/"상한가"를 설정할 수 있고, 수집한 현재가가 그 값을
   벗어나면(이하로 내려가거나 이상으로 올라가면) 카카오톡("나에게 보내기")
   으로 알림 메시지를 전송한다.
3. 카카오톡 알림을 사용하려면 카카오 개발자 계정에서 발급받은 REST API 키와
   OAuth 인증을 통해 얻은 액세스 토큰이 필요하다. 프로그램 내에서 인증 URL을
   열고, 인가 코드를 붙여넣어 토큰을 발급/갱신할 수 있는 버튼을 제공한다.

주의: 이 스크립트는 네이버 금융 페이지의 HTML 구조 및 카카오 API 정책에
의존한다. 두 서비스의 정책/마크업이 바뀌면 동작하지 않을 수 있다.
"""

import json
import os
import re
import sys
import threading
import webbrowser
from datetime import datetime, timedelta

import pandas as pd
import requests
from bs4 import BeautifulSoup
from tkinter import (
    Tk, Toplevel, Frame, Label, Entry, Button, Text, Scrollbar, END, filedialog, messagebox, DISABLED, NORMAL
)
from tkinter import ttk

MAX_STOCKS = 10
DEFAULT_START_TIME = "09:00"
DEFAULT_END_TIME = "15:30"
DEFAULT_INTERVAL_MIN = "10"
DEFAULT_CODE_EXAMPLE = "122630"  # KODEX 레버리지 예시
DEFAULT_KAKAO_REDIRECT_URI = "https://localhost.com"

def get_app_dir() -> str:
    """실행 파일(exe) 또는 스크립트가 실제로 위치한 폴더를 반환한다.

    PyInstaller로 만든 exe(특히 onefile)를 실행하면 __file__은 실행할 때마다
    새로 생성되고 종료 시 삭제되는 임시 압축 해제 폴더(sys._MEIPASS)를 가리키므로,
    거기에 설정 파일을 저장하면 프로그램을 재시작할 때마다 사라진다.
    frozen 상태(exe로 빌드된 경우)에는 sys.executable이 위치한 폴더를 사용해
    설정이 exe 옆에 영구적으로 저장되도록 한다.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_FILE = os.path.join(get_app_dir(), "stock_monitor_config.json")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}

KAKAO_AUTHORIZE_URL = "https://kauth.kakao.com/oauth/authorize"
KAKAO_TOKEN_URL = "https://kauth.kakao.com/oauth/token"
KAKAO_MEMO_SEND_URL = "https://kapi.kakao.com/v2/api/talk/memo/default/send"


def fetch_price(code: str):
    """네이버 금융에서 종목명과 현재가를 조회한다.

    Returns:
        (name, price) 성공 시. 실패 시 (None, None).
    """
    url = f"https://finance.naver.com/item/main.naver?code={code}"
    try:
        response = requests.get(url, headers=HEADERS, timeout=5)
        response.raise_for_status()
    except requests.RequestException:
        return None, None

    soup = BeautifulSoup(response.text, "html.parser")

    # 종목명 파싱
    name = None
    name_tag = soup.select_one("div.wrap_company h2 a")
    if name_tag and name_tag.get_text(strip=True):
        name = name_tag.get_text(strip=True)
    else:
        title_tag = soup.select_one("title")
        if title_tag and title_tag.get_text(strip=True):
            name = title_tag.get_text(strip=True).split(":")[0].strip()

    # 현재가 파싱 (디지트 스프라이트 마크업 변화에 안전하도록 정규식으로 숫자만 추출)
    price = None
    price_tag = soup.select_one("p.no_today")
    if price_tag:
        match = re.search(r"[0-9][0-9,]*", price_tag.get_text())
        if match:
            try:
                price = int(match.group().replace(",", ""))
            except ValueError:
                price = None

    if name is None or price is None:
        return None, None
    return name, price


def load_or_init_df(file_path: str, columns: list) -> pd.DataFrame:
    if os.path.exists(file_path):
        df = pd.read_excel(file_path, engine="openpyxl")
        # 새로 추가된 종목 컬럼이 있으면 채워준다
        for col in columns:
            if col not in df.columns:
                df[col] = pd.NA
        return df[columns] if list(df.columns) == columns else df.reindex(columns=columns)
    return pd.DataFrame(columns=columns)


def append_row(file_path: str, columns: list, row: dict):
    df = load_or_init_df(file_path, columns)
    new_row = pd.DataFrame([row], columns=columns)
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(file_path, index=False, engine="openpyxl")


def build_ticks(start_time: str, end_time: str, interval_min: int):
    today = datetime.now().date()
    start_dt = datetime.combine(today, datetime.strptime(start_time, "%H:%M").time())
    end_dt = datetime.combine(today, datetime.strptime(end_time, "%H:%M").time())
    ticks = []
    t = start_dt
    while t <= end_dt:
        ticks.append(t)
        t += timedelta(minutes=interval_min)
    return ticks


def load_config(path: str) -> dict:
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_config(path: str, data: dict):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def send_kakao_message(access_token: str, text: str):
    """카카오톡 '나에게 보내기'로 텍스트 메시지를 전송한다.

    Returns:
        (success: bool, info: str)
    """
    template_object = {
        "object_type": "text",
        "text": text,
        "link": {
            "web_url": "https://finance.naver.com",
            "mobile_web_url": "https://finance.naver.com",
        },
    }
    headers = {"Authorization": f"Bearer {access_token}"}
    data = {"template_object": json.dumps(template_object, ensure_ascii=False)}
    try:
        response = requests.post(KAKAO_MEMO_SEND_URL, headers=headers, data=data, timeout=5)
    except requests.RequestException as exc:
        return False, str(exc)

    if response.status_code == 200:
        return True, "OK"
    return False, f"{response.status_code} {response.text}"


def issue_kakao_token(rest_api_key: str, redirect_uri: str, auth_code: str):
    data = {
        "grant_type": "authorization_code",
        "client_id": rest_api_key,
        "redirect_uri": redirect_uri,
        "code": auth_code,
    }
    response = requests.post(KAKAO_TOKEN_URL, data=data, timeout=5)
    return response.status_code == 200, response.json() if response.content else {}


def refresh_kakao_token(rest_api_key: str, refresh_token: str):
    data = {
        "grant_type": "refresh_token",
        "client_id": rest_api_key,
        "refresh_token": refresh_token,
    }
    response = requests.post(KAKAO_TOKEN_URL, data=data, timeout=5)
    return response.status_code == 200, response.json() if response.content else {}


class StockPriceMonitorApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("네이버 금융 종목 시세 자동 수집기 (카카오톡 알림)")

        self.config_data = load_config(CONFIG_FILE)

        self.code_entries = []
        self.lower_entries = []
        self.upper_entries = []
        self.stop_event = None
        self.worker_thread = None
        self.alert_state = {}
        # 실행 중(모니터링 중)에 하한가/상한가 입력칸을 수정하면 즉시 반영되도록,
        # 종목코드 -> (하한가 Entry, 상한가 Entry) 매핑을 저장해두고 매 조회 시점마다
        # 다시 읽어온다.
        self.threshold_widgets = {}
        # 수집 데이터 미리보기는 종목 수가 많아지면 메인 창을 옆으로 과도하게 늘리므로,
        # 별도의 창(Toplevel)에서 보여준다.
        self.preview_window = None
        self.tree = None

        self._build_gui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ GUI
    def _build_gui(self):
        stocks_cfg = self.config_data.get("stocks", [])

        codes_frame = Frame(self.root, padx=10, pady=10)
        codes_frame.grid(row=0, column=0, sticky="n")

        Label(codes_frame, text=f"종목코드 / 알림 가격 (최대 {MAX_STOCKS}개)", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 5)
        )
        Label(codes_frame, text="").grid(row=1, column=0)
        Label(codes_frame, text="종목코드").grid(row=1, column=1)
        Label(codes_frame, text="하한가(이하 알림)").grid(row=1, column=2)
        Label(codes_frame, text="상한가(이상 알림)").grid(row=1, column=3)

        for i in range(MAX_STOCKS):
            saved = stocks_cfg[i] if i < len(stocks_cfg) else {}

            Label(codes_frame, text=f"종목 {i + 1}").grid(row=i + 2, column=0, sticky="w")

            code_entry = Entry(codes_frame, width=12)
            code_entry.grid(row=i + 2, column=1, padx=2, pady=1)
            default_code = saved.get("code", "")
            if not default_code and i == 0 and not stocks_cfg:
                default_code = DEFAULT_CODE_EXAMPLE
            code_entry.insert(0, default_code)
            self.code_entries.append(code_entry)

            lower_entry = Entry(codes_frame, width=12)
            lower_entry.grid(row=i + 2, column=2, padx=2, pady=1)
            lower_entry.insert(0, saved.get("lower", ""))
            self.lower_entries.append(lower_entry)

            upper_entry = Entry(codes_frame, width=12)
            upper_entry.grid(row=i + 2, column=3, padx=2, pady=1)
            upper_entry.insert(0, saved.get("upper", ""))
            self.upper_entries.append(upper_entry)

        settings_frame = Frame(self.root, padx=10, pady=10)
        settings_frame.grid(row=0, column=1, sticky="n")

        Label(settings_frame, text="수집 시간 설정", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 5)
        )

        Label(settings_frame, text="시작시간 (HH:MM)").grid(row=1, column=0, sticky="w")
        self.start_time_entry = Entry(settings_frame, width=10)
        self.start_time_entry.insert(0, self.config_data.get("start_time", DEFAULT_START_TIME))
        self.start_time_entry.grid(row=1, column=1, pady=2)

        Label(settings_frame, text="종료시간 (HH:MM)").grid(row=2, column=0, sticky="w")
        self.end_time_entry = Entry(settings_frame, width=10)
        self.end_time_entry.insert(0, self.config_data.get("end_time", DEFAULT_END_TIME))
        self.end_time_entry.grid(row=2, column=1, pady=2)

        Label(settings_frame, text="간격(분)").grid(row=3, column=0, sticky="w")
        self.interval_entry = Entry(settings_frame, width=10)
        self.interval_entry.insert(0, self.config_data.get("interval_min", DEFAULT_INTERVAL_MIN))
        self.interval_entry.grid(row=3, column=1, pady=2)

        Label(settings_frame, text="저장 파일").grid(row=4, column=0, sticky="w", pady=(10, 0))
        self.file_path_entry = Entry(settings_frame, width=28)
        default_file = self.config_data.get("file_path") or f"stock_prices_{datetime.now():%Y%m%d}.xlsx"
        self.file_path_entry.insert(0, default_file)
        self.file_path_entry.grid(row=5, column=0, columnspan=2, sticky="w")

        Button(settings_frame, text="찾아보기", command=self._browse_file).grid(
            row=6, column=0, columnspan=2, sticky="w", pady=(2, 10)
        )

        button_frame = Frame(settings_frame)
        button_frame.grid(row=7, column=0, columnspan=2, sticky="w")

        self.start_button = Button(button_frame, text="시작", width=10, command=self._on_start)
        self.start_button.grid(row=0, column=0, padx=(0, 5))

        self.stop_button = Button(button_frame, text="정지", width=10, command=self._on_stop, state=DISABLED)
        self.stop_button.grid(row=0, column=1)

        # -------------------------------------------------------- 카카오 설정
        kakao_cfg = self.config_data.get("kakao", {})
        kakao_frame = Frame(self.root, padx=10, pady=10)
        kakao_frame.grid(row=0, column=2, sticky="n")

        Label(kakao_frame, text="카카오톡 알림 설정", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 5)
        )

        Label(kakao_frame, text="REST API 키").grid(row=1, column=0, sticky="w")
        self.kakao_rest_key_entry = Entry(kakao_frame, width=30)
        self.kakao_rest_key_entry.insert(0, kakao_cfg.get("rest_api_key", ""))
        self.kakao_rest_key_entry.grid(row=1, column=1, pady=2)

        Label(kakao_frame, text="Redirect URI").grid(row=2, column=0, sticky="w")
        self.kakao_redirect_entry = Entry(kakao_frame, width=30)
        self.kakao_redirect_entry.insert(0, kakao_cfg.get("redirect_uri", DEFAULT_KAKAO_REDIRECT_URI))
        self.kakao_redirect_entry.grid(row=2, column=1, pady=2)

        Button(kakao_frame, text="1) 인증 URL 열기", command=self._open_kakao_auth_url).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(6, 2)
        )

        Label(kakao_frame, text="인가 코드(code)").grid(row=4, column=0, sticky="w")
        self.kakao_auth_code_entry = Entry(kakao_frame, width=30)
        self.kakao_auth_code_entry.grid(row=4, column=1, pady=2)

        Button(kakao_frame, text="2) 토큰 발급", command=self._issue_kakao_token).grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(2, 10)
        )

        Label(kakao_frame, text="Access Token").grid(row=6, column=0, sticky="w")
        self.kakao_access_token_entry = Entry(kakao_frame, width=30)
        self.kakao_access_token_entry.insert(0, kakao_cfg.get("access_token", ""))
        self.kakao_access_token_entry.grid(row=6, column=1, pady=2)

        Label(kakao_frame, text="Refresh Token").grid(row=7, column=0, sticky="w")
        self.kakao_refresh_token_entry = Entry(kakao_frame, width=30)
        self.kakao_refresh_token_entry.insert(0, kakao_cfg.get("refresh_token", ""))
        self.kakao_refresh_token_entry.grid(row=7, column=1, pady=2)

        kakao_button_frame = Frame(kakao_frame)
        kakao_button_frame.grid(row=8, column=0, columnspan=2, sticky="w", pady=(4, 0))

        Button(kakao_button_frame, text="토큰 갱신", width=10, command=self._refresh_kakao_token).grid(
            row=0, column=0, padx=(0, 5)
        )
        Button(kakao_button_frame, text="테스트 발송", width=10, command=self._test_kakao_message).grid(
            row=0, column=1
        )

        table_frame = Frame(self.root, padx=10)
        table_frame.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 10))

        Label(table_frame, text="수집 데이터 미리보기", font=("", 10, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 5)
        )
        Label(
            table_frame,
            text="종목 수가 많아져도 메인 창이 옆으로 늘어나지 않도록, 별도의 창에서 표시합니다.",
            fg="gray30",
        ).grid(row=1, column=0, sticky="w")
        Button(table_frame, text="미리보기 창 열기", command=self._show_preview_window).grid(
            row=2, column=0, sticky="w", pady=(4, 0)
        )

        log_frame = Frame(self.root, padx=10, pady=10)
        log_frame.grid(row=2, column=0, columnspan=3, sticky="nsew")

        self.log_text = Text(log_frame, width=100, height=10, state=DISABLED)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.config(yscrollcommand=scrollbar.set)

    def _browse_file(self):
        path = filedialog.asksaveasfilename(
            title="저장할 엑셀 파일 선택",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.file_path_entry.delete(0, END)
            self.file_path_entry.insert(0, path)

    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=NORMAL)
        self.log_text.insert(END, f"[{timestamp}] {message}\n")
        self.log_text.see(END)
        self.log_text.config(state=DISABLED)

    def _ensure_preview_window(self):
        """수집 데이터 미리보기용 Treeview를 별도의 창(Toplevel)에 만든다.

        종목 수가 늘어날수록 열(column) 개수가 늘어나 표가 옆으로 넓어지는데,
        이를 메인 창 안에 두면 메인 창 자체가 모니터 밖으로 벗어나 버린다.
        따라서 미리보기는 독립적인 창으로 분리하고, 창 자체의 스크롤바로
        넓어진 표를 감당하도록 한다.
        """
        if self.preview_window is not None and self.preview_window.winfo_exists():
            return

        self.preview_window = Toplevel(self.root)
        self.preview_window.title("수집 데이터 미리보기")
        self.preview_window.geometry("900x400")
        # 닫기 버튼을 누르면 완전히 없애지 않고 숨겨서, 다음 수집 때 재사용한다.
        self.preview_window.protocol("WM_DELETE_WINDOW", self.preview_window.withdraw)

        tree_frame = Frame(self.preview_window, padx=10, pady=10)
        tree_frame.pack(fill="both", expand=True)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_frame, show="headings", height=15)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y = Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x = Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        tree_scroll_x.grid(row=1, column=0, sticky="ew")
        self.tree.config(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

    def _show_preview_window(self):
        self._ensure_preview_window()
        self.preview_window.deiconify()
        self.preview_window.lift()

    def _reset_table(self, columns: list):
        self._ensure_preview_window()
        self.preview_window.deiconify()
        self.preview_window.lift()
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            width = 70 if col == "시간" else 120
            self.tree.column(col, width=width, anchor="center")

    def _add_table_row(self, row: dict, columns: list):
        values = [("" if row.get(col) is None else row.get(col)) for col in columns]
        self.tree.insert("", END, values=values)
        children = self.tree.get_children()
        if children:
            self.tree.see(children[-1])

    # --------------------------------------------------------------- 카카오
    def _open_kakao_auth_url(self):
        rest_key = self.kakao_rest_key_entry.get().strip()
        redirect_uri = self.kakao_redirect_entry.get().strip()
        if not rest_key or not redirect_uri:
            messagebox.showerror("입력 오류", "REST API 키와 Redirect URI를 먼저 입력하세요.")
            return
        url = (
            f"{KAKAO_AUTHORIZE_URL}?client_id={rest_key}"
            f"&redirect_uri={redirect_uri}&response_type=code"
        )
        webbrowser.open(url)
        self.log("카카오 인증 URL을 브라우저에서 열었습니다. 로그인 후 리다이렉트된 주소의 'code' 값을 복사해 입력하세요.")

    def _issue_kakao_token(self):
        rest_key = self.kakao_rest_key_entry.get().strip()
        redirect_uri = self.kakao_redirect_entry.get().strip()
        auth_code = self.kakao_auth_code_entry.get().strip()
        if not rest_key or not redirect_uri or not auth_code:
            messagebox.showerror("입력 오류", "REST API 키, Redirect URI, 인가 코드를 모두 입력하세요.")
            return
        ok, result = issue_kakao_token(rest_key, redirect_uri, auth_code)
        if not ok:
            messagebox.showerror("토큰 발급 실패", f"토큰 발급에 실패했습니다: {result}")
            self.log(f"카카오 토큰 발급 실패: {result}")
            return
        self.kakao_access_token_entry.delete(0, END)
        self.kakao_access_token_entry.insert(0, result.get("access_token", ""))
        self.kakao_refresh_token_entry.delete(0, END)
        self.kakao_refresh_token_entry.insert(0, result.get("refresh_token", ""))
        self.log("카카오 액세스/리프레시 토큰을 발급받았습니다.")

    def _refresh_kakao_token(self):
        rest_key = self.kakao_rest_key_entry.get().strip()
        refresh_token = self.kakao_refresh_token_entry.get().strip()
        if not rest_key or not refresh_token:
            messagebox.showerror("입력 오류", "REST API 키와 Refresh Token을 입력하세요.")
            return
        ok, result = refresh_kakao_token(rest_key, refresh_token)
        if not ok:
            messagebox.showerror("토큰 갱신 실패", f"토큰 갱신에 실패했습니다: {result}")
            self.log(f"카카오 토큰 갱신 실패: {result}")
            return
        self.kakao_access_token_entry.delete(0, END)
        self.kakao_access_token_entry.insert(0, result.get("access_token", ""))
        if result.get("refresh_token"):
            self.kakao_refresh_token_entry.delete(0, END)
            self.kakao_refresh_token_entry.insert(0, result.get("refresh_token", ""))
        self.log("카카오 액세스 토큰을 갱신했습니다.")

    def _test_kakao_message(self):
        access_token = self.kakao_access_token_entry.get().strip()
        if not access_token:
            messagebox.showerror("입력 오류", "Access Token을 먼저 발급받으세요.")
            return
        ok, info = send_kakao_message(access_token, "[테스트] 종목 시세 알림 프로그램이 정상적으로 연결되었습니다.")
        if ok:
            messagebox.showinfo("전송 완료", "카카오톡으로 테스트 메시지를 전송했습니다.")
            self.log("카카오톡 테스트 메시지 전송 성공")
        else:
            messagebox.showerror("전송 실패", f"테스트 메시지 전송에 실패했습니다: {info}")
            self.log(f"카카오톡 테스트 메시지 전송 실패: {info}")

    # --------------------------------------------------------------- 설정 저장
    def _collect_stock_rows(self):
        """종목코드/하한가/상한가 입력값을 리스트[dict]로 수집한다(빈 종목코드는 건너뜀)."""
        rows = []
        for code_entry, lower_entry, upper_entry in zip(self.code_entries, self.lower_entries, self.upper_entries):
            code = code_entry.get().strip()
            if not code:
                continue
            rows.append({
                "code": code,
                "lower": lower_entry.get().strip(),
                "upper": upper_entry.get().strip(),
            })
        return rows

    def _save_current_config(self):
        data = {
            "stocks": self._collect_stock_rows(),
            "start_time": self.start_time_entry.get().strip(),
            "end_time": self.end_time_entry.get().strip(),
            "interval_min": self.interval_entry.get().strip(),
            "file_path": self.file_path_entry.get().strip(),
            "kakao": {
                "rest_api_key": self.kakao_rest_key_entry.get().strip(),
                "redirect_uri": self.kakao_redirect_entry.get().strip(),
                "access_token": self.kakao_access_token_entry.get().strip(),
                "refresh_token": self.kakao_refresh_token_entry.get().strip(),
            },
        }
        save_config(CONFIG_FILE, data)

    def _on_close(self):
        self._save_current_config()
        if self.preview_window is not None and self.preview_window.winfo_exists():
            self.preview_window.destroy()
        self.root.destroy()

    # --------------------------------------------------------------- 시작/정지
    def _on_start(self):
        codes = []
        thresholds = {}
        self.threshold_widgets = {}
        for code_entry, lower_entry, upper_entry in zip(self.code_entries, self.lower_entries, self.upper_entries):
            code = code_entry.get().strip()
            if not code:
                continue

            lower_text = lower_entry.get().strip()
            upper_text = upper_entry.get().strip()
            lower_value = None
            upper_value = None
            try:
                if lower_text:
                    lower_value = float(lower_text.replace(",", ""))
                if upper_text:
                    upper_value = float(upper_text.replace(",", ""))
            except ValueError:
                messagebox.showerror("입력 오류", f"종목 '{code}'의 하한가/상한가는 숫자로 입력하세요.")
                return

            if lower_value is not None and upper_value is not None and lower_value >= upper_value:
                messagebox.showerror("입력 오류", f"종목 '{code}'의 하한가는 상한가보다 작아야 합니다.")
                return

            codes.append(code)
            thresholds[code] = {"lower": lower_value, "upper": upper_value}
            # 모니터링 중 하한가/상한가 입력값을 바로 반영하기 위해 위젯 자체를 기억해둔다.
            self.threshold_widgets[code] = (lower_entry, upper_entry)

        if not codes:
            messagebox.showerror("입력 오류", "최소 1개 이상의 종목코드를 입력하세요.")
            return
        if len(codes) > MAX_STOCKS:
            messagebox.showerror("입력 오류", f"종목코드는 최대 {MAX_STOCKS}개까지 입력할 수 있습니다.")
            return

        start_time = self.start_time_entry.get().strip()
        end_time = self.end_time_entry.get().strip()
        interval_text = self.interval_entry.get().strip()
        file_path = self.file_path_entry.get().strip()

        try:
            start_dt = datetime.strptime(start_time, "%H:%M")
            end_dt = datetime.strptime(end_time, "%H:%M")
        except ValueError:
            messagebox.showerror("입력 오류", "시작/종료 시간은 HH:MM 형식으로 입력하세요. (예: 09:00)")
            return

        if end_dt <= start_dt:
            messagebox.showerror("입력 오류", "종료시간은 시작시간보다 이후여야 합니다.")
            return

        try:
            interval_min = int(interval_text)
            if interval_min <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("입력 오류", "간격(분)은 1 이상의 정수로 입력하세요.")
            return

        if not file_path:
            messagebox.showerror("입력 오류", "저장할 파일 경로를 입력하세요.")
            return

        file_path = os.path.abspath(file_path)

        self.log("입력된 종목코드의 유효성을 확인하는 중...")
        names_by_code = {}
        for code in codes:
            name, _ = fetch_price(code)
            if name is None:
                messagebox.showerror(
                    "종목코드 오류",
                    f"종목코드 '{code}'의 정보를 가져올 수 없습니다. 코드를 확인하세요.",
                )
                self.log(f"종목코드 '{code}' 조회 실패로 시작을 취소했습니다.")
                return
            names_by_code[code] = name
            threshold_info = thresholds[code]
            self.log(
                f"종목코드 '{code}' -> '{name}' 확인 완료 "
                f"(하한가: {threshold_info['lower']}, 상한가: {threshold_info['upper']})"
            )

        ticks = build_ticks(start_time, end_time, interval_min)
        columns = ["시간"] + [f"{names_by_code[c]}({c})" for c in codes]

        pending_ticks = [t for t in ticks if t >= datetime.now()]
        if not pending_ticks:
            messagebox.showerror(
                "시간 오류",
                f"설정한 수집 시간대({start_time}~{end_time})가 이미 지났습니다. "
                "시간을 다시 설정하세요.",
            )
            return

        self._reset_table(columns)

        try:
            if not os.path.exists(file_path):
                pd.DataFrame(columns=columns).to_excel(file_path, index=False, engine="openpyxl")
            self.log(f"저장 파일을 준비했습니다: {file_path}")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("파일 오류", f"엑셀 파일을 생성할 수 없습니다: {exc}")
            return

        # 종목코드/알림가격/설정값을 다음 실행을 위해 저장해둔다.
        self._save_current_config()

        # 알림 상태를 초기화한다. 설정값(하한가/상한가)이 바뀌지 않는 한
        # 동일 조건으로는 다시 알리지 않도록 마지막 설정값도 함께 기록한다.
        self.alert_state = {
            code: {
                "lower_last_value": None,
                "lower_alerted": False,
                "upper_last_value": None,
                "upper_alerted": False,
            }
            for code in codes
        }

        self.stop_event = threading.Event()
        self.worker_thread = threading.Thread(
            target=self._collection_loop,
            args=(codes, names_by_code, columns, ticks, file_path, self.stop_event),
            daemon=True,
        )
        self.worker_thread.start()

        self.start_button.config(state=DISABLED)
        self.stop_button.config(state=NORMAL)
        self.log(f"수집을 시작합니다. ({start_time} ~ {end_time}, {interval_min}분 간격, 총 {len(pending_ticks)}회)")

    def _on_stop(self):
        if self.stop_event:
            self.stop_event.set()
            self.log("정지 요청을 보냈습니다. 진행 중인 대기가 끝나면 종료됩니다.")
        self.stop_button.config(state=DISABLED)

    # --------------------------------------------------------------- 알림 체크
    def _get_current_threshold(self, code: str) -> dict:
        """하한가/상한가 입력칸의 현재 값을 읽어온다.

        모니터링이 진행되는 동안에도 사용자가 입력칸을 수정하면 다음 조회부터
        바로 반영되도록, 시작 시점에 저장해둔 고정값이 아니라 매번 위젯에서
        직접 읽어온다.
        """
        widgets = self.threshold_widgets.get(code)
        if not widgets:
            return {"lower": None, "upper": None}

        lower_entry, upper_entry = widgets
        lower_value = None
        upper_value = None
        try:
            lower_text = lower_entry.get().strip()
            if lower_text:
                lower_value = float(lower_text.replace(",", ""))
        except ValueError:
            lower_value = None
        try:
            upper_text = upper_entry.get().strip()
            if upper_text:
                upper_value = float(upper_text.replace(",", ""))
        except ValueError:
            upper_value = None
        return {"lower": lower_value, "upper": upper_value}

    def _check_and_notify(self, code: str, name: str, price: float, threshold: dict):
        lower = threshold.get("lower")
        upper = threshold.get("upper")
        state = self.alert_state.setdefault(
            code,
            {
                "lower_last_value": None,
                "lower_alerted": False,
                "upper_last_value": None,
                "upper_alerted": False,
            },
        )
        now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 설정값(하한가/상한가) 자체가 바뀌었을 때만 알림 가능 상태를 다시 켠다.
        # 값이 그대로면 가격이 범위를 여러 번 왔다갔다 해도 알림은 최초 1회만 보낸다.
        if lower != state["lower_last_value"]:
            state["lower_last_value"] = lower
            state["lower_alerted"] = False
        if upper != state["upper_last_value"]:
            state["upper_last_value"] = upper
            state["upper_alerted"] = False

        if lower is not None and price <= lower and not state["lower_alerted"]:
            self._send_alert(
                f"[하한가 알림] {name}({code})\n발생시각: {now_text}\n현재가 {price:,.0f}원이 "
                f"설정한 하한가 {lower:,.0f}원 이하로 내려갔습니다."
            )
            state["lower_alerted"] = True

        if upper is not None and price >= upper and not state["upper_alerted"]:
            self._send_alert(
                f"[상한가 알림] {name}({code})\n발생시각: {now_text}\n현재가 {price:,.0f}원이 "
                f"설정한 상한가 {upper:,.0f}원 이상으로 올라갔습니다."
            )
            state["upper_alerted"] = True

    def _send_alert(self, text: str):
        self.log(f"알림 조건 충족: {text.splitlines()[0] if text else ''}")

        # 수집 루프는 별도 스레드에서 실행되므로, UI(팝업)는 반드시 메인 스레드에서
        # root.after()를 통해 예약해야 한다.
        self.root.after(0, lambda t=text: self._show_alert_popup(t))

        access_token = self.kakao_access_token_entry.get().strip()
        if not access_token:
            self.log("카카오 Access Token이 없어 알림을 전송하지 못했습니다.")
            return

        ok, info = send_kakao_message(access_token, text)
        if ok:
            self.log("카카오톡 알림을 전송했습니다.")
            return

        # 토큰 만료(401) 가능성이 있으면 리프레시 토큰으로 갱신 후 1회 재시도한다.
        rest_key = self.kakao_rest_key_entry.get().strip()
        refresh_token = self.kakao_refresh_token_entry.get().strip()
        if rest_key and refresh_token and "401" in info:
            self.log("액세스 토큰이 만료된 것으로 보입니다. 토큰을 갱신하고 재시도합니다.")
            refreshed_ok, result = refresh_kakao_token(rest_key, refresh_token)
            if refreshed_ok:
                new_access_token = result.get("access_token", "")
                self.kakao_access_token_entry.delete(0, END)
                self.kakao_access_token_entry.insert(0, new_access_token)
                retry_ok, retry_info = send_kakao_message(new_access_token, text)
                if retry_ok:
                    self.log("토큰 갱신 후 카카오톡 알림을 전송했습니다.")
                    return
                self.log(f"카카오톡 알림 재전송 실패: {retry_info}")
                return
            self.log(f"카카오 토큰 갱신 실패: {result}")
            return

        self.log(f"카카오톡 알림 전송 실패: {info}")

    def _show_alert_popup(self, text: str):
        """가격 알림 발생 시 카카오톡 전송과 별도로 UI에 팝업창을 띄운다.

        수집 루프(백그라운드 스레드)에서 직접 호출하면 안 되며,
        반드시 root.after()를 통해 메인 스레드에서 실행되어야 한다.
        """
        popup = Toplevel(self.root)
        popup.title("가격 알림")
        popup.attributes("-topmost", True)
        popup.resizable(False, False)

        Label(
            popup, text=text, justify="left", padx=20, pady=15, font=("", 11),
        ).pack()
        Button(popup, text="확인", width=10, command=popup.destroy).pack(pady=(0, 12))

        popup.bell()
        popup.focus_force()

    # --------------------------------------------------------------- 수집 루프
    def _collection_loop(self, codes, names_by_code, columns, ticks, file_path, stop_event):
        now = datetime.now()
        pending_ticks = [t for t in ticks if t >= now]
        skipped = len(ticks) - len(pending_ticks)
        if skipped > 0:
            self.log(f"이미 지나간 {skipped}개의 수집 시각은 건너뜁니다.")

        for tick in pending_ticks:
            while not stop_event.is_set():
                remaining = (tick - datetime.now()).total_seconds()
                if remaining <= 0:
                    break
                stop_event.wait(min(1.0, remaining))

            if stop_event.is_set():
                break

            row = {"시간": tick.strftime("%H:%M")}
            for code in codes:
                name, price = fetch_price(code)
                col = f"{names_by_code[code]}({code})"
                if price is None:
                    row[col] = None
                    self.log(f"{tick.strftime('%H:%M')} - '{code}' 가격 조회 실패")
                else:
                    row[col] = price
                    # 하한가/상한가는 시작 시점 값이 아니라 지금 입력칸에 있는 값을 사용한다.
                    self._check_and_notify(code, names_by_code[code], price, self._get_current_threshold(code))

            self._add_table_row(row, columns)

            try:
                append_row(file_path, columns, row)
                self.log(f"{tick.strftime('%H:%M')} 시세를 저장했습니다. ({file_path})")
            except Exception as exc:  # noqa: BLE001
                self.log(f"엑셀 저장 중 오류가 발생했습니다: {exc}")

        if stop_event.is_set():
            self.log("수집이 정지되었습니다.")
        else:
            self.log("모든 수집이 완료되었습니다.")

        self.start_button.config(state=NORMAL)
        self.stop_button.config(state=DISABLED)


def main():
    root = Tk()
    StockPriceMonitorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
