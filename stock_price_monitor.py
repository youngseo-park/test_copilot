# -*- coding: utf-8 -*-
"""
네이버 금융(finance.naver.com) 종목 시세 자동 수집기

- 최대 15개 종목코드를 등록하면, 지정한 시간대(기본 09:00~15:30) 동안
  지정한 간격(기본 10분)마다 현재가를 조회하여 엑셀에 누적 저장한다.
- 엑셀은 "시간(행) x 종목명(코드)(열)" 형태의 피벗 테이블로 저장된다.
- 실행 방식: 스크립트를 한 번 실행한 뒤 "시작" 버튼을 누르면, 프로그램이
  종료시간까지 상주하면서 자동으로 수집을 반복한다. ("정지" 버튼으로 중단 가능)

주의: 이 스크립트는 네이버 금융 페이지의 HTML 구조에 의존한다.
네이버가 마크업을 변경하면 파싱이 실패할 수 있다.
"""

import os
import re
import threading
from datetime import datetime, timedelta

import pandas as pd
import requests
from bs4 import BeautifulSoup
from tkinter import (
    Tk, Frame, Label, Entry, Button, Text, Scrollbar, END, filedialog, messagebox, DISABLED, NORMAL
)
from tkinter import ttk

MAX_STOCKS = 15
DEFAULT_START_TIME = "09:00"
DEFAULT_END_TIME = "15:30"
DEFAULT_INTERVAL_MIN = "10"
DEFAULT_CODE_EXAMPLE = "122630"  # KODEX 레버리지 예시

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}


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


class StockPriceMonitorApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("네이버 금융 종목 시세 자동 수집기")

        self.code_entries = []
        self.stop_event = None
        self.worker_thread = None

        self._build_gui()

    def _build_gui(self):
        codes_frame = Frame(self.root, padx=10, pady=10)
        codes_frame.grid(row=0, column=0, sticky="n")

        Label(codes_frame, text="종목코드 (최대 15개)", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 5)
        )
        for i in range(MAX_STOCKS):
            Label(codes_frame, text=f"종목 {i + 1}").grid(row=i + 1, column=0, sticky="w")
            entry = Entry(codes_frame, width=12)
            entry.grid(row=i + 1, column=1, pady=1)
            if i == 0:
                entry.insert(0, DEFAULT_CODE_EXAMPLE)
            self.code_entries.append(entry)

        settings_frame = Frame(self.root, padx=10, pady=10)
        settings_frame.grid(row=0, column=1, sticky="n")

        Label(settings_frame, text="수집 시간 설정", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 5)
        )

        Label(settings_frame, text="시작시간 (HH:MM)").grid(row=1, column=0, sticky="w")
        self.start_time_entry = Entry(settings_frame, width=10)
        self.start_time_entry.insert(0, DEFAULT_START_TIME)
        self.start_time_entry.grid(row=1, column=1, pady=2)

        Label(settings_frame, text="종료시간 (HH:MM)").grid(row=2, column=0, sticky="w")
        self.end_time_entry = Entry(settings_frame, width=10)
        self.end_time_entry.insert(0, DEFAULT_END_TIME)
        self.end_time_entry.grid(row=2, column=1, pady=2)

        Label(settings_frame, text="간격(분)").grid(row=3, column=0, sticky="w")
        self.interval_entry = Entry(settings_frame, width=10)
        self.interval_entry.insert(0, DEFAULT_INTERVAL_MIN)
        self.interval_entry.grid(row=3, column=1, pady=2)

        Label(settings_frame, text="저장 파일").grid(row=4, column=0, sticky="w", pady=(10, 0))
        self.file_path_entry = Entry(settings_frame, width=28)
        default_file = f"stock_prices_{datetime.now():%Y%m%d}.xlsx"
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

        table_frame = Frame(self.root, padx=10)
        table_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 10))

        Label(table_frame, text="수집 데이터 미리보기", font=("", 10, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 5)
        )

        self.tree = ttk.Treeview(table_frame, show="headings", height=8)
        self.tree.grid(row=1, column=0, sticky="nsew")
        tree_scroll_y = Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_y.grid(row=1, column=1, sticky="ns")
        tree_scroll_x = Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        tree_scroll_x.grid(row=2, column=0, sticky="ew")
        self.tree.config(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        log_frame = Frame(self.root, padx=10, pady=10)
        log_frame.grid(row=2, column=0, columnspan=2, sticky="nsew")

        self.log_text = Text(log_frame, width=80, height=10, state=DISABLED)
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

    def _reset_table(self, columns: list):
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

    def _on_start(self):
        codes = []
        for entry in self.code_entries:
            code = entry.get().strip()
            if code:
                codes.append(code)

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
            self.log(f"종목코드 '{code}' -> '{name}' 확인 완료")

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
