import gspread
import qrcode
import os
import urllib.parse
from google.oauth2.service_account import Credentials

# -------------------------
# ① 구글 시트 설정
# -------------------------
SHEET_KEY = "15q9wgugYHKXbaYE5tZibAd1TGUG_bkQbnicpGT6AYq4"  # 시트 URL 중 /d/와 /edit 사이에 있는 긴 문자열
SHEET_NAME = "도서목록"  # 시트 탭 이름

# -------------------------
# ② 구글 폼 및 entry ID
# -------------------------
FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSdGuoHeUW-37RFCBgMH7ZUNL0tt_yHQiIFMrif85mrV428Omg/viewform"
ENTRY_IDS = {
    "code": "entry.32105598",     # 도서코드
    "title": "entry.1234176416",  # 도서명
    "author": "entry.628826984",  # 저자명
    "status": "entry.12641564"    # 대출/반납
}

# -------------------------
# ③ 인증 및 시트 불러오기
# -------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly"
]
creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
client = gspread.authorize(creds)

sheet = client.open_by_key(SHEET_KEY).worksheet(SHEET_NAME)
rows = sheet.get_all_records(
    expected_headers=["코드번호", "제목", "지은이", "출판사", "상태", "기타"]
)


# -------------------------
# ④ QR 생성
# -------------------------
os.makedirs("qr_codes", exist_ok=True)

for row in rows:
    code = str(row['코드번호'])
    title = str(row['제목'])
    author = str(row['지은이'])
    status = str(row['상태']).strip()

    # 상태 한글 → 정확한 인코딩 매핑
    status_encoded = {
        "대출": "%EB%8C%80%EC%B6%9C",
        "반납": "%EB%B0%98%EB%82%A9"
    }.get(status, "")

    # 직접 조합
    qr_url = (
        f"{FORM_URL}?usp=pp_url"
        f"&{ENTRY_IDS['code']}={urllib.parse.quote(code)}"
        f"&{ENTRY_IDS['title']}={urllib.parse.quote(title)}"
        f"&{ENTRY_IDS['author']}={urllib.parse.quote(author)}"
        f"&{ENTRY_IDS['status']}={status_encoded}"
    )

    img = qrcode.make(qr_url)
    img.save(f"qr_codes/{code}.png")
    print(f"{code} → QR 생성 완료 ({status})")

print("\n✅ 모든 QR코드 생성 완료! 폴더: qr_codes/")
