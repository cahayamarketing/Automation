import os
import sys
from pathlib import Path
from datetime import datetime, timedelta

import requests
from dotenv import load_dotenv

# ============================================================
# CONFIG
# ============================================================

BASE_URL = "https://dis.fifgroup.co.id/dis2"

LOGIN_URL = f"{BASE_URL}/authenticate"
DEALER_URL = f"{BASE_URL}/mstdealer/findlookup"
DOWNLOAD_URL = f"{BASE_URL}/ordertracking/downloadxls"

# Folder hasil download
OUTPUT_DIR = Path(r"D:\H1\PBI\PBI\db_DIS_CSM")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

load_dotenv()


# ============================================================
# AKUN
# ============================================================

ACCOUNTS = [
    {
        "name": "AKUN 1",
        "username": os.getenv("USERNAME_1"),
        "password": os.getenv("PASSWORD_1"),
    },
    {
        "name": "AKUN 2",
        "username": os.getenv("USERNAME_2"),
        "password": os.getenv("PASSWORD_2"),
    },
]


# ============================================================
# HEADER
# ============================================================

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/151.0.0.0 Safari/537.36"
)


# ============================================================
# SESSION
# ============================================================

def create_session():

    session = requests.Session()

    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "application/json, text/plain, */*",
        "X-Requested-With": "XMLHttpRequest",
    })

    return session


# ============================================================
# LOGIN DIS
# ============================================================

def login(session, account):

    username = account["username"]
    password = account["password"]

    if not username or not password:
        raise RuntimeError(
            f"{account['name']}: username/password belum diisi di .env"
        )

    print()
    print("=" * 70)
    print(f"LOGIN {account['name']}")
    print("=" * 70)

    params = {
        "username": username,
        "password": password,
    }

    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": "https://dis.fifgroup.co.id",
        "Referer": f"{BASE_URL}/login",
    }

    response = session.post(
        LOGIN_URL,
        params=params,
        headers=headers,
        timeout=60,
    )

    response.raise_for_status()

    print("HTTP Status :", response.status_code)
    print("JSESSIONID  :", session.cookies.get("JSESSIONID"))

    if not session.cookies.get("JSESSIONID"):
        raise RuntimeError(
            "JSESSIONID tidak ditemukan. "
            "Login kemungkinan gagal."
        )

    print(f"Login berhasil : {username}")


# ============================================================
# AMBIL DEALER
# ============================================================

def get_dealers(session):

    url = f"{BASE_URL}/mstdealer/findlookup"

    data = {
        "filter": None,
        "kodeProduk": None,
        "office": None,
    }

    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "Referer": f"{BASE_URL}/ordertracking",
    }

    response = session.post(
        url,
        json=data,
        headers=headers,
        timeout=60,
    )

    response.raise_for_status()

    result = response.json()

    # Berdasarkan response yang kamu kirim sebelumnya
    if isinstance(result, dict):

        dealers = result.get("data", [])

    elif isinstance(result, list):

        dealers = result

    else:

        dealers = []

    return dealers


# ============================================================
# TAMPILKAN DEALER
# ============================================================

def show_dealers(dealers):

    print()
    print("=" * 70)
    print("DAFTAR DEALER")
    print("=" * 70)

    if not dealers:

        print("Dealer tidak ditemukan.")

        return

    for i, dealer in enumerate(dealers, start=1):

        print(
            f"{i:3}. "
            f"{dealer.get('dealerId', '')} | "
            f"{dealer.get('dealerName', '')} | "
            f"{dealer.get('dealerLob', '')}"
        )


# ============================================================
# TANGGAL
# ============================================================

def validate_date(value):

    try:

        parse_date(value)

        return True

    except (TypeError, ValueError):

        return False


def parse_date(value):
    """Return a validated date value in the format used by the DIS API."""

    if not isinstance(value, str):
        raise ValueError("Tanggal harus berupa teks.")

    return datetime.strptime(
        value.strip(),
        "%d-%m-%Y"
    )


def ask_date(label, default):

    while True:

        value = input(
            f"{label} [{default}]: "
        ).strip()

        if not value:

            value = default

        if validate_date(value):

            return value

        print(
            "Format tanggal harus DD-MM-YYYY."
        )


# ============================================================
# DOWNLOAD ORDER TRACKING
# ============================================================

def download_order_tracking(
    session,
    account_name,
    dealer,
    date_from,
    date_to,
):

    dealer_id = dealer.get("dealerId")
    source_input = dealer.get("dealerLob", "NMC")

    params = {
        "dateFrom": datetime.strptime(date_from, "%d-%m-%Y").strftime("%Y-%m-%d"),
        "dateTo": datetime.strptime(date_to, "%d-%m-%Y").strftime("%Y-%m-%d"),
        "dealerId": dealer_id,
        "sourceInput": source_input,
    }

    headers = {
        "Accept": "*/*",
        "Referer": f"{BASE_URL}/ordertracking",
    }

    print()
    print("-" * 70)
    print("DOWNLOAD ORDER TRACKING")
    print("-" * 70)

    print("Akun       :", account_name)
    print("Dealer ID  :", dealer_id)
    print("Dealer     :", dealer.get("dealerName"))
    print("LOB        :", source_input)
    print("FROM       :", date_from)
    print("TO         :", date_to)

    response = session.get(
        DOWNLOAD_URL,
        params=params,
        headers=headers,
        timeout=300,
    )

    response.raise_for_status()

    content_type = response.headers.get(
        "Content-Type",
        ""
    ).lower()

    print("HTTP Status:", response.status_code)
    print("Content-Type:", content_type)

    # --------------------------------------------------------
    # CEK JIKA SERVER MENGEMBALIKAN HTML
    # --------------------------------------------------------

    if "text/html" in content_type:

        preview = response.text[:500]

        raise RuntimeError(
            "Server mengembalikan HTML, bukan file download.\n"
            f"Preview:\n{preview}"
        )

    # --------------------------------------------------------
    # NAMA FILE
    # --------------------------------------------------------

    filename = (
        f"{dealer_id}_"
        f"{date_from}_"
        f"{date_to}.xls"
    )

    # --------------------------------------------------------
    # BERSIHKAN NAMA FILE
    # --------------------------------------------------------

    invalid_chars = '<>:"/\\|?*'

    for char in invalid_chars:

        filename = filename.replace(
            char,
            "_"
        )

    month_folder = datetime.strptime(
            date_from,
            "%d-%m-%Y"
    ).strftime("%Y%m")

    output1 = OUTPUT_DIR / month_folder
    output1.mkdir(parents=True, exist_ok=True)

    # --------------------------------------------------------
    # SIMPAN
    # --------------------------------------------------------

    output_file = output1 / filename

    output_file.write_bytes(
        response.content
    )

    print()
    print("BERHASIL DOWNLOAD")
    print("File :", output_file)
    print(
        "Size :",
        f"{len(response.content):,}",
        "bytes"
    )

    return output_file


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 70)
    print("DIS ORDER TRACKING DOWNLOADER")
    print("=" * 70)

    # --------------------------------------------------------
    # TANGGAL DEFAULT = KEMARIN
    # --------------------------------------------------------

    yesterday = (
        datetime.now()
        - timedelta(days=1)
    ).strftime("%d-%m-%Y")

    date_from = ask_date(
        "Tanggal FROM",
        yesterday
    )

    date_to = ask_date(
        "Tanggal TO",
        yesterday
    )

    if parse_date(date_from) > parse_date(date_to):

        print(
            "\nERROR: Tanggal FROM "
            "tidak boleh lebih besar dari TO."
        )

        sys.exit(1)

    print()
    print("=" * 70)
    print("AKAN MEMPROSES 2 AKUN")
    print("=" * 70)

    print("FROM :", date_from)
    print("TO   :", date_to)
    print("FOLDER:", OUTPUT_DIR)

    # ========================================================
    # PROSES AKUN
    # ========================================================

    for account in ACCOUNTS:

        if not account["username"]:
            print(
                f"\n{account['name']} dilewati: "
                "USERNAME belum ada."
            )
            continue

        session = None

        try:
            session = create_session()
            login(session, account)

            dealers = get_dealers(session)

            if not dealers:
                print(
                    f"{account['name']}: "
                    "dealer tidak ditemukan."
                )
                continue

            print(f"Jumlah dealer ditemukan: {len(dealers)}")

            for i, dealer in enumerate(dealers, start=1):
                print()
                print("=" * 70)
                print(f"DEALER {i}/{len(dealers)}")
                print("=" * 70)

                try:
                    download_order_tracking(
                        session,
                        account["name"],
                        dealer,
                        date_from,
                        date_to,
                    )
                except Exception as exc:
                    print(
                        f"ERROR dealer "
                        f"{dealer.get('dealerId')}: {exc}"
                    )

        except Exception as exc:
            print(f"ERROR pada {account['name']}: {exc}")

        finally:
            if session is not None:
                session.close()


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()