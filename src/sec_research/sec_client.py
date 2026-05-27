from __future__ import annotations

import json
import time
import zipfile
from pathlib import Path
from typing import Any

import requests


class SecClient:
    def __init__(self, cache_dir: Path, user_agent: str, sleep_seconds: float = 0.12):
        self.cache_dir = cache_dir
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": user_agent,
            "Accept-Encoding": "gzip, deflate",
        })
        self.sleep_seconds = sleep_seconds

    def download(self, url: str, filename: str, refresh: bool = False) -> Path:
        path = self.cache_dir / filename
        if path.exists() and not refresh:
            return path
        response = self.session.get(url, timeout=120)
        response.raise_for_status()
        path.write_bytes(response.content)
        time.sleep(self.sleep_seconds)
        return path

    def get_json(self, url: str) -> dict[str, Any] | None:
        response = self.session.get(url, timeout=60)
        if response.status_code == 404:
            return None
        response.raise_for_status()
        time.sleep(self.sleep_seconds)
        return response.json()

    def load_company_tickers(self, url: str) -> list[dict[str, Any]]:
        path = self.download(url, "company_tickers_exchange.json")
        data = json.loads(path.read_text(encoding="utf-8"))
        fields = data["fields"]
        return [dict(zip(fields, row)) for row in data["data"]]

    def read_submission_from_zip(self, zip_path: Path, cik: str) -> dict[str, Any] | None:
        member = f"CIK{cik.zfill(10)}.json"
        with zipfile.ZipFile(zip_path) as zf:
            try:
                with zf.open(member) as f:
                    return json.load(f)
            except KeyError:
                return None

    def read_companyfacts_from_zip(self, zip_path: Path, cik: str) -> dict[str, Any] | None:
        member = f"CIK{cik.zfill(10)}.json"
        with zipfile.ZipFile(zip_path) as zf:
            try:
                with zf.open(member) as f:
                    return json.load(f)
            except KeyError:
                return None


def sec_company_url(cik: str, accession: str | None = None) -> str:
    if not accession:
        return f"https://www.sec.gov/edgar/browse/?CIK={int(cik)}"
    acc_no_dash = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc_no_dash}/"


def sec_submission_api_url(cik: str) -> str:
    return f"https://data.sec.gov/submissions/CIK{cik.zfill(10)}.json"


def sec_companyfacts_api_url(cik: str) -> str:
    return f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik.zfill(10)}.json"
