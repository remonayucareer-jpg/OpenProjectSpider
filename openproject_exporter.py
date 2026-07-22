import base64
import html
import json
import os
import re
import ssl
from datetime import datetime, timedelta
from difflib import get_close_matches
from io import BytesIO
from urllib.parse import quote
from urllib.request import HTTPSHandler, ProxyHandler, Request, build_opener, urlopen

import pandas as pd


BASE_URL = "https://pmo.cemsmart.com"
TARGET_TYPE_ID = "11"
TARGET_TYPE_NAME = "AI运营工单"
DEFAULT_START_DATE = "2026-07-08"
PAGE_SIZE = 100
VERIFY_SSL = False
DISABLE_SYSTEM_PROXY = True
HOTEL_MAP_FILE = "AI项目管理.xlsx"

EXPORT_COLUMNS = [
    "工单类别",
    "工单编号",
    "机器人编号",
    "酒店ID",
    "所属集团",
    "酒店名称",
    "非酒店名称",
    "具体需求",
    "需求数量",
    "问题方",
    "AI应用场景",
    "需求/问题",
    "具体需求/问题",
    "工单创建时间",
    "提交人",
    "工单开始处理时间",
    "工单解决时间",
    "工单状态",
]


def build_authorization_header(api_key=None, auth_header=None):
    if auth_header:
        return auth_header

    api_key = api_key or os.getenv("OPENPROJECT_API_KEY")
    if api_key:
        token = base64.b64encode(f"apikey:{api_key}".encode("utf-8")).decode("ascii")
        return f"Basic {token}"

    auth_header = os.getenv("OPENPROJECT_AUTH_HEADER")
    if auth_header:
        return auth_header

    raise ValueError("缺少 OpenProject API Key，请配置 OPENPROJECT_API_KEY。")


class OpenProjectClient:
    def __init__(self, api_key=None, auth_header=None):
        self.headers = {"Authorization": build_authorization_header(api_key, auth_header)}
        self.ssl_context = ssl._create_unverified_context() if not VERIFY_SSL else None

        handlers = []
        if DISABLE_SYSTEM_PROXY:
            handlers.append(ProxyHandler({}))
        if self.ssl_context:
            handlers.append(HTTPSHandler(context=self.ssl_context))
        self.opener = build_opener(*handlers) if handlers else None

    def get_json(self, path, params=None):
        query = ""
        if params:
            query = "?" + "&".join(
                f"{quote(str(key))}={quote(str(value))}" for key, value in params.items()
            )
        request = Request(f"{BASE_URL}{path}{query}", headers=self.headers)

        if self.opener:
            response = self.opener.open(request, timeout=60)
        else:
            response = urlopen(request, timeout=60, context=self.ssl_context)

        with response:
            charset = response.headers.get_content_charset() or "utf-8"
            return json.loads(response.read().decode(charset))


def load_hotel_map(path=HOTEL_MAP_FILE):
    try:
        df_map = pd.read_excel(path)
        hotel_map = dict(zip(df_map["酒店名称"], df_map["编号"]))
        return hotel_map, list(hotel_map.keys())
    except Exception:
        return {}, []


HOTEL_MAP, HOTEL_NAMES = load_hotel_map()


def clean_subject_for_hotel(subject):
    subject = subject or ""
    subject = re.sub(r"【.*?】|\[.*?\]", "", subject)
    subject = re.sub(r"(--|-)AI.*", "", subject, flags=re.IGNORECASE)
    subject = re.sub(r"\s+", " ", subject)
    return subject.strip()


def match_hotel(subject):
    clean_subject = clean_subject_for_hotel(subject)
    if HOTEL_NAMES:
        matches = get_close_matches(clean_subject, HOTEL_NAMES, n=1, cutoff=0.5)
        if matches:
            hotel_name = matches[0]
            return str(HOTEL_MAP[hotel_name]), hotel_name
    return "未找到", clean_subject or "未找到"


def strip_description_markup(text):
    text = text or ""
    text = re.sub(r"!\[.*?\]\(.*?\)", "", text)
    text = re.sub(r"<img[^>]*>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</p\s*>|</div\s*>|</li\s*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_requirement(description_raw):
    description = strip_description_markup(description_raw)
    if not description:
        return "暂无描述"

    match_req = re.search(r"【问题描述】([\s\S]+)", description)
    if match_req:
        return match_req.group(1).strip()

    match_req = re.search(r"【具体需求】([\s\S]+)", description)
    if match_req:
        return match_req.group(1).strip()

    return description


def parse_issue_source(wp):
    cf11_obj = wp.get("customField11") or wp.get("_links", {}).get("customField11")
    if isinstance(cf11_obj, dict):
        title = cf11_obj.get("title") or ""
        if title == "客户沟通群":
            return "酒店"
        if "内部" in title:
            return "华客"
        return title or "未知"
    return "未知"


def parse_current_status(wp):
    status_obj = wp.get("_links", {}).get("status") or wp.get("status")
    if isinstance(status_obj, dict):
        return status_obj.get("title") or status_obj.get("name") or "未知"
    return "未知"


def parse_openproject_time(value):
    if not value:
        return None
    try:
        dt_utc = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt_utc + timedelta(hours=8)
    except ValueError:
        return None


def format_time(value):
    dt = parse_openproject_time(value)
    if not dt:
        return "未知" if not value else value
    return dt.strftime("%Y-%m-%d %H:%M")


def parse_filter_date(value):
    if not value:
        return None
    if hasattr(value, "strftime"):
        return value
    return datetime.strptime(str(value), "%Y-%m-%d").date()


def is_created_in_date_range(created_at, start_date, end_date):
    created_dt = parse_openproject_time(created_at)
    if not created_dt:
        return False

    created_date = created_dt.date()
    start = parse_filter_date(start_date)
    end = parse_filter_date(end_date) if end_date else None

    if start and created_date < start:
        return False
    if end and created_date > end:
        return False
    return True


def detail_new_value_text(detail):
    new_value = detail.get("newValue")
    if isinstance(new_value, dict):
        return new_value.get("title") or new_value.get("name") or str(new_value)
    if new_value is None:
        return ""
    return str(new_value)


def is_status_changed_to(detail, target_status):
    raw_text = detail.get("raw", "")
    property_name = detail.get("property")
    new_value_text = detail_new_value_text(detail)

    if property_name == "status" and target_status in new_value_text:
        return True

    patterns = [
        rf"状态\s*设置为\s*{re.escape(target_status)}",
        rf"状态\s*已从\s*.+?\s*更改为\s*{re.escape(target_status)}",
    ]
    return any(re.search(pattern, raw_text) for pattern in patterns)


def fetch_activities(client, wp_id):
    payload = client.get_json(
        f"/api/v3/work_packages/{wp_id}/activities",
        params={"pageSize": PAGE_SIZE},
    )
    return payload.get("_embedded", {}).get("elements", [])


def get_status_time(client, wp_id, target_status, author_href=None, skip_author=False):
    activities = fetch_activities(client, wp_id)
    for act in sorted(activities, key=lambda x: x.get("createdAt", "")):
        current_user_href = act.get("_links", {}).get("user", {}).get("href", "")
        if skip_author and author_href and current_user_href == author_href:
            continue

        for detail in act.get("details", []):
            if is_status_changed_to(detail, target_status):
                return format_time(act.get("createdAt"))

    return ""


def fetch_work_packages(client):
    offset = 1
    while True:
        params = {
            "filters": json.dumps(
                [{"type": {"operator": "=", "values": [TARGET_TYPE_ID]}}],
                ensure_ascii=False,
            ),
            "pageSize": PAGE_SIZE,
            "offset": offset,
            "sortBy": json.dumps([["createdAt", "desc"]]),
        }
        payload = client.get_json("/api/v3/work_packages", params=params)
        elements = payload.get("_embedded", {}).get("elements", [])

        for wp in elements:
            yield wp

        count = payload.get("count", len(elements))
        total = payload.get("total")
        if not elements or count <= 0:
            break
        offset += count
        if total is not None and offset > total:
            break


def build_row(client, wp):
    wp_id = wp.get("id")
    author_href = wp.get("_links", {}).get("author", {}).get("href", "")
    robot_id, hotel_name = match_hotel(wp.get("subject", ""))

    return {
        "工单类别": TARGET_TYPE_NAME,
        "工单编号": wp_id,
        "机器人编号": robot_id,
        "酒店ID": "",
        "所属集团": "",
        "酒店名称": hotel_name,
        "非酒店名称": "",
        "具体需求": extract_requirement(wp.get("description", {}).get("raw", "")),
        "需求数量": "",
        "问题方": parse_issue_source(wp),
        "AI应用场景": "",
        "需求/问题": "",
        "具体需求/问题": "",
        "工单创建时间": format_time(wp.get("createdAt")),
        "提交人": wp.get("_links", {}).get("author", {}).get("title", "未知"),
        "工单开始处理时间": get_status_time(
            client,
            wp_id,
            "进行中",
            author_href=author_href,
            skip_author=True,
        ),
        "工单解决时间": get_status_time(client, wp_id, "已解决"),
        "工单状态": parse_current_status(wp),
    }


def build_dataframe(start_date, end_date, api_key=None, auth_header=None, exclude_keywords=None):
    client = OpenProjectClient(api_key=api_key, auth_header=auth_header)
    rows = []
    raw_subjects = []  # 保存原始主题用于过滤
    for wp in fetch_work_packages(client):
        if is_created_in_date_range(wp.get("createdAt"), start_date, end_date):
            rows.append(build_row(client, wp))
            raw_subjects.append(wp.get("subject", ""))
    df_all = pd.DataFrame(rows, columns=EXPORT_COLUMNS)
    # 将原始主题附加到 df_all 中（用于被排除工单的展示，导出时不会包含）
    df_all["_原始主题"] = raw_subjects if len(raw_subjects) == len(df_all) else [""] * len(df_all)

    # 根据原始主题关键词过滤
    if exclude_keywords and len(raw_subjects) == len(df_all):
        mask = pd.Series([True] * len(df_all))
        for kw in exclude_keywords:
            kw = kw.strip()
            if kw:
                mask &= ~pd.Series(raw_subjects).str.contains(kw, case=False, na=False)
        df_filtered = df_all[mask]
    else:
        df_filtered = df_all

    return df_filtered, df_all, len(df_all)


def build_excel_bytes(start_date, end_date, api_key=None, auth_header=None, exclude_keywords=None):
    df_filtered, _, total_count = build_dataframe(start_date, end_date, api_key=api_key, auth_header=auth_header, exclude_keywords=exclude_keywords)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_filtered.to_excel(writer, index=False, sheet_name=TARGET_TYPE_NAME)
    output.seek(0)
    return output.read(), len(df_filtered)


def make_output_filename(start_date, end_date):
    start_text = start_date.strftime("%Y-%m-%d") if hasattr(start_date, "strftime") else str(start_date)
    end_text = end_date.strftime("%Y-%m-%d") if hasattr(end_date, "strftime") else str(end_date)
    return f"openproject_ai_work_packages_{start_text}_to_{end_text}.xlsx"
