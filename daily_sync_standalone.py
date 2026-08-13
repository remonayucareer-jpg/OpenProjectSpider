"""
每日定时抓取脚本 - 每天18:00运行
抓取前一天的18:01 ~ 当天的18:00 的工单数据，上传到飞书多维表格

独立版本，不依赖 openproject_exporter.py
"""
import base64
import html
import json
import os
import re
import ssl
import sys
from datetime import datetime, timedelta
from difflib import get_close_matches
from io import BytesIO
from urllib.parse import quote
from urllib.request import HTTPSHandler, ProxyHandler, Request, build_opener, urlopen

import pandas as pd
import requests

# 尝试导入 curl_cffi（用于解决 OpenSSL 与服务器 TLS 不兼容的问题）
# curl_cffi 使用预编译的 wheel 包，不需要编译，可以在 Streamlit Cloud 上安装
try:
    from curl_cffi import requests as curl_requests
    HAS_CURL_CFFI = True
except ImportError:
    HAS_CURL_CFFI = False

# 尝试导入 pycurl（备选方案）
try:
    import pycurl
    HAS_PYCURL = True
except ImportError:
    HAS_PYCURL = False




# ============ OpenProject 配置 ============
BASE_URL = "https://pmo.cemsmart.com"
TARGET_TYPE_ID = "11"
TS_TYPE_ID = "1"
PAGE_SIZE = 100
VERIFY_SSL = False
DISABLE_SYSTEM_PROXY = True
HOTEL_MAP_FILE = "AI项目管理.xlsx"

# ============ 飞书配置 ============
FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")
FEISHU_BITABLE_ID = os.environ.get("FEISHU_BITABLE_ID", "OUtfbjipOaGb1osjgktcG3zknae")
FEISHU_TABLE_ID = os.environ.get("FEISHU_TABLE_ID", "tblvZWf9dKkPpXb6")

# ============ 飞书API ============
FEISHU_TOKEN_URL = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
FEISHU_BITABLE_URL = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{FEISHU_BITABLE_ID}/tables/{FEISHU_TABLE_ID}/records"


# ==================== OpenProject 客户端 ====================

def build_authorization_header(api_key=None):
    api_key = api_key or os.getenv("OPENPROJECT_API_KEY")
    if api_key:
        token = base64.b64encode(f"apikey:{api_key}".encode("utf-8")).decode("ascii")
        return f"Basic {token}"
    raise ValueError("缺少 OpenProject API Key，请配置 OPENPROJECT_API_KEY。")


class OpenProjectClient:
    def __init__(self, api_key=None):
        self.headers = {"Authorization": build_authorization_header(api_key)}
        self.ssl_context = ssl._create_unverified_context() if not VERIFY_SSL else None
        handlers = []
        if DISABLE_SYSTEM_PROXY:
            handlers.append(ProxyHandler({}))
        if self.ssl_context:
            handlers.append(HTTPSHandler(context=self.ssl_context))
        self.opener = build_opener(*handlers) if handlers else None

    def _get_json_with_curl_cffi(self, url):
        """使用 curl_cffi 发起请求（解决 OpenSSL 与服务器 TLS 不兼容的问题）"""
        resp = curl_requests.get(
            url,
            headers=self.headers,
            timeout=60,
            verify=False,
            impersonate="chrome",
        )
        if resp.status_code >= 400:
            raise Exception(f"HTTP {resp.status_code}: {resp.text[:500]}")
        return resp.json()

    def _get_json_with_pycurl(self, url):
        """使用 pycurl 发起请求（备选方案）"""
        buf = BytesIO()
        c = pycurl.Curl()
        try:
            c.setopt(pycurl.URL, url)
            c.setopt(pycurl.HTTPHEADER, [f"{k}: {v}" for k, v in self.headers.items()])
            c.setopt(pycurl.SSL_VERIFYPEER, 0)
            c.setopt(pycurl.SSL_VERIFYHOST, 0)
            c.setopt(pycurl.WRITEDATA, buf)
            c.setopt(pycurl.TIMEOUT, 60)
            if DISABLE_SYSTEM_PROXY:
                c.setopt(pycurl.PROXY, "")
            c.perform()
            status = c.getinfo(pycurl.RESPONSE_CODE)
            if status >= 400:
                raise Exception(f"HTTP {status}: {buf.getvalue().decode('utf-8', errors='replace')[:500]}")
            return json.loads(buf.getvalue().decode("utf-8"))
        finally:
            c.close()

    def get_json(self, path, params=None):
        query = ""
        if params:
            query = "?" + "&".join(
                f"{quote(str(key))}={quote(str(value))}" for key, value in params.items()
            )
        url = f"{BASE_URL}{path}{query}"

        # 1. 优先使用 curl_cffi（使用预编译 wheel，无需编译，可在 Streamlit Cloud 上安装）
        if HAS_CURL_CFFI:
            try:
                return self._get_json_with_curl_cffi(url)
            except Exception as e:
                # curl_cffi 失败时继续尝试其他方案
                pass

        # 2. 其次使用 pycurl
        if HAS_PYCURL:
            try:
                return self._get_json_with_pycurl(url)
            except Exception as e:
                # pycurl 失败时回退到 urllib
                pass

        # 3. 最后回退到 urllib
        request = Request(url, headers=self.headers)
        if self.opener:
            response = self.opener.open(request, timeout=60)
        else:
            response = urlopen(request, timeout=60, context=self.ssl_context)
        with response:
            charset = response.headers.get_content_charset() or "utf-8"
            return json.loads(response.read().decode(charset))




# ==================== 工具函数 ====================

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


def format_time(value):
    if not value:
        return "未知"
    try:
        dt_utc = datetime.fromisoformat(value.replace("Z", "+00:00"))
        dt = dt_utc + timedelta(hours=8)
        return dt.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return value


def parse_openproject_time(value):
    if not value:
        return None
    try:
        dt_utc = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt_utc + timedelta(hours=8)
    except ValueError:
        return None


def fetch_work_packages_by_type(client, type_id):
    """根据 type_id 抓取工单"""
    offset = 1
    while True:
        params = {
            "filters": json.dumps(
                [{"type": {"operator": "=", "values": [type_id]}}],
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


# ==================== 飞书相关函数 ====================

def get_feishu_token():
    """获取飞书 tenant_access_token"""
    resp = requests.post(
        FEISHU_TOKEN_URL,
        json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET},
        timeout=30,
    )
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取飞书Token失败: {data}")
    return data["tenant_access_token"]


def extract_contact_person(description_raw):
    """从description中提取【问题反馈人】后的名字"""
    desc = strip_description_markup(description_raw)
    if not desc:
        return ""
    match = re.search(r"【问题反馈人】\s*[:：]?\s*([^\n]+)", desc)
    if match:
        return match.group(1).strip()
    match = re.search(r"反馈人\s*[:：]?\s*([^\n]+)", desc)
    if match:
        return match.group(1).strip()
    return ""


def extract_problem_description(description_raw):
    """从description中提取【问题描述】后的内容"""
    desc = strip_description_markup(description_raw)
    if not desc:
        return ""
    match = re.search(r"【问题描述】([\s\S]+)", desc)
    if match:
        return match.group(1).strip()
    return desc


def get_feishu_source(custom_field_11):
    """根据customField11判断跟进方式"""
    if isinstance(custom_field_11, dict):
        title = custom_field_11.get("title", "")
        if "客户沟通群" in title:
            return "企微"
    return "企微"


def build_feishu_record(wp):
    """构建飞书多维表格的记录"""
    wp_id = wp.get("id")
    desc_raw = wp.get("description", {}).get("raw", "")
    author_title = wp.get("_links", {}).get("author", {}).get("title", "")
    created_at = wp.get("createdAt", "")
    _, hotel_name = match_hotel(wp.get("subject", ""))

    # 获取工单类型
    type_title = wp.get("_links", {}).get("type", {}).get("title", "")

    # 格式化时间
    created_formatted = format_time(created_at)
    created_date_only = created_formatted.split(" ")[0] if " " in created_formatted else created_formatted

    fields = {
        "展示名称": "",
        "跟进类型": "被动咨询",
        "主动服务场景": "",
        "服务项目": "",
        "任务": "",
        "集团": "",
        "跟进方式": get_feishu_source(wp.get("_links", {}).get("customField11")),
        "联系人": extract_contact_person(desc_raw),
        "联系人角色": "",
        "客户反馈/问题": extract_problem_description(desc_raw),
        "本次跟进记录": "反馈已提交",
        "待跟进事项": str(wp_id),
        "跟进总结": "",
        "附件": "",
        "跟进日期": created_date_only,
        "本次跟进情况": "",
        "跟进分类": "",
        "跟进人": author_title,
        "创建人": author_title,
        "创建时间": created_formatted,
        "父记录ID": "",
        "酒店名称": hotel_name,
        "酒店ID": "",
        "产品名称": "",
    }
    return fields


def upload_to_feishu(records, token):
    """批量上传记录到飞书多维表格（每次最多500条）"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    batch_size = 500
    total = len(records)
    success_count = 0

    for i in range(0, total, batch_size):
        batch = records[i : i + batch_size]
        payload = {
            "records": [{"fields": record} for record in batch],
        }

        resp = requests.post(
            FEISHU_BITABLE_URL + "/batch_create",
            headers=headers,
            json=payload,
            timeout=60,
        )
        data = resp.json()
        if data.get("code") != 0:
            print(f"上传批次 {i//batch_size + 1} 失败: {data}")
            raise Exception(f"飞书上传失败: {data}")
        success_count += len(batch)
        print(f"已上传 {success_count}/{total} 条")

    return success_count


# ==================== 主函数 ====================

def fetch_and_sync():
    """主函数：抓取数据并上传到飞书"""
    print("=" * 60)
    print(f"开始每日同步 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 计算时间范围：前一天的18:01 ~ 当天的18:00
    now = datetime.now()
    today_18 = now.replace(hour=18, minute=0, second=0, microsecond=0)
    yesterday_18_01 = today_18 - timedelta(days=1) + timedelta(minutes=1)

    # 如果当前时间还没到18:00，则使用昨天的18:00作为结束
    if now < today_18:
        today_18 = today_18 - timedelta(days=1)
        yesterday_18_01 = today_18 - timedelta(days=1) + timedelta(minutes=1)

    start_time = yesterday_18_01.strftime("%Y-%m-%d %H:%M")
    end_time = today_18.strftime("%Y-%m-%d %H:%M")
    print(f"抓取时间范围: {start_time} ~ {end_time}")

    # 连接OpenProject
    api_key = os.environ.get("OPENPROJECT_API_KEY", "")
    if not api_key:
        raise ValueError("缺少 OPENPROJECT_API_KEY 环境变量")

    client = OpenProjectClient(api_key=api_key)

    # 获取飞书Token
    print("获取飞书Token...")
    token = get_feishu_token()
    print("飞书Token获取成功")

    all_records = []
    type_stats = {"AI运营": 0, "TS": 0}

    # 1. 抓取AI运营工单
    print("抓取AI运营工单...")
    for wp in fetch_work_packages_by_type(client, TARGET_TYPE_ID):
        created_at = wp.get("createdAt", "")
        created_dt = parse_openproject_time(created_at)
        if created_dt and yesterday_18_01 <= created_dt <= today_18:
            fields = build_feishu_record(wp)
            fields["展示名称"] = f"AI运营-{wp.get('id')}"
            all_records.append(fields)
            type_stats["AI运营"] += 1

    # 2. 抓取TS工单
    print("抓取TS工单...")
    for wp in fetch_work_packages_by_type(client, TS_TYPE_ID):
        created_at = wp.get("createdAt", "")
        created_dt = parse_openproject_time(created_at)
        if created_dt and yesterday_18_01 <= created_dt <= today_18:
            fields = build_feishu_record(wp)
            fields["展示名称"] = f"TS-{wp.get('id')}"
            all_records.append(fields)
            type_stats["TS"] += 1

    print(f"\n抓取完成:")
    print(f"  AI运营工单: {type_stats['AI运营']} 条")
    print(f"  TS工单: {type_stats['TS']} 条")
    print(f"  总计: {len(all_records)} 条")

    if not all_records:
        print("没有需要上传的数据")
        return

    # 3. 上传到飞书
    print("\n上传到飞书多维表格...")
    success_count = upload_to_feishu(all_records, token)
    print(f"\n✅ 同步完成！成功上传 {success_count} 条记录到飞书")


if __name__ == "__main__":
    try:
        fetch_and_sync()
    except Exception as e:
        print(f"❌ 同步失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
