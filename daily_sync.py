"""
每日定时抓取脚本 - 每天18:00运行
抓取前一天的18:01 ~ 当天的18:00 的工单数据，上传到飞书多维表格
"""
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone

import requests

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openproject_exporter import (
    OpenProjectClient,
    TARGET_TYPE_ID,
    TS_TYPE_ID,
    match_hotel,
    strip_description_markup,
    parse_issue_source,
    format_time,
)

# ============ 飞书配置 ============
FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")
FEISHU_BITABLE_ID = os.environ.get("FEISHU_BITABLE_ID", "OUtfbjipOaGb1osjgktcG3zknae")
FEISHU_TABLE_ID = os.environ.get("FEISHU_TABLE_ID", "tblvZWf9dKkPpXb6")

# ============ 飞书API ============
FEISHU_TOKEN_URL = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
FEISHU_BITABLE_URL = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{FEISHU_BITABLE_ID}/tables/{FEISHU_TABLE_ID}/records"


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
    return "企微"  # 默认填企微


def build_feishu_record(wp, client):
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
        "跟进总结": "",
    }
    return fields


def upload_to_feishu(records, token):
    """批量上传记录到飞书多维表格（每次最多500条）"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    # 飞书批量写入每次最多500条
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
    from openproject_exporter import fetch_work_packages_by_type

    for wp in fetch_work_packages_by_type(client, TARGET_TYPE_ID):
        created_at = wp.get("createdAt", "")
        created_dt = parse_openproject_time(created_at)
        if created_dt and yesterday_18_01 <= created_dt <= today_18:
            fields = build_feishu_record(wp, client)
            fields["展示名称"] = f"AI运营-{wp.get('id')}"
            all_records.append(fields)
            type_stats["AI运营"] += 1

    # 2. 抓取TS工单
    print("抓取TS工单...")
    for wp in fetch_work_packages_by_type(client, TS_TYPE_ID):
        created_at = wp.get("createdAt", "")
        created_dt = parse_openproject_time(created_at)
        if created_dt and yesterday_18_01 <= created_dt <= today_18:
            fields = build_feishu_record(wp, client)
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


def parse_openproject_time(value):
    """解析OpenProject时间字符串为datetime对象"""
    if not value:
        return None
    try:
        dt_utc = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt_utc + timedelta(hours=8)
    except ValueError:
        return None


if __name__ == "__main__":
    try:
        fetch_and_sync()
    except Exception as e:
        print(f"❌ 同步失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
