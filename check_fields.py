import sys, json
sys.path.insert(0, 'c:\\Users\\Administrator\\Desktop\\openproject_spider')
from openproject_exporter import OpenProjectClient

api_key = '5c229cdb014a2cfc140dae69462458f2f31d16147f855831aa1890ac0eabd10f'
client = OpenProjectClient(api_key=api_key)

# 抓取多个TS工单，找包含【问题反馈人】的
data = client.get_json('/api/v3/work_packages', params={
    'pageSize': 20,
    'filters': json.dumps([{"type": {"operator": "=", "values": ["1"]}}], ensure_ascii=False)
})
wps = data['_embedded']['elements']

with open('c:\\Users\\Administrator\\Desktop\\openproject_spider\\fields_output.txt', 'w', encoding='utf-8') as f:
    for wp in wps:
        desc = wp.get('description', {}).get('raw', '')
        if '问题反馈人' in desc or '反馈人' in desc:
            f.write(f"=== TS工单ID: {wp.get('id')} ===\n")
            f.write(f"subject: {wp.get('subject')}\n")
            f.write(f"description:\n{desc[:1500]}\n\n")
            f.write(f"customField11: {wp.get('_links', {}).get('customField11', {}).get('title', 'N/A')}\n")
            f.write(f"author: {wp.get('_links', {}).get('author', {}).get('title', 'N/A')}\n")
            f.write("-" * 50 + "\n\n")

print("Done")
