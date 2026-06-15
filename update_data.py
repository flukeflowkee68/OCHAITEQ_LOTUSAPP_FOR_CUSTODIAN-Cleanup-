import os
import re
import json
import requests
import pandas as pd

# ดึงค่ารหัสลับที่เซฟไว้ใน GitHub Secrets ออกมาใช้งาน
TENANT_ID = os.environ.get("MICROSOFT_TENANT_ID")
CLIENT_ID = os.environ.get("MICROSOFT_CLIENT_ID")
CLIENT_SECRET = os.environ.get("MICROSOFT_CLIENT_SECRET")

# ลิงก์ตรงและที่อยู่ของไฟล์บน SharePoint ของ OCHA
SHAREPOINT_SITE = "https://unitednations.sharepoint.com/sites/OCHAROAP" 
FILE_PATH = "/Shared Documents/08 Inventory/OCHAITEQ_LOTUSAPP_FOR_CUSTODIAN-Cleanup 2023.xlsx"

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()["access_token"]

def fetch_excel_from_sharepoint(token):
    site_url_clean = SHAREPOINT_SITE.replace("https://", "")
    graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_url_clean}/drive/root:{FILE_PATH}:/content"
    
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(graph_url, headers=headers)
    response.raise_for_status()
    
    with open("temp_inventory.xlsx", "wb") as f:
        f.write(response.content)

def process_data_to_json():
    # ดึงข้อมูลจาก Sheet1 ขึ้นมาจัดฟอร์แมต
    df = pd.read_excel("temp_inventory.xlsx", sheet_name="Sheet1")
    
    processed_data = []
    for _, row in df.iterrows():
        # แมปคอลัมน์จาก Excel (Inventory number, Manufacture, Model, KIT type, Asset Subtype)
        processed_data.append({
            "inventory": str(row.get("Inventory number", "")),
            "brand": str(row.get("Manufacture", "")),
            "model": str(row.get("Model", "")),
            "status": str(row.get("KIT type", "Assigned")), 
            "location": str(row.get("Asset Subtype", ""))   
        })
    return processed_data

def update_html_file(data_list):
    with open("index.html", "r", encoding="utf-8") as f:
        html_content = f.read()
    
    json_data_str = json.dumps(data_list, indent=8, ensure_ascii=False)
    new_data_segment = f"const data = {json_data_str};"
    
    updated_html = re.sub(
        r"const data\s*=\s*\[[\s\S]*?\];", 
        new_data_segment, 
        html_content
    )
    
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(updated_html)

if __name__ == "__main__":
    try:
        print("🔄 กำลังตรวจสอบสิทธิ์กุญแจผ่าน Microsoft Graph API...")
        token = get_access_token()
        print("📥 กำลังดาวน์โหลดไฟล์ Excel มาจาก SharePoint...")
        fetch_excel_from_sharepoint(token)
        print("📊 กำลังประมวลผลข้อมูล...")
        data = process_data_to_json()
        print(f"✍️ กำลังบันทึกข้อมูล {len(data)} รายการ ลงใน index.html...")
        update_html_file(data)
        print("✅ ดึงข้อมูลและอัปเดตโค้ดเว็บเสร็จสิ้น!")
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาด: {str(e)}")
