import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

st.set_page_config(page_title="ABB 物品领用", layout="centered")

# ====================== 配置 ======================
st.title("🎁 ABB 键盘鼠标网线领用系统")

# Google Sheets 配置（后面会教你怎么填）
SHEET_URL = st.secrets["SHEET_URL"] if "SHEET_URL" in st.secrets else ""

if not SHEET_URL:
    st.error("尚未配置 Google Sheets 链接，请联系管理员")
    st.stop()

# ====================== 员工提交页面 ======================
page = st.sidebar.radio("请选择", ["提交领用申请", "查看所有记录"])

if page == "提交领用申请":
    st.subheader("请填写领用信息")
    
    with st.form("request_form"):
        email = st.text_input("公司邮箱", placeholder="your.name@cn.abb.com")
        cost_center = st.text_input("Cost Center (CC)", placeholder="12345678")
        
        item = st.radio("请选择领取物品", 
                       ["无线鼠标", "有线鼠标", "2米网线", "5米网线"])
        
        submitted = st.form_submit_button("提交申请")
        
        if submitted:
            if not email or not cost_center:
                st.error("邮箱和 Cost Center 不能为空！")
            else:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # 连接 Google Sheets 并写入
                try:
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(SHEET_URL)
                    worksheet = sh.sheet1
                    
                    worksheet.append_row([now, email.strip(), cost_center.strip(), item])
                    
                    st.success("✅ 提交成功！感谢您的申请。")
                    st.balloons()
                except Exception as e:
                    st.error(f"提交失败：{str(e)}")

# ====================== 查看记录页面 ======================
elif page == "查看所有记录":
    st.subheader("📋 所有领用记录")
    
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(SHEET_URL)
        worksheet = sh.sheet1
        data = worksheet.get_all_records()
        
        if data:
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True)
            
            # 一键下载 Excel
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 下载 Excel 文件",
                data=csv,
                file_name=f"ABB领用记录_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        else:
            st.info("目前还没有领用记录")
    except Exception as e:
        st.error(f"读取数据失败：{str(e)}")

st.caption("ABB IT 支持 | 数据存储于 Google Sheets")
