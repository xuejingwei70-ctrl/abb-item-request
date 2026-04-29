import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="ABB 物品领用", layout="centered")

st.title("🎁 ABB 键盘鼠标网线领用系统")

# ====================== 员工提交 ======================
page = st.sidebar.radio("请选择功能", ["提交领用申请", "查看所有记录"])

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
                st.error("❌ 邮箱和 Cost Center 不能为空！")
            else:
                try:
                    # 从 secrets 中读取服务账号信息
                    credentials_dict = st.secrets["gcp_service_account"]
                    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
                    creds = Credentials.from_service_account_info(credentials_dict, scopes=scopes)
                    
                    client = gspread.authorize(creds)
                    sh = client.open_by_url(st.secrets["SHEET_URL"])
                    worksheet = sh.sheet1
                    
                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    worksheet.append_row([now, email.strip(), cost_center.strip(), item])
                    
                    st.success("✅ 提交成功！感谢您的申请。")
                    st.balloons()
                except Exception as e:
                    st.error(f"提交失败，请联系IT：{str(e)}")

# ====================== 查看记录 ======================
elif page == "查看所有记录":
    st.subheader("📋 所有领用记录")
    
    try:
        credentials_dict = st.secrets["gcp_service_account"]
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(credentials_dict, scopes=scopes)
        
        client = gspread.authorize(creds)
        sh = client.open_by_url(st.secrets["SHEET_URL"])
        worksheet = sh.sheet1
        data = worksheet.get_all_records()
        
        if data:
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True, hide_index=True)
            
            # 下载为 CSV（公司电脑上可以直接另存为 Excel）
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 下载领用记录 (CSV)",
                data=csv,
                file_name=f"ABB领用记录_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        else:
            st.info("目前还没有任何领用记录")
    except Exception as e:
        st.error(f"读取失败，请联系IT：{str(e)}")

st.caption("ABB IT 支持 | 数据存储于 Google Sheets")
