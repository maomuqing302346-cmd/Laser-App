import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime
import os

# ==========================================
# 1. 初始化设置 (页面标题、数据库)
# ==========================================
st.set_page_config(page_title="激光器维修系统", page_icon="🔧")

# 初始化：模拟云端数据库 (如果内存里没有db，就建一个空的)
if 'db' not in st.session_state:
    st.session_state['db'] = pd.DataFrame(columns=[
        "id", "sn", "model", "operator", "date", "problem", "action", "current", "power_1064"
    ])

# 初始化：登录状态
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False
if 'current_user' not in st.session_state:
    st.session_state['current_user'] = None
if 'role' not in st.session_state:
    st.session_state['role'] = None

# ==========================================
# 2. 核心功能函数
# ==========================================
def generate_doc(record):
    """生成 Word 文档的函数"""
    # 检查模板是否存在
    if not os.path.exists("template.docx"):
        return None
    
    doc = DocxTemplate("template.docx")
    
    # 准备填空数据
    context = {
        'sn': record['sn'],
        'model': record['model'],
        'date': record['date'],
        'operator': record['operator'],
        'problem': record['problem'],
        'action': record['action'],
        'current': record['current'],
        'power_1064': record.get('power_1064', ''), # 防止旧数据没有这个字段报错
        # 其他默认填充，防止模板报错
        'voltage': "24V", 'obs_case': "正常", 'obs_mech': "无",
        'work_hours': "", 'alarms': "无", 'hv': "", 'pulse': ""
    }
    
    doc.render(context)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def login_page():
    """登录页面"""
    st.markdown("<h1 style='text-align: center;'>🔐 激光器维修系统登录</h1>", unsafe_allow_html=True)
    st.write("") # 空行
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("用户名 (admin / user)")
            password = st.text_input("密码 (admin123)", type="password")
            submitted = st.form_submit_button("登录", use_container_width=True)
            
            if submitted:
                if username == "admin" and password == "admin123":
                    st.session_state['authenticated'] = True
                    st.session_state['role'] = 'admin'
                    st.session_state['current_user'] = 'admin'
                    st.rerun() # 登录成功，强制刷新页面进入主界面
                elif username == "user":
                    st.session_state['authenticated'] = True
                    st.session_state['role'] = 'user'
                    st.session_state['current_user'] = 'user'
                    st.rerun()
                else:
                    st.error("用户名或密码错误")

def main_app():
    """主程序界面 (只有登录后才会执行这里)"""
    
    # --- 侧边栏 ---
    st.sidebar.markdown(f"👤 **当前用户:** {st.session_state['current_user']}")
    st.sidebar.markdown(f"🛡️ **权限:** {st.session_state['role']}")
    
    if st.sidebar.button("🚪 退出登录"):
        st.session_state['authenticated'] = False
        st.session_state['current_user'] = None
        st.session_state['role'] = None
        st.rerun() # 退出后，强制刷新回登录页

    st.sidebar.divider()
    menu = st.sidebar.radio("功能导航", ["📝 录入新单", "🔍 历史查询"])

    # --- 页面 1: 录入 ---
    if menu == "📝 录入新单":
        st.title("📝 新建维修记录")
        st.info("填写下方表单，点击保存即可归档。")
        
        with st.form("repair_form"):
            c1, c2 = st.columns(2)
            sn = c1.text_input("序列号 ({{sn}})")
            model = c2.selectbox("型号 ({{model}})", ["WYP-Series", "Other", "Unknown"])
            
            operator = st.text_input("操作员 ({{operator}})", value=st.session_state['current_user'])
            
            st.markdown("---")
            problem = st.text_area("故障描述 ({{problem}})", height=100)
            action = st.text_area("维修措施 ({{action}})", height=100)
            
            st.markdown("---")
            d1, d2 = st.columns(2)
            current = d1.number_input("电流值 (A) ({{current}})", step=0.1)
            power = d2.number_input("1064nm 功率 (W) ({{power_1064}})", step=0.1)
            
            submitted = st.form_submit_button("💾 保存到云端", type="primary")
            
            if submitted:
                if not sn:
                    st.error("❌ 序列号不能为空")
                else:
                    new_id = len(st.session_state['db']) + 1
                    new_data = {
                        "id": new_id,
                        "sn": sn, "model": model, "operator": operator,
                        "date": datetime.now().strftime("%Y-%m-%d"),
                        "problem": problem, "action": action, 
                        "current": current, "power_1064": power
                    }
                    # 追加数据
                    st.session_state['db'] = pd.concat([st.session_state['db'], pd.DataFrame([new_data])], ignore_index=True)
                    st.success(f"✅ 序列号 {sn} 保存成功！")

    # --- 页面 2: 查询 ---
    elif menu == "🔍 历史查询":
        st.title("🔍 维修档案库")
        
        search_sn = st.text_input("输入序列号进行搜索 (留空显示所有):")
        
        if not st.session_state['db'].empty:
            # 过滤逻辑
            if search_sn:
                df_show = st.session_state['db'][st.session_state['db']['sn'].str.contains(search_sn, case=False)]
            else:
                df_show = st.session_state['db']
            
            st.write(f"共找到 {len(df_show)} 条记录")

            # 遍历显示
            for index, row in df_show.iterrows():
                # 使用 expander 收纳详细信息
                with st.expander(f"📅 {row['date']} | SN: {row['sn']} | 操作员: {row['operator']}"):
                    st.markdown(f"**故障:** {row['problem']}")
                    st.markdown(f"**措施:** {row['action']}")
                    st.markdown(f"**数据:** 电流 {row['current']}A | 功率 {row['power_1064']}W")
                    
                    col_down, col_del = st.columns([1, 1])
                    
                    # 下载按钮
                    with col_down:
                        doc_file = generate_doc(row)
                        if doc_file:
                            st.download_button(
                                label="📥 下载 Word 报告",
                                data=doc_file,
                                file_name=f"Report_{row['sn']}_{row['date']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"dl_{row['id']}"
                            )
                        else:
                            st.warning("⚠️ 未检测到 template.docx，无法下载")

                    # 删除按钮 (仅管理员可见)
                    if st.session_state['role'] == 'admin':
                        with col_del:
                            if st.button("🗑️ 删除此条", key=f"del_{row['id']}"):
                                # 删除逻辑：通过ID找到并删除
                                st.session_state['db'] = st.session_state['db'][st.session_state['db']['id'] != row['id']]
                                st.rerun() # 立即刷新列表
        else:
            st.info("📭 数据库暂时为空，请先录入数据。")

# ==========================================
# 3. 主逻辑控制 (路由)
# ==========================================
if __name__ == "__main__":
    if st.session_state['authenticated']:
        main_app()
    else:
        login_page()
