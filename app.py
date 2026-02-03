import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime

# --- 模拟云端数据库 (实际部署时这里会替换成 Google Sheets 连接代码) ---
if 'db' not in st.session_state:
    st.session_state['db'] = pd.DataFrame(columns=[
        "id", "sn", "model", "operator", "date", "problem", "action", "current"
    ])

# --- 权限管理 (简单的登录逻辑) ---
# --- 权限管理 (简单的登录逻辑) ---
def check_password():
    """返回 True 代表已登录，False 代表未登录"""
    
    # 1. 如果没有这两个变量，先初始化（防止报错）
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if "current_user" not in st.session_state:
        st.session_state["current_user"] = None
    if "role" not in st.session_state:
        st.session_state["role"] = None

    # 2. 定义登录点击后的动作
    def password_entered():
        # 注意：这里我们读取临时输入框的 username，存到永久变量 current_user 里
        user = st.session_state.get("username_input", "") # 使用 get 防止报错
        pwd = st.session_state.get("password_input", "")

        if user in ["admin", "user"]:
            # 验证逻辑
            if user == "admin" and pwd == "admin123":
                st.session_state["authenticated"] = True
                st.session_state["role"] = "admin"
                st.session_state["current_user"] = user # 【关键】存到这里！
            elif user == "user": # 假设普通用户不校验复杂密码
                st.session_state["authenticated"] = True
                st.session_state["role"] = "user"
                st.session_state["current_user"] = user # 【关键】存到这里！
            else:
                st.error("密码错误")
        else:
            st.error("用户不存在")

    # 3. 如果已登录，直接返回 True
    if st.session_state["authenticated"]:
        return True

    # 4. 如果没登录，显示输入框
    st.text_input("用户名 (admin/user)", key="username_input") # 【关键】改个名，叫 username_input
    st.text_input("密码 (admin的密码是 admin123)", type="password", key="password_input")
    st.button("登录", on_click=password_entered)
    
    return False

# --- Word 生成功能 (核心) ---
def generate_doc(record):
    # 加载您的模板
    doc = DocxTemplate("template.docx")
    
    # 准备要填入的数据 (Context)
    context = {
        'sn': record['sn'],
        'model': record['model'],
        'date': record['date'],
        'operator': record['operator'],
        'problem': record['problem'],
        'action': record['action'],
        'current': record['current'],
        # 这里对应您 Word 里 {{xxx}} 的所有标签
    }
    
    doc.render(context)
    
    # 保存到内存
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- 主程序 ---
if check_password():
    # 【注意】这里改成 st.session_state['current_user']
    st.sidebar.write(f"当前用户: {st.session_state['current_user']} (权限: {st.session_state['role']})")
    
    if st.sidebar.button("退出登录"):
        st.session_state["authenticated"] = False
        st.session_state["current_user"] = None # 退出时清空
        st.rerun()

    menu = st.sidebar.radio("菜单", ["📝 录入新单", "🔍 历史查询"])

    if menu == "📝 录入新单":
        st.title("新建维修记录")
        with st.form("repair_form"):
            c1, c2 = st.columns(2)
            sn = c1.text_input("序列号 {{sn}}")
            model = c2.selectbox("型号 {{model}}", ["WYP-Series", "Other"])
            operator = st.text_input("操作员 {{operator}}", value=st.session_state['username'])
            problem = st.text_area("故障描述 {{problem}}")
            action = st.text_area("维修措施 {{action}}")
            current = st.number_input("电流值 {{current}}", step=0.1)
            
            submitted = st.form_submit_button("保存到云端")
            
            if submitted:
                # 构建新数据
                new_data = {
                    "id": len(st.session_state['db']) + 1,
                    "sn": sn, "model": model, "operator": operator,
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "problem": problem, "action": action, "current": current
                }
                # 追加到数据库
                st.session_state['db'] = pd.concat([st.session_state['db'], pd.DataFrame([new_data])], ignore_index=True)
                st.success("保存成功！")

    elif menu == "🔍 历史查询":
        st.title("维修档案库")
        
        # 搜索框
        search_sn = st.text_input("搜索序列号:")
        
        # 过滤数据
        if search_sn:
            df_show = st.session_state['db'][st.session_state['db']['sn'].str.contains(search_sn)]
        else:
            df_show = st.session_state['db']

        # 展示每一行
        for index, row in df_show.iterrows():
            with st.expander(f"{row['date']} - SN: {row['sn']} (操作员: {row['operator']})"):
                st.write(f"**故障:** {row['problem']}")
                st.write(f"**措施:** {row['action']}")
                
                col_a, col_b = st.columns([1, 1])
                
                # 下载按钮
                with col_a:
                    doc_file = generate_doc(row)
                    st.download_button(
                        label="📥 下载 Word 报告",
                        data=doc_file,
                        file_name=f"Report_{row['sn']}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{row['id']}"
                    )
                
                # 删除按钮 (仅管理员可见)
                if st.session_state['role'] == "admin":
                    with col_b:
                        if st.button("🗑️ 删除记录", key=f"del_{row['id']}"):
                            # 这里写删除逻辑
                            st.warning("模拟删除成功 (实际需要连接云数据库)")
