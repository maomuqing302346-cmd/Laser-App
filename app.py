import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime
import os

# ==========================================
# 1. 初始化设置
# ==========================================
st.set_page_config(page_title="激光器维修系统 (完整版)", page_icon="🔧", layout="wide")

# 定义所有需要的字段 (对应模板里的标签)
ALL_COLUMNS = [
    "id", "sn", "model", "voltage", "operator", "date",
    "obs_case", "obs_mech",
    "work_hours", "alarms",
    "tec1_set", "tec1_read", "tec1_peltier",
    "tec2_set", "tec2_read", "tec2_peltier",
    "hv", "current", "pulse",
    "current_1", "pulse_1", "nm_1", "power_1", # 二极管测量(第一行)
    "power_355", "power_532", "power_1064",
    "problem", "action", "note",
    # 维修记录表 (3行)
    "action_1", "operator_1", "date_1",
    "action_2", "operator_2", "date_2",
    "action_3", "operator_3", "date_3"
]

# 初始化数据库
if 'db' not in st.session_state:
    st.session_state['db'] = pd.DataFrame(columns=ALL_COLUMNS)

# 初始化登录状态
for key in ['authenticated', 'current_user', 'role']:
    if key not in st.session_state:
        st.session_state[key] = None if key != 'authenticated' else False

# ==========================================
# 2. 核心功能函数
# ==========================================
def generate_doc(record):
    """生成 Word 文档"""
    if not os.path.exists("template.docx"):
        return None
    
    doc = DocxTemplate("template.docx")
    
    # 将记录转换为字典，并处理空值为 ""，防止 Word 报错
    context = record.to_dict()
    for k, v in context.items():
        if pd.isna(v) or v is None:
            context[k] = ""
    
    # 渲染模板
    doc.render(context)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def login_page():
    """登录页面"""
    st.markdown("<h1 style='text-align: center;'>🔐 激光器维修系统登录</h1>", unsafe_allow_html=True)
    st.write("")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("用户名")
            password = st.text_input("密码", type="password")
            submitted = st.form_submit_button("登录", use_container_width=True)
            
            if submitted:
                if username == "admin" and password == "admin123":
                    st.session_state['authenticated'] = True
                    st.session_state['role'] = 'admin'
                    st.session_state['current_user'] = 'admin'
                    st.rerun()
                elif username == "user":
                    st.session_state['authenticated'] = True
                    st.session_state['role'] = 'user'
                    st.session_state['current_user'] = 'user'
                    st.rerun()
                else:
                    st.error("用户名或密码错误")

def main_app():
    """主程序界面"""
    
    # --- 侧边栏 ---
    st.sidebar.markdown(f"👤 **操作员:** {st.session_state['current_user']}")
    if st.sidebar.button("🚪 退出登录"):
        st.session_state['authenticated'] = False
        st.rerun()
    
    st.sidebar.divider()
    menu = st.sidebar.radio("导航", ["📝 录入新单", "🔍 历史查询"])

    # --- 录入页面 ---
    if menu == "📝 录入新单":
        st.title("📝 新建维修工单")
        
        with st.form("full_repair_form"):
            # 1. 基础信息
            st.subheader("1. 基础信息")
            c1, c2, c3 = st.columns(3)
            sn = c1.text_input("序列号 {{sn}}")
            model = c2.text_input("型号 {{model}}", value="WYP-")
            voltage = c3.text_input("电压 {{voltage}}", value="24V")
            
            # 2. 外观检查
            st.subheader("2. 外观检查")
            c1, c2 = st.columns(2)
            obs_case = c1.text_input("外壳/包装 {{obs_case}}", value="完好 Normal")
            obs_mech = c2.text_input("机械损伤 {{obs_mech}}", value="无 None")
            
            # 3. 电子参数与TEC (使用折叠面板节省空间)
            with st.expander("3. 电子参数与 TEC 设置 (点击展开)", expanded=True):
                e1, e2 = st.columns(2)
                work_hours = e1.text_input("工作时长 {{work_hours}}")
                alarms = e2.text_input("报警状态 {{alarms}}", value="No Alarm")
                
                st.markdown("---")
                st.caption("TEC 1 设置")
                t1_1, t1_2, t1_3 = st.columns(3)
                tec1_set = t1_1.text_input("Set {{tec1_set}}")
                tec1_read = t1_2.text_input("Read {{tec1_read}}")
                tec1_peltier = t1_3.text_input("Peltier {{tec1_peltier}}")

                st.caption("TEC 2 设置")
                t2_1, t2_2, t2_3 = st.columns(3)
                tec2_set = t2_1.text_input("Set {{tec2_set}}")
                tec2_read = t2_2.text_input("Read {{tec2_read}}")
                tec2_peltier = t2_3.text_input("Peltier {{tec2_peltier}}")
                
                st.markdown("---")
                h1, h2, h3 = st.columns(3)
                hv = h1.text_input("高压 HV {{hv}}")
                current = h2.text_input("I Peak {{current}}")
                pulse = h3.text_input("Tau Pulse {{pulse}}")

            # 4. 功率测量
            with st.expander("4. 功率测量数据", expanded=True):
                st.caption("二极管功率测量 (Row 1)")
                d1, d2, d3, d4 = st.columns(4)
                current_1 = d1.text_input("电流 I [A] {{current_1}}")
                pulse_1 = d2.text_input("脉宽 [us] {{pulse_1}}")
                nm_1 = d3.text_input("波长 λ {{nm_1}}")
                power_1 = d4.text_input("功率 P [W] {{power_1}}")
                
                st.caption("输出功率 (Output Power)")
                p1, p2, p3 = st.columns(3)
                power_355 = p1.text_input("355nm {{power_355}}")
                power_532 = p2.text_input("532nm {{power_532}}")
                power_1064 = p3.text_input("1064nm {{power_1064}}")

            # 5. 故障分析与措施
            st.subheader("5. 故障分析与维修日志")
            problem = st.text_area("故障描述 {{problem}}", height=80)
            action = st.text_area("采取措施 (总体) {{action}}", height=80)
            
            st.caption("详细维修步骤记录 (Repair Actions Table)")
            r1_1, r1_2, r1_3 = st.columns([3, 1, 1])
            action_1 = r1_1.text_input("步骤1 内容 {{action_1}}")
            operator_1 = r1_2.text_input("操作员1", value=st.session_state['current_user'])
            date_1 = r1_3.text_input("日期1", value=datetime.now().strftime("%Y-%m-%d"))
            
            r2_1, r2_2, r2_3 = st.columns([3, 1, 1])
            action_2 = r2_1.text_input("步骤2 内容 {{action_2}}")
            operator_2 = r2_2.text_input("操作员2")
            date_2 = r2_3.text_input("日期2")

            r3_1, r3_2, r3_3 = st.columns([3, 1, 1])
            action_3 = r3_1.text_input("步骤3 内容 {{action_3}}")
            operator_3 = r3_2.text_input("操作员3")
            date_3 = r3_3.text_input("日期3")
            
            note = st.text_area("备注 (NOTES) {{note}}")

            # 提交按钮
            submitted = st.form_submit_button("💾 保存完整记录", type="primary")
            
            if submitted:
                if not sn:
                    st.error("❌ 序列号必填！")
                else:
                    new_id = len(st.session_state['db']) + 1
                    # 收集所有数据
                    new_data = {
                        "id": new_id,
                        "sn": sn, "model": model, "voltage": voltage,
                        "operator": st.session_state['current_user'],
                        "date": datetime.now().strftime("%Y-%m-%d"),
                        "obs_case": obs_case, "obs_mech": obs_mech,
                        "work_hours": work_hours, "alarms": alarms,
                        "tec1_set": tec1_set, "tec1_read": tec1_read, "tec1_peltier": tec1_peltier,
                        "tec2_set": tec2_set, "tec2_read": tec2_read, "tec2_peltier": tec2_peltier,
                        "hv": hv, "current": current, "pulse": pulse,
                        "current_1": current_1, "pulse_1": pulse_1, "nm_1": nm_1, "power_1": power_1,
                        "power_355": power_355, "power_532": power_532, "power_1064": power_1064,
                        "problem": problem, "action": action, "note": note,
                        "action_1": action_1, "operator_1": operator_1, "date_1": date_1,
                        "action_2": action_2, "operator_2": operator_2, "date_2": date_2,
                        "action_3": action_3, "operator_3": operator_3, "date_3": date_3,
                    }
                    st.session_state['db'] = pd.concat([st.session_state['db'], pd.DataFrame([new_data])], ignore_index=True)
                    st.success(f"✅ SN: {sn} 记录已保存！")

    # --- 查询页面 ---
    elif menu == "🔍 历史查询":
        st.title("🔍 维修档案库")
        search_sn = st.text_input("输入序列号搜索:")
        
        if not st.session_state['db'].empty:
            df = st.session_state['db']
            if search_sn:
                df = df[df['sn'].str.contains(search_sn, case=False, na=False)]
            
            st.write(f"共找到 {len(df)} 条记录")
            
            for idx, row in df.iterrows():
                with st.expander(f"{row['date']} | SN: {row['sn']} | 故障: {row['problem'][:20]}..."):
                    c1, c2 = st.columns(2)
                    c1.write(f"**操作员:** {row['operator']}")
                    c1.write(f"**型号:** {row['model']}")
                    c2.write(f"**措施:** {row['action']}")
                    c2.write(f"**1064nm功率:** {row['power_1064']}")
                    
                    # 下载按钮
                    doc_file = generate_doc(row)
                    if doc_file:
                        st.download_button(
                            label="📥 下载完整版 Word",
                            data=doc_file,
                            file_name=f"Report_{row['sn']}_{row['date']}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_{row['id']}"
                        )
                    else:
                        st.error("未找到 template.docx，请上传模板文件！")
                    
                    # 删除按钮
                    if st.session_state['role'] == 'admin':
                        if st.button("🗑️ 删除", key=f"del_{row['id']}"):
                            st.session_state['db'] = st.session_state['db'][st.session_state['db']['id'] != row['id']]
                            st.rerun()
        else:
            st.info("暂无数据")

# ==========================================
# 3. 启动
# ==========================================
if __name__ == "__main__":
    if st.session_state['authenticated']:
        main_app()
    else:
        login_page()
