import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime
import os

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(page_title="激光器维修系统 (修复版)", page_icon="🔋", layout="wide")

# 初始化数据库
if 'db' not in st.session_state:
    st.session_state['db'] = []

# 初始化管理员状态
if 'is_admin' not in st.session_state:
    st.session_state['is_admin'] = False

# ==========================================
# 2. 状态管理与清空逻辑 (关键修复)
# ==========================================
# 为了实现“不使用Form也能在保存后清空”，我们需要手动管理这些输入框的状态
def init_input_states():
    defaults = {
        "sn_input": "", "model_input": "WYP-", "voltage_input": "24V", "operator_input": "Guest",
        "obs_case_input": "完好 Normal", "obs_mech_input": "无 None",
        "work_hours_input": "", "alarms_input": "No Alarm",
        "tec1_set_input": "", "tec1_read_input": "", "tec1_peltier_input": "",
        "tec2_set_input": "", "tec2_read_input": "", "tec2_peltier_input": "",
        "hv_input": "", "current_input": "", "pulse_input": "",
        "problem_input": "", "action_summary_input": "", "note_input": ""
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

    # 初始化表格数据 (用于DataEditor的重置)
    if "df_power" not in st.session_state:
        st.session_state.df_power = pd.DataFrame([{"电流 I [A]": "", "脉宽 [us]": "", "波长 λ": "", "功率 P [W]": ""}])
    if "df_output" not in st.session_state:
        st.session_state.df_output = pd.DataFrame([{"355nm": "", "532nm": "", "1064nm": ""}])
    if "df_action" not in st.session_state:
        st.session_state.df_action = pd.DataFrame([{"维修措施": "", "操作员": "Guest", "日期": datetime.now().strftime("%Y-%m-%d")}])

def clear_all_inputs():
    """保存成功后调用此函数，强制重置所有输入框"""
    # 重置文本框
    st.session_state["sn_input"] = ""
    st.session_state["model_input"] = "WYP-"
    st.session_state["problem_input"] = ""
    st.session_state["action_summary_input"] = ""
    st.session_state["note_input"] = ""
    # ... 您可以根据需要重置更多字段，这里重置了核心字段
    
    # 重置表格
    st.session_state.df_power = pd.DataFrame([{"电流 I [A]": "", "脉宽 [us]": "", "波长 λ": "", "功率 P [W]": ""}])
    st.session_state.df_action = pd.DataFrame([{"维修措施": "", "操作员": st.session_state.get("operator_input", "Guest"), "日期": datetime.now().strftime("%Y-%m-%d")}])

# 运行初始化
init_input_states()

# ==========================================
# 3. 核心逻辑函数
# ==========================================

def flatten_data_for_template(record):
    """
    数据拍平处理：解决变量冲突的关键步骤
    """
    # 1. 基础复制 (包含 action, problem 等)
    context = record.copy()
    
    # 2. 处理功率表 (Power Table) -> {{ current_1 }}, {{ current_2 }} ...
    power_data = record.get('power_table', [])
    for i, row in enumerate(power_data):
        suffix = f"_{i+1}"
        context[f"current{suffix}"] = row.get("电流 I [A]", "")
        context[f"pulse{suffix}"] = row.get("脉宽 [us]", "")
        context[f"nm{suffix}"] = row.get("波长 λ", "")
        context[f"power{suffix}"] = row.get("功率 P [W]", "")
    
    # 3. 处理输出功率 (Output Table)
    output_data = record.get('output_table', [])
    for i, row in enumerate(output_data):
        suffix = f"_{i+1}"
        context[f"power_355{suffix}"] = row.get("355nm", "")
        context[f"power_532{suffix}"] = row.get("532nm", "")
        context[f"power_1064{suffix}"] = row.get("1064nm", "")

    # 4. 处理维修步骤表 (Action Table) -> {{ action_1 }}, {{ action_2 }} ...
    # 【重点】这里生成的 key 是 action_1, action_2，绝对不会覆盖 record['action'] (这是总体描述)
    action_data = record.get('action_table', [])
    for i, row in enumerate(action_data):
        suffix = f"_{i+1}"
        context[f"action{suffix}"] = row.get("维修措施", "")
        context[f"operator{suffix}"] = row.get("操作员", "")
        context[f"date{suffix}"] = row.get("日期", "")
        
    return context

def generate_doc(record):
    if not os.path.exists("template.docx"):
        return None
    doc = DocxTemplate("template.docx")
    final_context = flatten_data_for_template(record)
    try:
        doc.render(final_context)
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"生成文档出错: {e}")
        return None

# ==========================================
# 4. 侧边栏：管理员登录
# ==========================================
with st.sidebar:
    st.header("🔧 系统菜单")
    with st.expander("👮‍♂️ 管理员登录"):
        if not st.session_state['is_admin']:
            adm_user = st.text_input("账号")
            adm_pwd = st.text_input("密码", type="password")
            if st.button("登录"):
                if adm_user == "admin" and adm_pwd == "admin":
                    st.session_state['is_admin'] = True
                    st.rerun()
                else:
                    st.error("账号或密码错误")
        else:
            st.success("已登录为管理员")
            if st.button("退出管理员"):
                st.session_state['is_admin'] = False
                st.rerun()

# ==========================================
# 5. 主界面
# ==========================================
st.title("🔋 激光器维修档案系统")

tab1, tab2 = st.tabs(["📝 录入新记录", "🔍 历史档案库"])

# --- TAB 1: 录入界面 (无Form模式) ---
with tab1:
    st.info("💡 提示：所有输入框现在按 Enter 不会自动保存。只有点击最底部的按钮才会提交。")
    
    # 我们直接使用 columns 布局，绑定 key 到 session_state
    
    # Section 1: 基础信息
    st.subheader("1. 基础信息")
    c1, c2, c3, c4 = st.columns(4)
    sn = c1.text_input("序列号 (Serial No.)", key="sn_input")
    model = c2.text_input("型号 (Model)", key="model_input")
    voltage = c3.text_input("电压 (Voltage)", key="voltage_input")
    operator = c4.text_input("当前操作员", key="operator_input")
    
    # Section 2: 外观
    st.subheader("2. 外观检查")
    c1, c2 = st.columns(2)
    obs_case = c1.text_input("外壳/包装状态", key="obs_case_input")
    obs_mech = c2.text_input("机械损伤", key="obs_mech_input")

    # Section 3: 电子与TEC
    with st.expander("3. 电子参数与 TEC 设置 (点击展开)", expanded=False):
        e1, e2 = st.columns(2)
        work_hours = e1.text_input("工作时长 (Hours)", key="work_hours_input")
        alarms = e2.text_input("报警状态 (Alarms)", key="alarms_input")
        
        st.markdown("**TEC 1 设置**")
        t1_1, t1_2, t1_3 = st.columns(3)
        tec1_set = t1_1.text_input("TEC1 设定值", key="tec1_set_input")
        tec1_read = t1_2.text_input("TEC1 回读值", key="tec1_read_input")
        tec1_peltier = t1_3.text_input("TEC1 电流", key="tec1_peltier_input")

        st.markdown("**TEC 2 设置**")
        t2_1, t2_2, t2_3 = st.columns(3)
        tec2_set = t2_1.text_input("TEC2 设定值", key="tec2_set_input")
        tec2_read = t2_2.text_input("TEC2 回读值", key="tec2_read_input")
        tec2_peltier = t2_3.text_input("TEC2 电流", key="tec2_peltier_input")
        
        st.markdown("**驱动参数**")
        h1, h2, h3 = st.columns(3)
        hv = h1.text_input("高压 (HV)", key="hv_input")
        current = h2.text_input("峰值电流 (I Peak)", key="current_input")
        pulse = h3.text_input("脉宽 (Pulse)", key="pulse_input")

    # Section 4: 动态表格 (绑定 Session State 数据源)
    st.subheader("4. 功率测量数据 (支持多行)")
    
    # 【重要】DataEditor 必须绑定到 session_state 才能实现保存后重置
    edited_power_df = st.data_editor(st.session_state.df_power, num_rows="dynamic", use_container_width=True, key="editor_power")

    st.markdown("**输出功率 (Output Power)**")
    edited_output_df = st.data_editor(st.session_state.df_output, num_rows="fixed", use_container_width=True, key="editor_output")

    # Section 5: 故障与动态维修记录
    st.subheader("5. 故障分析与维修日志")
    problem = st.text_area("故障描述 (Description)", height=80, key="problem_input")
    
    # 【注意】这里是总体描述，对应模板 {{ action }}
    action_summary = st.text_area("采取措施总体描述 (Action Taken)", height=80, key="action_summary_input")
    
    st.markdown("**详细维修步骤记录 (Repair Actions Table)**")
    # 【注意】这里是详细步骤，对应模板 {{ action_1 }}, {{ action_2 }}...
    edited_action_df = st.data_editor(st.session_state.df_action, num_rows="dynamic", use_container_width=True, key="editor_action")
    
    note = st.text_area("备注 (Notes)", key="note_input")

    # ================= 保存逻辑 =================
    st.markdown("---")
    # 使用普通的 button，不使用 form_submit_button
    if st.button("💾 保存完整记录", type="primary"):
        # 1. 验证
        if not sn:
            st.error("❌ 保存失败：序列号不能为空！")
        else:
            # 2. 收集数据
            new_record = {
                "id": len(st.session_state['db']) + 1,
                "sn": sn, "model": model, "voltage": voltage, "operator": operator,
                "date": datetime.now().strftime("%Y-%m-%d"),
                "obs_case": obs_case, "obs_mech": obs_mech,
                "work_hours": work_hours, "alarms": alarms,
                "tec1_set": tec1_set, "tec1_read": tec1_read, "tec1_peltier": tec1_peltier,
                "tec2_set": tec2_set, "tec2_read": tec2_read, "tec2_peltier": tec2_peltier,
                "hv": hv, "current": current, "pulse": pulse,
                "problem": problem, 
                "action": action_summary, # 存为 'action' 供模板使用
                "note": note,
                # 收集表格数据
                "power_table": edited_power_df.to_dict('records'),
                "output_table": edited_output_df.to_dict('records'),
                "action_table": edited_action_df.to_dict('records')
            }
            
            # 3. 存入数据库
            st.session_state['db'].append(new_record)
            st.success(f"✅ 序列号 {sn} 的记录已成功保存！")
            
            # 4. 清空输入框并刷新页面
            clear_all_inputs()
            st.rerun()

# --- TAB 2: 查询界面 ---
with tab2:
    st.header("🗄️ 维修档案库")
    
    search_term = st.text_input("🔍 输入序列号搜索：")
    
    display_data = st.session_state['db']
    if search_term:
        display_data = [d for d in display_data if search_term.lower() in d['sn'].lower()]

    if not display_data:
        st.info("暂无数据。")
    else:
        for i, record in enumerate(reversed(display_data)):
            with st.expander(f"📅 {record['date']} | SN: {record['sn']} | 操作员: {record['operator']}"):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.markdown(f"**故障:** {record['problem']}")
                    st.markdown(f"**措施(总体):** {record['action']}") 
                with col2:
                    doc_file = generate_doc(record)
                    if doc_file:
                        st.download_button(
                            label="📥 下载 Word",
                            data=doc_file,
                            file_name=f"Report_{record['sn']}_{record['date']}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_{record['id']}" # 使用唯一ID作为key
                        )
                    else:
                        st.warning("缺少模板文件")
                    
                    if st.session_state['is_admin']:
                        if st.button("🗑️ 删除记录", key=f"del_{record['id']}"):
                            st.session_state['db'] = [d for d in st.session_state['db'] if d['id'] != record['id']]
                            st.rerun()
