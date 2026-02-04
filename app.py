import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime
import os

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(page_title="激光器维修系统 (稳定版)", page_icon="🔋", layout="wide")

# 初始化数据库
if 'db' not in st.session_state:
    st.session_state['db'] = []

# 初始化管理员状态
if 'is_admin' not in st.session_state:
    st.session_state['is_admin'] = False

# 初始化消息提示状态 (用于Callback反馈)
if 'msg_type' not in st.session_state:
    st.session_state['msg_type'] = None # success / error
if 'msg_content' not in st.session_state:
    st.session_state['msg_content'] = ""

# ==========================================
# 2. 状态管理与清空逻辑
# ==========================================
def init_input_states():
    # 定义所有输入框的默认值
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

    # 初始化表格数据源
    if "df_power" not in st.session_state:
        st.session_state.df_power = pd.DataFrame([{"电流 I [A]": "", "脉宽 [us]": "", "波长 λ": "", "功率 P [W]": ""}])
    if "df_output" not in st.session_state:
        st.session_state.df_output = pd.DataFrame([{"355nm": "", "532nm": "", "1064nm": ""}])
    if "df_action" not in st.session_state:
        st.session_state.df_action = pd.DataFrame([{"维修措施": "", "操作员": "Guest", "日期": datetime.now().strftime("%Y-%m-%d")}])

# 运行初始化
init_input_states()

# ==========================================
# 3. 核心逻辑函数
# ==========================================
def flatten_data_for_template(record):
    """数据拍平处理"""
    context = record.copy()
    
    # Power Table
    power_data = record.get('power_table', [])
    for i, row in enumerate(power_data):
        suffix = f"_{i+1}"
        context[f"current{suffix}"] = row.get("电流 I [A]", "")
        context[f"pulse{suffix}"] = row.get("脉宽 [us]", "")
        context[f"nm{suffix}"] = row.get("波长 λ", "")
        context[f"power{suffix}"] = row.get("功率 P [W]", "")
    
    # Output Table
    output_data = record.get('output_table', [])
    for i, row in enumerate(output_data):
        suffix = f"_{i+1}"
        context[f"power_355{suffix}"] = row.get("355nm", "")
        context[f"power_532{suffix}"] = row.get("532nm", "")
        context[f"power_1064{suffix}"] = row.get("1064nm", "")

    # Action Table
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
        return None

# ==========================================
# 4. 【关键修复】保存回调函数
# ==========================================
def save_data_callback():
    """
    这是一个回调函数。
    它会在点击按钮后、页面重新刷新前执行。
    只有在这里，我们才能安全地清空输入框。
    """
    # 1. 从 session_state 获取当前输入框的值
    sn = st.session_state.sn_input
    
    # 2. 验证
    if not sn:
        st.session_state['msg_type'] = 'error'
        st.session_state['msg_content'] = "❌ 保存失败：序列号不能为空！"
        return # 验证失败，直接结束，不清空输入框

    # 3. 收集数据
    new_record = {
        "id": len(st.session_state['db']) + 1,
        "sn": sn, 
        "model": st.session_state.model_input, 
        "voltage": st.session_state.voltage_input, 
        "operator": st.session_state.operator_input,
        "date": datetime.now().strftime("%Y-%m-%d"),
        "obs_case": st.session_state.obs_case_input, 
        "obs_mech": st.session_state.obs_mech_input,
        "work_hours": st.session_state.work_hours_input, 
        "alarms": st.session_state.alarms_input,
        "tec1_set": st.session_state.tec1_set_input, 
        "tec1_read": st.session_state.tec1_read_input, 
        "tec1_peltier": st.session_state.tec1_peltier_input,
        "tec2_set": st.session_state.tec2_set_input, 
        "tec2_read": st.session_state.tec2_read_input, 
        "tec2_peltier": st.session_state.tec2_peltier_input,
        "hv": st.session_state.hv_input, 
        "current": st.session_state.current_input, 
        "pulse": st.session_state.pulse_input,
        "problem": st.session_state.problem_input, 
        "action": st.session_state.action_summary_input, # 总体描述
        "note": st.session_state.note_input,
        # 获取表格数据 (DataEditor 的数据会自动同步到绑定的 session_state key 中，但这里我们需要取它的 value)
        # 注意：DataEditor 绑定的 key 在 session_state 中就是修改后的 DataFrame
        "power_table": st.session_state.editor_power.to_dict('records'),
        "output_table": st.session_state.editor_output.to_dict('records'),
        "action_table": st.session_state.editor_action.to_dict('records')
    }

    # 4. 存入数据库
    st.session_state['db'].append(new_record)
    
    # 5. 设置成功消息
    st.session_state['msg_type'] = 'success'
    st.session_state['msg_content'] = f"✅ 序列号 {sn} 的记录已成功保存！"

    # 6. 【安全清空】直接修改 session_state，准备下一次渲染
    st.session_state.sn_input = ""
    st.session_state.problem_input = ""
    st.session_state.action_summary_input = ""
    st.session_state.note_input = ""
    # 重置其他字段为默认值...
    st.session_state.model_input = "WYP-"
    
    # 重置表格数据源 (这样 DataEditor 重新加载时就是空的)
    st.session_state.df_power = pd.DataFrame([{"电流 I [A]": "", "脉宽 [us]": "", "波长 λ": "", "功率 P [W]": ""}])
    st.session_state.df_action = pd.DataFrame([{"维修措施": "", "操作员": st.session_state.operator_input, "日期": datetime.now().strftime("%Y-%m-%d")}])
    # 注意：Output 表格一般不需要重置为空，保留默认结构即可

# ==========================================
# 5. 侧边栏与主界面
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

st.title("🔋 激光器维修档案系统")

# 顶部消息提示区 (处理 Callback 的反馈)
if st.session_state['msg_type'] == 'success':
    st.success(st.session_state['msg_content'])
    st.session_state['msg_type'] = None # 显示一次后重置
elif st.session_state['msg_type'] == 'error':
    st.error(st.session_state['msg_content'])
    st.session_state['msg_type'] = None

tab1, tab2 = st.tabs(["📝 录入新记录", "🔍 历史档案库"])

# --- TAB 1: 录入界面 ---
with tab1:
    st.info("💡 提示：所有输入框现在按 Enter 不会自动保存。只有点击最底部的按钮才会提交。")
    
    # 基础信息
    st.subheader("1. 基础信息")
    c1, c2, c3, c4 = st.columns(4)
    st.text_input("序列号 (Serial No.)", key="sn_input")
    st.text_input("型号 (Model)", key="model_input")
    st.text_input("电压 (Voltage)", key="voltage_input")
    st.text_input("当前操作员", key="operator_input")
    
    # 外观
    st.subheader("2. 外观检查")
    c1, c2 = st.columns(2)
    st.text_input("外壳/包装状态", key="obs_case_input")
    st.text_input("机械损伤", key="obs_mech_input")

    # 电子与TEC
    with st.expander("3. 电子参数与 TEC 设置 (点击展开)", expanded=False):
        e1, e2 = st.columns(2)
        st.text_input("工作时长 (Hours)", key="work_hours_input")
        st.text_input("报警状态 (Alarms)", key="alarms_input")
        
        st.markdown("**TEC 1 设置**")
        t1_1, t1_2, t1_3 = st.columns(3)
        st.text_input("TEC1 设定值", key="tec1_set_input")
        st.text_input("TEC1 回读值", key="tec1_read_input")
        st.text_input("TEC1 电流", key="tec1_peltier_input")

        st.markdown("**TEC 2 设置**")
        t2_1, t2_2, t2_3 = st.columns(3)
        st.text_input("TEC2 设定值", key="tec2_set_input")
        st.text_input("TEC2 回读值", key="tec2_read_input")
        st.text_input("TEC2 电流", key="tec2_peltier_input")
        
        st.markdown("**驱动参数**")
        h1, h2, h3 = st.columns(3)
        st.text_input("高压 (HV)", key="hv_input")
        st.text_input("峰值电流 (I Peak)", key="current_input")
        st.text_input("脉宽 (Pulse)", key="pulse_input")

    # 动态表格
    st.subheader("4. 功率测量数据 (支持多行)")
    # 绑定 st.session_state.df_power 确保重置生效
    st.data_editor(st.session_state.df_power, num_rows="dynamic", use_container_width=True, key="editor_power")

    st.markdown("**输出功率 (Output Power)**")
    st.data_editor(st.session_state.df_output, num_rows="fixed", use_container_width=True, key="editor_output")

    # 故障与维修
    st.subheader("5. 故障分析与维修日志")
    st.text_area("故障描述 (Description)", height=80, key="problem_input")
    st.text_area("采取措施总体描述 (Action Taken)", height=80, key="action_summary_input")
    
    st.markdown("**详细维修步骤记录 (Repair Actions Table)**")
    st.data_editor(st.session_state.df_action, num_rows="dynamic", use_container_width=True, key="editor_action")
    
    st.text_area("备注 (Notes)", key="note_input")

    st.markdown("---")
    # 【关键修改】使用 on_click 绑定回调函数
    # 这样点击按钮时，先执行 save_data_callback 清空数据，然后再刷新页面，就不会报错了
    st.button("💾 保存完整记录", type="primary", on_click=save_data_callback)

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
                            key=f"dl_{record['id']}"
                        )
                    else:
                        st.warning("缺少模板文件")
                    
                    if st.session_state['is_admin']:
                        if st.button("🗑️ 删除记录", key=f"del_{record['id']}"):
                            # 删除逻辑
                            st.session_state['db'] = [d for d in st.session_state['db'] if d['id'] != record['id']]
                            st.rerun()
