import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime
import os

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(page_title="激光器维修系统 (完美版)", page_icon="🔋", layout="wide")

# 初始化数据库
if 'db' not in st.session_state:
    st.session_state['db'] = []

# 初始化管理员状态
if 'is_admin' not in st.session_state:
    st.session_state['is_admin'] = False

# 初始化重置标志位 (用于清空输入框)
if 'reset_trigger' not in st.session_state:
    st.session_state['reset_trigger'] = False

# ==========================================
# 2. 核心逻辑函数
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
# 3. 侧边栏与主界面
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

tab1, tab2 = st.tabs(["📝 录入新记录", "🔍 历史档案库"])

# --- TAB 1: 录入界面 ---
with tab1:
    st.info("💡 提示：所有输入框现在按 Enter 不会自动保存。只有点击最底部的按钮才会提交。")
    
    # 【核心技巧】使用 reset_trigger 来控制是否清空
    # 如果刚保存完，reset_trigger 为 True，我们就不给 default value，或者给空值
    # 但由于 Streamlit 的 text_input 没有直接的 "clear" 方法，我们通过 key 的变化来强制重置组件
    # 或者简单点：我们手动检查 reset_trigger，如果是 True，就用空字符串初始化，否则用 session state
    
    if st.session_state['reset_trigger']:
        # 刚刚保存过，需要重置所有默认值
        default_sn = ""
        default_problem = ""
        default_action = ""
        default_note = ""
        # 强制重置 DataFrame
        df_power_source = pd.DataFrame([{"电流 I [A]": "", "脉宽 [us]": "", "波长 λ": "", "功率 P [W]": ""}])
        df_output_source = pd.DataFrame([{"355nm": "", "532nm": "", "1064nm": ""}])
        df_action_source = pd.DataFrame([{"维修措施": "", "操作员": "Guest", "日期": datetime.now().strftime("%Y-%m-%d")}])
        # 重置标志位，防止死循环 (但要等到页面渲染完，所以在最后重置)
    else:
        # 正常状态，保持用户输入（这里其实不需要做太多，Streamlit 会自动保持，除非我们想回填数据）
        # 为了简单，我们每次都给默认值，依靠 st.session_state 自动保持输入
        default_sn = st.session_state.get("_sn_val", "")
        default_problem = st.session_state.get("_prob_val", "")
        default_action = st.session_state.get("_act_val", "")
        default_note = st.session_state.get("_note_val", "")
        
        # 表格数据源需要持久化，否则每次刷新都空了
        if 'df_power_cache' not in st.session_state:
            st.session_state.df_power_cache = pd.DataFrame([{"电流 I [A]": "", "脉宽 [us]": "", "波长 λ": "", "功率 P [W]": ""}])
        df_power_source = st.session_state.df_power_cache
        
        if 'df_output_cache' not in st.session_state:
            st.session_state.df_output_cache = pd.DataFrame([{"355nm": "", "532nm": "", "1064nm": ""}])
        df_output_source = st.session_state.df_output_cache
        
        if 'df_action_cache' not in st.session_state:
            st.session_state.df_action_cache = pd.DataFrame([{"维修措施": "", "操作员": "Guest", "日期": datetime.now().strftime("%Y-%m-%d")}])
        df_action_source = st.session_state.df_action_cache

    # --- 开始绘制表单 (直接使用返回值) ---
    
    st.subheader("1. 基础信息")
    c1, c2, c3, c4 = st.columns(4)
    # 使用 key="_xxx_val" 来让 Streamlit 自动管理状态，但在 key 变化时会重置
    # 为了实现清空，我们这里使用 value 参数 + key
    # 这里的技巧是：当 reset_trigger 为 True 时，我们不传 value (或者传空)，但 key 必须变一下才能强制刷新？
    # 不，更简单的办法是：使用 st.empty() 或者回调清空 session state 对应的 key。
    # 让我们回到最稳妥的 key 绑定法，但在保存时，手动清空 session_state[key]
    
    sn = st.text_input("序列号 (Serial No.)", key="sn_key")
    model = st.text_input("型号 (Model)", value="WYP-", key="model_key")
    voltage = st.text_input("电压 (Voltage)", value="24V", key="voltage_key")
    operator = st.text_input("当前操作员", value="Guest", key="operator_key")
    
    st.subheader("2. 外观检查")
    c1, c2 = st.columns(2)
    obs_case = c1.text_input("外壳/包装状态", value="完好 Normal", key="case_key")
    obs_mech = c2.text_input("机械损伤", value="无 None", key="mech_key")

    with st.expander("3. 电子参数与 TEC 设置", expanded=False):
        e1, e2 = st.columns(2)
        wh = st.text_input("工作时长", key="wh_key")
        alarms = st.text_input("报警状态", value="No Alarm", key="alarm_key")
        
        st.markdown("**TEC 1**")
        t1_1, t1_2, t1_3 = st.columns(3)
        tec1_s = st.text_input("TEC1 设定", key="t1s_key")
        tec1_r = st.text_input("TEC1 回读", key="t1r_key")
        tec1_p = st.text_input("TEC1 电流", key="t1p_key")

        st.markdown("**TEC 2**")
        t2_1, t2_2, t2_3 = st.columns(3)
        tec2_s = st.text_input("TEC2 设定", key="t2s_key")
        tec2_r = st.text_input("TEC2 回读", key="t2r_key")
        tec2_p = st.text_input("TEC2 电流", key="t2p_key")
        
        st.markdown("**驱动**")
        h1, h2, h3 = st.columns(3)
        hv = st.text_input("高压 (HV)", key="hv_key")
        curr = st.text_input("峰值电流", key="curr_key")
        puls = st.text_input("脉宽", key="puls_key")

    st.subheader("4. 功率测量数据")
    # 【关键】直接获取编辑后的 DataFrame
    edited_power_df = st.data_editor(df_power_source, num_rows="dynamic", use_container_width=True, key="editor_power_new")
    # 实时更新缓存，防止刷新丢失
    st.session_state.df_power_cache = edited_power_df 

    st.markdown("**输出功率**")
    edited_output_df = st.data_editor(df_output_source, num_rows="fixed", use_container_width=True, key="editor_output_new")
    st.session_state.df_output_cache = edited_output_df

    st.subheader("5. 故障分析与维修日志")
    problem = st.text_area("故障描述", height=80, key="prob_key")
    action_sum = st.text_area("采取措施 (总体)", height=80, key="act_key")
    
    st.markdown("**详细维修步骤**")
    edited_action_df = st.data_editor(df_action_source, num_rows="dynamic", use_container_width=True, key="editor_action_new")
    st.session_state.df_action_cache = edited_action_df
    
    note = st.text_area("备注", key="note_key")

    st.markdown("---")
    
    # ================= 保存逻辑 (无需回调，直接写在按钮逻辑里) =================
    if st.button("💾 保存完整记录", type="primary"):
        if not sn:
            st.error("❌ 序列号不能为空！")
        else:
            # 1. 收集数据 (直接使用上面的变量)
            new_record = {
                "id": len(st.session_state['db']) + 1,
                "sn": sn, "model": model, "voltage": voltage, "operator": operator,
                "date": datetime.now().strftime("%Y-%m-%d"),
                "obs_case": obs_case, "obs_mech": obs_mech,
                "work_hours": wh, "alarms": alarms,
                "tec1_set": tec1_s, "tec1_read": tec1_r, "tec1_peltier": tec1_p,
                "tec2_set": tec2_s, "tec2_read": tec2_r, "tec2_peltier": tec2_p,
                "hv": hv, "current": curr, "pulse": puls,
                "problem": problem, "action": action_sum, "note": note,
                # 2. 表格数据 (直接用 edited_power_df.to_dict，因为它是真的 DataFrame)
                "power_table": edited_power_df.to_dict('records'),
                "output_table": edited_output_df.to_dict('records'),
                "action_table": edited_action_df.to_dict('records')
            }
            
            st.session_state['db'].append(new_record)
            st.success(f"✅ 序列号 {sn} 保存成功！")
            
            # 3. 清空数据 (简单粗暴法：直接删 key 或置空 session state)
            st.session_state["sn_key"] = ""
            st.session_state["prob_key"] = ""
            st.session_state["act_key"] = ""
            st.session_state["note_key"] = ""
            # 清空表格缓存
            st.session_state.df_power_cache = pd.DataFrame([{"电流 I [A]": "", "脉宽 [us]": "", "波长 λ": "", "功率 P [W]": ""}])
            st.session_state.df_action_cache = pd.DataFrame([{"维修措施": "", "操作员": "Guest", "日期": datetime.now().strftime("%Y-%m-%d")}])
            
            # 4. 强制重跑一次，让清空生效
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
                            key=f"dl_{record['id']}"
                        )
                    else:
                        st.warning("缺少模板文件")
                    
                    if st.session_state['is_admin']:
                        if st.button("🗑️ 删除记录", key=f"del_{record['id']}"):
                            st.session_state['db'] = [d for d in st.session_state['db'] if d['id'] != record['id']]
                            st.rerun()
