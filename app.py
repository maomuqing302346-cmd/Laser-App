import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime
import os

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(page_title="激光器维修系统 (表格版)", page_icon="🔋", layout="wide")

# 初始化数据库
if 'db' not in st.session_state:
    st.session_state['db'] = []

# 初始化管理员状态
if 'is_admin' not in st.session_state:
    st.session_state['is_admin'] = False

# ==========================================
# 2. 初始化表格数据源 (用于清空和默认值)
# ==========================================
def init_dataframes():
    # 1. 基础信息表 (单行)
    if 'df_basic' not in st.session_state:
        st.session_state.df_basic = pd.DataFrame([
            {"序列号": "", "型号": "C-WEDG", "电压": "9V/15V", "操作员": ""}
        ])
    
    # 2. 外观检查表 (单行)
    if 'df_inspect' not in st.session_state:
        st.session_state.df_inspect = pd.DataFrame([
            {"外壳/包装": "完好 Normal", "机械损伤": "无 None"}
        ])

    # 3. 电子参数表 (单行)
    if 'df_elec' not in st.session_state:
        st.session_state.df_elec = pd.DataFrame([
            {"工作时长": "", "报警状态": "No Alarm"}
        ])

    # 4. TEC 参数表 (2行: TEC1, TEC2)
    if 'df_tec' not in st.session_state:
        st.session_state.df_tec = pd.DataFrame([
            {"名称": "TEC 1（Pump）", "设定值": "", "回读值": "", "电流": ""},
            {"名称": "TEC 2(Res)", "设定值": "", "回读值": "", "电流": ""}
        ])

    # 5. 驱动参数表 (单行)
    if 'df_driver' not in st.session_state:
        st.session_state.df_driver = pd.DataFrame([
            {"高压 (HV)": "", "峰值电流": "", "脉宽": ""}
        ])

    # 6. 功率测量表 (动态)
    if 'df_power' not in st.session_state:
        st.session_state.df_power = pd.DataFrame([
            {"电流 I [A]": "", "脉宽 [us]": "", "波长 λ": "", "功率 P [W]": ""}
        ])

    # 7. 输出功率表 (单行)
    if 'df_output' not in st.session_state:
        st.session_state.df_output = pd.DataFrame([
            {"355nm": "", "532nm": "", "1064nm": ""}
        ])
    
    # 8. 详细维修步骤 (动态)
    if 'df_action' not in st.session_state:
        st.session_state.df_action = pd.DataFrame([
            {"维修措施": "", "操作员": "Guest", "日期": datetime.now().strftime("%Y-%m-%d")}
        ])

    # 文本域状态
    if 'txt_problem' not in st.session_state: st.session_state.txt_problem = ""
    if 'txt_summary' not in st.session_state: st.session_state.txt_summary = ""
    if 'txt_note' not in st.session_state: st.session_state.txt_note = ""

def reset_all_data():
    """强制重置所有表格为默认状态"""
    del st.session_state.df_basic
    del st.session_state.df_inspect
    del st.session_state.df_elec
    del st.session_state.df_tec
    del st.session_state.df_driver
    del st.session_state.df_power
    del st.session_state.df_output
    del st.session_state.df_action
    st.session_state.txt_problem = ""
    st.session_state.txt_summary = ""
    st.session_state.txt_note = ""
    init_dataframes()

# 运行初始化
init_dataframes()

# ==========================================
# 3. 文档生成逻辑
# ==========================================
def flatten_data_for_template(record):
    context = record.copy()
    
    # 处理功率表
    for i, row in enumerate(record.get('power_table', [])):
        suffix = f"_{i+1}"
        context[f"current{suffix}"] = row.get("电流 I [A]", "")
        context[f"pulse{suffix}"] = row.get("脉宽 [us]", "")
        context[f"nm{suffix}"] = row.get("波长 λ", "")
        context[f"power{suffix}"] = row.get("功率 P [W]", "")
    
    # 处理输出功率
    for i, row in enumerate(record.get('output_table', [])):
        suffix = f"_{i+1}"
        context[f"power_355{suffix}"] = row.get("355nm", "")
        context[f"power_532{suffix}"] = row.get("532nm", "")
        context[f"power_1064{suffix}"] = row.get("1064nm", "")

    # 处理维修步骤
    for i, row in enumerate(record.get('action_table', [])):
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
# 4. 侧边栏：管理员
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
st.caption("全表格交互模式：在表格内按 Enter 仅确认输入，不会误提交。")

tab1, tab2 = st.tabs(["📝 录入工单", "🔍 历史档案"])

with tab1:
    # 1. 基础信息区 (使用表格代替输入框)
    st.subheader("1. 基础信息 & 外观")
    col1, col2 = st.columns([1.5, 1])
    with col1:
        st.caption("基础参数")
        # 这里的 num_rows="fixed" 禁止添加新行，只能修改第一行
        basic_df = st.data_editor(st.session_state.df_basic, num_rows="fixed", use_container_width=True, hide_index=True, key="ed_basic")
    with col2:
        st.caption("外观检查")
        inspect_df = st.data_editor(st.session_state.df_inspect, num_rows="fixed", use_container_width=True, hide_index=True, key="ed_inspect")

    # 2. 电子参数区
    st.subheader("2. 电子参数 & TEC")
    elec_df = st.data_editor(st.session_state.df_elec, num_rows="fixed", use_container_width=True, hide_index=True, key="ed_elec")
    
    c1, c2 = st.columns([1.5, 1])
    with c1:
        st.caption("TEC 参数 (请直接在表格内填写)")
        # TEC 表格预设了2行，用户直接填
        tec_df = st.data_editor(st.session_state.df_tec, num_rows="fixed", use_container_width=True, hide_index=True, key="ed_tec")
    with c2:
        st.caption("驱动参数 (Driver)")
        driver_df = st.data_editor(st.session_state.df_driver, num_rows="fixed", use_container_width=True, hide_index=True, key="ed_driver")

    # 3. 功率测量
    st.subheader("3. 功率测量 (支持多行)")
    power_df = st.data_editor(st.session_state.df_power, num_rows="dynamic", use_container_width=True, key="ed_power")
    
    st.caption("输出功率 (Output Power)")
    output_df = st.data_editor(st.session_state.df_output, num_rows="fixed", use_container_width=True, hide_index=True, key="ed_output")

    # 4. 故障描述 (保留文本域，支持回车换行)
    st.subheader("4. 故障与措施")
    problem = st.text_area("故障描述 ", value=st.session_state.txt_problem, height=100, key="area_problem")
    action_sum = st.text_area("采取措施-总体描述 ", value=st.session_state.txt_summary, height=100, key="area_summary")
    
    st.caption("详细维修步骤 ")
    action_df = st.data_editor(st.session_state.df_action, num_rows="dynamic", use_container_width=True, hide_index=True, key="ed_action")
    
    note = st.text_area("备注", value=st.session_state.txt_note, height=68, key="area_note")

    st.markdown("---")
    
    # ================= 保存按钮 =================
    if st.button("💾 保存完整记录", type="primary"):
        # 1. 从表格提取数据 (取第一行数据作为单值)
        try:
            # 基础信息取第0行
            sn_val = basic_df.iloc[0]["序列号"]
            
            if not sn_val:
                st.error("❌ 保存失败：【序列号】不能为空！")
            else:
                # 提取单行数据
                record = {
                    "id": len(st.session_state['db']) + 1,
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    
                    # 基础表
                    "sn": sn_val,
                    "model": basic_df.iloc[0]["型号"],
                    "voltage": basic_df.iloc[0]["电压"],
                    "operator": basic_df.iloc[0]["操作员"],
                    
                    # 外观表
                    "obs_case": inspect_df.iloc[0]["外壳/包装"],
                    "obs_mech": inspect_df.iloc[0]["机械损伤"],
                    
                    # 电子表
                    "work_hours": elec_df.iloc[0]["工作时长"],
                    "alarms": elec_df.iloc[0]["报警状态"],
                    
                    # 驱动表
                    "hv": driver_df.iloc[0]["高压 (HV)"],
                    "current": driver_df.iloc[0]["峰值电流"],
                    "pulse": driver_df.iloc[0]["脉宽"],
                    
                    # TEC表 (需要取第0行和第1行)
                    "tec1_set": tec_df.iloc[0]["设定值"], "tec1_read": tec_df.iloc[0]["回读值"], "tec1_peltier": tec_df.iloc[0]["电流"],
                    "tec2_set": tec_df.iloc[1]["设定值"], "tec2_read": tec_df.iloc[1]["回读值"], "tec2_peltier": tec_df.iloc[1]["电流"],
                    
                    # 文本域
                    "problem": problem,
                    "action": action_sum,
                    "note": note,
                    
                    # 动态表格 (转字典)
                    "power_table": power_df.to_dict('records'),
                    "output_table": output_df.to_dict('records'),
                    "action_table": action_df.to_dict('records')
                }
                
                # 保存
                st.session_state['db'].append(record)
                st.success(f"✅ 序列号 {sn_val} 保存成功！")
                
                # 重置所有数据
                reset_all_data()
                st.rerun()
                
        except Exception as e:
            st.error(f"数据提取错误: {e}")

# --- TAB 2: 历史记录 ---
with tab2:
    st.header("🗄️ 维修档案库")
    search_term = st.text_input("🔍 搜索序列号:")
    
    display_data = st.session_state['db']
    if search_term:
        display_data = [d for d in display_data if search_term.lower() in d['sn'].lower()]

    if not display_data:
        st.info("暂无数据。")
    else:
        for i, record in enumerate(reversed(display_data)):
            with st.expander(f"📅 {record['date']} | SN: {record['sn']} | {record['operator']}"):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"**故障:** {record['problem']}")
                    st.write(f"**措施:** {record['action']}")
                with col2:
                    doc_file = generate_doc(record)
                    if doc_file:
                        st.download_button("📥 下载 Word", doc_file, f"Report_{record['sn']}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_{record['id']}")
                    
                    if st.session_state['is_admin']:
                        if st.button("🗑️ 删除", key=f"del_{record['id']}"):
                            st.session_state['db'] = [d for d in st.session_state['db'] if d['id'] != record['id']]
                            st.rerun()
