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

# ==========================================
# 2. 核心逻辑：数据处理与文档生成
# ==========================================

def flatten_data_for_template(record):
    """
    将复杂的数据结构拍平，适配 Word 模板的 {{ tag_1 }} 格式
    """
    # 1. 复制基础字段 (sn, model, action, problem 等)
    context = record.copy()
    
    # 2. 处理功率测量表 (Power Table)
    # 对应模板: {{ current_1 }}, {{ pulse_1 }}, {{ nm_1 }}, {{ power_1 }}
    power_data = record.get('power_table', [])
    for i, row in enumerate(power_data):
        suffix = f"_{i+1}"
        # 注意：这里要用 .get() 防止表格里有空值导致报错
        context[f"current{suffix}"] = row.get("电流 I [A]", "")
        context[f"pulse{suffix}"] = row.get("脉宽 [us]", "")
        context[f"nm{suffix}"] = row.get("波长 λ", "")
        context[f"power{suffix}"] = row.get("功率 P [W]", "")
    
    # 3. 处理输出功率表 (Output Table)
    # 对应模板: {{ power_355_1 }} ...
    output_data = record.get('output_table', [])
    for i, row in enumerate(output_data):
        suffix = f"_{i+1}"
        context[f"power_355{suffix}"] = row.get("355nm", "")
        context[f"power_532{suffix}"] = row.get("532nm", "")
        context[f"power_1064{suffix}"] = row.get("1064nm", "")

    # 4. 处理维修步骤表 (Action Table)
    # 对应模板: {{ action_1 }}, {{ operator_1 }} ...
    action_data = record.get('action_table', [])
    for i, row in enumerate(action_data):
        suffix = f"_{i+1}"
        # 这里使用了 action_1，绝对不会和外面的 action (总体描述) 冲突
        context[f"action{suffix}"] = row.get("维修措施", "")
        context[f"operator{suffix}"] = row.get("操作员", "")
        context[f"date{suffix}"] = row.get("日期", "")
        
    return context

def generate_doc(record):
    if not os.path.exists("template.docx"):
        return None
    
    doc = DocxTemplate("template.docx")
    
    # 数据转换
    final_context = flatten_data_for_template(record)
    
    try:
        doc.render(final_context)
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    except Exception as e:
        # 这里记录错误但不中断程序
        print(f"Word生成错误: {e}")
        return None

# ==========================================
# 3. 侧边栏：管理员登录
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
# 4. 主界面
# ==========================================
st.title("🔋 激光器维修档案系统")

tab1, tab2 = st.tabs(["📝 录入新记录", "🔍 历史档案库"])

# --- TAB 1: 录入界面 ---
with tab1:
    # 【关键】使用 st.form 解决“填一个数刷新一下”的问题
    # clear_on_submit=True 解决“保存后需要手动清空”的问题
    with st.form("main_form", clear_on_submit=True):
        st.info("💡 提示：在表格中按 Enter 是确认输入，不会提交表单。只有点击最底部的“保存”按钮才会提交并清空。")
        
        # 1. 基础信息
        st.subheader("1. 基础信息")
        c1, c2, c3, c4 = st.columns(4)
        sn = c1.text_input("序列号 (Serial No.)")
        model = c2.text_input("型号 (Model)", value="WYP-")
        voltage = c3.text_input("电压 (Voltage)", value="24V")
        operator = c4.text_input("当前操作员", value="Guest")
        
        # 2. 外观
        st.subheader("2. 外观检查")
        c1, c2 = st.columns(2)
        obs_case = c1.text_input("外壳/包装状态", value="完好 Normal")
        obs_mech = c2.text_input("机械损伤", value="无 None")

        # 3. 电子与TEC
        with st.expander("3. 电子参数与 TEC 设置", expanded=False):
            e1, e2 = st.columns(2)
            work_hours = e1.text_input("工作时长")
            alarms = e2.text_input("报警状态", value="No Alarm")
            
            st.markdown("**TEC 1 设置**")
            t1_1, t1_2, t1_3 = st.columns(3)
            tec1_set = t1_1.text_input("TEC1 设定值")
            tec1_read = t1_2.text_input("TEC1 回读值")
            tec1_peltier = t1_3.text_input("TEC1 电流")

            st.markdown("**TEC 2 设置**")
            t2_1, t2_2, t2_3 = st.columns(3)
            tec2_set = t2_1.text_input("TEC2 设定值")
            tec2_read = t2_2.text_input("TEC2 回读值")
            tec2_peltier = t2_3.text_input("TEC2 电流")
            
            st.markdown("**驱动参数**")
            h1, h2, h3 = st.columns(3)
            hv = h1.text_input("高压 (HV)")
            current = h2.text_input("峰值电流 (I Peak)")
            pulse = h3.text_input("脉宽 (Pulse)")

        # 4. 动态表格 (功率)
        st.subheader("4. 功率测量数据")
        st.caption("👇 在下方表格直接编辑，支持多行。")
        
        # 定义初始数据结构
        # num_rows="dynamic" 允许用户自由添加行
        default_power = pd.DataFrame([{"电流 I [A]": "", "脉宽 [us]": "", "波长 λ": "", "功率 P [W]": ""}])
        edited_power_df = st.data_editor(default_power, num_rows="dynamic", use_container_width=True, key="power_editor")

        st.markdown("**输出功率 (Output Power)**")
        default_output = pd.DataFrame([{"355nm": "", "532nm": "", "1064nm": ""}])
        edited_output_df = st.data_editor(default_output, num_rows="dynamic", use_container_width=True, key="output_editor")

        # 5. 故障与维修
        st.subheader("5. 故障分析与维修日志")
        problem = st.text_area("故障描述", height=80)
        action_summary = st.text_area("采取措施总体描述 (对应模板 {{ action }})", height=80)
        
        st.markdown("**详细维修步骤记录 (对应模板 {{ action_1 }} 等)**")
        default_action = pd.DataFrame([{"维修措施": "", "操作员": operator, "日期": datetime.now().strftime("%Y-%m-%d")}])
        edited_action_df = st.data_editor(default_action, num_rows="dynamic", use_container_width=True, key="action_editor")
        
        note = st.text_area("备注 (Notes)")

        st.markdown("---")
        # 提交按钮
        submitted = st.form_submit_button("💾 保存完整记录", type="primary")

        # ================== 保存逻辑 ==================
        if submitted:
            if not sn:
                st.error("❌ 保存失败：序列号不能为空！")
            else:
                # 1. 提取表格数据 (这里直接用变量，不再去 session_state 找 key，避免报错)
                power_records = edited_power_df.to_dict('records')
                output_records = edited_output_df.to_dict('records')
                action_records = edited_action_df.to_dict('records')

                # 2. 构建记录字典
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
                    "action": action_summary, # 总体描述
                    "note": note,
                    # 动态表格数据
                    "power_table": power_records,
                    "output_table": output_records,
                    "action_table": action_records
                }
                
                # 3. 存入数据库
                st.session_state['db'].append(new_record)
                st.success(f"✅ 序列号 {sn} 已保存！(表单已自动清空)")
                
                # 4. 这里的 clear_on_submit=True 会在下次刷新时自动清空所有框
                # 不需要额外写 clear() 代码

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
            # 倒序显示，最新的在最上面
            with st.expander(f"📅 {record['date']} | SN: {record['sn']} | 操作员: {record['operator']}"):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.markdown(f"**故障:** {record['problem']}")
                    st.markdown(f"**措施(总体):** {record['action']}")
                with col2:
                    # 下载 Word
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
                        st.warning("⚠️ 缺少模板文件")
                    
                    # 删除按钮 (仅管理员)
                    if st.session_state['is_admin']:
                        if st.button("🗑️ 删除记录", key=f"del_{record['id']}"):
                            st.session_state['db'] = [d for d in st.session_state['db'] if d['id'] != record['id']]
                            st.rerun()
