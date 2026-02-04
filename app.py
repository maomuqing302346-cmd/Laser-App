import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
from datetime import datetime
import os

# ==========================================
# 1. 页面配置与 CSS 美化
# ==========================================
st.set_page_config(page_title="激光器维修系统 (最终版)", page_icon="🔋", layout="wide")

# 初始化数据库
if 'db' not in st.session_state:
    # 定义基础字段
    cols = ["id", "sn", "model", "voltage", "operator", "date", 
            "obs_case", "obs_mech", "work_hours", "alarms",
            "tec1_set", "tec1_read", "tec1_peltier",
            "tec2_set", "tec2_read", "tec2_peltier",
            "hv", "current", "pulse", 
            "problem", "action", "note"]
    # 动态表格的数据将以 JSON 或 字符串形式存储，或者在生成 Word 时动态解析
    # 这里为了简单，我们存储整个记录字典
    st.session_state['db'] = [] # 使用列表存储字典，比DataFrame更灵活处理嵌套结构

# 初始化管理员状态
if 'is_admin' not in st.session_state:
    st.session_state['is_admin'] = False

# ==========================================
# 2. 核心逻辑函数
# ==========================================

def flatten_data_for_template(record):
    """
    将动态表格的数据（列表格式）拍平，适配 Word 模板的 {{ tag_1 }}, {{ tag_2 }} 格式
    """
    context = record.copy()
    
    # 1. 处理功率测量表 (Power Table)
    # 假设模板里是 current_1, current_2 ... 
    power_data = record.get('power_table', [])
    for i, row in enumerate(power_data):
        suffix = f"_{i+1}" # 生成 _1, _2, _3
        context[f"current{suffix}"] = row.get("电流 I [A]", "")
        context[f"pulse{suffix}"] = row.get("脉宽 [us]", "")
        context[f"nm{suffix}"] = row.get("波长 λ", "")
        context[f"power{suffix}"] = row.get("功率 P [W]", "")
    
    # 2. 处理输出功率表 (Output Table) - 假设只有一行，直接取值
    # 如果您希望输出功率也是多行的，逻辑同上。这里假设是单行多列结构。
    output_data = record.get('output_table', [])
    if output_data:
        first_row = output_data[0]
        context["power_355"] = first_row.get("355nm", "")
        context["power_532"] = first_row.get("532nm", "")
        context["power_1064"] = first_row.get("1064nm", "")

    # 3. 处理维修步骤表 (Action Table)
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
    
    # 数据预处理：把动态表格转成模板能认的 _1, _2 格式
    final_context = flatten_data_for_template(record)
    
    # 渲染
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
# 3. 侧边栏：管理员登录
# ==========================================
with st.sidebar:
    st.header("🔧 系统菜单")
    
    # 权限开关
    with st.expander("👮‍♂️ 管理员登录 (仅用于删除/编辑)"):
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

tab1, tab2 = st.tabs(["📝 录入新记录 (所有人员)", "🔍 历史档案库 (仅管理可删)"])

# --- TAB 1: 录入界面 ---
with tab1:
    with st.form("entry_form", clear_on_submit=True):
        st.info("💡 提示：所有内容填写完毕后，请点击底部的“保存完整记录”按钮提交。表格支持点击添加多行。")
        
        # Section 1: 基础信息
        st.subheader("1. 基础信息")
        c1, c2, c3, c4 = st.columns(4)
        sn = c1.text_input("序列号 (Serial No.)")
        model = c2.text_input("型号 (Model)", value="WYP-")
        voltage = c3.text_input("电压 (Voltage)", value="24V")
        operator = c4.text_input("当前操作员", value="Guest")
        
        # Section 2: 外观
        st.subheader("2. 外观检查")
        c1, c2 = st.columns(2)
        obs_case = c1.text_input("外壳/包装状态", value="完好 Normal")
        obs_mech = c2.text_input("机械损伤", value="无 None")

        # Section 3: 电子与TEC
        with st.expander("3. 电子参数与 TEC 设置 (点击展开)", expanded=False):
            e1, e2 = st.columns(2)
            work_hours = e1.text_input("工作时长 (Hours)")
            alarms = e2.text_input("报警状态 (Alarms)", value="No Alarm")
            
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

        # Section 4: 动态表格 (功率)
        st.subheader("4. 功率测量数据 (支持多行)")
        st.caption("👇 直接在表格中修改，点击表格下方的 + 号添加新行")
        
        # 定义初始表格结构
        default_power_df = pd.DataFrame([
            {"电流 I [A]": "", "脉宽 [us]": "", "波长 λ": "", "功率 P [W]": ""}
        ])
        # 使用 data_editor 实现动态增删
        edited_power_df = st.data_editor(default_power_df, num_rows="dynamic", use_container_width=True, key="editor_power")

        st.markdown("**输出功率 (Output Power)**")
        default_output_df = pd.DataFrame([{"355nm": "", "532nm": "", "1064nm": ""}])
        edited_output_df = st.data_editor(default_output_df, num_rows="fixed", use_container_width=True, key="editor_output")

        # Section 5: 故障与动态维修记录
        st.subheader("5. 故障分析与维修日志")
        problem = st.text_area("故障描述 (Description)", height=80)
        action_summary = st.text_area("采取措施总体描述 (Action Taken)", height=80)
        
        st.markdown("**详细维修步骤记录 (Repair Actions Table)**")
        default_action_df = pd.DataFrame([
            {"维修措施": "", "操作员": operator, "日期": datetime.now().strftime("%Y-%m-%d")}
        ])
        edited_action_df = st.data_editor(default_action_df, num_rows="dynamic", use_container_width=True, key="editor_action")
        
        note = st.text_area("备注 (Notes)")

        # 保存按钮
        submitted = st.form_submit_button("💾 保存完整记录", type="primary")
        
        if submitted:
            if not sn:
                st.error("❌ 保存失败：序列号不能为空！")
            else:
                # 收集所有数据打包成字典
                new_record = {
                    "id": len(st.session_state['db']) + 1,
                    "sn": sn, "model": model, "voltage": voltage, "operator": operator,
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "obs_case": obs_case, "obs_mech": obs_mech,
                    "work_hours": work_hours, "alarms": alarms,
                    "tec1_set": tec1_set, "tec1_read": tec1_read, "tec1_peltier": tec1_peltier,
                    "tec2_set": tec2_set, "tec2_read": tec2_read, "tec2_peltier": tec2_peltier,
                    "hv": hv, "current": current, "pulse": pulse,
                    "problem": problem, "action": action_summary, "note": note,
                    # 将 DataFrame 转为字典列表存储
                    "power_table": edited_power_df.to_dict('records'),
                    "output_table": edited_output_df.to_dict('records'),
                    "action_table": edited_action_df.to_dict('records')
                }
                
                # 保存到 Session State (模拟数据库)
                st.session_state['db'].append(new_record)
                st.success(f"✅ 序列号 {sn} 的记录已成功保存！")

# --- TAB 2: 查询界面 ---
with tab2:
    st.header("🗄️ 维修档案库")
    
    # 搜索功能
    search_term = st.text_input("🔍 输入序列号搜索：")
    
    # 过滤数据
    display_data = st.session_state['db']
    if search_term:
        display_data = [d for d in display_data if search_term.lower() in d['sn'].lower()]

    if not display_data:
        st.info("暂无数据。请在“录入新记录”页面添加。")
    else:
        # 倒序显示，最新的在前面
        for i, record in enumerate(reversed(display_data)):
            with st.expander(f"📅 {record['date']} | SN: {record['sn']} | 操作员: {record['operator']}"):
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    st.markdown(f"**故障:** {record['problem']}")
                    st.markdown(f"**措施:** {record['action']}")
                    st.caption("表格数据包含在导出的 Word 中")

                with col2:
                    # 下载按钮 (所有人可见)
                    doc_file = generate_doc(record)
                    if doc_file:
                        st.download_button(
                            label="📥 下载 Word",
                            data=doc_file,
                            file_name=f"Report_{record['sn']}_{record['date']}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_{i}"
                        )
                    else:
                        st.warning("缺少模板文件")
                    
                    # 删除按钮 (仅管理员可见)
                    if st.session_state['is_admin']:
                        if st.button("🗑️ 删除记录", key=f"del_{i}"):
                            # 从原始列表中移除
                            # 注意：这里需要根据 id 或内容去原始 db 列表中找，因为 display_data 是过滤过的
                            st.session_state['db'] = [d for d in st.session_state['db'] if d['id'] != record['id']]
                            st.rerun()
