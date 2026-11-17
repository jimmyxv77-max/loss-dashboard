import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="锂电池电极制造损耗与利用率自动计算系统",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ 锂电池电极制造工序损耗与利用率自动生成系统")
st.markdown("---")

# 文件上传区
st.sidebar.header("📁 数据文件上传")
uploaded_log = st.sidebar.file_uploader(
    "上传现场记录表 (material loss log sheet.xlsx)", type=["xlsx"]
)
uploaded_calc = st.sidebar.file_uploader(
    "上传损耗计算标准模板 (material loss calculation.xlsx)", type=["xlsx"]
)

if uploaded_log and uploaded_calc:
  try:
    # 1. 读取参数表
    df_param = pd.read_excel(uploaded_calc, sheet_name="parameter")

    # 2. 读取现场记录表各Sheet
    xls_log = pd.ExcelFile(uploaded_log)

    df_main_anode = (
        pd.read_excel(xls_log, sheet_name="main material log sheet Anode")
        if "main material log sheet Anode" in xls_log.sheet_names
        else None
    )
    df_st = (
        pd.read_excel(xls_log, sheet_name="ST log sheet")
        if "ST log sheet" in xls_log.sheet_names
        else None
    )
    df_foil_log = (
        pd.read_excel(xls_log, sheet_name="Foil log sheet")
        if "Foil log sheet" in xls_log.sheet_names
        else None
    )
    df_mix_chk = (
        pd.read_excel(xls_log, sheet_name="7am check list for mixing")
        if "7am check list for mixing" in xls_log.sheet_names
        else None
    )
    df_coat_chk = (
        pd.read_excel(xls_log, sheet_name="7am check list for coating")
        if "7am check list for coating" in xls_log.sheet_names
        else None
    )

    st.success("✅ 文件读取成功！")

    # 展示基础参数
    with st.expander("查看工艺参数表 (Parameter)", expanded=False):
      st.dataframe(df_param)

    # 核心计算逻辑预览选项卡
    tab1, tab2, tab3 = st.tabs([
        "匀浆浆料利用率 (Mixing Slurry)",
        "涂布箔材利用率 (Foil Usage)",
        "涂布浆料利用率 (Coating Slurry)",
    ])

    with tab1:
      st.subheader("制浆/匀浆利用率计算")
      st.info(
          "💡 规则说明：每一行的‘前一日交接班数据’自动继承上一行的‘当日交接班数据’。"
      )
      if df_mix_chk is not None:
        st.dataframe(df_mix_chk.head(15))
      else:
        st.warning("未找到 7am check list for mixing 表格")

    with tab2:
      st.subheader("涂布箔材利用率计算")
      if df_foil_log is not None:
        st.dataframe(df_foil_log.head(15))
      else:
        st.warning("未找到 Foil log sheet 表格")

    with tab3:
      st.subheader("涂布浆料利用率计算")
      if df_st is not None:
        st.dataframe(df_st.head(15))
      else:
        st.warning("未找到 ST log sheet 表格")

    # 导出按钮
    st.markdown("---")
    st.subheader("📥 导出完整损耗计算报表")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_param.to_excel(
          writer, sheet_name="parameter_checked", index=False
      )
      if df_mix_chk is not None:
        df_mix_chk.to_excel(writer, sheet_name="mixing_processed", index=False)
    processed_data = output.getvalue()

    st.download_button(
        label="下载计算完成的完整损耗报表 (.xlsx)",
        data=processed_data,
        file_name="Material_Loss_Calculation_Result.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  except Exception as e:
    st.error(f"❌ 计算或读取过程中发生错误: {e}")
else:
  st.info("👆 请在左侧边栏同时上传**现场记录表**与**损耗计算模板**以启动自动化程序。")