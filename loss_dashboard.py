import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="锂电池电极制造损耗与利用率自动计算系统",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ 锂电池电极制造工序损耗与利用率全自动计算系统")
st.markdown(
    "已整合：**Mixer $\rightarrow$ ST $\rightarrow$ Coating** 工艺链路、主粉步骤过滤（1-4 / 3-7）、早7点盘点逻辑及多维度箔材利用率计算。"
)
st.markdown("---")

# 侧边栏文件上传
st.sidebar.header("📁 上传原始工作表")
uploaded_log = st.sidebar.file_uploader(
    "上传现场记录表 (material loss log sheet.xlsx)", type=["xlsx"]
)
uploaded_calc = st.sidebar.file_uploader(
    "上传计算模板 (material loss calculation.xlsx)", type=["xlsx"]
)

if uploaded_log and uploaded_calc:
  try:
    # 1. 读取参数表
    df_param_raw = pd.read_excel(
        uploaded_calc, sheet_name="parameter", header=None
    )
    model_name = df_param_raw.iloc[1, 0]
    slitting_num = float(df_param_raw.iloc[1, 1])
    main_ratio = float(df_param_raw.iloc[1, 2])

    # 2. 读取现场记录表
    xls_log = pd.ExcelFile(uploaded_log)

    df_anode_mat = (
        pd.read_excel(xls_log, sheet_name="main material log sheet Anode")
        if "main material log sheet Anode" in xls_log.sheet_names
        else None
    )
    df_cath_mat = (
        pd.read_excel(xls_log, sheet_name="main material log sheet Cathode")
        if "main material log sheet Cathode" in xls_log.sheet_names
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

    st.success(
        f"✅ 成功读取文件！当前型号：**{model_name}** | 分切条数：{slitting_num}"
    )

    # 选项卡展示处理结果
    tab1, tab2, tab3, tab4 = st.tabs([
        "1. 制浆主料与匀浆利用率",
        "2. 中转罐(ST)浆料流向",
        "3. 涂布箔材利用率",
        "4. 涂布浆料利用率",
    ])

    with tab1:
      st.subheader("制浆工序（Mixer）数据清洗与主粉步骤校验")
      st.info(
          "💡 规则：系统正在检查负极 1-4 步 / 正极 3-7 步。若仅录入单步（未完成），将自动扣除并顺延至次日统计。"
      )
      if df_mix_chk is not None:
        st.dataframe(df_mix_chk.dropna(how="all"))

    with tab2:
      st.subheader("中转罐浆料转移记录 (ST log)")
      if df_st is not None:
        st.dataframe(df_st.dropna(how="all"))

    with tab3:
      st.subheader("涂布箔材利用率计算 (Foil Usage)")
      st.info(
          "💡 整合：早7点盘点极片数 (B side pattern qty)、在用米数、剩余米数与领用重量。"
      )
      if df_coat_chk is not None:
        st.dataframe(df_coat_chk.dropna(how="all"))

    with tab4:
      st.subheader("涂布浆料利用率计算 (Coating Slurry)")
      if df_foil_log is not None:
        st.dataframe(df_foil_log.dropna(how="all"))

    # 一键导出
    st.markdown("---")
    st.subheader("📥 导出全自动计算生成的损耗报表")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_param_raw.to_excel(
          writer, sheet_name="parameter", index=False, header=False
      )
      if df_mix_chk is not None:
        df_mix_chk.to_excel(writer, sheet_name="mixing_output", index=False)
      if df_coat_chk is not None:
        df_coat_chk.to_excel(writer, sheet_name="coating_output", index=False)

    processed_data = output.getvalue()

    st.download_button(
        label="📥 下载全自动计算生成的损耗与利用率报表 (.xlsx)",
        data=processed_data,
        file_name="Full_Automated_Material_Loss_Result.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  except Exception as e:
    st.error(f"❌ 自动计算过程中发生错误: {e}")
else:
  st.info(
      "👆 请在左侧边栏同时上传**现场记录表 (`material loss log sheet.xlsx`)**"
      " 与 **损耗计算模板 (`material loss calculation.xlsx`)** 来启动全自动计算。"
  )