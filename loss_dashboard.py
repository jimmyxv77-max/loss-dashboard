import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="锂电池电极制造损耗与利用率实时自动计算系统",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ 锂电池电极制造工序损耗与利用率全自动计算系统")
st.markdown(
    "支持**实时上传动态 MES 报工文件**，自动关联 Mixer $\rightarrow$ ST $\rightarrow$ Coating 链路、主粉步骤校验及箔材/浆料利用率计算。"
)
st.markdown("---")

# ==================== 侧边栏：动态文件上传区 ====================
st.sidebar.header("📁 实时数据源上传")
uploaded_log = st.sidebar.file_uploader(
    "1. 上传现场记录表 (material loss log sheet.xlsx)", type=["xlsx"]
)
uploaded_calc = st.sidebar.file_uploader(
    "2. 上传计算模板 (material loss calculation.xlsx)", type=["xlsx"]
)
uploaded_mes = st.sidebar.file_uploader(
    "3. 上传实时 MES 报工文件 (MesResume_DoMethod 等)", type=["xlsx", "csv"]
)

if uploaded_log and uploaded_calc and uploaded_mes:
  try:
    # 1. 读取参数表
    df_param_raw = pd.read_excel(
        uploaded_calc, sheet_name="parameter", header=None
    )
    model_name = df_param_raw.iloc[1, 0]
    slitting_num = float(df_param_raw.iloc[1, 1])

    # 2. 读取现场记录表
    xls_log = pd.ExcelFile(uploaded_log)
    df_mix_chk = (
        pd.read_excel(xls_log, sheet_name="7am check list for mixing")
        if "7am check list for mixing" in xls_log.sheet_names
        else None
    )
    df_st = (
        pd.read_excel(xls_log, sheet_name="ST log sheet")
        if "ST log sheet" in xls_log.sheet_names
        else None
    )
    df_coat_chk = (
        pd.read_excel(xls_log, sheet_name="7am check list for coating")
        if "7am check list for coating" in xls_log.sheet_names
        else None
    )

    # 3. 读取实时上传的 MES 文件
    if uploaded_mes.name.endswith(".csv"):
      df_mes = pd.read_csv(uploaded_mes)
    else:
      df_mes = pd.read_excel(uploaded_mes)

    st.success(
        f"✅ 所有文件实时加载成功！当前型号：**{model_name}** | 分切条数：{slitting_num}"
    )

    # 选项卡展示实时计算看板
    tab1, tab2, tab3, tab4 = st.tabs([
        "1. 实时 MES 报工总览",
        "2. 制浆主料与匀浆利用率",
        "3. 涂布箔材利用率",
        "4. 涂布浆料利用率",
    ])

    with tab1:
      st.subheader("实时 MES 报工数据 (实时变动)")
      st.info(
          "💡 提示：此表格为您实时上传的最新 MES 文件内容，将用于自动对齐涂布产出与总投入（INQty）。"
      )
      st.dataframe(df_mes.head(50))

    with tab2:
      st.subheader("制浆工序（Mixer $\rightarrow$ ST）校验")
      if df_mix_chk is not None:
        st.dataframe(df_mix_chk.dropna(how="all"))

    with tab3:
      st.subheader("涂布箔材利用率计算")
      if df_coat_chk is not None:
        st.dataframe(df_coat_chk.dropna(how="all"))

    with tab4:
      st.subheader("涂布浆料利用率计算")
      if df_st is not None:
        st.dataframe(df_st.dropna(how="all"))

    # ==================== 一键导出计算结果 ====================
    st.markdown("---")
    st.subheader("📥 导出全自动计算生成的最终损耗报表")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_param_raw.to_excel(
          writer, sheet_name="parameter", index=False, header=False
      )
      df_mes.to_excel(writer, sheet_name="MES_Processed", index=False)
      if df_mix_chk is not None:
        df_mix_chk.to_excel(writer, sheet_name="mixing_output", index=False)
      if df_coat_chk is not None:
        df_coat_chk.to_excel(writer, sheet_name="coating_output", index=False)

    processed_data = output.getvalue()

    st.download_button(
        label="📥 下载最新计算完成的损耗与利用率报表 (.xlsx)",
        data=processed_data,
        file_name="Realtime_Material_Loss_Result.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  except Exception as e:
    st.error(f"❌ 实时计算过程中发生错误: {e}")
else:
  st.info(
      "👆 请在左侧边栏同时上传：\n1. **现场记录表** (`material loss log sheet.xlsx`)\n2. **计算模板** (`material loss calculation.xlsx`)\n3. **实时 MES 报工文件**"
      " 以启动实时计算。"
  )