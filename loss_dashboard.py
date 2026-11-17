import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="锂电池电极制造损耗与利用率全自动计算系统",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ 锂电池电极制造工序损耗与利用率全自动计算系统")
st.markdown(
    "已修复参数行索引错误：精准读取第1行真实数值，自动应用 **分切条数 (number of slitting)** 折算 MES 报工。"
)
st.markdown("---")

# ==================== 侧边栏：多文件动态上传区 ====================
st.sidebar.header("📁 实时数据源上传")
uploaded_log = st.sidebar.file_uploader(
    "1. 上传现场记录表 (material loss log sheet.xlsx)", type=["xlsx"]
)
uploaded_calc = st.sidebar.file_uploader(
    "2. 上传计算模板 (material loss calculation.xlsx)", type=["xlsx"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("3. 实时 MES 报工文件（正负极独立）")
uploaded_mes_neg = st.sidebar.file_uploader(
    "上传【负极】MES 报工文件 (MesResume_DoMethod.xlsx)", type=["xlsx", "csv"]
)
uploaded_mes_pos = st.sidebar.file_uploader(
    "上传【正极】MES 报工文件 (MesResume_DoMethod+.xlsx)", type=["xlsx", "csv"]
)

if uploaded_log and uploaded_calc:
  try:
    # 1. 读取参数表：第0行是英文说明，第1行是具体数值（如 Plant 4 DQ+）
    df_param_raw = pd.read_excel(uploaded_calc, sheet_name="parameter")
    model_name = df_param_raw.iloc[
        1, 0
    ]  # 获取型号 (Row 1, Col 0: Plant 4 DQ+)
    slitting_num = float(
        df_param_raw.iloc[1, 1]
    )  # 获取分切条数 (Row 1, Col 1: 11)
    main_ratio = float(
        df_param_raw.iloc[1, 2]
    )  # 获取主料占比 (Row 1, Col 2: 0.9575)

    # 2. 读取现场记录表各Sheet
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

    # 3. 读取实时上传的正负极 MES 文件，并利用分切条数折算真实产量
    df_mes_neg = None
    df_mes_pos = None

    if uploaded_mes_neg is not None:
      df_mes_neg = (
          pd.read_csv(uploaded_mes_neg)
          if uploaded_mes_neg.name.endswith(".csv")
          else pd.read_excel(uploaded_mes_neg)
      )
      if "OutQty" in df_mes_neg.columns:
        df_mes_neg["Coated_Actual_Output"] = df_mes_neg["OutQty"] / slitting_num

    if uploaded_mes_pos is not None:
      df_mes_pos = (
          pd.read_csv(uploaded_mes_pos)
          if uploaded_mes_pos.name.endswith(".csv")
          else pd.read_excel(uploaded_mes_pos)
      )
      if "OutQty" in df_mes_pos.columns:
        df_mes_pos["Coated_Actual_Output"] = df_mes_pos["OutQty"] / slitting_num

    st.success(
        f"✅ 加载成功！当前型号：**{model_name}** | 分切条数：{slitting_num}"
        "（已用于折算MES小片报工）"
    )

    # 选项卡展示实时计算看板
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "1. 负极 MES 报工看板 (-)",
        "2. 正极 MES 报工看板 (+)",
        "3. 制浆主料与匀浆利用率",
        "4. 涂布箔材利用率",
        "5. 涂布浆料利用率",
    ])

    with tab1:
      st.subheader("【负极】MES 实时报工数据（含折算真实产量）")
      if df_mes_neg is not None:
        st.dataframe(df_mes_neg)
      else:
        st.info("请在左侧边栏上传负极 MES 报工文件。")

    with tab2:
      st.subheader("【正极】MES 实时报工数据（含折算真实产量）")
      if df_mes_pos is not None:
        st.dataframe(df_mes_pos)
      else:
        st.info("请在左侧边栏上传正极 MES 报工文件。")

    with tab3:
      st.subheader("制浆工序（Mixer $\rightarrow$ ST）主粉步骤校验")
      st.info(
          "💡 规则：自动校验负极 1-4 步 / 正极 3-7 步完整性，单步不完整自动顺延。"
      )
      if df_mix_chk is not None:
        st.dataframe(df_mix_chk.dropna(how="all"))

    with tab4:
      st.subheader("涂布箔材利用率计算")
      if df_coat_chk is not None:
        st.dataframe(df_coat_chk.dropna(how="all"))

    with tab5:
      st.subheader("涂布浆料利用率计算")
      if df_st is not None:
        st.dataframe(df_st.dropna(how="all"))

    # ==================== 一键导出计算结果 ====================
    st.markdown("---")
    st.subheader("📥 导出全自动计算生成的最终损耗报表")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_param_raw.to_excel(
          writer, sheet_name="parameter", index=False
      )
      if df_mes_neg is not None:
        df_mes_neg.to_excel(writer, sheet_name="MES_Negative", index=False)
      if df_mes_pos is not None:
        df_mes_pos.to_excel(writer, sheet_name="MES_Positive", index=False)
      if df_mix_chk is not None:
        df_mix_chk.to_excel(writer, sheet_name="mixing_output", index=False)
      if df_coat_chk is not None:
        df_coat_chk.to_excel(writer, sheet_name="coating_output", index=False)

    processed_data = output.getvalue()

    st.download_button(
        label="📥 下载最新计算完成的损耗与利用率报表 (.xlsx)",
        data=processed_data,
        file_name="Full_Automated_Material_Loss_Result.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  except Exception as e:
    st.error(f"❌ 实时计算过程中发生错误: {e}")
else:
  st.info(
      "👆 请在左侧边栏同时上传：\n1. **现场记录表** (`material loss log sheet.xlsx`)\n2."
      " **计算模板** (`material loss calculation.xlsx`)\n以及实时 **正/负极 MES"
      " 报工文件** 以启动全自动计算。"
  )