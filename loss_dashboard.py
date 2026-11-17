import io
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

# 设置页面支持中文及宽屏
st.set_page_config(
    page_title="锂电池电极制造损耗与利用率可视化系统", page_icon="⚡", layout="wide"
)

# 解决matplotlib中文乱码问题
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

st.title("⚡ 锂电池电极制造工序损耗与利用率自动化分析看板")
st.markdown(
    "已升级：完美兼容 `(-)` / `(+)` 与 `Anode` / `Cathode` 双重命名后缀，实现全自动跨表关联与可视化分析。"
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
    # 1. 读取参数表
    df_param_raw = pd.read_excel(uploaded_calc, sheet_name="parameter")
    model_name = df_param_raw.iloc[1, 0]
    slitting_num = float(df_param_raw.iloc[1, 1])  # 分切条数
    main_ratio = float(df_param_raw.iloc[1, 2])

    # 2. 智能读取现场记录表（兼容 (-)/(+) 与 Anode/Cathode）
    xls_log = pd.ExcelFile(uploaded_log)
    sheet_names = xls_log.sheet_names


    def get_sheet_safely(negative_keys, positive_keys, is_negative=True):
      """智能匹配：支持 (-)/Anode 或 (+)/Cathode"""
      keys = negative_keys if is_negative else positive_keys
      for name in sheet_names:
        name_lower = name.lower()
        if any(k.lower() in name_lower for k in keys):
          return pd.read_excel(xls_log, sheet_name=name)
      return None


    # 获取负极和正极的各类日志
    df_mix_neg = get_sheet_safely(
        ["mixing", "-", "anode"], ["mixing", "+", "cathode"], is_negative=True
    )
    df_mix_pos = get_sheet_safely(
        ["mixing", "-", "anode"], ["mixing", "+", "cathode"], is_negative=False
    )
    df_coat_neg = get_sheet_safely(
        ["coating", "-", "anode"], ["coating", "+", "cathode"], is_negative=True
    )
    df_coat_pos = get_sheet_safely(
        ["coating", "-", "anode"], ["coating", "+", "cathode"], is_negative=False
    )
    df_st_neg = get_sheet_safely(
        ["st", "-", "anode"], ["st", "+", "cathode"], is_negative=True
    )
    df_st_pos = get_sheet_safely(
        ["st", "-", "anode"], ["st", "+", "cathode"], is_negative=False
    )

    # 3. 读取实时上传的正负极 MES 文件并折算产量
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
        f"✅ 智能匹配成功！当前型号：**{model_name}** | 分切条数：{slitting_num}"
    )

    # ==================== 选项卡看板展示 ====================
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 生产总览与可视化分析",
        "1. 负极 MES 报工 (-/Anode)",
        "2. 正极 MES 报工 (+/Cathode)",
        "3. 制浆与中转校验",
        "4. 涂布与箔材盘点",
    ])

    with tab1:
      st.subheader("📈 涂布产出与损耗多维可视化分析")
      col1, col2 = st.columns(2)
      with col1:
        st.markdown("#### 负极产出与报废量趋势 (-)")
        if df_mes_neg is not None and "OutQty" in df_mes_neg.columns:
          fig, ax = plt.subplots(figsize=(6, 4))
          ax.plot(
              df_mes_neg.index,
              df_mes_neg["OutQty"],
              marker="o",
              label="实际产出 (OutQty)",
              color="blue",
          )
          if "ScrapQty" in df_mes_neg.columns:
            ax.plot(
                df_mes_neg.index,
                df_mes_neg["ScrapQty"],
                marker="x",
                label="报废量 (ScrapQty)",
                color="red",
            )
          ax.set_title("Negative Electrode Output")
          ax.legend()
          st.pyplot(fig)
        else:
          st.info("请上传负极 MES 文件以查看图表")

      with col2:
        st.markdown("#### 正极产出与报废量趋势 (+)")
        if df_mes_pos is not None and "OutQty" in df_mes_pos.columns:
          fig, ax = plt.subplots(figsize=(6, 4))
          ax.plot(
              df_mes_pos.index,
              df_mes_pos["OutQty"],
              marker="o",
              label="实际产出 (OutQty)",
              color="green",
          )
          if "ScrapQty" in df_mes_pos.columns:
            ax.plot(
                df_mes_pos.index,
                df_mes_pos["ScrapQty"],
                marker="x",
                label="报废量 (ScrapQty)",
                color="orange",
            )
          ax.set_title("Positive Electrode Output")
          ax.legend()
          st.pyplot(fig)
        else:
          st.info("请上传正极 MES 文件以查看图表")

    with tab2:
      st.subheader("【负极】MES 实时报工数据 (Anode / -)")
      if df_mes_neg is not None:
        st.dataframe(df_mes_neg)
      else:
        st.info("请在左侧边栏上传负极 MES 报工文件。")

    with tab3:
      st.subheader("【正极】MES 实时报工数据 (Cathode / +)")
      if df_mes_pos is not None:
        st.dataframe(df_mes_pos)
      else:
        st.info("请在左侧边栏上传正极 MES 报工文件。")

    with tab4:
      st.subheader("制浆工序与中转罐校验 (Mixer $\rightarrow$ ST)")
      col_m1, col_m2 = st.columns(2)
      with col_m1:
        st.markdown("#### 负极制浆盘点记录")
        if df_mix_neg is not None:
          st.dataframe(df_mix_neg.dropna(how="all"))
        else:
          st.info("未找到负极制浆记录")
      with col_m2:
        st.markdown("#### 正极制浆盘点记录")
        if df_mix_pos is not None:
          st.dataframe(df_mix_pos.dropna(how="all"))
        else:
          st.info("未找到正极制浆记录")

    with tab5:
      st.subheader("涂布与箔材盘点记录 (Coating & Foil)")
      col_c1, col_c2 = st.columns(2)
      with col_c1:
        st.markdown("#### 负极涂布盘点记录")
        if df_coat_neg is not None:
          st.dataframe(df_coat_neg.dropna(how="all"))
        else:
          st.info("未找到负极涂布盘点表")
      with col_c2:
        st.markdown("#### 正极涂布盘点记录")
        if df_coat_pos is not None:
          st.dataframe(df_coat_pos.dropna(how="all"))
        else:
          st.info("未找到正极涂布盘点表")

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
      " 报工文件** 以启动全自动看板。"
  )