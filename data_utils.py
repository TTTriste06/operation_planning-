import pandas as pd
import re

def extract_required_columns(sheet_name: str, df: pd.DataFrame) -> pd.DataFrame:
    """
    提取指定字段，清洗列名并根据字段名自动类型转换（日期 -> datetime、数量/金额 -> float）。
    """
    df = df.copy()

    # 🔄 去除列名中的中文/英文括号注释
    cleaned_columns = {col: re.sub(r"[（(].*?[）)]", "", col).strip() for col in df.columns}
    df.rename(columns=cleaned_columns, inplace=True)

    # 📋 每张表所需字段（静态部分）
    required_fields_map = {
        "赛卓-未交订单": ["预交货日", "品名", "规格", "晶圆品名", "未交订单数量", "已交订单数量", "订单数量"],
        "赛卓-成品在制": ["产品规格", "产品品名", "晶圆型号", "封装形式", "工作中心", "未交"],
        "赛卓-成品库存": ["品名", "WAFER品名", "规格", "仓库名称", "数量"],
        "赛卓-到货明细": ["到货日期", "品名", "规格", "允收数量"],
        "赛卓-下单明细": ["下单日期", "供应商名称", "回货明细_回货品名", "回货明细_回货规格", "回货明细_回货数量"],
        "赛卓-销货明细": ["交易日期", "品名", "规格", "数量", "原币金额"],
        "赛卓-安全库存": ["WaferID", "OrderInformation", "ProductionNO.", "InvWaf", "InvPart"],
        "赛卓-预测": ["产品型号", "生产料号"]
    }

    # 📌 添加预测列中动态包含“预测”的字段
    required_fields = required_fields_map.get(sheet_name, [])
    if sheet_name == "赛卓-预测":
        forecast_cols = [col for col in df.columns if "预测" in col]
        required_fields += forecast_cols

    # ✅ 实际存在的列
    present_fields = [col for col in required_fields if col in df.columns]
    missing_fields = [col for col in required_fields if col not in df.columns]
    if missing_fields:
        print(f"⚠️ `{sheet_name}` 缺少字段: {missing_fields}")

    # ✨ 类型转换：推断字段类型
    for col in present_fields:
        if "日期" in col:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        elif any(keyword in col for keyword in ["数量", "金额", "预测", "Inv", "未交"]):
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df[present_fields]
