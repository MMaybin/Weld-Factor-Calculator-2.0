
#!/usr/bin/env python3
import sys
import pandas as pd

DATA_XLSX = r"""/mnt/data/Weld_Factor_Calculator.xlsx"""

def load_table(path=DATA_XLSX):
    df = pd.read_excel(path, sheet_name="Data")
    return df

def compute_key(args, criteria_cols):
    parts = []
    for c in criteria_cols:
        v = args.get(c, "")
        try:
            import math
            if isinstance(v, float) and math.isnan(v):
                v = ""
        except Exception:
            pass
        parts.append(str(v).strip())
    return " | ".join(parts)

def main():
    df = load_table()
    # Identify factor column
    factor_col = None
    for c in df.columns:
        if "factor" in str(c).lower():
            factor_col = c
            break
    if factor_col is None:
        numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        factor_col = numeric_cols[0] if numeric_cols else df.columns[-2]

    criteria_cols = [c for c in df.columns if c not in [factor_col, "__KEY__"]]

    # Parse CLI args as key=value
    args = {}
    for arg in sys.argv[1:]:
        if "=" in arg:
            k, v = arg.split("=", 1)
            args[k] = v

    # Ensure all criteria present
    for c in criteria_cols:
        if c not in args:
            print("Missing input: " + c + ". Pass as " + c + "=value")
            return 1

    key = compute_key(args, criteria_cols)
    match = df.loc[df["__KEY__"] == key]
    if match.empty:
        print("No matching combination found.")
        return 2
    val = match.iloc[0][factor_col]
    print(val)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
