"""Generate C:\TradeReview\colors.xlsx (3 color versions).
Run once: python make_colors_xlsx.py
After creation, edit rows in Excel anytime; app re-reads the file at startup.
Up to 4 versions will be read by the app.

Border columns are optional — leave empty for no border ring on endpoints.
"""
import pandas as pd, os
OUT_DIR = r"C:\TradeReview"
OUT = os.path.join(OUT_DIR, "colors.xlsx")
os.makedirs(OUT_DIR, exist_ok=True)
df = pd.DataFrame([
    {"Version": "1",
     "App_Long": "#5BA0FF", "App_Short": "#B0B0B0", "App_Hold": "#9B59B6",
     "Export_Long": "#0040C0", "Export_Short": "#000000", "Export_Hold": "#B07CC6",
     "App_Long_Border": "", "App_Short_Border": "", "App_Hold_Border": "",
     "Export_Long_Border": "", "Export_Short_Border": "", "Export_Hold_Border": ""},
    {"Version": "2",
     "App_Long": "#FF6600", "App_Short": "#9900FF", "App_Hold": "#3399FF",
     "Export_Long": "#FF6600", "Export_Short": "#9900FF", "Export_Hold": "#3399FF",
     "App_Long_Border": "", "App_Short_Border": "", "App_Hold_Border": "",
     "Export_Long_Border": "", "Export_Short_Border": "", "Export_Hold_Border": ""},
    {"Version": "3",
     "App_Long": "#F5D76E", "App_Short": "#9900FF", "App_Hold": "#3399FF",
     "Export_Long": "#F5D76E", "Export_Short": "#9900FF", "Export_Hold": "#3399FF",
     "App_Long_Border": "#B8860B", "App_Short_Border": "", "App_Hold_Border": "",
     "Export_Long_Border": "#B8860B", "Export_Short_Border": "", "Export_Hold_Border": ""},
])
df.to_excel(OUT, index=False)
print(f"Wrote {OUT}")
