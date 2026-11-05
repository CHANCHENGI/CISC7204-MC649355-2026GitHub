"""Create preliminary visualization for project proposal."""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import os

BASE = os.path.dirname(__file__)

# --- Load price indices ---
idx_path = os.path.join(BASE, "rvd_price_indices.csv")
df_idx = pd.read_csv(idx_path, skiprows=1)
df_idx = df_idx.rename(columns={"Month": "month"})
df_idx["date"] = pd.to_datetime(df_idx["month"], format="%m-%Y")
df_idx = df_idx.sort_values("date")

classes = ["Class A", "Class B", "Class C", "Class D", "Class E"]
df_recent = df_idx[df_idx["date"] >= "1999-01-01"].copy()

fig, axes = plt.subplots(1, 2, figsize=(14, 5))

# Chart 1: Price indices by class (territory-wide)
ax1 = axes[0]
for cls in classes:
    ax1.plot(df_recent["date"], df_recent[cls], label=cls, linewidth=1.5)
ax1.set_title("Hong Kong Private Residential Price Indices\nby Housing Class (1999–Present)")
ax1.set_xlabel("Year")
ax1.set_ylabel("Price Index (1999 = 100)")
ax1.legend(loc="upper left", fontsize=8)
ax1.xaxis.set_major_locator(mdates.YearLocator(5))
ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
ax1.grid(True, alpha=0.3)

# --- Load average prices by region ---
price_path = os.path.join(BASE, "rvd_prices.csv")
df_price = pd.read_csv(price_path, skiprows=1)
df_price = df_price.rename(columns={"Month": "month"})
df_price["date"] = pd.to_datetime(df_price["month"], format="%m-%Y")
df_price = df_price.sort_values("date")
df_price_recent = df_price[df_price["date"] >= "2010-01-01"].copy()

# Chart 2: Class B average prices by region
ax2 = axes[1]
regions = {
    "Class B Hong Kong": "Hong Kong Island",
    "Class B Kowloon": "Kowloon",
    "Class B New Territories": "New Territories",
}
for col, label in regions.items():
    series = pd.to_numeric(df_price_recent[col], errors="coerce")
    ax2.plot(df_price_recent["date"], series, label=label, linewidth=1.5)
ax2.set_title("Class B Average Prices by Region\n(2010–Present, HK$/m²)")
ax2.set_xlabel("Year")
ax2.set_ylabel("Average Price (HK$/m²)")
ax2.legend(loc="upper left", fontsize=8)
ax2.xaxis.set_major_locator(mdates.YearLocator(2))
ax2.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
ax2.grid(True, alpha=0.3)

plt.tight_layout()
fig_path = os.path.join(BASE, "proposal_visualization.png")
plt.savefig(fig_path, dpi=150, bbox_inches="tight")
print(f"Saved: {fig_path}")
plt.close()
