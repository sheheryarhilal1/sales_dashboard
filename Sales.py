import pandas as pd
import matplotlib.pyplot as plt

# =========================
# 1. LOAD DATA
# =========================
df = pd.read_excel("raw_sales.xlsx")

print("Initial Shape:", df.shape)

# =========================
# 2. CLEAN DATA
# =========================
df = df.dropna()
df = df.drop_duplicates()
df["Date"] = pd.to_datetime(df["Date"])

# =========================
# 3. FEATURE ENGINEERING
# =========================
df["Revenue"] = df["Quantity"] * df["Unit_Price"] * (1 - df["Discount"])
df["Month"] = df["Date"].dt.to_period("M").astype(str)

# =========================
# 4. SUMMARY PRINT
# =========================
print("\nTOTAL SALES:", df["Revenue"].sum())
print("TOTAL ORDERS:", df["Order_ID"].nunique())
print("TOTAL CUSTOMERS:", df["Customer_ID"].nunique())

# =========================
# 5. PLOTS (SINGLE SCREEN OUTPUT)
# =========================

plt.figure(figsize=(10,5))
df.groupby("Region")["Revenue"].sum().plot(kind="bar")
plt.title("Sales by Region")
plt.ylabel("Revenue")
plt.show()

plt.figure(figsize=(10,5))
df.groupby("Product")["Revenue"].sum().plot(kind="bar", color="orange")
plt.title("Sales by Product")
plt.ylabel("Revenue")
plt.show()

plt.figure(figsize=(10,5))
df.groupby("Month")["Revenue"].sum().plot(kind="line", marker="o")
plt.title("Monthly Sales Trend")
plt.xticks(rotation=45)
plt.ylabel("Revenue")
plt.show()

plt.figure(figsize=(10,5))
df.groupby("Customer_ID")["Revenue"].sum().sort_values(ascending=False).head(10).plot(kind="bar")
plt.title("Top 10 Customers")
plt.ylabel("Revenue")
plt.show()

# =========================
# 6. SAVE CLEAN DATA
# =========================
df.to_excel("cleaned_sales.xlsx", index=False)

print("\n✅ Done: Clean data saved + Analysis complete")