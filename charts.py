import pandas as pd
import matplotlib.pyplot as plt

# Data load
df = pd.read_excel('data_sheet.xlsx')
total = df[df['Area'] == 'Total'].copy()

# Clean 
clean = total[['States/UTs',
               df.columns[7],
               df.columns[18],
               df.columns[13],
               df.columns[16]]].copy()

clean.columns = ['State', 'Sex_Ratio',
                 'Female_Literacy',
                 'Sanitation', 'Health_Insurance']

for col in clean.columns[1:]:
    clean[col] = pd.to_numeric(clean[col], errors='coerce')

states = clean[clean['State'] != 'India'].reset_index(drop=True)

# ── Chart 1 — Female Literacy Top 10 ──
top10 = states.nlargest(10, 'Female_Literacy')
bot10 = states.nsmallest(10, 'Female_Literacy')

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
ax1.barh(top10['State'], top10['Female_Literacy'], color='steelblue')
ax1.set_xlabel('Female Literacy (%)')
ax1.set_title('Top 10 States')
ax2.barh(bot10['State'], bot10['Female_Literacy'], color='tomato')
ax2.set_xlabel('Female Literacy (%)')
ax2.set_title('Bottom 10 States')
fig.suptitle('Female Literacy — Top 10 vs Bottom 10 States', fontsize=14)
plt.tight_layout()
plt.savefig('chart_literacy.png')
plt.show()
print("Chart 1 ban gaya!")

# ── Chart 2 — Health Insurance ──
hi = states.sort_values('Health_Insurance', ascending=False)

plt.figure(figsize=(12, 6))
colors = ['green' if v > 41 else 'tomato' for v in hi['Health_Insurance']]
plt.bar(hi['State'], hi['Health_Insurance'], color=colors)
plt.xticks(rotation=45, ha='right', fontsize=7)
plt.ylabel('Coverage (%)')
plt.title('Health Insurance Coverage — All States')
plt.axhline(y=41, color='navy', linestyle='--', label='India Avg: 41%')
plt.legend()
plt.tight_layout()
plt.savefig('chart_insurance.png')
plt.show()
print("Chart 2 ban gaya!")

# ── Chart 3 — Sanitation ──
san = states.sort_values('Sanitation', ascending=True).head(15)

plt.figure(figsize=(10, 6))
plt.barh(san['State'], san['Sanitation'],
         color=['red' if v < 60 else 'orange' for v in san['Sanitation']])
plt.xlabel('Households with Sanitation (%)')
plt.title('Sanitation Access — 15 Lowest States')
plt.tight_layout()
plt.savefig('chart_sanitation.png')
plt.show()
print(" 3 chart are ready !")

print("\n all charths are ready !")