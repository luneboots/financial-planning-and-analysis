import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import numpy as np

file_path = "sample-dataset.csv"
dataset = pd.read_csv(file_path)

dataset.columns = dataset.columns.str.strip()

#Create variance in days for project end dates to see which was on time/late/early
dataset['Project Phase Actual Start Date'] = pd.to_datetime(dataset['Project Phase Actual Start Date'], errors='coerce')
dataset['Project Phase Planned End Date'] = pd.to_datetime(dataset['Project Phase Actual Start Date'], errors='coerce')
dataset['Project Phase Actual End Date'] = pd.to_datetime(dataset['Project Phase Actual End Date'], errors='coerce')

filtered_data = dataset.dropna(subset=['Project Phase Planned End Date', 'Project Phase Actual End Date']).copy()

filtered_data['Variance (Days)'] = (filtered_data['Project Phase Actual End Date'] - filtered_data['Project Phase Planned End Date']).dt.days

filtered_data['Category'] = filtered_data['Variance (Days)'].apply(
    lambda x: 'Ended Early' if x < 0 else ('Ended Late' if x > 0 else 'On Time')
)

trends_analysis = filtered_data[[
    'Project Description',
    'Project Type',
    'Project Phase Planned End Date',
    'Project Phase Actual End Date',
    'Variance (Days)',
    'Category'
]]

#Prepare line graph that illustrates cost in 6-month periods, Green = ended early, Red = ended late
filtered_data['6-Month Period'] = pd.cut(
    filtered_data['Project Phase Actual End Date'],
    bins=pd.date_range(
        start=filtered_data['Project Phase Actual End Date'].min(),
        end=filtered_data['Project Phase Planned End Date'].max(),
        freq='6ME'
    ),
    right=False
)

histogram_data = filtered_data.groupby(['6-Month Period', 'Category'], observed=False).size().unstack(fill_value=0)

plt.figure(figsize=(12,6))
x = np.arange(len(histogram_data.index))
width = 0.5

plt.bar(x - width/2, histogram_data.get('On Time', 0), width=width, label='On Time', color='green')
plt.bar(x + width/2, histogram_data.get('Ended Late', 0), width=width, label='Ended Late', color='red')

for i, count in enumerate(histogram_data.get('On Time', 0)):
    plt.text(x[i] - width/2, count + 0.5, str(count), ha='center', va='bottom', color='green', fontsize=10)
for i, count in enumerate(histogram_data.get('Ended Late', 0)):
    plt.text(x[i] + width/2, count + 0.5, str(count), ha='center', va='bottom', color='red', fontsize=10)

plt.title('Project Count by 6-Month Period')
plt.xlabel('6-Month Period')
plt.ylabel('Number of Projects')
plt.xticks(x, [f"{period.left:%Y-%m} to {period.right:%Y-%m}" for period in histogram_data.index], rotation=45)
plt.legend()

graph_image_path = 'trends_histogram.png'
plt.tight_layout()
plt.savefig(graph_image_path)
plt.close()

output_file = "Financial Planning and Analysis.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    
    trends_analysis.to_excel(writer, index=False, sheet_name='trends-end-date')

    histogram_data.to_excel(writer, index=False, sheet_name='trends-cost')

workbook = load_workbook(output_file)
worksheet = workbook['trends-cost']

img = Image(graph_image_path)
worksheet.add_image(img, 'H2')
workbook.save(output_file)

print(f"Trends analysis added to '{output_file}' in FPA workbook on the sheet trends-X.")














