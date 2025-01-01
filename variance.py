import pandas as pd

file_path = "sample-dataset.csv"
dataset = pd.read_csv(file_path)

dataset.columns = dataset.columns.str.strip()

dataset['Project Budget Amount'] = pd.to_numeric(dataset['Project Budget Amount'], errors = 'coerce')
dataset['Total Phase Actual Spending Amount'] = pd.to_numeric(dataset['Total Phase Actual Spending Amount'], errors='coerce')

#ensure greater than 0, there are many un-budgeted projects
filtered_data = dataset[
    (dataset['Project Budget Amount'] > 0) |
    (dataset['Total Phase Actual Spending Amount'] > 0)
]

#variance analysis by project desc && type
variance_analysis = filtered_data.groupby(['Project Description', 'Project Type']).agg(
    total_budget=pd.NamedAgg(column='Project Budget Amount', aggfunc='sum'),
    total_spending=pd.NamedAgg(column='Total Phase Actual Spending Amount', aggfunc='sum')
).reset_index()

variance_analysis['variance'] = variance_analysis['total_budget'] - variance_analysis['total_spending']

#variance analysis by only project type
variance_analysis_type = filtered_data.groupby(['Project Type']).agg(
    total_budget=pd.NamedAgg(column='Project Budget Amount', aggfunc='sum'),
    total_spending=pd.NamedAgg(column='Total Phase Actual Spending Amount', aggfunc='sum')
).reset_index()

variance_analysis_type['variance'] = variance_analysis_type['total_budget'] - variance_analysis_type['total_spending']

#Show total only for project type
gross_totals = {
    'Project Type': 'TOTAL',
    'total_budget': variance_analysis_type['total_budget'].sum(),
    'total_spending': variance_analysis_type['total_spending'].sum(),
    'variance': variance_analysis_type['variance'].sum(),
    }

variance_analysis_type = variance_analysis_type._append(gross_totals, ignore_index=True)

#Split the sheets to include detailed and type breakdown
output_file = "Financial Planning and Analysis.xlsx"
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    variance_analysis.to_excel(writer, index=False, sheet_name='variance-analysis-desc')

    variance_analysis_type.to_excel(writer, index=False, sheet_name='variance-analysis-type')

    workbook = writer.book
    worksheet_detailed = writer.sheets['variance-analysis-desc']
    worksheet_type = writer.sheets['variance-analysis-type']

    #add formatting for currency
    currency_format = workbook.add_format({'num_format': '$#,##0', 'align': 'right'})

    monetary_columns_detailed = ['total_budget', 'total_spending', 'variance']
    for col_num, column in enumerate(variance_analysis.columns):
        if column in monetary_columns_detailed:
            worksheet_detailed.set_column(col_num, col_num, 15, currency_format)

    monetary_columns_type = ['total_budget', 'total_spending', 'variance']
    for col_num, column in enumerate(variance_analysis_type.columns):
         if column in monetary_columns_type:
            worksheet_type.set_column(col_num, col_num, 15, currency_format)

print(f"Excel document '{output_file}' created with variance analysis sheet.")