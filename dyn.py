
import pandas as pd
import win32com.client as win32

def merge_csv_to_excel_and_copy_to_cloud(csv_files, cloud_workbook):
    try:
        #read and merge
        dfs = [pd.read_csv(file).set_index('Month').transpose() for file in csv_files]
        merged_df = pd.concat(dfs, axis=0)
        
        #reset index and rename them
        merged_df = merged_df.reset_index().rename(columns={'index': 'Subscription'})

        #arrange months in order
        column_headers = sorted(merged_df.columns[1:], key=lambda x: pd.to_datetime(x, format='%B %Y'))

        #reorder columns
        merged_df = merged_df[['Subscription'] + column_headers]

        #replace 0 with blank spaces
        merged_df.replace(0, '', inplace=True)
        merged_df.insert(0, '', '')
        merged_df.iloc[0, 0] = 'AWS'
        merged_df.iloc[len(dfs[0]), 0] = 'Azure'
        merged_df.iloc[len(dfs[0]) + len(dfs[1]), 0] = 'GCP'

        output_file = '"C:/Users/v.jeevinee/Downloads/Consolidate Cloud Expense April 2024 3.xlsx"'
        merged_df.to_excel(output_file, index=False)

        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False

        wb = excel.Workbooks.Open(output_file)
        ws = wb.ActiveSheet

        #autofit columns
        ws.Columns.AutoFit()
        wb.Save()



        print(f"Data merged and saved to {output_file}.")

        #copy merged data to cloud workbook
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True

        source_wb = excel.Workbooks.Open(output_file)
        source_ws = source_wb.ActiveSheet
        cloud_wb = excel.Workbooks.Open(cloud_workbook)
        last_sheet_index = cloud_wb.Sheets.Count
        source_ws.Copy(Before=cloud_wb.Sheets(last_sheet_index))
        copied_ws = cloud_wb.Sheets(last_sheet_index) 
        copied_ws.Name = "Monthly Summary"
        cloud_wb.Save()

        print(f"Data copied from '{output_file}' to '{cloud_workbook}' successfully.")
        excel.Quit()

    except Exception as e:
        print(f"Error: {e}")

#main function
csv_files = ["C:/Users/v.jeevinee/Downloads/gcp.csv",
             "C:/Users/v.jeevinee/Downloads/azure.csv",
             "C:/Users/v.jeevinee/Downloads/aws.csv"]
cloud_workbook = "C:/Users/v.jeevinee/Downloads/Consolidate Cloud Expense April 2024 3.xlsx"

merge_csv_to_excel_and_copy_to_cloud(csv_files, cloud_workbook)
