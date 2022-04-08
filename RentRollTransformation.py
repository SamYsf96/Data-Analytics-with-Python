import pandas as pd
from datetime import date, timedelta
import datetime
import openpyxl


def file_transformation(file):
    ### Variables containing arrays of column headers
    drop_columns = ['Resident First Name','Resident Last Name', 'Gender', 'Resident Status', 'Room Type',
                    'Potential Occupancy', 'Size', 'Deposit Required', 'Deposit Received','Lease/Rent Start','Lease/Rent End','Anniversary Date',
                    'Estimated Discharge','Total Stay','Room Status','Bed Status','Available Days','Occupied Days','Rate Type','Market Rate']

    column_list = ['Discount', 'Inspired Forever Program IL','Maintenance Fee ', 'Second Occupant ', 'Utilities', 'Utilities.1']

    mylist = ['Period',	'Property','Unit Number','Unit Sqft','Unit Type','Unit Service Type','Private/Semi-Private Indicator',	
            'Resident ID',	'Physical Move-In Date','Physical Move-Out Date','Days Vacant','Contract Type','Base Rate Type','Care Rate Type',	 
            'Base Actual Rate', 'Base Market Rate', 'Care Actual Rate', 'Care Market Rate', 'Medication Actual Rate', 	 
            'Medication Market Rate', 'Continence Actual Rate', 'Continence Market Rate', 'Other Actual Rate', 'Other Market Rate','Total Actual Rate',
            'Total Market Rate', 'Total Variance']

    ### Open the file
    df = pd.read_csv(file)

    ### Remove facility name column and add period column with last day of previous month
    df = df.drop("Facility Name",1)
    first_day = date.today().replace(day=1)
    last_day_of_month = first_day - timedelta(days=1)
    df.insert(0, "Period", last_day_of_month)

    ## Drop all rows where Reisdent Number is null
    df.dropna(subset=['Resident Number'], inplace=True)

    ### Drop all columns listed in drop_columns array
    for i in drop_columns:
        df = df.drop(i,1)
        date_now = datetime.datetime.now().replace(day=1)

    ###Drop all columns that have the month name of previous month eg. March, April etc.
    date_before = date_now - timedelta(days=1)
    month_name = date_before.strftime("%B")
    for i in df.columns:
        if month_name in i:
            df.drop(i,axis=1,inplace=True)

    ### Remove Monthly Forecast column
    df = df.drop('Monthly Forecast',1)

    ### Add new column named base actual rate that sums values from the columns listed below     
    base_actual_rate = df.loc[0:300,['Actual Rate','Inspired Forever Program IL','Second Occupant ','Maintenance Fee ','Utilities','Utilities.1']].sum(axis=1)
    df.insert(9,"Base Actual Rate",base_actual_rate)

    ### Remove columns in columns_list array
    for i in column_list:
        df.drop(i,axis=1,inplace=True)

    df.drop('Bed',axis=1,inplace=True)

    ### Insert columns below
    df.insert(0, "Unit Sqft", 0)
    df.insert(1, "Unit Service Type", "Independent Living")
    df.insert(2, "Private/Semi-Private Indicator", "Private")
    df.insert(3, "Days Vacant","")
    df.insert(4, "Contract Type","Permanent")
    df.insert(5,"Base Rate Type","Monthly")
    df.insert(7, "Care Rate Type", "")
    df.insert(8, "Care Actual Rate",0)
    df.insert(9, "Care Market Rate",0)
    df.insert(10, "Medication Actual Rate",0)
    df.insert(11, "Medication Market Rate",0)
    df.insert(12, "Continence Actual Rate",0)
    df.insert(13, "Continence Market Rate",0)
    df.insert(14,"Other Market Rate",0)

    ### Rename columns
    df.rename(columns={'Facility Code':'Property','Unit':'Unit Type','Room':'Unit Number','Resident Number':'Resident ID','Admission':'Physical Move-In Date',
                    'Actual Discharge':'Physical Move-Out Date','Actual Rate':'Base Market Rate','Concession IL':'Other Actual Rate'}, inplace = True)

    ###Add 2 new columns which sum values from other columns, add variance column to subtract actual - market base rate
    columns_with_actual = df.loc[0:300,['Base Actual Rate', 'Other Actual Rate']].sum(axis=1)
    df.insert(0,"Total Actual Rate",columns_with_actual)
    columns_with_market = df.loc[0:300,['Base Market Rate']].sum(axis=1)
    df.insert(0,"Total Market Rate", columns_with_market)
    df['Total Variance'] = df['Total Actual Rate'] - df['Total Market Rate']
    
    ### Reorder columns to match the ones in mylist array
    df = df.reindex(columns=list([a for a in mylist]))

    ### save file in designated directory
    writer = pd.ExcelWriter('Y:\\Reporting\\WELL Templates\\Source\\3004 WELL Rent Roll Template.xlsx')
    df.to_excel(writer)
    writer.save()

file = 'C:\\Users\\syousefi\\PythonScripts\\Pandas\\CompletedScripts\\venv\\3004 Rent Roll.csv'
file_transformation(file)
