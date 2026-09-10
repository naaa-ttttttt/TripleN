import calendar
from datetime import datetime
import pandas as pd

def generate_adashe_ledger(customer_data, year, month, filename="Adashe_ledger.xlsx"):
    """Generate a structured Excel ledger for 'Triple N plus one' Adashe savings.

    Args: 
        customer_data (list): List of dicts with 'names' and 'daily_rate'.
        year (int): Year of the contribution cycle, e.g(,,2026).
        month (int): Month of the cycle (1 - 12).
        filename (str): Output Excel file name.
    """
    
    _, total_days = calendar.monthrange(year, month)
    month_name = calendar.month_name[month]

    processed_records = []

    for customer in customer_data:
        name = customer["name"]
        daily_rate = customer["daily_rate"]

        total_collected = daily_rate * total_days
        organizer_earnings = daily_rate
        customer_payout = total_collected - organizer_earnings

        processed_records.append({
            "Customer Name": name,
            "Daily rate (₦)": daily_rate,
            "Total Days": total_days,
            "Total Collected (₦)": total_collected,
            "Company's Earnings (₦)": organizer_earnings,
            "Customer Payout (₦)": customer_payout,

        })

        excel_format = pd.DataFrame(processed_records)
        excel_title = f"Adashe_{month_name}_{year}"

        company_total_earnings = excel_format["Company's Earnings (₦)"].sum()

        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            excel_format.to_excel(writer, sheet_name=excel_title[:31], index=False)

            worksheet = writer.sheets[excel_title[:31]]
            last_row = len(excel_format) + 2  # +1 for header, +1 for next empty row
            worksheet.cell(row=last_row, column=5, value="Total")
            worksheet.cell(row=last_row, column=5, value=company_total_earnings)

            print(f"saved ledger to '{filename}' under sheet '{excel_title[:31]}'")            

def process_all_workers(worker_files, year, month):
    for worker_name, csv_path in worker_files.items():
        customer_data = pd.read_csv(csv_path).to_dict(orient="records")
        output_filename = f"{worker_name}_ledger_{month}_{year}.xlsx"
        generate_adashe_ledger(customer_data, year, month, filename=output_filename)
        print(f"Done: {worker_name} -> {output_filename}")

 

worker_files = {
    "mary": "mary.csv",
    "grace": "grace.csv",
}

process_all_workers(worker_files, 2026, 9)
