import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import openpyxl
import os
import json

def parse_file(file_path):
    _, file_extension = os.path.splitext(file_path)
    if file_extension == '.csv':
        return pd.read_csv(file_path)
    elif file_extension == '.json':
        return pd.read_json(file_path)
    elif file_extension in ['.xls', '.xlsx']:
        return pd.read_excel(file_path)
    else:
        raise ValueError(f"Unsupported file format: {file_extension}")

def parse_and_analyze(file_paths):
    # Aggregate data from multiple files
    data_frames = []
    for file_path in file_paths:
        try:
            df = parse_file(file_path)
            data_frames.append(df)
        except Exception as e:
            print(f"Error parsing {file_path}: {e}")

    if not data_frames:
        raise ValueError("No valid data found in the provided files.")
    
    # Concatenate all data frames
    df = pd.concat(data_frames, ignore_index=True)

    # Clean and process data
    df['ExecutionTime'] = df['ExecutionTime'].str.replace('s', '').astype(float)
    
    # Handle missing values
    df.fillna('N/A', inplace=True)
    
    # Summary statistics
    total_tests = df.shape[0]
    passed_tests = df[df['Status'] == 'Pass'].shape[0]
    failed_tests = df[df['Status'] == 'Fail'].shape[0]
    average_execution_time = df['ExecutionTime'].mean()
    
    analysis = {
        'total_tests': total_tests,
        'passed_tests': passed_tests,
        'failed_tests': failed_tests,
        'average_execution_time': average_execution_time
    }
    
    return df, analysis

def generate_pdf_report(data, analysis, output_file):
    c = canvas.Canvas(output_file, pagesize=letter)
    width, height = letter
    
    # Title
    c.setFont("Helvetica-Bold", 16)
    c.drawString(100, height - 100, "QA Test Results Report")
    
    # Analysis summary
    c.setFont("Helvetica", 12)
    c.drawString(100, height - 140, f"Total Tests: {analysis['total_tests']}")
    c.drawString(100, height - 160, f"Passed Tests: {analysis['passed_tests']}")
    c.drawString(100, height - 180, f"Failed Tests: {analysis['failed_tests']}")
    c.drawString(100, height - 200, f"Average Execution Time: {analysis['average_execution_time']:.2f} seconds")
    
    # Detailed results
    c.drawString(100, height - 240, "Detailed Results:")
    y = height - 260
    for index, row in data.iterrows():
        c.drawString(100, y, f"TestCaseID: {row['TestCaseID']}, TestCaseName: {row['TestCaseName']}, Status: {row['Status']}, ExecutionTime: {row['ExecutionTime']}s, ErrorDetails: {row['ErrorDetails']}")
        y -= 20
    
    c.save()

def generate_excel_report(data, analysis, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "QA Test Results"
    
    # Analysis summary
    sheet.append(["Total Tests", analysis['total_tests']])
    sheet.append(["Passed Tests", analysis['passed_tests']])
    sheet.append(["Failed Tests", analysis['failed_tests']])
    sheet.append(["Average Execution Time", f"{analysis['average_execution_time']:.2f} seconds"])
    sheet.append([])  # Empty row
    
    # Column headers for detailed results
    sheet.append(["TestCaseID", "TestCaseName", "Status", "ExecutionTime", "ErrorDetails"])
    
    # Data rows
    for _, row in data.iterrows():
        sheet.append(list(row))
    
    workbook.save(output_file)

def main():
    # Load configuration
    with open('config.json') as config_file:
        config = json.load(config_file)
    
    # Paths
    input_files = config['input_files']
    pdf_report_path = config['output_pdf']
    excel_report_path = config['output_excel']
    
    # Parse and analyze data
    data, analysis = parse_and_analyze(input_files)
    
    # Generate reports
    generate_pdf_report(data, analysis, pdf_report_path)
    generate_excel_report(data, analysis, excel_report_path)
    
    print("Reports generated successfully!")

if __name__ == "__main__":
    main()
