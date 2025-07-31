import os
import pdfplumber
import pandas as pd
from flask import Flask, render_template, request, send_file
import webbrowser
from threading import Thread

app = Flask(__name__)

# Ensure uploads folder exists (this will handle folder creation automatically)
upload_folder = os.path.join(os.getcwd(), 'uploads')  # Absolute path
if not os.path.exists(upload_folder):
    os.makedirs(upload_folder)  # Create the folder if it doesn't exist

# Function to extract data from a single PDF file
def extract_data_from_pdf(pdf_file, section_name):
    extracted_data = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row in table[1:]:  # Skip header row
                    product_id = row[0]  # Product ID
                    product_name = row[1] if row[1] else 'Unknown Product'  # Handle missing product names
                    quantity_sold = row[2]  # Quantity Sold

                    # Ensure the quantity is a valid number (handle missing or non-numeric values)
                    try:
                        quantity_sold = float(quantity_sold)
                    except (ValueError, TypeError):
                        quantity_sold = 0  # Default to 0 if conversion fails

                    extracted_data.append([product_id, product_name, quantity_sold, section_name])

    return pd.DataFrame(extracted_data, columns=['Product ID', 'Product Name', 'Quantity Sold', 'Section'])

# Function to read the stock data from the Excel file
def read_excel_for_stock(file_path):
    stock_data = pd.read_excel(file_path)
    
    # Check if the required columns exist
    required_columns = ['Product ID', 'Product Name', 'Opening Stock', 'Issuance Stock', 'Physical Stock']
    missing_columns = [col for col in required_columns if col not in stock_data.columns]
    
    if missing_columns:
        raise ValueError(f"Missing columns in stock file: {', '.join(missing_columns)}")
    
    # Select the required columns
    stock_data = stock_data[['Product ID', 'Product Name', 'Opening Stock', 'Issuance Stock', 'Physical Stock']]
    
    return stock_data

# Function to open browser automatically
def open_browser():
    webbrowser.open("http://127.0.0.1:5000/")  # Automatically opens the browser

@app.route('/')
def index():
    return render_template('index.html')  # Home page with file upload

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'food_court_file' not in request.files or 'restaurant_file' not in request.files or 'delivery_file' not in request.files or 'stock_file' not in request.files:
        return "No file part", 400
    
    food_court_file = request.files['food_court_file']
    restaurant_file = request.files['restaurant_file']
    delivery_file = request.files['delivery_file']
    stock_file = request.files['stock_file']
    
    # Save the uploaded files in the uploads folder
    food_court_file_path = os.path.join(upload_folder, 'food_court_report.pdf')
    restaurant_file_path = os.path.join(upload_folder, 'restaurant_report.pdf')
    delivery_file_path = os.path.join(upload_folder, 'delivery_report.pdf')
    stock_file_path = os.path.join(upload_folder, 'stock_report.xlsx')

    food_court_file.save(food_court_file_path)
    restaurant_file.save(restaurant_file_path)
    delivery_file.save(delivery_file_path)
    stock_file.save(stock_file_path)
    
    # Process the files and generate the combined report
    food_court_data = extract_data_from_pdf(food_court_file_path, 'Food Court')
    restaurant_data = extract_data_from_pdf(restaurant_file_path, 'Restaurant')
    delivery_data = extract_data_from_pdf(delivery_file_path, 'Delivery')
    
    # Read the stock data from the Excel file
    stock_data = read_excel_for_stock(stock_file_path)

    # Merge the stock data with the sales data
    combined_data = pd.concat([food_court_data, restaurant_data, delivery_data], ignore_index=True)

    # Convert Quantity Sold to numeric, forcing errors to NaN (which can be filled with 0 later)
    combined_data['Quantity Sold'] = pd.to_numeric(combined_data['Quantity Sold'], errors='coerce')

    # Create pivot table (only Product ID and Product Name)
    pivot_df = combined_data.pivot_table(
        index=['Product ID', 'Product Name'],  # Product ID and Name as index
        columns='Section',
        values='Quantity Sold',
        aggfunc='sum',
        fill_value=0
    )
    
    # Ensure the stock data columns are numeric, coercing errors to NaN
    stock_data['Opening Stock'] = pd.to_numeric(stock_data['Opening Stock'], errors='coerce')
    stock_data['Issuance Stock'] = pd.to_numeric(stock_data['Issuance Stock'], errors='coerce')
    stock_data['Physical Stock'] = pd.to_numeric(stock_data['Physical Stock'], errors='coerce')

    # Merge the stock data with the pivot table on 'Product ID'
    final_df = pd.merge(pivot_df, stock_data, on='Product ID', how='left')

    # Calculate the Total Quantity Sold for each product (sum across sections)
    final_df['Total Quantity Sold'] = final_df['Food Court'] + final_df['Restaurant'] + final_df['Delivery']

    # Calculate the Difference (Physical Stock - Total Quantity Sold)
    final_df['Stock Difference'] = final_df['Physical Stock'] - final_df['Total Quantity Sold']

    # Reset index to keep the product ID and Name as columns
    final_df.reset_index(inplace=True)

    # Reorder columns to place Product Name after Product ID
    final_df = final_df[['Product ID', 'Product Name', 'Food Court', 'Restaurant', 'Delivery', 'Opening Stock', 'Issuance Stock', 'Physical Stock', 'Total Quantity Sold', 'Stock Difference']]

    # Save the final report in uploads folder
    output_file = os.path.join(upload_folder, 'final-Combined-Difference-report.xlsx')
    final_df.to_excel(output_file, index=False)

    # Return the file to the client for download
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    # Open the browser in the background
    Thread(target=open_browser).start()
    app.run(debug=True, use_reloader=False)  # Disable reloader since we're using threading
