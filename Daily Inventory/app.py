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
                    product_name = row[1]  # Product Name
                    quantity_sold = row[2]  # Quantity Sold
                    extracted_data.append([product_id, product_name, quantity_sold, section_name])
    
    return pd.DataFrame(extracted_data, columns=['Product ID', 'Product Name', 'Quantity Sold', 'Section'])

# Function to open browser automatically
def open_browser():
    webbrowser.open("http://127.0.0.1:5000/")  # Automatically opens the browser

@app.route('/')
def index():
    return render_template('index.html')  # Home page with file upload

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'food_court_file' not in request.files or 'restaurant_file' not in request.files or 'delivery_file' not in request.files:
        return "No file part", 400
    
    food_court_file = request.files['food_court_file']
    restaurant_file = request.files['restaurant_file']
    delivery_file = request.files['delivery_file']
    
    # Save the uploaded files in the uploads folder
    food_court_file_path = os.path.join(upload_folder, 'food_court_report.pdf')
    restaurant_file_path = os.path.join(upload_folder, 'restaurant_report.pdf')
    delivery_file_path = os.path.join(upload_folder, 'delivery_report.pdf')

    food_court_file.save(food_court_file_path)
    restaurant_file.save(restaurant_file_path)
    delivery_file.save(delivery_file_path)
    
    # Process the files and generate the combined report
    food_court_data = extract_data_from_pdf(food_court_file_path, 'Food Court')
    restaurant_data = extract_data_from_pdf(restaurant_file_path, 'Restaurant')
    delivery_data = extract_data_from_pdf(delivery_file_path, 'Delivery')

    combined_data = pd.concat([food_court_data, restaurant_data, delivery_data], ignore_index=True)
    
    # Convert Quantity Sold to numeric, forcing errors to NaN (which can be filled with 0 later)
    combined_data['Quantity Sold'] = pd.to_numeric(combined_data['Quantity Sold'], errors='coerce')
    
    # Create pivot table
    pivot_df = combined_data.pivot_table(
        index=['Product ID', 'Product Name'],
        columns='Section',
        values='Quantity Sold',
        aggfunc='sum',
        fill_value=0
    )
    
    # Calculate the Total Quantity Sold for each product
    pivot_df['Total Quantity Sold'] = pivot_df.sum(axis=1)  # Sum along the row (axis=1)

    # Reset index to keep the product ID and Name as columns
    pivot_df.reset_index(inplace=True)

    # Filter out any unwanted "Total Qty" rows that are mistakenly included
    pivot_df = pivot_df[~pivot_df['Product ID'].str.contains('Total Qty', na=False)]
    
    # Save the final report in uploads folder
    output_file = os.path.join(upload_folder, 'final_combined_report_updated.xlsx')
    pivot_df.to_excel(output_file, index=False)

    # Return the file to the client for download
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    # Open the browser in the background
    Thread(target=open_browser).start()
    app.run(debug=True, use_reloader=False)  # Disable reloader since we're using threading