import tkinter as tk
import pandas as pd

# Function to read and display the stock data
def update_stock_data():
    try:
        # Read the Excel file
        data = pd.read_excel('finalized_orders.xlsx', sheet_name='StockData')  # Ensure 'StockData' matches your sheet name
        
        # Clear the frame before updating with new data
        for widget in frame.winfo_children():
            widget.destroy()

        # Display stock for each item
        for index, row in data.iterrows():
            item = row['Item']
            stock = row['Stock']
            
            # Create a label for each item with its stock amount
            color = 'red' if stock <= 10 else 'black'
            item_label = tk.Label(frame, text=f"{item}: {stock}", font=("Helvetica", 60), fg=color)
            item_label.pack()

    except FileNotFoundError:
        error_label = tk.Label(frame, text="File not found.", font=("Helvetica", 32), fg='red')
        error_label.pack()
    except Exception as e:
        error_label = tk.Label(frame, text=f"Error: {str(e)}", font=("Helvetica", 32), fg='red')
        error_label.pack()

    # Schedule the function to run again after 5 seconds (5000 milliseconds)
    root.after(1000, update_stock_data)

# Create the main window
root = tk.Tk()
root.title("Stock Data Display")

# Create a frame to hold the item labels
frame = tk.Frame(root)
frame.pack(pady=20)

# Call the function to update the stock data initially
update_stock_data()

# Start the Tkinter event loop
root.mainloop()
