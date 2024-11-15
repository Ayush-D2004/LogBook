import pandas as pd
import os
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as ticker  # Import ticker for setting y-axis ticks

EXCEL_FILE = os.path.join(os.path.dirname(__file__), 'logbook.xlsx')

def initialize_logbook():
    """Initialize the logbook Excel file if it doesn't exist."""
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=['Date', 'Title', 'Content', 'Tag'])
        df.to_excel(EXCEL_FILE, index=False)

def add_entry(date_time, title, content, tag):
    """Add a new log entry."""
    df = pd.read_excel(EXCEL_FILE)
    new_entry = pd.DataFrame([[date_time, title, content, tag]], columns=['Date', 'Title', 'Content', 'Tag'])
    df = pd.concat([df, new_entry], ignore_index=True)
    df = df.sort_values(by='Date')  # Sort by date
    df.to_excel(EXCEL_FILE, index=False)

def view_entries():
    """View all log entries."""
    df = pd.read_excel(EXCEL_FILE)
    
    # Convert relevant columns to string for better alignment
    df['Date'] = df['Date'].astype(str)
    df['Title'] = df['Title'].astype(str)
    df['Content'] = df['Content'].astype(str)
    df['Tag'] = df['Tag'].astype(str)
    
    print(df.sort_values(by='Date'))  # Display sorted entries

def update_entry(date_time, title, new_content):
    """Update an existing log entry."""
    df = pd.read_excel(EXCEL_FILE)
    df.loc[(df['Date'] == date_time) & (df['Title'] == title), 'Content'] = new_content
    df = df.sort_values(by='Date')  # Sort by date
    df.to_excel(EXCEL_FILE, index=False)

def delete_entry(date_time, title):
    """Delete a log entry."""
    df = pd.read_excel(EXCEL_FILE)
    df = df[~((df['Date'] == date_time) & (df['Title'] == title))]
    df = df.sort_values(by='Date')  # Sort by date
    df.to_excel(EXCEL_FILE, index=False)

def search_entries(keyword):
    """Search for entries containing a keyword."""
    df = pd.read_excel(EXCEL_FILE)
    results = df[df['Title'].str.contains(keyword, case=False) | df['Content'].str.contains(keyword, case=False)]
    print(results)

def get_current_date_time():
    """Get the current date and time in DD-MM-YYYY HH:MM format."""
    return datetime.now().strftime('%d-%m-%Y %H:%M')

def visualize_entries_per_day():
    """Visualize the number of entries made per day using a scatter plot and save the data to Excel."""
    df = pd.read_excel(EXCEL_FILE)
    
    # Convert 'Date' column to datetime
    df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y %H:%M')
    
    # Group by day and count entries
    daily_counts = df.groupby(df['Date'].dt.date).size()
    
    # Prepare data for scatter plot
    dates = daily_counts.index
    counts = daily_counts.values.tolist()  # Convert to list
    
    # Plotting
    plt.figure(figsize=(10, 6))
    plt.bar(dates, counts, color='skyblue', width=0.5)
    plt.title('Number of Log Entries per Day')
    plt.xlabel(str(datetime.now().year))  # Set to current year as a string
    plt.ylabel('Number of Entries')
    
    # Format the x-axis dates
    plt.gca().xaxis.set_major_locator(mdates.DayLocator())  # Set major ticks to each day
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d-%m'))  # Set date format
    plt.gcf().autofmt_xdate()  # Auto-format the x-axis dates for better visibility
    
    # Set y-axis to show only integer values
    plt.gca().yaxis.set_major_locator(ticker.MaxNLocator(integer=True))  # Ensure y-axis ticks are integers
    
    plt.grid(axis='y')
    plt.tight_layout()
    # plt.show()
    plt.savefig(os.path.join(os.path.dirname(__file__), 'log.png')) 
    plt.close()  

    # Save the counts to the Excel file
    counts_df = pd.DataFrame({'Date': dates, 'Entry Count': counts})
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        counts_df.to_excel(writer, sheet_name='Daily Entry Counts', index=False)

    print("Visualization saved as 'log.png' in the LogBook folder and data logged in 'logbook.xlsx'.")

def main():
    initialize_logbook()
    
    while True:
        print("\n--- LogBook Menu ---")
        print("1. Add Entry")
        print("2. View Entries")
        print("3. Update Entry")
        print("4. Delete Entry")
        print("5. Search Entries")
        print("6. Visualize Entries per Day")
        print("7. Exit")
        
        choice = input("Select an option (1-7): ")
        
        match choice:
            case '1':
                # Prompt user for date and time
                date_time_input = input("Enter date and time (DD-MM-YYYY HH:MM) or press Enter for current time: ")
                if not date_time_input:
                    date_time = get_current_date_time()
                else:
                    # Validate the input date and time
                    try:
                        datetime.strptime(date_time_input, '%d-%m-%Y %H:%M')
                        date_time = date_time_input
                    except ValueError:
                        print("Invalid date and time format. Please use DD-MM-YYYY HH:MM.")
                        continue
                
                title = input("Enter title: ")
                content = input("Enter content: ")
                tag = input("Enter tag (e.g., work, personal, idea): ")
                add_entry(date_time, title, content, tag)
                print(f"Entry added successfully for date and time: {date_time}.")
                
            case '2':
                print("Log Entries:")
                view_entries()
                
            case '3':
                date_time = input("Enter date and time of the entry to update (DD-MM-YYYY HH:MM): ")
                title = input("Enter title of the entry to update: ")
                new_content = input("Enter new content: ")
                update_entry(date_time, title, new_content)
                print("Entry updated successfully.")
                
            case '4':
                date_time = input("Enter date and time of the entry to delete (DD-MM-YYYY HH:MM): ")
                title = input("Enter title of the entry to delete: ")
                delete_entry(date_time, title)
                print("Entry deleted successfully.")
                
            case '5':
                keyword = input("Enter keyword to search: ")
                print("Search Results:")
                search_entries(keyword)
                
            case '6':
                visualize_entries_per_day()
                
            case '7':
                print("Exiting the LogBook application.")
                break
                
            case _:
                print("Invalid choice. Please select a valid option.")

if __name__ == "__main__":
    main()
