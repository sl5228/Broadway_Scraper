import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import re

def scrape_weekly_show_data(url):
    """
    Scrapes weekly show data from a website and saves to Excel
    
    Args:
        url (str): The website URL to scrape
    
    Returns:
        pandas.DataFrame: DataFrame containing the scraped data
    """
    try:
        # Fetch the webpage
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Raises an HTTPError for bad responses
        
        # Parse the HTML
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Extract data - Method 1: Using regex patterns
        text_content = soup.get_text()
        
        # Define regex patterns for all data points
        week_pattern = r'Week Ending:\s*(\d{1,2}/\d{1,2}/\d{4})'
        shows_pattern = r'Number of Shows:\s*(\d+)'
        gross_pattern = r'Gross Gross:\s*\$?([\d,]+)'
        attendance_pattern = r'Total Attendance:\s*([\d,]+)'
        
        # Find all instances of each data type
        weeks = re.findall(week_pattern, text_content)
        shows = re.findall(shows_pattern, text_content)
        gross = re.findall(gross_pattern, text_content)
        attendance = re.findall(attendance_pattern, text_content)
        
        # Debug: Print what was found
        print(f"Found {len(weeks)} weeks, {len(shows)} shows, {len(gross)} gross, {len(attendance)} attendance")
        
        # Alternative Method 2: If data is in specific HTML elements
        # Uncomment and modify these lines if the above regex doesn't work well
        """
        weeks = []
        shows = []
        gross = []
        attendance = []
        
        # Method 2a: Look for text patterns in HTML elements
        for element in soup.find_all(text=re.compile(r'Week Ending:')):
            week_match = re.search(r'Week Ending:\s*(\d{1,2}/\d{1,2}/\d{4})', element)
            if week_match:
                weeks.append(week_match.group(1))
        
        for element in soup.find_all(text=re.compile(r'Number of Shows:')):
            shows_match = re.search(r'Number of Shows:\s*(\d+)', element)
            if shows_match:
                shows.append(shows_match.group(1))
        
        for element in soup.find_all(text=re.compile(r'Gross Gross:')):
            gross_match = re.search(r'Gross Gross:\s*\$?([\d,]+)', element)
            if gross_match:
                gross.append(gross_match.group(1))
        
        for element in soup.find_all(text=re.compile(r'Total Attendance:')):
            attendance_match = re.search(r'Total Attendance:\s*([\d,]+)', element)
            if attendance_match:
                attendance.append(attendance_match.group(1))
        
        # Method 2b: If data is in structured HTML (like tables or specific divs)
        # Example for table structure:
        # for row in soup.find_all('tr'):  # or 'div', 'li', etc.
        #     cells = row.find_all('td')  # or whatever contains the data
        #     if len(cells) >= 4:
        #         weeks.append(extract_date_from_cell(cells[0]))
        #         shows.append(extract_number_from_cell(cells[1]))
        #         gross.append(extract_currency_from_cell(cells[2]))
        #         attendance.append(extract_number_from_cell(cells[3]))
        """
        
        # Check if any data was found
        if not any([weeks, shows, gross, attendance]):
            print("No data found with current regex patterns. The website structure may have changed.")
            # Save the HTML content for debugging
            with open('debug_html.txt', 'w', encoding='utf-8') as f:
                f.write(text_content[:5000])  # First 5000 characters
            print("First 5000 characters of webpage saved to debug_html.txt for inspection")
            return pd.DataFrame()
        
        # Create DataFrame with all four columns
        data = []
        
        # Check if all lists have the same length
        if not (len(weeks) == len(shows) == len(gross) == len(attendance)):
            print(f"Warning: Mismatched data counts - Weeks: {len(weeks)}, Shows: {len(shows)}, Gross: {len(gross)}, Attendance: {len(attendance)}")
            print("Using the minimum count to avoid mismatched data.")
        
        # Use minimum length to ensure data integrity
        min_length = min(len(weeks), len(shows), len(gross), len(attendance))
        
        if min_length == 0:
            print("No matching data found for all required fields.")
            return pd.DataFrame()
        
        for i in range(min_length):
            try:
                # Clean up the data
                clean_gross = gross[i].replace(',', '')  # Remove commas from gross
                clean_attendance = attendance[i].replace(',', '')  # Remove commas from attendance
                
                data.append({
                    'Week_Ending': weeks[i],
                    'Gross_Gross': int(clean_gross),
                    'Total_Attendance': int(clean_attendance),
                    'Number_of_Shows': int(shows[i]),
                    'Scraped_Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })
            except ValueError as e:
                print(f"Error converting data at index {i}: {e}")
                continue
        
        df = pd.DataFrame(data)
        
        # Convert Week_Ending to datetime for better Excel formatting
        if not df.empty:
            try:
                df['Week_Ending'] = pd.to_datetime(df['Week_Ending'])
                # Sort by week ending date (newest first)
                df = df.sort_values('Week_Ending', ascending=False).reset_index(drop=True)
            except Exception as e:
                print(f"Warning: Could not convert dates: {e}")
        
        return df
        
    except requests.RequestException as e:
        print(f"Error fetching the webpage: {e}")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error processing data: {e}")
        return pd.DataFrame()

def save_to_excel(df, filename=None):
    """
    Saves DataFrame to Excel file
    
    Args:
        df (pandas.DataFrame): Data to save
        filename (str): Optional filename, defaults to timestamped name
    """
    if filename is None:
        filename = f"weekly_show_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    try:
        # Save to Excel with formatting
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Weekly Shows', index=False)
            
            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Weekly Shows']
            
            # Format currency and number columns in Excel
            from openpyxl.styles import NamedStyle
            
            # Create currency style
            currency_style = NamedStyle(name='currency')
            currency_style.number_format = '"$"#,##0'
            
            # Create number style with commas
            number_style = NamedStyle(name='number_comma')
            number_style.number_format = '#,##0'
            
            # Apply formatting to columns
            for row in range(2, len(df) + 2):  # Start from row 2 (skip header)
                worksheet[f'B{row}'].style = currency_style  # Gross_Gross column
                worksheet[f'C{row}'].style = number_style    # Total_Attendance column
            
            # Set column widths
            column_widths = {
                'A': 15,  # Week_Ending
                'B': 15,  # Gross_Gross
                'C': 18,  # Total_Attendance
                'D': 15,  # Number_of_Shows
                'E': 20   # Scraped_Date
            }
            
            for col_letter, width in column_widths.items():
                worksheet.column_dimensions[col_letter].width = width
        
        print(f"Data saved to {filename}")
        return filename
        
    except Exception as e:
        print(f"Error saving to Excel: {e}")
        return None

def main():
    # Replace with your actual URL
    url = "https://www.broadwayleague.com/research/grosses-broadway-nyc/"
    
    print("Starting web scraping...")
    
    # Scrape the data
    df = scrape_weekly_show_data(url)
    
    if not df.empty:
        print(f"Successfully scraped {len(df)} records")
        print("\nPreview of scraped data:")
        print(df.head())
        
        # Save to Excel
        filename = save_to_excel(df)
        
        if filename:
            print(f"\nData successfully saved to {filename}")
        else:
            print("\nFailed to save data to Excel")
    else:
        print("No data was scraped. Please check the URL and data format.")

if __name__ == "__main__":
    main()