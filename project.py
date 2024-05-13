import mysql
from mysql import connector
import csv
import pandas
import os
import matplotlib
from matplotlib import pyplot
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter
from tkinter import scrolledtext
from tkinter import messagebox

# File paths - Insert your own
file_path = r"C:\General Programming\Python Programs\hotel_booking.csv"
modified_file_path = r"C:\General Programming\Python Programs\modified_hotel_booking.csv"

# Function which stores the statistics collected from the csv in the database
def insert_statistics():
    try:
        # Establish a connection to the MySQL server
        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            password="Matsaniarakos9",
            database="hoteldb"
        )

        # Switch to the database
        cursor = connection.cursor()

        # Create table for max bookings month per year
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS max_bookings_month_per_year (
                year INT,
                month VARCHAR(20)
            )
        """)
        
        # Create table for max cancellations month per year
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS max_cancellations_month_per_year (
                year INT,
                month VARCHAR(20)
            )
        """)
        
        # Create table for month trend
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS month_trend (
                month VARCHAR(20),
                trend VARCHAR(20)
            )
        """)
        
        # Create table for traveling groups
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS traveling_groups (
                families INT,
                couples INT,
                singles INT
            )
        """)
        
        # Create table for room type distribution
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS room_type_distribution (
                room_type VARCHAR(1),
                count INT
            )
        """)
        
        # Create table for month reservations distribution
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS month_reservations_distribution (
                year INT,
                month VARCHAR(20),
                count INT
            )
        """)
        
        # Create table for season reservations distribution
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS season_reservations_distribution (
                year INT,
                season VARCHAR(20),
                count INT
            )
        """)
        
        # Create table for stay-in average
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS stay_in_average (
                hotel_type VARCHAR(20),
                average_nights FLOAT
            )
        """)
        
        # Create table for hotel cancellations
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS hotel_cancellations (
                hotel_type VARCHAR(20),
                actual_bookings INT,
                cancelled_bookings INT,
                cancellation_percentage FLOAT
            )
        """)

        # Fetch data from hotel_seasonability function
        max_bookings_month_per_year, max_cancellations_month_per_year = hotel_seasonability()

        # Insert data into max_bookings_month_per_year table
        for year, month in max_bookings_month_per_year.items():
            cursor.execute("INSERT INTO max_bookings_month_per_year (year, month) VALUES (%s, %s)", (int(year), str(month)))
        
        # Insert data into max_cancellations_month_per_year table
        for year, month in max_cancellations_month_per_year.items():
            cursor.execute("INSERT INTO max_cancellations_month_per_year (year, month) VALUES (%s, %s)", (int(year), str(month)))

        # Fetch data from hotel_check_trend function
        trend_dict = hotel_check_trend()

        # Insert data into month_trend table
        for month, trend in trend_dict.items():
            cursor.execute("INSERT INTO month_trend (month, trend) VALUES (%s, %s)", (str(month), str(trend)))

        # Fetch data from hotel_traveling_group function
        families, couples, singles = hotel_traveling_group()

        # Insert data into traveling_groups table
        cursor.execute("INSERT INTO traveling_groups (families, couples, singles) VALUES (%s, %s, %s)", (families, couples, singles))
        
        # Fetch data from hotel_room_type_distribution function
        room_type_counts = hotel_room_type_distribution()

        # Insert data into room_type_distribution table
        cursor.execute("INSERT INTO room_type_distribution (room_type, count) VALUES ('A', %s)", (room_type_counts[0],))
        cursor.execute("INSERT INTO room_type_distribution (room_type, count) VALUES ('B', %s)", (room_type_counts[1],))
        cursor.execute("INSERT INTO room_type_distribution (room_type, count) VALUES ('C', %s)", (room_type_counts[2],))
        cursor.execute("INSERT INTO room_type_distribution (room_type, count) VALUES ('D', %s)", (room_type_counts[3],))
        cursor.execute("INSERT INTO room_type_distribution (room_type, count) VALUES ('E', %s)", (room_type_counts[4],))
        cursor.execute("INSERT INTO room_type_distribution (room_type, count) VALUES ('F', %s)", (room_type_counts[5],))
        cursor.execute("INSERT INTO room_type_distribution (room_type, count) VALUES ('G', %s)", (room_type_counts[6],))
        cursor.execute("INSERT INTO room_type_distribution (room_type, count) VALUES ('H', %s)", (room_type_counts[7],))
        
        # Fetch data from hotel_month_reservations_distribution function
        month_counts, season_counts = hotel_month_reservations_distribution()

        # Insert data into month_reservations_distribution table
        for (year, month), count in month_counts.items():
            cursor.execute("INSERT INTO month_reservations_distribution (year, month, count) VALUES (%s, %s, %s)", (year, month, count))
        
        # Insert data into season_reservations_distribution table
        for (year, season), count in season_counts.items():
            cursor.execute("INSERT INTO season_reservations_distribution (year, season, count) VALUES (%s, %s, %s)", (year, season, count))

        # Fetch data from hotel_stay_in_average function
        resort_average, city_average = hotel_stay_in_average()

        # Insert data into stay_in_average table
        cursor.execute("INSERT INTO stay_in_average (hotel_type, average_nights) VALUES ('Resort Hotel', %s)", (resort_average,))
        cursor.execute("INSERT INTO stay_in_average (hotel_type, average_nights) VALUES ('City Hotel', %s)", (city_average,))
        
        # Fetch data from hotel_cancellations function
        actual_resort_bookings, resort_count_cancelled, actual_city_bookings, city_count_cancelled = hotel_cancellations()

        # Calculate cancellation percentage
        resort_cancellation_percentage = resort_count_cancelled / (actual_resort_bookings + resort_count_cancelled) * 100 if actual_resort_bookings + resort_count_cancelled > 0 else 0
        city_cancellation_percentage = city_count_cancelled / (actual_city_bookings + city_count_cancelled) * 100 if actual_city_bookings + city_count_cancelled > 0 else 0
        
        # Insert data into hotel_cancellations table
        cursor.execute("INSERT INTO hotel_cancellations (hotel_type, actual_bookings, cancelled_bookings, cancellation_percentage) VALUES ('Resort Hotel', %s, %s, %s)", (actual_resort_bookings, resort_count_cancelled, resort_cancellation_percentage))
        cursor.execute("INSERT INTO hotel_cancellations (hotel_type, actual_bookings, cancelled_bookings, cancellation_percentage) VALUES ('City Hotel', %s, %s, %s)", (actual_city_bookings, city_count_cancelled, city_cancellation_percentage))

        # Commit the transaction
        connection.commit()

        # Close cursor and connection
        cursor.close()
        connection.close()

    except Exception as e:
        print("Error:", e)
        return None
    
def export_data_to_csv():
    try:
        # Establish a connection to the MySQL server
        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            password="Matsaniarakos9",
            database="hoteldb"
        )

        # Switch to the database
        cursor = connection.cursor()

        # List of tables to export
        tables = [
            "max_bookings_month_per_year",
            "max_cancellations_month_per_year",
            "month_trend",
            "traveling_groups",
            "room_type_distribution",
            "month_reservations_distribution",
            "season_reservations_distribution",
            "stay_in_average",
            "hotel_cancellations"
        ]

        # Export data from each table to a CSV file
        for table in tables:
            # Fetch data from the table
            cursor.execute(f"SELECT * FROM {table}")
            data = cursor.fetchall()

            # Write data to CSV file
            with open(f"{table}.csv", "w", newline="") as csvfile:
                csv_writer = csv.writer(csvfile)
                # Write column headers
                csv_writer.writerow([i[0] for i in cursor.description])
                # Write data rows
                csv_writer.writerows(data)

        # Close cursor and connection
        cursor.close()
        connection.close()

    except Exception as e:
        print("Error:", e)

# Function which returns info about the seasonability of bookings and cancellations per month
def hotel_seasonability():
    try:
        # Open the file and handle the newlines
        with open(file_path, newline='') as data:

            # Read the data
            csv_opener = csv.reader(data)

            # Skip the header row
            next(csv_opener)

            # Define dictionaries to store bookings and cancellations for each month-year combination
            month_year_info = {}

            # Check all the rows
            for row in csv_opener:
                year = int(row[3])
                month = row[4]
                cancellations = int(row[1])

                # Combination of month and year
                key_month = (year, month)

                # Increment month-year bookings
                if key_month not in month_year_info:
                    month_year_info[key_month] = {'bookings': 0, 'cancellations': 0}
                month_year_info[key_month]['bookings'] += 1
                month_year_info[key_month]['cancellations'] += cancellations

        # Find the month with the most bookings for each year
        max_bookings_month_per_year = {}
        for key_year in month_year_info:
            year = key_year[0]
            max_bookings = max(month_year_info, key=lambda x: month_year_info[x]['bookings'] if x[0] == year else 0)[1]
            max_bookings_month_per_year[year] = max_bookings

        # Find the month with the most cancellations for each year
        max_cancellations_month_per_year = {}
        for key_year in month_year_info:
            year = key_year[0]
            max_cancellations = max(month_year_info, key=lambda x: month_year_info[x]['cancellations'] if x[0] == year else 0)[1]
            max_cancellations_month_per_year[year] = max_cancellations

        return max_bookings_month_per_year, max_cancellations_month_per_year
        
    except Exception as e:
        print("Error:", e)
        return None


def hotel_check_trend():
    # Dictionary that contains the bookings count for every month and season
    booking_data = hotel_month_reservations_distribution()
    # Extract just the month data
    month_counts = booking_data[0]

    # Dictionary which indicates the trend for every month
    trend_dict = {}

    # Add the months to a list
    months = list(month_counts.keys())
    # Initialize the first month to be constant
    trend_dict[months[0]] = "constant"

    # Iterate through every month starting from the second month
    for i in range(1, len(months)):
        current_month = months[i]
        previous_month = months[i - 1]

        # Compare the booking counts of the current and the previous month to determine the trend
        if month_counts[current_month] > month_counts[previous_month]:
            trend_dict[current_month] = "upward"
        elif month_counts[current_month] < month_counts[previous_month]:
            trend_dict[current_month] = "downward"
        else:
            trend_dict[current_month] = "constant"

    return trend_dict

# Function which returns the type people who booked (families, couples, single travelers)
def hotel_traveling_group():
    try:
        # Open the file and handle the newlines
        with open(file_path, newline='') as data:

            # Read the data
            csv_opener = csv.reader(data)

            # Skip the header row
            next(csv_opener)

            # Variables that count the traveling groups
            families = couples = singles = 0

            # Check all the rows
            for row in csv_opener:
                
                # Type of traveling group
                adults = int(float(row[9])) if row[9] else 0
                children = int(float(row[10])) if row[10] else 0
                babies = int(float(row[11])) if row[11] else 0
                kids = children + babies

                if adults > 2 or kids >= 1:
                    families += 1
                if adults == 2 and kids == 0:
                    couples += 1
                if adults == 1 and kids == 0:
                    singles += 1
            
            return families, couples, singles
        
    except Exception as e:
        print("Error:", e)
        return None

# Function which returns the number of reservations per room type
def hotel_room_type_distribution():
    try:
        # Open the file and handle the newlines
        with open(file_path, newline='') as data:

            # Read the data
            csv_opener = csv.reader(data)

            # Skip the header row
            next(csv_opener)

            # Variables that count the reservations per room type
            a = b = c = d = e = f = g = h = 0

            # Check all the rows
            for row in csv_opener:
                # Type of the reserved room
                room_type = row[19]

                if room_type == "A":
                    a += 1
                elif room_type == "B":
                    b += 1
                elif room_type == "C":
                    c += 1
                elif room_type == "D":
                    d += 1
                elif room_type == "E":
                    e += 1
                elif room_type == "F":
                    f += 1
                elif room_type == "G":
                    g += 1
                elif room_type == "H":
                    h += 1
            
            return a, b, c, d, e, f, g, h
        
    except Exception as e:
        print("Error:", e)
        return None

# Function which returns the number of reservations for every month and season
def hotel_month_reservations_distribution():
    try:
        # Open the file and handle the newlines
        with open(file_path, newline='') as data:

            # Read the data
            csv_opener = csv.reader(data)

            # Skip the header row
            next(csv_opener)

            # Define dictionaries to store counts for each month-year combination and each season-year combination
            month_year_counts = {}
            season_year_counts = {}

            # Check all the rows
            for row in csv_opener:
                year = int(row[3])
                month = row[4]

                # Combination of month and year
                key_month = (year, month)

                season = ""
                if month in ["December", "January", "February"]:
                    season = "Winter"
                elif month in ["March", "April", "May"]:
                    season = "Spring"
                elif month in ["June", "July", "August"]:
                    season = "Summer"
                elif month in ["September", "October", "November"]:
                    season = "Autumn"

                # Combination of season and year
                key_season = (year, season)

                # Increment month-year count
                if key_month not in month_year_counts:
                    month_year_counts[key_month] = 0
                month_year_counts[key_month] += 1

                # Increment season-year count
                if key_season not in season_year_counts:
                    season_year_counts[key_season] = 0
                season_year_counts[key_season] += 1

            
        return month_year_counts, season_year_counts
        
    except Exception as e:
        print("Error:", e)
        return None

# Function which returns the average stay-in nights in a hotel
def hotel_stay_in_average():
    try:
        # Open the file and handle the newlines
        with open(file_path, newline='') as data:

            # Read the data
            csv_opener = csv.reader(data)

            # Skip the header row
            next(csv_opener)

            # Variables that show the sum of the days all people stayed in each hotel
            resort_total_nights = 0 
            city_total_nights = 0

            # Variables that count the number of rows each hotel appears in
            resort_count = 0
            city_count = 0

            # Check all the rows
            for row in csv_opener:
                # Stayed-in weekend nights
                weekend_nights = int(row[7])
                # Stayed-in weekday nights
                weekday_nights = int(row[8])
                # Name of the hotel
                hotel = row[0]

                if hotel == "Resort Hotel":
                    resort_sum_nights = weekend_nights + weekday_nights
                    resort_total_nights += resort_sum_nights
                    resort_count += 1
                elif hotel == "City Hotel":
                    city_sum_nights = weekend_nights + weekday_nights
                    city_total_nights += city_sum_nights
                    city_count += 1

                        
            # Calculate the averages
            resort_average = resort_total_nights / resort_count
            city_average = city_total_nights / city_count
            
            return resort_average, city_average
        
    except Exception as e:
        print("Error:", e)
        return None

# Function which returns the number of the cancelled bookings in the two hotels
def hotel_cancellations():
    try:
        # Open the file and handle the newlines
        with open(file_path, newline='') as data:

            # Read the data
            csv_opener = csv.reader(data)

            # Skip the header row
            next(csv_opener)

            # Variables that count the number of actual bookings and cancelled ones
            resort_count_cancelled = 0
            city_count_cancelled = 0
            resort_count_all = 0
            city_count_all = 0

            # Check all the rows
            for row in csv_opener:
                # Binary value which indicates a booking cancellation
                cell_value = int(row[1])
                # Name of the hotel
                hotel = row[0]

                if hotel == "Resort Hotel":
                    resort_count_all += 1
                    if cell_value == 1:
                        resort_count_cancelled += 1
                elif hotel == "City Hotel":
                    city_count_all += 1
                    if cell_value == 1:
                        city_count_cancelled += 1

                # Calculate the number of actual bookings
                actual_resort_bookings = resort_count_all - resort_count_cancelled
                actual_city_bookings = city_count_all - city_count_cancelled

            return actual_resort_bookings, resort_count_cancelled, actual_city_bookings, city_count_cancelled
        
    except Exception as e:
        print("Error:", e)
        return None

def main():
    try:
        # Establish a connection to the MySQL server
        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            password="Matsaniarakos9"
        )

        # Create a database called "hotelDB" to store the records
        cursor = connection.cursor()
        cursor.execute("CREATE DATABASE IF NOT EXISTS hoteldb")

        # Switch to the database
        cursor.execute("USE hoteldb")

        # Create a main table for the data
        create_table_query = """
        CREATE TABLE IF NOT EXISTS records (
        hotel VARCHAR(255),
        is_canceled VARCHAR(255),
        lead_time VARCHAR(255),
        arrival_date_year VARCHAR(255),
        arrival_date_month VARCHAR(255),
        arrival_date_week_number VARCHAR(255),
        arrival_date_day_of_month VARCHAR(255),
        stays_in_weekend_nights VARCHAR(255),
        stays_in_week_nights VARCHAR(255),
        adults VARCHAR(255),
        children VARCHAR(255),
        babies VARCHAR(255),
        meal VARCHAR(255),
        country VARCHAR(255),
        market_segment VARCHAR(255),
        distribution_channel VARCHAR(255),
        is_repeated_guest VARCHAR(255),
        previous_cancellations VARCHAR(255),
        previous_bookings_not_canceled VARCHAR(255),
        reserved_room_type VARCHAR(255),
        assigned_room_type VARCHAR(255),
        booking_changes VARCHAR(255),
        deposit_type VARCHAR(255),
        agent VARCHAR(255),
        company VARCHAR(255),
        days_in_waiting_list VARCHAR(255),
        customer_type VARCHAR(255),
        adr VARCHAR(255),
        required_car_parking_spaces VARCHAR(255),
        total_of_special_requests VARCHAR(255),
        reservation_status VARCHAR(255),
        reservation_status_date VARCHAR(255),
        name VARCHAR(255),
        email VARCHAR(255),
        phone_number VARCHAR(255),
        credit_card VARCHAR(255)
        )
        """
        cursor.execute(create_table_query)

        # Preprocess the file to fill empty cells with "NA"
        # Read the CSV file into a DataFrame
        df = pandas.read_csv(file_path)

        # Cast all columns to object data type
        df = df.astype(object)

        # Opt-in to future behavior to suppress the warning
        pandas.set_option('future.no_silent_downcasting', True)

        # Replace empty cells with "NA" using infer_objects, to handle future update of pandas
        df = df.fillna("NA").infer_objects(copy=False)

        # Write the modified DataFrame into a new CSV file
        df.to_csv(modified_file_path, index=False)

        # Open the new, modified CSV file for reading
        with open(modified_file_path) as data:
            csv_opener = csv.reader(data)

            # Skip the header row
            next(csv_opener)

            # Loop through each row of data in the CSV
            for row in csv_opener:
                insert_query = """ 
                INSERT INTO records (hotel, is_canceled, lead_time, arrival_date_year,
                  arrival_date_month, arrival_date_week_number, arrival_date_day_of_month,
                  stays_in_weekend_nights, stays_in_week_nights, adults, children, babies, meal,
                  country, market_segment, distribution_channel, is_repeated_guest, previous_cancellations,
                  previous_bookings_not_canceled, reserved_room_type, assigned_room_type, booking_changes,
                  deposit_type, agent, company, days_in_waiting_list, customer_type, adr, required_car_parking_spaces,
                  total_of_special_requests, reservation_status, reservation_status_date, name, email, phone_number,
                  credit_card) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                  %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """
            
                # Execute the insert query
                cursor.execute(insert_query, row)

        # Commit the changes to the database
        connection.commit()

        # Delete the modified CSV file
        os.remove(modified_file_path)

        # BELOW THE GUI IS IMPLEMENTED

        # Switches to the analytics window
        def show_analytics_window():
            # Hide the analytics frame
            main_frame.grid_forget()
            # Show the main menu frame
            analytics_frame.grid(row=0, column=0, sticky="nsew")

        # Switches to the main window
        def show_main_window():
            # Hide the main menu frame
            analytics_frame.grid_forget()
            # Show the analytics frame
            main_frame.grid(row=0, column=0, sticky="nsew")

        # Function to reset the analytics window
        def analytics_reset():
            # Destroy all widgets in the analytics frame except for the original buttons
            for widget in analytics_frame.winfo_children():
                if widget not in (back_button, button_frame):
                    widget.destroy()

        # Function to be called when the "Cancellation and Booking Seasonability" button is clicked
        def show_cancellation_and_booking_seasonability():

            # Reset the analytics window
            analytics_reset()
            # Create a new frame within the analytics frame
            seasonability_frame = tkinter.Frame(analytics_frame)
            seasonability_frame.pack(expand=True, fill="both")
            # Retrieve results from the hotel_seasonability function
            max_bookings_month_per_year, max_cancellations_month_per_year = hotel_seasonability()

            # Check if results are available
            if max_bookings_month_per_year and max_cancellations_month_per_year:
                # Check if the label widget already exists
                existing_label = None
                for widget in seasonability_frame.winfo_children():
                    if isinstance(widget, tkinter.Label):
                        existing_label = widget
                        break

                # Create a new label or update the existing one with the updated data
                label_text = ""
                for year, month in max_bookings_month_per_year.items():
                    label_text += f"Month with the most bookings for {year}: {month}\n\n\n"
                for year, month in max_cancellations_month_per_year.items():
                    label_text += f"Month with the most cancellations for {year}: {month}\n\n\n"

                if existing_label:
                    # Update the text of the existing label
                    existing_label.config(text=label_text)
                else:
                    # Create a new label to display the results
                    label = tkinter.Label(seasonability_frame, text=label_text, font=("Arial", 12))
                    label.pack(padx=20, pady=20, anchor="center")
            else:
                # If no results are available, show an error message
                messagebox.showerror("Error", "Failed to retrieve seasonability data.")

            # Run the analytics frame's event loop
            analytics_frame.mainloop()

        # Function to be called when the "Booking Trend" button is clicked
        def show_booking_trend():
            # Reset the analytics window
            analytics_reset()
            
            # Call hotel_check_trend() function to get trend data
            trend_dict = hotel_check_trend()
            
            # Create a new frame within the analytics frame
            booking_trend_frame = tkinter.Frame(analytics_frame)
            booking_trend_frame.pack(expand=True, fill="both")
            
            # Check if results are available
            if trend_dict:
                # Check if the label widget already exists
                existing_label = None
                for widget in booking_trend_frame.winfo_children():
                    if isinstance(widget, tkinter.Label):
                        existing_label = widget
                        break
                
                # Create a new label or update the existing one with the trend data
                label_text = "Booking Trend:\n"
                for month, trend in trend_dict.items():
                    label_text += f"{month}: {trend}\n"
                
                if existing_label:
                    # Update the text of the existing label
                    existing_label.config(text=label_text)
                else:
                    # Create a new label to display the trend data
                    label = tkinter.Label(booking_trend_frame, text=label_text, font=("Arial", 12))
                    label.pack(padx=20, pady=20, anchor="center")  # Center the label within the frame
            else:
                # If no results are available, show an error message
                messagebox.showerror("Error", "Failed to retrieve booking trend data.")
            
            # Run the analytics frame's event loop
            analytics_frame.mainloop()

        # Function to be called when the "Resident Groups" button is clicked
        def show_resident_groups():
            # Reset the analytics window
            analytics_reset()
            
            # Call hotel_traveling_group() function to get traveling group data
            traveling_group_data = hotel_traveling_group()
            
            # Check if data is available
            if traveling_group_data:
                # Extract data
                families, couples, singles = traveling_group_data
                
                # Create a new frame within the analytics frame
                resident_groups_frame = tkinter.Frame(analytics_frame)
                resident_groups_frame.pack(expand=True, fill="both")
                
                # Plotting the bar graph
                fig, ax = matplotlib.pyplot.subplots(figsize=(8, 6))
                labels = ['Families', 'Couples', 'Single Travelers']
                values = [families, couples, singles]
                ax.bar(labels, values, color=['blue', 'green', 'orange'])
                ax.set_xlabel('Traveling Group')
                ax.set_ylabel('Number of Bookings')
                ax.set_title('Traveling Group Distribution')
                ax.grid(True)
                
                # Embed the plot into the frame
                canvas = FigureCanvasTkAgg(fig, master=resident_groups_frame)
                canvas.draw()
                canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=True)
            else:
                # If no data is available, show an error message
                messagebox.showerror("Error", "Failed to retrieve traveling group data.")

        # Function to be called when the "Room Type" button is clicked
        def show_room_type():
            # Reset the analytics window
            analytics_reset()
            
            # Call hotel_room_type_distribution() function to get room type distribution data
            room_type_data = hotel_room_type_distribution()
            
            # Check if data is available
            if room_type_data:
                # Extract data
                a, b, c, d, e, f, g, h = room_type_data
                
                # Create a new frame within the analytics frame
                room_type_frame = tkinter.Frame(analytics_frame)
                room_type_frame.pack(expand=True, fill="both")
                
                # Plotting the bar graph
                fig, ax = matplotlib.pyplot.subplots(figsize=(8, 6))
                labels = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
                values = [a, b, c, d, e, f, g, h]
                ax.bar(labels, values, color=['blue', 'green', 'orange', 'red', 'purple', 'brown', 'pink', 'gray'])
                ax.set_xlabel('Room Type')
                ax.set_ylabel('Number of Reservations')
                ax.set_title('Room Type Distribution')
                ax.grid(True)
                
                # Embed the plot into the frame
                canvas = FigureCanvasTkAgg(fig, master=room_type_frame)
                canvas.draw()
                canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=True)
            else:
                # If no data is available, show an error message
                messagebox.showerror("Error", "Failed to retrieve room type distribution data.")

        # Function to be called when the "Booking Distribution" button is clicked
        def show_booking_distribution():
            try:

                # Reset the analytics window
                analytics_reset()

                # Call hotel_month_reservations_distribution() function to get data
                data = hotel_month_reservations_distribution()
                
                # Check if data is available
                if data:
                    # Extract data
                    month_year_counts, season_year_counts = data
                    
                    # Convert dictionaries to lists of tuples for easier plotting
                    month_labels, month_counts = zip(*month_year_counts.items())
                    season_labels, season_counts = zip(*season_year_counts.items())
                    
                    # Plotting the first graph for month reservations
                    matplotlib.pyplot.figure(figsize=(6, 5), dpi=100)
                    matplotlib.pyplot.bar(range(len(month_labels)), month_counts, color='skyblue')
                    matplotlib.pyplot.xlabel('Month-Year', fontsize=10)
                    matplotlib.pyplot.ylabel('Number of Reservations', fontsize=10)
                    matplotlib.pyplot.title('Month Reservations Distribution')
                    matplotlib.pyplot.xticks(range(len(month_labels)), month_labels, rotation=45, fontsize=8)
                    matplotlib.pyplot.grid(True)
                    
                    # Embed the first plot into the analytics window
                    canvas1 = FigureCanvasTkAgg(matplotlib.pyplot.gcf(), master=analytics_frame)
                    canvas1.draw()
                    canvas1.get_tk_widget().pack(side=tkinter.LEFT, fill=tkinter.BOTH, expand=True)
                    
                    # Plotting the second graph for season reservations
                    matplotlib.pyplot.figure(figsize=(6, 5), dpi=100)
                    matplotlib.pyplot.bar(range(len(season_labels)), season_counts, color='lightgreen')
                    matplotlib.pyplot.xlabel('Season-Year', fontsize=10)
                    matplotlib.pyplot.ylabel('Number of Reservations', fontsize=10)
                    matplotlib.pyplot.title('Season Reservations Distribution')
                    matplotlib.pyplot.xticks(range(len(season_labels)), season_labels, rotation=45, fontsize=8)
                    matplotlib.pyplot.grid(True)
                    
                    # Embed the second plot into the analytics window
                    canvas2 = FigureCanvasTkAgg(matplotlib.pyplot.gcf(), master=analytics_frame)
                    canvas2.draw()
                    canvas2.get_tk_widget().pack(side=tkinter.RIGHT, fill=tkinter.BOTH, expand=True)
                else:
                    # If no data is available, show an error message
                    print("Failed to retrieve booking distribution data.")
            except Exception as e:
                print("Error:", e)

        # Function to be called when the "Average Stay-Ins" button is clicked
        def show_average_stay_ins():
            try:
                # Call hotel_stay_in_average() function to get data
                data = hotel_stay_in_average()
                
                # Check if data is available
                if data:
                    # Extract data
                    resort_average, city_average = data
                    
                    # Clear any previous plots on the analytics window
                    analytics_reset()
                    
                    # Plotting the bar plot for average stay-in nights
                    matplotlib.pyplot.figure(figsize=(6, 5), dpi=100)
                    hotel_types = ['Resort Hotel', 'City Hotel']
                    averages = [resort_average, city_average]
                    matplotlib.pyplot.bar(hotel_types, averages, color=['lightcoral', 'lightblue'])
                    matplotlib.pyplot.xlabel('Hotel Type', fontsize=10)
                    matplotlib.pyplot.ylabel('Average Stay-in Nights', fontsize=10)
                    matplotlib.pyplot.title('Average Stay-in Nights in Each Hotel Type')
                    matplotlib.pyplot.xticks(fontsize=8)
                    matplotlib.pyplot.grid(True)
                    
                    # Embed the plot into the analytics window
                    canvas = FigureCanvasTkAgg(matplotlib.pyplot.gcf(), master=analytics_frame)
                    canvas.draw()
                    canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=True)
                else:
                    # If no data is available, show an error message
                    print("Failed to retrieve average stay-in nights data.")
            except Exception as e:
                print("Error:", e)

        # Function to be called when the "Cancellations" button is clicked
        def show_cancellations():
            try:
                # Call the hotel_cancellations() function to get the data
                data = hotel_cancellations()

                # Check if data is available
                if data:
                    # Extract data
                    actual_resort_bookings, resort_count_cancelled, actual_city_bookings, city_count_cancelled = data
                    
                    # Create labels for the pie charts
                    labels_resort = ['Actual Bookings', 'Cancelled Bookings']
                    labels_city = ['Actual Bookings', 'Cancelled Bookings']
                    
                    # Create sizes for the pie charts
                    sizes_resort = [actual_resort_bookings, resort_count_cancelled]
                    sizes_city = [actual_city_bookings, city_count_cancelled]
                    
                    # Colors for the pie charts
                    colors = ['lightcoral', 'lightblue']
                    
                    # Reset the analytics window
                    analytics_reset()
                    
                    # Create a frame within the analytics frame
                    cancellation_pie_frame = tkinter.Frame(analytics_frame)
                    cancellation_pie_frame.pack(expand=True, fill="both")
                    
                    # Create pie chart for the Resort Hotel
                    matplotlib.pyplot.figure(figsize=(6, 5))
                    matplotlib.pyplot.pie(sizes_resort, labels=labels_resort, colors=colors, autopct='%1.1f%%', startangle=140)
                    matplotlib.pyplot.title('Resort Hotel Bookings')
                    matplotlib.pyplot.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
                    
                    # Embed the pie chart into the frame
                    canvas_resort = FigureCanvasTkAgg(matplotlib.pyplot.gcf(), master=cancellation_pie_frame)
                    canvas_resort.draw()
                    canvas_resort.get_tk_widget().pack(side=tkinter.LEFT, fill=tkinter.BOTH, expand=True)
                    
                    # Create pie chart for the City Hotel
                    matplotlib.pyplot.figure(figsize=(6, 5))
                    matplotlib.pyplot.pie(sizes_city, labels=labels_city, colors=colors, autopct='%1.1f%%', startangle=140)
                    matplotlib.pyplot.title('City Hotel Bookings')
                    matplotlib.pyplot.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
                    
                    # Embed the pie chart into the frame
                    canvas_city = FigureCanvasTkAgg(matplotlib.pyplot.gcf(), master=cancellation_pie_frame)
                    canvas_city.draw()
                    canvas_city.get_tk_widget().pack(side=tkinter.RIGHT, fill=tkinter.BOTH, expand=True)
                else:
                    print("Data is not available.")
            except Exception as e:
                print("Error:", e)

        # Create the main window
        window = tkinter.Tk()

        # Set the title of the window
        window.title("Hotel Analytics")

        # Main menu frame
        main_frame = tkinter.Frame(window)

        # Analytics frame
        analytics_frame = tkinter.Frame(window)
        analytics_frame.grid(row=0, column=0, sticky="nsew")

        # Create a button to switch to the analytics window
        analytics_button = tkinter.Button(main_frame, text="Show Analytics", command=show_analytics_window)
        analytics_button.pack(pady=10)

        # Create a text widget with both vertical and horizontal scrollbars
        text_widget = scrolledtext.ScrolledText(main_frame, wrap=tkinter.NONE)
        text_widget.pack(pady=10, expand=True, fill="both")
        text_widget.insert("1.0", df.to_string(index=False))
        text_widget.configure(state="disabled")

        # Create horizontal scrollbar
        xscrollbar = tkinter.Scrollbar(main_frame, orient=tkinter.HORIZONTAL, command=text_widget.xview)
        xscrollbar.pack(side="bottom", fill="x")
        text_widget.config(xscrollcommand=xscrollbar.set)

        # Create vertical scrollbar
        yscrollbar = tkinter.Scrollbar(main_frame, orient=tkinter.VERTICAL, command=text_widget.yview)
        yscrollbar.pack(side="right", fill="y")
        text_widget.config(yscrollcommand=yscrollbar.set)

        # Create a button to switch back to the main menu
        back_button = tkinter.Button(analytics_frame, text="Back to Main Window", command=show_main_window)
        back_button.pack(pady=10)

        # Create a list of names for the buttons
        names = [
            "Cancellation and Booking Seasonability",
            "Booking Trend",
            "Resident Groups",
            "Room Type",
            "Booking Distribution",
            "Average Stay-Ins",
            "Cancellations"
        ]

        # Create a frame to contain the buttons and place it at the bottom center
        button_frame = tkinter.Frame(analytics_frame)
        button_frame.pack(side="bottom", pady=10)

        # Create buttons for each name and arrange them in a row inside the button frame
        for name in names:
            button_command = locals().get("show_" + name.lower().replace(" ", "_").replace("-", "_"))
            if button_command:
                button = tkinter.Button(button_frame, text=name, command=button_command)
                # Arrange them side by size
                button.pack(side="left", padx=10)

        # Center the button frame horizontally
        button_frame.pack_configure(anchor="center")

        # Configure row and column weights to make the widgets stretch along with the window
        window.grid_rowconfigure(0, weight=1)
        window.grid_columnconfigure(0, weight=1)

        # Show the main window
        show_main_window()

        # Run the tkinter event loop
        window.mainloop()

        # After the window closes, the statistics are inserted into the database
        insert_statistics()

        # Export the statistics to a csv
        export_data_to_csv()

    except mysql.connector.Error as err:
        # Report any connector errors
        print("Error:", err)
    finally:
        # Close the cursor and the connection
        if 'connection' in locals():
            connection.close()
        if 'cursor' in locals():
            cursor.close()

main()
