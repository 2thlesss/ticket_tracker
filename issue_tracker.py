import pandas as pd
import os


# Check if the data file exists; if not, create it
data_file = 'issue_tracker.xlsx'
if not os.path.exists(data_file):
    df = pd.DataFrame(columns=['Number', 'Tech', 'Date', 'Issue', 'Resolution', 'Follow Up'], dtype=str)
    df.to_excel(data_file, index=False)

# Load the existing data into a DataFrame
df = pd.read_excel(data_file)


# Function to add a new issue
def add_issue(tech, date, issue, resolution, follow_up):
    issue_number = len(df) + 1
    df.loc[issue_number - 1] = [issue_number, tech, date, issue, resolution, follow_up]
    df.to_excel(data_file, index=False)


# Function to view all issues
def view_issues():
    if len(df) == 0:
        print("No issues found.")
    else:
        # Set 'Number' as the index and reset it to remove the count
        df.set_index('Number', inplace=True)
        df.reset_index(inplace=True)
        print(df)


# Function to search by tech name
def search_by_tech(tech_name):
    tech_df = df[df['Tech'].str.contains(tech_name, case=False, na=False)]
    if tech_df.empty:
        print("No issues found for the specified tech.")
    else:
        print(tech_df)

    # funtion to follow up with an issue at a later date
    # this function needs to append the row of the ticket, searched by ticket number.
    # ideally, the user will search for an issue by tech name, then follow up with that ticket number.
    # followups will go into the follow_up column of the spread sheet.
def follow_up_issue():
    ticket_number = input("Enter the ticket number to follow up: ")
    follow_up_text = input("Enter the follow-up text: ")

    # Locate the ticket with the specified number
    ticket_row = df[df['Number'] == int(ticket_number)]

    if not ticket_row.empty:
        index = ticket_row.index[0]
        current_follow_up = str(ticket_row['Follow Up'].values[0])  # Convert to string
        updated_follow_up = current_follow_up + "\n" + follow_up_text
        df.at[index, 'Follow Up'] = updated_follow_up
        df.to_excel(data_file, index=False)
        print("Follow-up added successfully!")
    else:
        print(f"Ticket with number {ticket_number} not found.")


        # Prints out tickets by number in a readable format
def print_ticket_details(ticket_number):
    ticket_row = df[df['Number'] == ticket_number]

    if not ticket_row.empty:
        print("\nTicket Details:")
        print(f"Ticket Number: {ticket_row['Number'].values[0]}")
        print(f"Tech: {ticket_row['Tech'].values[0]}")
        print(f"Date: {ticket_row['Date'].values[0]}")
        print(f"Issue: \n{ticket_row['Issue'].values[0]}")
        print(f"Resolution: \n{ticket_row['Resolution'].values[0]}")
        print(f"Follow Up: \n{ticket_row['Follow Up'].values[0]}")
    else:
        print(f"Ticket with number {ticket_number} not found.")



        # mostly to be used to delete test tickets, or erroniously entered issues.(use with caution)
def delete_issue(ticket_number):
    try:
        # Check if the ticket number exists in the DataFrame
        if ticket_number in df['Number'].values:
            # Locate the index of the row with the specified ticket number
            index = df[df['Number'] == ticket_number].index[0]
            # Drop the row based on the index
            df.drop(index, inplace=True)
            # Reset the index
            df.reset_index(drop=True, inplace=True)
            # Save the updated DataFrame to the Excel file
            df.to_excel(data_file, index=False)
            print(f"Issue with ticket number {ticket_number} has been deleted.")
        else:
            print(f"Ticket with number {ticket_number} not found.")
    except Exception as e:
        print(str(e))



# Main loop
while True:
    print("\nIssue Tracker Menu:")
    print("1. Add Ticket")
    print("2. View Tickets")
    print("3. Follow Up on a Ticket (by number)")
    print("4. Delete a ticket (by number)")
    print("5. Print a ticket (by number)")
    print("6. Search by Tech Name")
    print("7. Quit")
    
    choice = input("Enter your choice: ")
    
    if choice == '1':
        tech = input("Enter tech's name: ")
        date = input("Enter the date (e.g., YYYY-MM-DD): ")
        issue = input("Enter the issue: ")
        resolution = input("Enter the resolution: ")
        follow_up = input ("Enter follow up information (if any): ")
        add_issue(tech, date, issue, resolution, follow_up)
        print("Issue added successfully!")
    elif choice == '2':
        view_issues()
    elif choice == '3':
        follow_up_issue()
    elif choice == '4':
        ticket_number = int(input("Enter the ticket number to delete: "))
        delete_issue(ticket_number)
    elif choice == '5':
        ticket_number = int(input("Enter the ticket number to print: "))
        print_ticket_details(ticket_number)
    elif choice == '6':
        tech_name = input("Enter tech's name to search: ")
        search_by_tech(tech_name)
    elif choice == '7':
        print("Goodbye!")
        break
    else:
        print("Invalid choice. Please try again.")



        # consider a function that will let you switch sheets in the workbook.  
        # this function could allow you to have a choice to switch to "notes sheet" or sheets specific to troubleshooting different issues
