import win32com.client
import os
import time
import datetime

# Run query

# Access Excel
x1 = win32com.client.DispatchEx('Excel.Application')

# Open wb
wb = x1.workbooks.open(r"") # Enter filepath in inverted commas here
x1.Visible = False

# Refresh wb
wb.RefreshAll()

# Wait for 20 seconds to give file time to refresh
time.sleep(20)

# Find specific date to enter in filename (e.g. previous Sunday)
today = datetime.date.today()
Sun = today - datetime.timedelta(5)
Sun_Formatted = datetime.date.today().strftime("%Y") + datetime.date.today().strftime("%m") + (Sun).strftime("%d")


# Create filename and specify path (also adds date to end of filename)
path = r"" # Enter filepath inverted commas here
wbName = "" + str(Sun_Formatted) # Enter filename in inverted commas here 

# Save new report 
wb.SaveAs(os.path.join(path, wbName))

# Send email that report is ready
outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "" # Enter email address here
mail.Subject = "" # Enter email subject here
mail.HTMLBody = (r"""The Report is ready to be sent:<br>
     <a href='Filepath'> 
     Filepath</a><br><br>""") # Replace 'Filepath' wtih actual filepath 
mail.Send()


wb.Close()
x1.Quit()
