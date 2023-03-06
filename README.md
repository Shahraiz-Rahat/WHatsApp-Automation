# WHatsApp-Automation
The project function is to access the google sheet and read write data in the specified sheets. it is a complete opreation tool where one can send automated messages.


This is the basic flow

1. Start reading rows from google sheet
2. Read name, phone number
3. Open whatsapp web
4. Search for name and phone number
5. IF Contact exists send message
- Update status column in google sheet - Status = sent , Sent Date = Date
6. IF Contact does not exist
- Update status column in google sheet - Status = contact not found

Next time when script runs if status = sent already then dont send message again
