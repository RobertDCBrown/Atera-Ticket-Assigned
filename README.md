# Atera-Ticket-Assigned
Uses Atera's and Twilio's API's to send a text message to your technicians when they are assigned a ticket.

Example:
When assigning a ticket to a technician, use a status they tells the script to send a text message (ie: "Send to Tech")
The script will look for that status and then alert the technician assigned via text message. Once the text is sent, it will update the ticket to another status (ie. "Assigned")

#Setup
1. Upon first launch, complete your Atera and Twilio API information.
2. Set your two status's for Atera. (I use "Send to Tech" and "Assigned")
3. Click "Save" (Saved to %AppData%\Roaming\OnIT\AssignedAlerts)
