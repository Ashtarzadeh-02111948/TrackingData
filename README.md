I needed a script that would check the tracking status of every tracking number that is inputted in my Google Sheets file.
Google Sheets offers scripting support that can interact with spreadsheets, so this is what I decided to use since it
would be the simplest way of achieving what I wanted. 

Almost all of the tracking I would receive would be from USPS, so I needed access to USPS APIs which I registered for
and was granted access to. Tracking numbers from any other carrier would not work since I currently do not have
access to their APIs. 

The script would go through the column that the tracking numbers are located in the spreadsheet, and check the status
for each tracking number. Then, the script will parse out a phrase from USPS API, that indicates the tracking status, 
and it would then paste that in a dedicated column called 'Tracking Status'. 
And the script will look for keywords in parsed phrases such as: 'delivered', 'departed usps', 'arrived at usps', 'action needed'

There would then be four general tracking statuses that the script will assign depending on the keywords.

These are delivered, delivering, pending, and pick up. There are two exception cases outside this if the tracking status contains 'invalid' or 'alert'.    

A color would also be assigned to each of these general statuses that would color code the tracking number cells. 
Green for delivered, light blue for delivering, yellow for pending, orange for pick up, red for tracking status containing 'alert'.
Yellow will also be assigned to invalid tracking numbers (those that arent USPS).

I also added regular expressions early on that would determine what type of carrier the tracking number is for, whether it was
FedEx, USPS, or UPS. This would allow for future support if I had access to those other carrier's APIs. As of now, the regular
expression function is mostly redundant. 

The script will run once the 'Check Tracking Status' button is clicked under 'Custom Menu' in the top toolbar on Google Sheets.






