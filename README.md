# TMS_CDR_Analysis
The scripts leverage TMS as source of information to calculate the number of concurrent calls give a CDR report. This CDR can comes from section such as Gateways, Gatekeepers and VCS or MCUs and Endpoints. 
The calls to considered can be filtered thanks to Regular Expressions (RegEx) and call minimum duration.
The result of the script will be a graphic with concurrent calls based on filter criteria and a Excel file (in the number of entries allows) with the number of concurrent calls every time a call starts or ends.
