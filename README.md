# TMS_CDR_Analysis
The scripts leverage TMS as source of information to calculate the number of concurrent calls give a CDR report. This CDR can comes from section such as Gateways, Gatekeepers and VCS or MCUs and Endpoints. 
The calls to considered can be filtered thanks to Regular Expressions (RegEx) and call minimum duration.
The result of the script will be a graphic with concurrent calls based on filter criteria and a Excel file (in the number of entries allows) with the number of concurrent calls every time a call starts or ends.

The program makes use of the following third party modules:
* appdirs==1.4.0
* cycler==0.10.0
* et-xmlfile==1.0.1
* jdcal==1.3
* matplotlib==2.0.0
* numpy==1.12.0
* openpyxl==2.4.2
* packaging==16.8
* pandas==0.19.2
* pyparsing==2.1.10
* python-dateutil==2.6.0
* pytz==2016.10
* scipy==0.18.1
* seaborn==0.7.1
* six==1.10.0
