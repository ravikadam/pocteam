# pocteam
Team POC
Please clone this repository into folder. You need to use NodeJS 12 or more to work with this.

Three important files

AppSettings.js files stores clientID - this is same as ClientID which is defined in Slidedge Azure Portal for Trial purpose. It is not used in production code anywhere. I used it to ensure I do not break any working code.

Helper file contains all Graph API calls wrapped by Javascript functions. 

index.js file contains UI for showing options and then when user selects option it gives a call to appropriate function in helper file which in turn calls API and returns response with subset of fields which might be useful.

Once authentication is done then please select options in exact sequence such as 1 for showing Tokem 2, for getting list of teams, 3 for getting channels and 4 for getting folder for last channel. Sequence is important because channel needs Team ID which is stored in a variable when Show Teams option is selected, same way API for folder needs channel and team both.

For implementation in slidedge we can simply call these API one after another or we may call them in sequence giving option to user which channel he needs to select.

Ultimate objective of these API is to get WebURL for Folder. This folder needs to be passed as parameter in odoptions in OneDrive Picker. 

I have also added one folder which contains one HTML and javascript file where this folder is manually typed (in javascript file). This was just to test if OneDrive picker opens directly in that folder.  This folder is optional to use because it is anyway picked up from Slidedge OneDrive picker code and modified to provide additional parameter and just to run in standalone mode without Angular
