**IMPORTANT**: this sample code is provided as is.

# QnA - A sample Microsoft Teams Meeting App

The QnA Meeting App is an example of how the existing Teams meeting extensibility points can be used to enable users, from different roles, to interact with meetings questions. The app leverages all existing meeting extensibility points, and provides different experiences according to the user role and also the meeting state, whether it is active or not. In addition, this app also connects to Microsoft Graph, based on a resource account, to expose a Power BI Report on meeting questions.


## YouTube Video

[![Apps in Teams Meeting YouTube Video](https://github.com/marcoszanre/meeting-apps-qna/blob/master/assets/githubImage8.png)](https://youtu.be/dH2OVc7D81A)


## How does it work?

When the user connects to the QnA app, the tab, which leverages Single Sign On, will authenticate the user and retrieve its current role, which could be of a presenter, organizer or attendee. After that, the tab will route the UI according to the user role and also the frameContext, which determines if the app is being opened from the Meetings Details page, side panel or as a task content.

A custom Meeting State manager has been used, also illustrating how additional variables can be configured, and updated throughout the different meeting phases, and this value is also leveraged to understand what type of experience and content should be presented to the user, again, according to its role and frameContext, before, during and after the meeting.

The available meeting actions are available as below:

Before a meeting (Details and Chat pane):
- Attendees can submit, edit and delete their own questions from the Details and Chat pane
- Organizers and presenters can promote/demote and delete not appropriate questions

During a meeting (Side Panel and Task Content pane):
- Attendees can review and vote on promoted questions
- Organizers and presenters can promote questions to all users, which will be displayed in a task pane

After a meeting (Details and Chat pane - called when organizer or presenter has closed the meeting):
- Attendees can only view a message stating that the meeting is closed
- Organizers and presenters can review and manage existing questions, review the Power BI report with a meeting summary and also download a CSV file with existing questions asked. 

A Power BI report has been configured, which leverages a resource account, and is exposed through the Embedded React Power BI Component, to allow organizers and presenters to review meeting questions. This Power BI report is published to a Teams workspace and is needs to be configured to be updated automatically.

In addition, the storage being used also stores audit information, which could be used to track when a user, from the different available roles, took a specific action.


## Components

The QnA Meeting App is comprised of the following components:

* A `Tab`, built with React, and initialized with the Yeoman Teams Generator, which os configured to use Single-Sign On, and is published to the store and needs to be installed manually in a Teams meeting by one of the authorized users.
* A `Bot registration`, which is used to retrieve the token and send the task content notification inside the meeting to all users
* An `Azure Table Storage`, to store the questions, meeting state, audit actions and like references of the app.
* An `Azure AD app registration`, to be able to configure Single-Sign On from the tab, and also expose the Power BI Rest API permissions for the embedded React component.

The QnA Meeting App can be run locally as well as from an Azure App Service (configured by default to Node running Linux).


## Enviroment variables

The following environment variables, including the ones already created by the Teams Yeoman Generator, have been used in this application and are provided in the sample_env file (these values need to be configured in Azure App Service):

STORAGE ACCOUNT
* `STORAGE_ACCOUNT_NAME=` Azure Storage account name
* `STORAGE_ACCOUNT_ACCESSKEY=` Azure Storage account access key

BOT SERVICE
* `MICROSOFT_APP_ID=` Bot app id credentials
* `MICROSOFT_APP_PASSWORD=` Bot app secret credentials
* `TENANT_ID=` Tenant identifier to retrieve bot tokens

TEAMS CATALOG REGISTRATION
* `APPLICATION_ID=` Unique ID of the app to be published to the Teams tenant catalog

POWER BI SERVICE
* `POWER_BI_USER=` Azure AD user with an applicable Power BI license, as well as workspace permission to access the report
* `POWER_BI_PASSWORD=` Azure AD user password


## How to run it locally

In order to run this repository locally, follow the steps below:

* Download a copy of this repository
* Run `ngrok` locally exposing the desired port (Yeoman Teams generator uses 3007, but that will be based on your `PORT` environment variable)
* Create in Azure a Bot Channel Registration and enable the Teams channel, and replace the endpoint with the ngrok address
* Create in Azure a Storage Account
* Customize the Power BI Reported provided in the assets folder, to point to your storage account, upload that to your workspace, retrieve its embedded url through any of the available options (e.g. through Docs get reports in a group try it experience - https://docs.microsoft.com/en-us/rest/api/power-bi/reports/getreportsingroup) and update the Organizer.tsx Power BI React Component embedUrl, id and page name value (lines 426, 432, 436) to point to your appropriate values (also make sure that the user that will use access this report is part of this group)
* Create a new Azure AD Application Registration as detailed in the Teams SSO documentation
* Assign the permissions specified in the Teams SSO documentation, as well as the Power BI Service Report.Read.All, and proceed to provide the admin consent to the application (or manual consent, if applicable)
* Upload the manifest available in the assets folder to App Studio to edit its values to point to the `ngrok` address as well as the Azure AD Application registration values
* Rename the sample_env file with .env and fill out required variables, which are: `HOSTNAME`, `MICROSOFT_APP_ID`, `MICROSOFT_APP_PASSWORD`, `STORAGE_ACCOUNT_NAME`, `STORAGE_ACCOUNT_ACCESSKEY`, `TENANT_ID`, `POWER_BI_USER` and `POWER_BI_PASSWORD`)
* Initiate the app running `gulp serve`, and make sure that all services (Table and React webserver) have initiated correctly
* Create a new meeting and, as an organizer, click on the `+` icon to add a new meeting. Click on `manage apps`, and on the bottom right hand corner, if sideload is enabled, click on `Upload a custom app`. It is also recommended to modify the roles, such as having a user as an attendee to be able to review the different app experiences.


## How to deploy it to Azure

In order to deploy this app to Azure, follow the steps below:

* Create a new Azure App Service running Node and Linux
* From Visual Studio Code, with the root project folder opened, open the Azure App Service extension, right click on your provisioned app service and click on `Configure Deployment source` and select LocalGit
* Still from Visual Studio Code, with the root project folder opened, open the Azure App Service extension, right click on your provisioned app service and click on `Deploy to Web App` (it should take from 1 to 3 minutes for the process to complete). In addition, if for any reason, the deployment fails, try to open the app service from the browser and retry the deployment from the deployment center.
* Update your Azure AD App registration with the new domain values from the app service for Single-Sign On
* Update your App Manifest with the new values from the app service
* Finish your manifest, sideload and submit that to the tenant app catalog.


## To Do

The following scenarios, though important, have not been implemented yet in this demo:

* Bot graph token management service to reuse existing valid tokens for the bot, rather than asking for a new one upon each request
* Graph Permissions to display question's author picture in the organizer and attendee in-meeting side panel app experience
* Complement Audit table action implementation to populate all fields
* Refactor the code to be more reusable towards React best practices and remove not relevant code
* Configure Express API routes to process requests, authorizing on the Graph Token as well as the role of the user making the call
* Configure a deep link for the organizer list items, to be able to initiate the chat. Also to achieve this, it is necessary to configure some routes to retrieve the UPN of the user, based on his Azure AD user id, which is what is needed for the deep link to work.
* Optionally, configure a Graph Notification service, to be notified whenever a meeting has ended, to update automatically the meeting state, without requiring the organizer to click on the `Close Meeting` action.
* Optionally, provide an initial recording field, where the organizer can provide when the meeting recorded started, and, as part of the download meeting questions, provide this calculated field (when question promoted time - when meeting recorded initiated) to allow users to understand when a particular question has been asked during a meeting.
* Optionally, configure the Power BI Report, to display audit actions as well as any other relevant content.


## Extension opportunities

As the concept of a meeting entity object, in this case, questions, has been used in this sample, try to understand what other entities, which make sense to your organization or to a particular industry or scenario, could also be used, such as a meeting:

* Agenda
* Statements
* Items to review
* Announcements
* Kudos

It is also very important to plan what actions will be available for each user role as well as consider any other extended meeting variables, like the cusom meeting state has been used in this demo.

## Demo screenshots

1. Before meeting organizer/presenter view of existing questions:

![Before meeting organizer/presenter view of existing questions screenshot](https://github.com/marcoszanre/meeting-apps-qna/blob/master/assets/githubImage1.png)


2. Before meeting attendee view of existing questions:

![Before meeting attendee view of existing questions screenshot](https://github.com/marcoszanre/meeting-apps-qna/blob/master/assets/githubImage2.png)


3. In-meeting organizer/presenter view of existing questions:

![In-meeting organizer/presenter view of existing questions screenshot](https://github.com/marcoszanre/meeting-apps-qna/blob/master/assets/githubImage3.png)


4. In-meeting organizer/presenter promoted questions task content:

![In-meeting organizer/presenter promoted questions task content screenshot](https://github.com/marcoszanre/meeting-apps-qna/blob/master/assets/githubImage4.png)


5. In-meeting attendee view of existing questions:

![In-meeting attendee view of existing questions screenshot](https://github.com/marcoszanre/meeting-apps-qna/blob/master/assets/githubImage5.png)


6. After meeting organizer/presenter view of questions reports:

![After meeting organizer/presenter view of questions reports screenshot](https://github.com/marcoszanre/meeting-apps-qna/blob/master/assets/githubImage6.png)


7. After meeting attendee view of closed meeting message:

![After meeting attendee view of closed meeting message screenshot](https://github.com/marcoszanre/meeting-apps-qna/blob/master/assets/githubImage7.png)


## References

Following are some of the references used in this project:

* [ngrok](https://ngrok.io)
* [Microsoft Teams official documentation](https://developer.microsoft.com/en-us/microsoft-teams)
* [Microsoft Teams Yeoman generator Wiki](https://github.com/PnP/generator-teams/wiki)
* [Create Apps for Teams Meeting](https://docs.microsoft.com/en-us/microsoftteams/platform/apps-in-teams-meetings/create-apps-for-teams-meetings?WT.mc_id=M365-MVP-5001530&tabs=json)
* [Power BI Get Reports try it experience for retrieving Embedded values](https://docs.microsoft.com/en-us/rest/api/power-bi/reports/getreportsingroup)
* [Power BI Client React](https://github.com/microsoft/powerbi-client-react)
* [Azure Table Storage Node reference](https://docs.microsoft.com/en-us/azure/cosmos-db/table-storage-how-to-use-nodejs)
* [YouTube presentation and demo](https://youtu.be/dH2OVc7D81A)


If you have any questions/suggestions, feel free to share them, **thanks**!