# MSAL JavaScript Sample Application with Bookings API

The purpose of this repo is to provide an example of how to call the Microsoft Bookings API.

The Microsoft Bookings API calls in this example were added to JavaScriptSPA/ui.js. 

The code will need to be modified for Production use. Microsoft's active-directory-javascript-graphapi-v2 tutorial was used for the example since it provides an excellent foundation for how to call the Graph API. 

Here is the link to the original code from Microsoft's Azure-Samples/active-directory-javascript-graphapi-v2 repo:  https://github.com/Azure-Samples/active-directory-javascript-graphapi-v2

Here is the link to the Bookings API documentation:
https://docs.microsoft.com/en-us/graph/api/resources/booking-api-overview?view=graph-rest-beta



![](https://raw.githubusercontent.com/shah2129/active-directory-javascript-graphapi-v2/quickstart/screenshot_showing_ui.png)
 
## Most of the README content below not been changed from the orginal active-directory-javascript-graphapi-v2 tutorial. Additonal updates will need to be made to this README.

 
---
page_type: sample
languages:
- javascript
products:
- microsoft-identity-platform
- azure-active-directory-v2
- microsoft-graph-api
description: "A simple JavaScript single-page application calling Microsoft Graph API using msal.js (w/ AAD v2 endpoint)"
urlFragment: "active-directory-javascript-graphapi-v2"
---

-------------------





A simple vanilla JavaScript single-page application which demonstrates how to configure [MSAL.JS Core](https://www.npmjs.com/package/msal) to login, logout, protect a route, and acquire an access token for a protected resource such as [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview).

**Note:** A detailed tutorial covering this sample can be found [here](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-v2-javascript).

## Contents

| File/folder       | Description                                |
|-------------------|--------------------------------------------|
| `AppCreationScripts`   | Contains automation scripts for Powershell users (can be safely removed if desired).|
| `JavaScriptSPA`   | Contains sample source files.  |
| `auth.js`   | Main application logic resides here.                     |
| `config.js`   | Contains configuration parameters for the sample. |
| `graph.js`   | Provides a helper function for calling MS Graph API.   |
| `index.html`   |  Contains the UI of the sample.                       |
| `.gitignore`      | Defines what to ignore at commit time.      |
| `CHANGELOG.md`    | List of changes to the sample.             |
| `CODE_OF_CONDUCT.md` | Code of Conduct information.            |
| `CONTRIBUTING.md` | Guidelines for contributing to the sample. |
| `LICENSE`         | The license for the sample.                |
| `package.json`    | Package manifest for npm.                   |
| `README.md`       | This README file.                          |
| `SECURITY.md`     | Security disclosures.                      |
| `server.js`     | Implements a simple Node server to serve index.html.  |

## Prerequisites

[Node](https://nodejs.org/en/) must be installed to run this sample.

## Setup

1. [Register a new application](https://docs.microsoft.com/azure/active-directory/develop/scenario-spa-app-registration) in the [Azure Portal](https://portal.azure.com). Ensure that the application is enabled for the [implicit flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-implicit-grant-flow).
2. Open the [/JavaScriptSPA/config.js](./JavaScriptSPA/config.js) file and provide the required configuration values.
3. On the command line, navigate to the root of the repository, and run `npm install` to install the project dependencies via npm.

## Running the sample

1. Configure authentication and authorization parameters:
   1. Open `authConfig.js`
   2. Replace the string `"Enter_the_Application_Id_Here"` with your app/client ID on AAD Portal.
   3. Replace the string `"Enter_the_Cloud_Instance_Id_HereEnter_the_Tenant_Info_Here"` with `"https://login.microsoftonline.com/common/"` (*note*: This is for multi-tenant applications located on the global Azure cloud. For more information, see the[documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-v2-javascript)).
   4. Replace the string `"Enter_the_Redirect_Uri_Here"` with the redirect uri you setup on AAD Portal.
2. Configure the parameters for calling MS Graph API:
   1. Open `graphConfig.js`.
   2. Replace the string `"Enter_the_Graph_Endpoint_Herev1.0/me"` with `"https://graph.microsoft.com/v1.0/me"`.
   3. Replace the string `"Enter_the_Graph_Endpoint_Herev1.0/me/messages"` with `"https://graph.microsoft.com/v1.0/me/messages"`.
3. To start the sample application, run `npm start`.
4. Next, open a browser to [http://localhost:30662](http://localhost:30662).

## Key points

This sample demonstrates the following MSAL workflows:

* How to configure application parameters.
* How to sign-in with popup and redirect methods.
* How to sign-out.
* How to get user consent incrementally.
* How to acquire an access token.
* How to make an API call with the access token.

## Contributing

If you'd like to contribute to this sample, see [CONTRIBUTING.MD](./CONTRIBUTING.md).

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
