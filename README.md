# Microsoft Graph sample for creating an O365 Excel  from CSV

## Introduction

This sample shows how to connect your .NET Console app to Office 365 using the Microsoft Graph API and generate an Excel workbook from CSV. It uses [Microsoft Graph .NET Client Library](https://github.com/microsoftgraph/msgraph-sdk-dotnet) as well as Azure AD v2.0 endpoint, which enables users to sign-in with either their personal or work or school Microsoft accounts.

## Prerequisites ##

This sample requires the following:  

  * [Visual Studio 2017](https://www.visualstudio.com/en-us/downloads) 
  * Either a [Microsoft](www.outlook.com) or [Office 365 for business account](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment#bk_Office365Account).

## Register and configure the app

1. Sign into the [App Registration Portal](https://apps.dev.microsoft.com/) using either your personal or work or school account.
2. Select **Add an app**.
3. Enter a name for the app, and select **Create application**.
4. Under **Platforms**, select **Add platform**.
5. Select **Native Application**.
6. Copy the Application Id value and replace it in ``AuthenticationHelper.ClientId``.
7. Select **Save**.

### Code

Most code is in ``Main()`` method, and is relatively straight-forward and self-explanatory. Calls for authentication actually occur in the ``AuthenticationHelper`` class.

## Additional resources ##

- [Microsoft Graph overview](http://graph.microsoft.io)
- [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) (Allows you to sign-in with your own O365 acount and query Graph APIs)
- [Working with Excel in Microsoft Graph](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/excel)