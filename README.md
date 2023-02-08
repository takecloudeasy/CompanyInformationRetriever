# CompanyInformationRetriever

This a Google Apps Script created to help Google Cloud Partners to retrieves any company information based on the company name. The company name is taken from a Google Sheet and the other details are filled in by the script. 

The details retrieved by the script include the company website URL, the company logo URL, the company's email provider (Google Workspace, Office 365 or Other), and check whether the company has a Google login page. The data is retrieved from the Clearbit Autocomplete API and Google's MX DNS service. The data is then written back to the Google Sheet in the corresponding columns.
