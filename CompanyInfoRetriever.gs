function CompanyInfoRetriever() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow-1, 1).getValues(); // Starting from the 2nd row
  var MAX_RETRIES = 5;
  var RETRY_DELAY = 1000; // milliseconds
  
  for (var i = 0; i < data.length; i++) {
    var companyName = data[i][0];
    var companyWebsiteUrl;
    var companyLogoUrl;
    var domain;
    var emailProvider;

    if (companyName) {
      
      // Use the company name to get the website and logo from Clearbit Autocomplete API
      var apiUrl = "https://autocomplete.clearbit.com/v1/companies/suggest?query=" + encodeURIComponent(companyName);
      var apiResponse = UrlFetchApp.fetch(apiUrl);
      var apiResponseJson = JSON.parse(apiResponse.getContentText());
      companyWebsiteUrl = apiResponseJson[0].domain;
      companyLogoUrl = apiResponseJson[0].logo;
      domain = companyWebsiteUrl;
    } else {
      companyWebsiteUrl = "Invalid Input";
      companyLogoUrl = "Invalid Input";
      domain = "Invalid Input";
    }
    
    // Write the company website URL and logo URL to the second and third column
    sheet.getRange(i+2, 2).setValue(companyWebsiteUrl);
    sheet.getRange(i+2, 3).setFormula('=image("'+companyLogoUrl+'")');
    

    // Use the provided company domain to get email provider either Google, Office 365 or othre
    var retries = 0;
    while (retries < MAX_RETRIES) {
      var url = "https://dns.google.com/resolve?name=%FQDN%&type=MX".replace("%FQDN%",domain);
      Utilities.sleep(100);
      var result = UrlFetchApp.fetch(url,{muteHttpExceptions:true});
      var rc = result.getResponseCode();
      if (rc !== 200) {
        retries++;
        Utilities.sleep(RETRY_DELAY);
        continue;
      }
      var response = JSON.parse(result.getContentText());
      if(response.status == "ERROR"){
        retries++;
        Utilities.sleep(RETRY_DELAY);
        continue;
      }
      if (response.Answer[0].data == null) {
        var mxRaw = response.Authority[0].data; 
      } else {
        var mxRaw = response.Answer[0].data;
      }
      var mx = mxRaw.toLowerCase();
      if (mx.indexOf("google.com") >= 0 || mx.indexOf("googlemail.com") >= 0)  { 
        emailProvider = "Google Workspace"; 
      }
      else if (mx.indexOf("outlook.com") >= 0) { 
        emailProvider = "Office 365"; 
      }
      else emailProvider = "Other";
      
      // Write the company email provider to the 4th column
      
      sheet.getRange(i+2, 4).setValue(emailProvider);
      break;
    }
    if(retries == MAX_RETRIES) {
      sheet.getRange(i+2, 4).setValue("ERROR");
    }

      // Use the provided company domain to check if there is a Goolge login page for this domain

    var googleLoginUrl = "https://mail.google.com/a/" + domain;
    var retries = 0;
    while (retries < MAX_RETRIES) {
      var result = UrlFetchApp.fetch(googleLoginUrl);
      var contentText = result.getContentText();
      // Write the result to the 5th column
      if(contentText.indexOf("Server error") !=-1){
        sheet.getRange(i+2, 5).setValue("Does not exist");
        break;
      }else{
        sheet.getRange(i+2, 5).setValue("Exists");
        break;
      }
    }
    if (retries == MAX_RETRIES) {
      sheet.getRange(i+2, 5).setValue("Error");
    }
  }
}

