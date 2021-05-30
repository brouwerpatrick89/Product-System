function updateCategories() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Dropdown");
  var range = sheet.getRange("A:A");
  range.clearContent();

  getCategories()
}

function getCategories() {
  var ck = "YOUR_API_KEY";
  var cs = "YOUR_API_SECRET";
  var website = "https://YOURWEBSITE.com/wp-json/wc/v3";
  var url = website + "/products/categories?consumer_key=" + ck + "&consumer_secret=" + cs + "&per_page=50"; 
  
  var options =
    {
      "method": "GET",
      "Content-Type": "application/json",
      "muteHttpExceptions": true,
    };
  
  var result = UrlFetchApp.fetch(url, options);
  Logger.log(result.getResponseCode());
  if (result.getResponseCode() == 200) {
    var params = JSON.parse(result.getContentText());
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Dropdown");
  var arrayLength = params.length;

  for (var i = 0; i < arrayLength; i++) {
    if (params[i]["parent"] == 0 || params[i]["parent"] == 50) {
      continue;
    }
    var listName = params[i]["name"];
    sheet.appendRow([listName]);
  }
}
