function getCategoryID() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ProductInfo");
  var categoryRange = sheet.getRange("B1");
  var category = categoryRange.getValue();
  
  var ck = "YOUR_API_KEY";
  var cs = "YOUR_API_SECRET";
  var website = "https://YOURWEBSITE.com/wp-json/wc/v3";
  var url = website + "/products/categories?consumer_key=" + ck + "&consumer_secret=" + cs + "&search=" + category; 
  
  var options =
    {
      "method": "GET",
      "contentType": "application/x-www-form-urlencoded;charset=UTF-8",
      "muteHttpExceptions": true,
    };
  
  var result = UrlFetchApp.fetch(url, options);
  Logger.log(result.getResponseCode());
  if (result.getResponseCode() == 200) {
    var params = JSON.parse(result.getContentText());
  }
  var categoryID = params[0]["id"];
  getProductInfo(categoryID);
}

function getProductInfo(categoryID) {
  var ck = "YOUR_API_KEY";
  var cs = "YOUR_API_SECRET";
  var website = "https://YOURWEBSITE.com/wp-json/wc/v3";
  var url = website + "/products?consumer_key=" + ck + "&consumer_secret=" + cs + "&category=" + categoryID + "&per_page=50";
  var options =
    {
      "method": "GET",
      "contentType": "application/x-www-form-urlencoded;charset=UTF-8",
      "muteHttpExceptions": true,
    };
  // get product data
  var result = UrlFetchApp.fetch(url, options);
  Logger.log(result.getResponseCode());
  if (result.getResponseCode() == 200) {
    var params = JSON.parse(result.getContentText());
  }
  
  if (params == "") {
    SpreadsheetApp.getUi().alert('No products in this category.');
  } else {
    var loopLength = params.length;
    var container = [];
    var count = 0;
    // loop through product data array
    for (var i = 0; i < loopLength; i++) {
      var id = params[i]["id"];
      var sku = params[i]["sku"];
      var name = params[i]["name"];
      var regularPrice = params[i]["regular_price"];
      var salePrice = params[i]["sale_price"];
      
      count += 1;
      container.push([id, sku, name, regularPrice, salePrice, ""]);
      
      var variations = params[i]["variations"];
      // if product has variations, retrieve variation data
      if (variations != "") {
        var variationURL = website + "/products/" + id + "/variations?consumer_key=" + ck + "&consumer_secret=" + cs;
        
        var variationResult = UrlFetchApp.fetch(variationURL, options);
        Logger.log(variationResult.getResponseCode());
        if (variationResult.getResponseCode() == 200) {
          var variationParams = JSON.parse(variationResult.getContentText());
        }

        var variationLoopLength = variationParams.length;

        for (var j = 0; j < variationLoopLength; j++) {
          var variationId = variationParams[j]["id"];
          var variationSku = variationParams[j]["sku"];
          var variationName = name + " - " + variationParams[j]["attributes"][0]["option"];
          var variationRegularPrice = variationParams[j]["regular_price"];
          var variationSalePrice = variationParams[j]["sale_price"];

          count += 1;
          container.push([variationId, variationSku, variationName, variationRegularPrice, variationSalePrice, id]);
        };
      } else {
        continue;
      };
    };
  };
    // get current product data in the sheet, delete it and input the newly requested product data
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ProductInfo");
    var lastRow = sheet.getLastRow();
    
    if (lastRow > 4) {
      var oldRange = sheet.getRange("A5:G" + lastRow)
      oldRange.clearContent();
      oldRange.clearFormat();
      oldRange.clearDataValidations();
    };
    
    if (params != ""){
      var inputRange = sheet.getRange(5,1,count,6);
      inputRange.setValues(container);
      inputRange.setBorder(true, true, true, true, true, true, "green", null);

      var checkboxRange = sheet.getRange(5,7,count,1);
      checkboxRange.insertCheckboxes();
      checkboxRange.setBorder(true, true, true, true, true, true, "green", null);
    };
};

function updateProduct() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ProductInfo");
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A5:F" + lastRow);
  var products = range.getValues();
  var ck = "YOUR_API_KEY";
  var cs = "YOUR_API_SECRET";
  var website = "https://YOURWEBSITE.com/wp-json/wc/v3";
  
  if (products.length > 1) {
    var loopLength = products.length;
    var data = {"update":[]};

    for (var j = 0; j < loopLength; j++) {
      // if parent id is not empty, update variation
      if (products[j][5] != "") {
        //variation options
        var variationUrl = website + "/products/" + products[j][5] + "/variations/" + products[j][0] + "?consumer_key=" + ck + "&consumer_secret=" + cs;

        var variationData = {
          "sku": products[j][1],
          "name": products[j][2],
          "regular_price": products[j][3].toString(),
          "sale_price": products[j][4].toString()
        };

        var variationOptions =
          {
            "method": "PUT",
            "contentType": "application/json",
            "muteHttpExceptions": true,
            "payload": JSON.stringify(variationData)
          };
        
        var result = UrlFetchApp.fetch(variationUrl, variationOptions);
        Logger.log(result.getResponseCode());
      }
      
      data.update[j] = {
        "id": products[j][0],
        "sku": products[j][1],
        "name": products[j][2],
        "regular_price": products[j][3].toString(),
        "sale_price": products[j][4].toString()
      };
    };
    
    var url = website + "/products/batch?consumer_key=" + ck + "&consumer_secret=" + cs;
    var options =
      {
        "method": "POST",
        "contentType": "application/json",
        "muteHttpExceptions": true,
        "payload": JSON.stringify(data)
      };
    
    var result = UrlFetchApp.fetch(url, options);
    Logger.log(result.getResponseCode());
  } else {
    var productID = products[0][0];
    var data = {
          "sku": products[0][1],
          "name": products[0][2],
          "regular_price": products[0][3].toString(),
          "sale_price": products[0][4].toString()
        };
    
    var url = website + "/products/" + productID + "?consumer_key=" + ck + "&consumer_secret=" + cs;
    var options =
      {
        "method": "PUT",
        "contentType": "application/json",
        "muteHttpExceptions": true,
        "payload": JSON.stringify(data)
      };
    
    var result = UrlFetchApp.fetch(url, options);
    Logger.log(result.getResponseCode());
  };
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Products updated. Press \'OK\' to refresh list.');
  getCategoryID()
}

function deleteProduct() {
  // get range values
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ProductInfo");
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A5:G" + lastRow);
  var products = range.getValues();
  var ck = "YOUR_API_KEY";
  var cs = "YOUR_API_SECRET";
  var website = "https://YOURWEBSITE.com/wp-json/wc/v3/products";

  // for every row in range: check if checkbox is checked
  var loopLength = products.length;

  for (var k = 0; k < loopLength; k++) {
    if (products[k][6] == true) {
      // if checkbox is checked: check if normal product or product variation (family camp)
      if (products[k][5] != "") {
        // get the correct 'ids'
        var productID = products[k][5];
        var variationID = products[k][0];

        // proceed to delete using correct endpoint
        var url = website + "/" + productID + "/variations/" + variationID + "?force=true&consumer_key=" + ck + "&consumer_secret=" + cs;
        var options =
          {
            "method": "DELETE",
            "muteHttpExceptions": true,
          };
        
        var result = UrlFetchApp.fetch(url, options);
        Logger.log(result.getResponseCode());
      } else {
        // get the correct 'id'
        var productID = products[k][0];

        // proceed to delete using correct endpoint
        var url = website + "/" + productID + "?force=true&consumer_key=" + ck + "&consumer_secret=" + cs;
        var options =
          {
            "method": "DELETE",
            "muteHttpExceptions": true,
          };
        
        var result = UrlFetchApp.fetch(url, options);
        Logger.log(result.getResponseCode());
      };
    };
  };
  
  // display alert when finished
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Products deleted. Press \'OK\' to refresh list.');

  // refresh list
  getCategoryID()
}
