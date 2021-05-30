function productsDelete() {
  var currentDate = new Date();
  currentDate.setUTCHours(currentDate.getHours() - 24);
  var formattedMonthYear = Utilities.formatDate(currentDate, "GMT+8", "MMMM yyyy");
  var formattedDateSlash = Utilities.formatDate(currentDate, "GMT+8", "dd/MM/yy");
  var day = currentDate.getDate().toLocaleString();

  var ck = "YOUR_API_KEY";
  var cs = "YOUR_API_SECRET";
  var website = "https://YOURWEBSITE.com/wp-json/wc/v3";
  var url = website + "/products?consumer_key=" + ck + "&consumer_secret=" + cs + "&search=" + day + "&per_page=50";

  var options =
    {
      "method": "GET",
      "muteHttpExceptions": true,
    };
  
  var result = UrlFetchApp.fetch(url, options);
  Logger.log(result.getResponseCode());
  if (result.getResponseCode() == 200) {
    var params = JSON.parse(result.getContentText());
  };

  var loopLength = params.length;

  for (var i = 0; i < loopLength; i++){
    var name = params[i]["name"].toLocaleString();
    // extract different date formats from product name
    var monthYear = name.match(/\w+\s[0-9]+$/g); // e.g. April 2021
    var dateSlash = name.match( /[0-9]+\/[0-9]+\/[0-9]+$/g); // e.g. 30/04/21
    var startDate = name.match(/[0-9]+(?=\-[0-9]+)/g); // e.g. 30
    // specifically for product names with two months
    var year = name.match(/[0-9]+$/g); // e.g. 2021
    var doubleMonth = name.match(/[a-zA-Z]+(?=\s\-)/g); // e.g. April
    var doubleDate = doubleMonth + " " + year; // e.g. April 2021
    var doubleStartDate = name.match(/[0-9]+(?=\s[a-zA-Z]+\s\-)/g); // e.g. 30
    

    // check if start date match or date with dd/MM/yy format match
    if (startDate == day | dateSlash == formattedDateSlash | doubleStartDate == day){
      // check if month & year match or date with dd/MM/yy format match
      if (monthYear == formattedMonthYear | dateSlash == formattedDateSlash | doubleDate == formattedMonthYear){
        // get the product id
        var productID = params[i]["id"];
        // create url and delete product
        var url = website + "/products/" + productID + "?force=true&consumer_key=" + ck + "&consumer_secret=" + cs;
        var options =
          {
            "method": "DELETE",
            "muteHttpExceptions": true,
          };
        
        var result = UrlFetchApp.fetch(url, options);
        Logger.log(result.getResponseCode());
      } else{
        continue;
      }
    } else{
      continue;
    };
  };
}
