function GeneratePages() {
  var ss = SpreadsheetApp.getActive();
  var inputSheet = ss.getSheetByName("InputSheet");
  var allCells = inputSheet.getRange("A1:A" + inputSheet.getLastRow()).getValues();
  var numberOfValues = allCells.filter(String).length;
  var URL_list = inputSheet.getRange(2, 1, numberOfValues - 1).getValues().flat();

  inputSheet.getRange('B1').setValue("Status");

  var valid_URL_list = [];

  // Check status for each URL and mark in sheet
  URL_list.forEach(function(url, index) {
    var cellPos = index + 2; // because sheet starts at row 2 for URLs
    var status = getStatusCode(url);

    if (status == 200 || status == 301) {
      inputSheet.getRange('B' + cellPos).setFontColor('black');
      inputSheet.getRange('B' + cellPos).setValue(status);
      valid_URL_list.push(url);
    } else {
      inputSheet.getRange('B' + cellPos).setFontColor('red');
      inputSheet.getRange('B' + cellPos).setValue(status);
    }
  });

  // Prepare Mobile sheet with all headers
  var mobileSheet = ss.getSheetByName('Mobile');
  if (!mobileSheet) {
    mobileSheet = ss.insertSheet("Mobile");
  }
  mobileSheet.getRange('A1').setValue("URL");
  mobileSheet.getRange('B1').setValue("Score");
  mobileSheet.getRange('C1').setValue("First Contentful Paint");
  mobileSheet.getRange('D1').setValue("Speed Index");
  mobileSheet.getRange('E1').setValue("Largest Contentful Paint");
  mobileSheet.getRange('F1').setValue("Time to Interactive");
  mobileSheet.getRange('G1').setValue("Total Blocking Time");
  mobileSheet.getRange('H1').setValue("Cumulative Layout Shift");

  // Prepare Desktop sheet with all headers
  var desktopSheet = ss.getSheetByName('Desktop');
  if (!desktopSheet) {
    desktopSheet = ss.insertSheet("Desktop");
  }
  desktopSheet.getRange('A1').setValue("URL");
  desktopSheet.getRange('B1').setValue("Score");
  desktopSheet.getRange('C1').setValue("First Contentful Paint");
  desktopSheet.getRange('D1').setValue("Speed Index");
  desktopSheet.getRange('E1').setValue("Largest Contentful Paint");
  desktopSheet.getRange('F1').setValue("Time to Interactive");
  desktopSheet.getRange('G1').setValue("Total Blocking Time");
  desktopSheet.getRange('H1').setValue("Cumulative Layout Shift");

  // Fetch PageSpeed Data for valid URLs and populate sheets
  valid_URL_list.forEach(function(url, index) {
    var pos = index + 2; // row in sheet
    try {
      var mobileData = fetchDataFromPSI('mobile', url);
      mobileSheet.getRange('A' + pos).setFontColor("black");
      mobileSheet.getRange('A' + pos).setValue(mobileData.url);
      mobileSheet.getRange('B' + pos).setValue(mobileData.score);
      mobileSheet.getRange('C' + pos).setValue(mobileData.firstContentfulPaint);
      mobileSheet.getRange('D' + pos).setValue(mobileData.speedIndex);
      mobileSheet.getRange('E' + pos).setValue(mobileData.largestContentfulPaint);
      mobileSheet.getRange('F' + pos).setValue(mobileData.timeToInteractive);
      mobileSheet.getRange('G' + pos).setValue(mobileData.totalBlockingTime);
      mobileSheet.getRange('H' + pos).setValue(mobileData.cumulativeLayoutShift);

      var desktopData = fetchDataFromPSI('desktop', url);
      desktopSheet.getRange('A' + pos).setFontColor("black");
      desktopSheet.getRange('A' + pos).setValue(desktopData.url);
      desktopSheet.getRange('B' + pos).setValue(desktopData.score);
      desktopSheet.getRange('C' + pos).setValue(desktopData.firstContentfulPaint);
      desktopSheet.getRange('D' + pos).setValue(desktopData.speedIndex);
      desktopSheet.getRange('E' + pos).setValue(desktopData.largestContentfulPaint);
      desktopSheet.getRange('F' + pos).setValue(desktopData.timeToInteractive);
      desktopSheet.getRange('G' + pos).setValue(desktopData.totalBlockingTime);
      desktopSheet.getRange('H' + pos).setValue(desktopData.cumulativeLayoutShift);
    } catch (error) {
      Logger.log("Invalid URL: " + url);
      mobileSheet.getRange('A' + pos).setFontColor("red").setValue(url);
      desktopSheet.getRange('A' + pos).setFontColor("red").setValue(url);
      Logger.log(error);
    }
  });
}

// Utility to build PSI API endpoint
function PageSpeedApiEndpointUrl(strategy, url) {
  const apiBaseUrl = 'https://www.googleapis.com/pagespeedonline/v5/runPagespeed';
  const apikey = 'Enter Your API key'; // Your API key
  const apiEndpointUrl = apiBaseUrl + '?url=' + encodeURIComponent(url) + '&key=' + apikey + '&strategy=' + strategy;
  return apiEndpointUrl;
}

// Fetch HTTP status code for a URL
function getStatusCode(url) {
  var options = {
    'muteHttpExceptions': true,
    'followRedirects': false
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
    return response.getResponseCode();
  } catch (error) {
    Logger.log(error);
    return -1; // Indicate error
  }
}

// Fetch PageSpeed data from PSI API
function fetchDataFromPSI(strategy, url) {
  var options = {
    'muteHttpExceptions': true,
  };
  const PageSpeedEndpointUrl = PageSpeedApiEndpointUrl(strategy, url);
  const response = UrlFetchApp.fetch(PageSpeedEndpointUrl, options);
  const parsedJson = JSON.parse(response.getContentText());

  if (!parsedJson['lighthouseResult']) {
    throw new Error("Invalid PSI response for URL: " + url);
  }
  const audits = parsedJson['lighthouseResult']['audits'];

  return {
    'url': url,
    'score': parsedJson['lighthouseResult']['categories']['performance']['score'] * 100,
    'firstContentfulPaint': audits['first-contentful-paint']['displayValue'],
    'speedIndex': audits['speed-index']['displayValue'],
    'largestContentfulPaint': audits['largest-contentful-paint']['displayValue'],
    'timeToInteractive': audits['interactive']['displayValue'],
    'totalBlockingTime': audits['total-blocking-time']['displayValue'],
    'cumulativeLayoutShift': audits['cumulative-layout-shift']['displayValue']
  };
}
