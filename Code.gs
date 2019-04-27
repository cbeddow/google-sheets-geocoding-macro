// Geocode Addresses
// Copyright (c) 2016 - 2017 Max Vilimpoc
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

// Bias the geocoding results in favor of these geographic regions.
// The regions are specified as ccTLD codes.
// 
// See: https://en.wikipedia.org/wiki/Country_code_top-level_domain
//
// Used:
// https://mbrownnyc.wordpress.com/misc/iso-3166-cctld-csv/
// http://www.convertcsv.com/csv-to-json.htm
// to generate the functions for menu item handling.

// Forward Geocoding -- convert address to GPS position.
function addressToPosition() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getActiveRange();
  
  var popup = SpreadsheetApp.getUi();
  
  // Must have selected at least 3 columns (Address, Lat, Lng).
  // Must have selected at least 1 row.
  
  var columnCount = cells.getNumColumns();
  var rowCount = cells.getNumRows();

  if (columnCount < 3) {
    popup.alert("Select at least 3 columns: Address in the leftmost column(s); the geocoded Latitude, Longitude will go into the last 2 columns.");
    return;
  }
  
  var addressRow;

//  var addressColumnStart = 1; // Address data is in columns [1 .. columnCount - 2].
//  var addressColumnStop  = columnCount - 2; 
  
  var addressColumn;
  
  var latColumn = columnCount - 1; // Latitude  goes into the next-to-last column.
  var lngColumn = columnCount;     // Longitude goes into the last column.
  
  var geocode = function(input) {
    var url = 'https://nominatim.openstreetmap.org/search/%22' + input + '%22?format=json&addressdetails=0&limit=1';
    if (window.XMLHttpRequest) {
        var request = new XMLHttpRequest();
    } else {
        var request = new ActiveXObject('Microsoft.XMLHTTP');
    }
    request.open('GET', url, false);
    request.send();
    var output = []
    var json = JSON.parse(request.responseText);
    var output_lon = json[0]['lon'];
    var output_lat = json[0]['lat'];
    output.push(output_lon);
    output.push(output_lat);
    return output;
  }
  var location;

  var addresses = sheet.getRange(cells.getRow(), cells.getColumn(), rowCount, columnCount - 2).getValues();
  
  // For each row of selected data...
  for (addressRow = 1; addressRow <= rowCount; ++addressRow) {
    var address = addresses[addressRow - 1].join(' ');

    // Replace problem characters.
    address = address.replace(/'/g, "%27");

    Logger.log(address);
    
    // Geocode the address and plug the lat, lng pair into the 
    // last 2 elements of the current range row.
    location = geocode(address);
   
    // Only change cells if geocoder seems to have gotten a 
    // valid response.
    lat = location[1];
    lng = location[0];     
    cells.getCell(addressRow, latColumn).setValue(lat);
    cells.getCell(addressRow, lngColumn).setValue(lng);
  }
};


/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 *
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Geocode', generateMenu());
  // SpreadsheetApp.getActiveSpreadsheet().addMenu('Region',  generateRegionMenu());
  // SpreadsheetApp.getUi()
  //   .createMenu();
};
