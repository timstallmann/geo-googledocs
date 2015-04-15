// Global variables
var ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = ss.getActiveSheet(),
    activeRange = ss.getActiveRange(),
    settings = {};
    
var geocoders = {
    mapquest: {
      query: function(query) {
        return 'http://open.mapquestapi.com/nominatim/v1/search?format=json&limit=1&q=' + query;
      },
      parse: function(r) {
        try {
          return {
            longitude: r[0].lon,
            latitude: r[0].lat,
            accuracy: r[0].type
          }
        } catch(e) {
          return { longitude: '', latitude: '', accuracy: 'failure' };
        }
      }
    }
};


// Parts of following is taken from a Google Apps Script example for
// [reading docs](http://goo.gl/TigQZ). It's modified to build a
// [GeoJSON](http://geojson.org/) object.

// Add menu for Geo functions
function onOpen() {
  ss.addMenu('Geo', [{
      name: 'Geocode Addresses',
      functionName: 'gcDialog'
  }, {
      name: 'Help',
      functionName: 'helpSite'
  }]);
}

// Help menu
function helpSite() {
  Browser.msgBox('Support available here: https://github.com/mapbox/geo-googledocs');
}

// Get headers within a sheet and range
function getHeaders(sheet, range, columnHeadersRowIndex) {
    var numColumns = range.getEndColumn() - range.getColumn() + 1;
    var headersRange = sheet.getRange(columnHeadersRowIndex,
        range.getColumn(), 1, numColumns);
    return headersRange.getValues()[0];
}

// Geocoding UI to select API and enter key
function gcDialog() {
  // Create a new UI
  var app = UiApp.createApplication()
    .setTitle('Geocode Addresses')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Create a grid to hold the form
  var grid = app.createGrid(3, 2);

  // Create a vertical panel...
  var panel = app.createVerticalPanel().setId('geocodePanel');

  panel.add(app.createLabel(
    'The selected cells will be joined together and sent to a geocoding service. '
    +'New columns will be added for longitude, latitude, and accuracy score. '
  ).setStyleAttribute('margin-bottom', '20'));

  // ...and add the grid to the panel
  panel.add(grid);

  // Create a button and click handler; pass in the grid object as a
  // callback element and the handler as a click handler
  // Identify the function b as the server click handler
  var button = app.createButton('Geocode')
      .setStyleAttribute('margin-top', '10')
      .setId('geocode');
  var handler = app.createServerClickHandler('geocode');
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);

  // Add the button to the panel and the panel to the application,
  // then display the application app in the Spreadsheet doc
  grid.setWidget(2, 1, button);
  app.add(panel);
  ss.show(app);
}

// Geocode selected range with user-selected api and key
function geocode(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sheet = ss.getActiveSheet(),
      activeRange = ss.getActiveRange(),
      address = '',
      api = 'mapquest',
      response = {},
      rowData = activeRange.getValues(),
      topRow = activeRange.getRow(),
      lastCol = activeRange.getLastColumn();

  
  // update UI
  updateUiGc();
  
  // Check to see if destination columns already exist
  
  var res = getDestCols();
  
  if (res.long >= 0 && res.lat  >= 0 && res.acc >= 0) {
    var longCol = (res.long+1),
        latCol = (res.lat+1),
        accCol = (res.acc+1);  
  } else {
   // Add new columns
    sheet.insertColumnsAfter(lastCol, 3);
    
    // Set new column headers
    sheet.getRange(1, lastCol + 1, 1, 1).setValue('geo_longitude');
    sheet.getRange(1, lastCol + 2, 1, 1).setValue('geo_latitude');
    sheet.getRange(1, lastCol + 3, 1, 1).setValue('geo_accuracy');
 
    // Set destination columns
    var longCol = (lastCol + 1),
        latCol = (lastCol + 2),
        accCol = (lastCol + 3);
  }

  // Don't geocode the first row!
  if (activeRange.getRow() == 1) {
    rowData.shift();
    topRow = topRow + 1;
  }
  
  // For each row, query the API and update the spreadsheet
  for (var i = 0; i < rowData.length; i++) {
    // Join all fields in selected row with a space
    address = rowData[i].join(' ');

    // Concatenate all geo columns
    if (longCol && latCol&& accCol) {
      var testString = sheet.getRange(i + topRow, longCol, 1, 1).getValues()
          + sheet.getRange(i + topRow, latCol, 1, 1).getValues() 
          + sheet.getRange(i + topRow, accCol, 1, 1).getValues();
    }
    // Test to see that all geo columns are empty    
    Logger.log(testString);
    if (testString == '') {
      // Send address to query the geocoding api
      response = getApiResponse(address, api);
  
      // Add responses to columns in the active spreadsheet
      try {
        sheet.getRange(i + topRow, longCol, 1, 1).setValue(response.longitude);
        sheet.getRange(i + topRow, latCol, 1, 1).setValue(response.latitude);
        sheet.getRange(i + topRow, accCol, 1, 1).setValue(response.accuracy);
      } catch(e) {
        Logger.log(e);
      }
    }
  }
  
  // Update UI to notify user the geocoding is done
  closeUiGc();
}

// Check the spreadsheet to see if geo columns exist
function getDestCols() {
  // Get all headers of the active spreadsheet
  var headers = getHeaders(sheet, sheet.getRange(1,1,1,sheet.getLastRow()), 1);
  
  // Search through array for geo cols
  var output = {
    'long': include(headers,'geo_longitude'),
    'lat': include(headers,'geo_latitude'),
    'acc': include(headers,'geo_accuracy')
  };
  
  Logger.log(output.long);
  return output;
}

// Find item in array, return its index
function include(arr,obj) {
    Logger.log(arr.indexOf(obj));
    return arr.indexOf(obj);
}

// Update the UI to show geocoding status
function updateUiGc() {
  // Create new UI
  var app = UiApp.createApplication()
    .setTitle('Geocode Addresses')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Show working message  
  app.add(app
    .createLabel('Geocoding these addresses...')
    .setStyleAttribute('margin-bottom', '10')
    .setId('geocodingLabel')
  );

  // Show the new ui
  ss.show(app);
}

// Update UI to show that geocoding is done
function closeUiGc() {
  Logger.log('starting updateUiGc');
  var app = UiApp.createApplication()
    .setTitle('Geocode Addresses')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Exporting message  
  app.add(app.createLabel(
    'Geocoding is done! You may close this window.')
    .setStyleAttribute('margin-bottom', '10')
    .setStyleAttribute('font-size', '150%')
    .setId('geocodingLabel'));

  ss.show(app);
}

// Send address to api
function getApiResponse(address, api) {
  var geocoder = geocoders[api],
      url = geocoder.query(encodeURI(address));
  
  // If the geocoder returns a response, parse it and return components
  // If the geocoder responds poorly or doesn't response, try again
  for (var i = 0; i < 5; i++) {
    try {
      var response = UrlFetchApp.fetch(url, {method:'get'});
    } catch(e) {
      Logger.log(e);
    }
    if (response && response.getResponseCode() == 200) {
      Logger.log(response.getResponseCode());
      return geocoder.parse(Utilities.jsonParse(response.getContentText()));
    } else {
      Logger.log('The geocoder service being used may be offline.');
    }
    // If no or bad response, sleep for 5 * i seconds and try again
    Logger.log('Something bad happened; retrying. Round: '+(i+1));
    for (var x = 0; x <= i; x++) {
      if (x < 3) { wait(5) };
      if (x = 3) { wait(60) };
      if (x = 4) { wait(120) };
    }
  }
  Logger.log('Tried 5 times, giving up.');
}

function wait(ms) {
  for (var i = 0; i < ms; i++) {
    Logger.log('Sleeping for '+(i+1)+' seconds.');
    Utilities.sleep(1000);
  }
}

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  if (range.getRowIndex() === 1) {
    range = range.offset(1, 0);
  }
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), headers.map(cleanCamel));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  var headers = getHeaders(sheet, activeRange, 1);

  // Zip an array of keys and an array of data into a single-level
  // object of `key[i]: data[i]`
  var zip = function(keys, data) {
    var obj = {};
    for (var i = 0; i < keys.length; i++) {
        obj[keys[i]] = data[i];
    }
    return obj;
  };

  // For each row
  for (var i = 0; i < data.length; i++) {
    var obj = zip(headers, data[i]);

    var lat = parseFloat(obj[settings.lat]),
      lon = parseFloat(obj[settings.lon]);

    var coordinates = (lat && lon) ? [lon, lat] : false;

    // If we have an id, lon, and lat
    if (obj[settings.id] && coordinates) {
      // Define a new GeoJSON feature object
      var feature = {
        type: 'Feature',
        // Get ID from UI
        id: obj[settings.id],
        geometry: {
          type: 'Point',
          // Get coordinates from UIr
          coordinates: coordinates
        },
        // Place holder for properties object
        properties: obj
      };
      objects.push(feature);
    }
  }
  return objects;
}

// Normalizes a string, by removing all non alphanumeric characters and using mixed case
// to separate words.
function cleanCamel(str) {
  return str
         .replace(/\s(.)/g, function($1) { return $1.toUpperCase(); })
         .replace(/\s/g, '')
         .replace(/[^\w]/g, '')
         .replace(/^(.)/, function($1) { return $1.toLowerCase(); });
}
