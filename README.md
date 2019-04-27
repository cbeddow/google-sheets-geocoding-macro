# Google Sheets Geocoding Macro

![How It Works](google-sheets-geocoding-macro.gif)

Geocode from addresses to latitude / longitude, using Nominatum API.

## Test Sheet

Try the macro out on a [Test Sheet](https://docs.google.com/spreadsheets/d/1PZGulsMOTAjJxjPDzMrkunTCXQlFYOga50m3ZouzxHg/edit?usp=sharing) with sample address data.

## Google Sheets Add-On

Unfortunately, you've got to add this script to each sheet you are using.

## Multicolumn Addresses &rarr; Latitude, Longitude

This tool supports geocoding using address data spread across multiple columns. 

The way this works is: You select a set of columns containing the data, and the geocoding process puts the latitude, longitude, and OSM URL data in the rightmost three columns. It will overwrite any data in those three columns.

Some care is needed, as it will concatenate all columns except the rightmost three columns to create the address string.

![Multicolumn Address Geocoding](google-sheets-geocoding-macro-forward.png)

## Credits

Original code is here: https://github.com/nuket/google-sheets-geocoding-macro

This is a fork with some modifications to migrate from Google geocoding to OSM Nominatum.
