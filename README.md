## A Net.Core development application to query latitude and longitude at once

    var _license = "PUT_YOUR_GOOGLEMAPS_KEY";
Enter your Google Maps key (you can get it from the Google Maps console> Credentials)

    var _path = "C:\\Pessoal\\GetManyLatLong.App";
Enter the path where your Excel file is

 - The file should preferably be in the ".xlsx" extension.
 - Close the Excel file before starting the application, as it will overwrite the file with the new columns containing the latitude and longitude data.
 - The file can contain several columns and each column will be a part of the address, the program will concatenate sequentially as each part of the address.