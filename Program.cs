using Newtonsoft.Json;
using OfficeOpenXml;
using RestSharp;
using System;
using System.IO;
using System.Reflection;

namespace GetManyLatLong.App
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var _license = "PUT_YOUR_GOOGLEMAPS_KEY";
            var _file  = "C:\\Pessoal\\GetManyLatLong.App\\address.xlsx"; // the path from your excel file
            var _region  = "BR"; // to help be more precise in the search

            if (!File.Exists(_file))
                throw new Exception("File not found");

            using (var package = new ExcelPackage(new FileInfo(_file)))
            {
                var sheet = package.Workbook.Worksheets[0];

                var totalRows = sheet.Dimension?.Rows ?? 0;
                var totalCollumns = sheet.Dimension?.Columns ?? 0;

                for (int r = 1; r <= totalRows; r++)
                {
                    var address = "#start";

                    for (int c = 1; c <= totalCollumns; c++)
                    {
                        var value = sheet.Cells[r, c].Value;
                        if (value != null && value.ToString() != "")
                            address += string.Format(", {0}", sheet.Cells[r, c].Value);
                    }

                    address = address.Replace("#start, ", string.Empty);

                    var url = $"https://maps.googleapis.com//maps/api/geocode/json?address={address}&region={_region}&key={_license}";

                    var client = new RestClient(url);
                    var request = new RestRequest(Method.GET);
                    var response = client.Execute(request);
                    var result = JsonConvert.DeserializeObject<dynamic>(response.Content);
                    if (result.status == "OK")
                    {
                        var _results = result.results;
                        if (_results != null)
                        {
                            var _geometry = _results[0].geometry;
                            if (_geometry != null)
                            {
                                var _location = _geometry.location;
                                if (_location != null)
                                {
                                    var latitude = Convert.ToString(_location.lat);
                                    var longitude = Convert.ToString(_location.lng);

                                    sheet.Cells[r, totalCollumns + 1].Value = latitude;
                                    sheet.Cells[r, totalCollumns + 2].Value = longitude;

                                    Console.WriteLine($"address: {address} | lat: {latitude} | lng: {longitude}");

                                    package.Save();
                                }
                            }
                        }
                    }

                }

            }
        }


    }
}
