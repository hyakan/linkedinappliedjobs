using System.Globalization;
using Microsoft.Extensions.Configuration;
using System.IO;
using ExcelDataReader;
using CsvHelper;
using System.Collections.Generic;
using System.ComponentModel;
using OfficeOpenXml;

namespace FlatenScrapedAppliedJobsspace
{
    public class Program
    {
        public static IConfiguration Configuration { get; set; }
        public static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

            Configuration = builder.Build();
            var inputFilePath = Configuration["InputFilePath"];
            var outFilePath = Configuration["OutputFilePath"];
            var outFileName = Configuration["OutputFileName"];
 
            // Deserialize the configuration OutputColumns section into the record
            var outputColumnLocations = Configuration.GetSection("OutputColumnsLocation").Get<ColumnsOutputColumnLocation>();


            var flattener = new FlattenLinkedInAppliedJobs();

            Directory.SetCurrentDirectory(outFilePath);  // i want to generate the file in the correct folder
            string filePathName = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-") + outFileName;
            flattener.ProcessExcelFile(inputFilePath, filePathName, outputColumnLocations);

            Console.WriteLine("Conversion completed.");
        }
    }

    public record ColumnsOutputColumnLocation
    {
        public int CompanyNameColumnNumber { get; init; }
        public int PositionColumnNumber { get; init; }
        public int LocationColumnNumber { get; init; }
    }


    public class FlattenLinkedInAppliedJobs
    {
        public void ProcessExcelFile(string inputFilePathName, string outputFilePathName, ColumnsOutputColumnLocation outputColumnLocations)
        {
            // Correctly specifying the EPPlus LicenseContext
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Jobs");
                int recordIndex = 1; // Start writing from the first row in the Excel sheet

                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = File.Open(inputFilePathName, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        while (reader.Read()) // Each read operation moves to the next row in the input Excel
                        {
                            // Assuming the first two rows are to be ignored for each record set
                            reader.Read(); // Skipping the second row (image with URL)
                            reader.Read(); // Now on the row with position data

                            // Extract the relevant information
                            var position = reader.GetValue(0)?.ToString(); // This should now be the position
                            reader.Read(); // Next row, expected to be the company name
                            var companyName = reader.GetValue(0)?.ToString();
                            reader.Read(); // Following row, expected to be the location
                            var location = reader.GetValue(0)?.ToString();

                            // Write the extracted information to the designated columns in the output Excel
                            worksheet.Cells[recordIndex, outputColumnLocations.CompanyNameColumnNumber].Value = companyName; // Company Name in column B
                            worksheet.Cells[recordIndex, outputColumnLocations.PositionColumnNumber].Value = position; // Position in column E
                            worksheet.Cells[recordIndex, outputColumnLocations.LocationColumnNumber].Value = location; // Location in column G

                            recordIndex++; // Move to the next row for the next set of data in the output Excel

                            reader.Read(); // Skip the "Applied X ago" row, moving to the next record
                        }
                    }
                }

 
                var fileInfo = new FileInfo(outputFilePathName);
                package.SaveAs(fileInfo); // Save the new Excel file
            }

            Console.WriteLine("File: '{0}' created successfully.", outputFilePathName);
        }
    }
 }