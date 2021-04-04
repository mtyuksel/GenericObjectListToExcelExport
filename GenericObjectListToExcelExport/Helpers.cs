using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GenericObjectListToExcelExport
{
    class Helpers
    {
        public static void ExportToExcel<T>(List<T> listT, string folderName)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(folderName);
                int currentRow = 1;
                int counter = 1;

                //Getting current class properties
                var propertiesOfT = typeof(T).GetProperties();

                //Create excel column names same as current class properties
                foreach (var prop in propertiesOfT)
                {
                    worksheet.Cell(currentRow, counter).Value = prop.Name.ToString();
                    counter++;
                }

                //Adding the lists values to excel
                foreach (var item in listT)
                {
                    currentRow++;
                    counter = 1;
                    foreach (var prop in propertiesOfT)
                    {
                        worksheet.Cell(currentRow, counter).Value = typeof(T).GetProperty(prop.Name.ToString()).GetValue(item, null);
                        counter++;
                    }
                }

                //Exporting excel to desktop MyFolder location
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    //Getting current users desktop location dynamically
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    //Creating excel file with given folder name.
                    FileStream file = new FileStream(desktopPath + @"\MyFolder\" + folderName + ".xlsx", FileMode.Create, FileAccess.Write);
                    stream.WriteTo(file);
                    stream.Close();
                    file.Close();
                }
            }
        }
    }
}
