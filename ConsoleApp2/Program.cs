using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;

class program
{
    public static Columns columns;
    static void Main(string[] args)
    {
        try
        {


            string path = $@"C:\Users\dorbs\OneDrive\שולחן העבודה\‏‏test1 - עותק.xlsx";
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open(path);
            Worksheet worksheet = workbook.ActiveSheet;

            // Add new column for distance
            int columnCount = worksheet.UsedRange.Columns.Count;
            int newColumnIndex = columnCount + 1;
            var newColumnRange = worksheet.Cells[1, newColumnIndex];
            newColumnRange.Value = "Distance";

           
            for (int i = 2; i <= worksheet.UsedRange.Rows.Count; i++)
            {
                string origin = worksheet.Cells[i, 1].Value?.ToString();
                string destination = worksheet.Cells[i, 2].Value?.ToString();
                int distance = 0;

                if (string.IsNullOrEmpty(origin) || string.IsNullOrEmpty(destination))
                {
                    if (string.IsNullOrEmpty(origin) && string.IsNullOrEmpty(destination))
                    {
                        distance = -1;
                    }
                }
                else
                {
                    distance = getDistance(origin, destination);
                }

                worksheet.Cells[i, newColumnIndex].Value = distance;
            }

            workbook.Save();
            workbook.Close();
            excel.Quit();
            Console.WriteLine("Success");
        }
        catch(Exception ex)
        {
            string path = $@"C:\Users\dorbs\OneDrive\שולחן העבודה\‏‏test1 - עותק.xlsx";
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open(path);
            Worksheet worksheet = workbook.ActiveSheet;
            workbook.Close();
            excel.Quit();
            Console.WriteLine($"Faild{ex.Message}");
        }

    }
    public static int getDistance(string origin, string destination)
    {
        //System.Threading.Thread.Sleep(1000);
        int distance = 0;
        string url = $"https://maps.googleapis.com/maps/api/directions/json?origin={origin}&destination={destination}&key=AIzaSyDdAd2YVG51DEku5h24PoGXIWh92SNscvQ";
        string content = fileGetContents(url);
        JObject o = JObject.Parse(content);
        try
        {
            distance = (int)o.SelectToken("routes[0].legs[0].distance.value");
            return distance;
        }
        catch
        {
            return distance;
        }

    }

    protected static string fileGetContents(string fileName)
    {
        string sContents = string.Empty;
        string me = string.Empty;
        try
        {
            if (fileName.ToLower().IndexOf("https:") > -1)
            {
                System.Net.WebClient wc = new System.Net.WebClient();
                byte[] response = wc.DownloadData(fileName);
                sContents = System.Text.Encoding.ASCII.GetString(response);

            }
            else
            {
                System.IO.StreamReader sr = new System.IO.StreamReader(fileName);
                sContents = sr.ReadToEnd();
                sr.Close();
            }
        }
        catch
        {
            sContents = "unable to connect to server ";
        }
        return sContents;
    }
    public static void updateExcel(string filePath)
    {


        string desti = null;
        
        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        Workbook wb;
        Worksheet ws;

        wb = app.Workbooks.Open(filePath);
        ws = wb.Worksheets[1];

        var o = ws.Range["A2:A9"];
        var d = ws.Range["B2:B9"];
        var di = ws.Range["C2:C9"];

        foreach (string des in d.Value)
        {
          if (des != null)
            {
                desti = des;
            }
         Console.WriteLine(desti);    
        }
       

    }
}
