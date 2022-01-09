using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace WorldQuant_Module3_CSA_SkeletonCode
{
    class Program
    {
        static Excel.Workbook workbook;
        static Excel.Application app;
        static Excel.Worksheet worksheet;
        static string filepath = Environment.CurrentDirectory + @"\property_pricing.xlsx";
        static string[] headers = { "Size", "Suburb", "City", "Market value" };
        static int nRows = 1;
        static float min = 1000000000000000;
        static float max = 0;
        static float sum = 0;

        static void Main(string[] args)
        {
            app = new Excel.Application();
            app.Visible = true;
            try
            {
                workbook = app.Workbooks.Open(filepath, ReadOnly: false);
                worksheet = workbook.Worksheets.get_Item(1);
                
                // get starting row count, sum, min, max
                string tempText;
                
                nRows++;
                tempText = worksheet.Cells[nRows, 1].Text as string;
                while (!string.IsNullOrEmpty(tempText))
                {
                    float value = (float)worksheet.Cells[nRows, 4].Value;
                    sum += value;
                    min = Math.Min(min, value);
                    max = Math.Max(max, value);
                    nRows++;
                    tempText = worksheet.Cells[nRows, 1].Text as string;
                }
            }
            catch
            {
                SetUp();
            }

            var input = "";
            while (input != "x")
            {
                PrintMenu();
                input = Console.ReadLine();
                try
                {
                    var option = int.Parse(input);
                    switch (option)
                    {
                        case 1:
                            try
                            {
                                Console.Write("Enter the size: ");
                                var size = float.Parse(Console.ReadLine());
                                Console.Write("Enter the suburb: ");
                                var suburb = Console.ReadLine();
                                Console.Write("Enter the city: ");
                                var city = Console.ReadLine();
                                Console.Write("Enter the market value: ");
                                var value = float.Parse(Console.ReadLine());

                                AddPropertyToWorksheet(size, suburb, city, value);
                            }
                            catch
                            {
                                Console.WriteLine("Error: couldn't parse input");
                            }
                            break;
                        case 2:
                            Console.WriteLine("Mean price: " + CalculateMean());
                            break;
                        case 3:
                            Console.WriteLine("Price variance: " + CalculateVariance());
                            break;
                        case 4:
                            Console.WriteLine("Minimum price: " + CalculateMinimum());
                            break;
                        case 5:
                            Console.WriteLine("Maximum price: " + CalculateMaximum());
                            break;
                        default:
                            break;
                    }
                }
                catch { }
            }

            // save before exiting
            workbook.Save();
            workbook.Close();
            app.Quit();
        }

        static void PrintMenu()
        {
            Console.WriteLine();
            Console.WriteLine("Select an option (1, 2, 3, 4, 5) " +
                              "or enter 'x' to quit...");
            Console.WriteLine("1: Add Property");
            Console.WriteLine("2: Calculate Mean");
            Console.WriteLine("3: Calculate Variance");
            Console.WriteLine("4: Calculate Minimum");
            Console.WriteLine("5: Calculate Maximum");
            Console.WriteLine();
        }

        static void SetUp()
        {
            // TODO: Implement this method
            workbook = app.Workbooks.Add();

            // set up headers
            workbook.SaveAs(filepath);
            worksheet = workbook.Worksheets.get_Item(1);
            for (int i = 0; i < headers.Length; i++) worksheet.Cells[1, i + 1] = headers[i];
            nRows++;

            workbook.Save();
        }

        static void AddPropertyToWorksheet(float size, string suburb, string city, float value)
        {
            // add data to the sheet then adds 1 to the row counter
            worksheet.Cells[nRows, 1] = size;
            worksheet.Cells[nRows, 2] = suburb;
            worksheet.Cells[nRows, 3] = city;
            worksheet.Cells[nRows, 4] = value;
            nRows++;
            
            sum += value; // for mean calculation

            // update min/max
            min = Math.Min(min, value);
            max = Math.Max(max, value);
            
            workbook.Save();
        }

        static float CalculateMean()
        {
            // mean price is the sum of all prices divided by the number of properties
            return sum / (nRows - 2);
        }

        static float CalculateVariance()
        {
            float mean = CalculateMean(); // need to get mean first to calculate sum of squared errors (sse)
            // initialize sse, go through each property, find squared error, add to counter
            double sse = 0;
            for (int row = 2; row < nRows; row++)
            {
                sse += Math.Pow(worksheet.Cells[row, 4].Value - mean, 2);
            }
            return (float)sse / (nRows - 3); // divide sse by #obs-1 for sample var
        }

        static float CalculateMinimum()
        {
            return min; // already calculated after each property is added
        }

        static float CalculateMaximum()
        {
            return max; // already calculated after each property is added
        }
    }
}
