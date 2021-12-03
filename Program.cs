using System;
using System.IO;
using Custom.Excel;

namespace CrossplatformExcelGenerationWithImages
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            var memoryStream = ExcelGenearator.GenerateExcel();
            if (memoryStream.Length > 0)
            {
                memoryStream.Position = 0;
                using (FileStream fileStream = new FileStream("generated.xlsx", FileMode.Create, System.IO.FileAccess.Write))
                {
                    byte[] bytes = new byte[memoryStream.Length];
                    memoryStream.Read(bytes, 0, (int)memoryStream.Length);
                    fileStream.Write(bytes, 0, bytes.Length);
                    memoryStream.Close();
                }
            }
        }
    }
}
