using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AISLab9
{
    class ExcelHelper
    {
        public void BuildGraph()
        {
            Excel.Application excelApp = new Excel.Application(); // Создаем экземпляр приложения
            Excel.Workbook workbook = excelApp.Workbooks.Add(); // Экземлпяр рабочей книги Excel 
            Excel.Worksheet workSheet = (Excel.Worksheet)workbook.Worksheets[1]; // Экземпляр рабочего листа

            int number = 60; // Заполняем первую строку числами от 1 до 10, вторую от 60 до 51
            for (int i = 1; i <= 10; i++)
            {
                workSheet.Cells[1, i] = i;
                workSheet.Cells[2, i] = number;
                number--;
            }

            // Вычисляем сумму этих чисел
            Excel.Range rng = workSheet.Range["A3"];
            rng.Formula = "=SUM(A1:J1)";
            rng.FormulaHidden = false;

            Excel.Borders border = rng.Borders; // Выделяем границы у этой ячейки 
            border.LineStyle = Excel.XlLineStyle.xlContinuous;

            // Строим диаграмму
            Excel.ChartObjects chartObjects = (Excel.ChartObjects)workSheet.ChartObjects(Type.Missing);
            Excel.ChartObject chartObject = chartObjects.Add(5, 50, 300, 300);
            Excel.Chart chart = chartObject.Chart;

            chart.SetSourceData(workSheet.Range["A1:J2"]);
            chart.ChartType = Excel.XlChartType.xlXYScatterSmooth;

            chart.HasTitle = true;
            chart.ChartTitle.Text = "График";


            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ExcelGraphExample.xlsx");
            workbook.SaveAs(filePath);

            excelApp.Visible = true;

            // Закрытие Excel
            Marshal.ReleaseComObject(chartObjects);
            Marshal.ReleaseComObject(chartObject);
            Marshal.ReleaseComObject(chart);
            Marshal.ReleaseComObject(workSheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

        }
    }
}
