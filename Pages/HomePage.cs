using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using OfficeOpenXml;
using System.IO;

namespace SpecFlowProjectVisionRt.Pages
{
    public class HomePage
    {
        private Application application;
        private UIA3Automation automation;
        private Window window;
        private ExcelPackage excel;
        private ExcelWorksheet excelWorksheet;
        private string applicationPath = @"C:\Users\rajudachi\AppData\Local\Apps\2.0\AC42NQKZ.MJR\V6DCYAVN.Y7W\wpfd..tion_2c1d60ad5a3725e1_07e5.0001_a3eaaf8e16dd8c94\WPF Demos.exe";
        private string excelFilepath = @"F:\OrderExcel\OrderData.xlsx";

        public HomePage()
        { }

        public void LaunchApplication()
        {
            this.application = FlaUI.Core.Application.Launch(applicationPath);
            this.automation = new UIA3Automation();
            this.window = application.GetMainWindow(automation);
        }

        public HomePage ClickGridViewButton()
        {
            var buttons = window.FindAllDescendants(cf => cf.ByAutomationId("goToControlButton"));
            var i = 0;
            foreach (var button in buttons)
            {
                if (i == 3)
                {
                    button.AsButton().Click(true);
                }
                i++;
            }
            return this;
        }

        public HomePage GetHeaderRowIntoExcelFile()
        {
            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            this.excel = new ExcelPackage();
            this.excelWorksheet = excel.Workbook.Worksheets.Add("Worksheet1");
            var headerGrids = window.FindAllDescendants(cf => cf.ByClassName("GridViewHeaderCell"));
            var headerGridLength = headerGrids.Length;
            var index = 0;
            foreach (var header in headerGrids)
            {
                if (index >= 2 && index <= headerGridLength - 1)
                {
                   var headervalue = header.AsGridCell().Value;
                   this.excelWorksheet.Cells[1, index - 1].Value = headervalue;
                }
                index++;
            }
            this.excelWorksheet.Cells[1, 7].Value = "Total Order Value";
            return this;
        }

        public HomePage ClickFirstTab()
        {
            var firstTab = window.FindFirstDescendant(cf => cf.ByAutomationId("Cell_0_0")).AsButton();
            firstTab.Click();
            return this;
        }
        public HomePage GetEmployeeData()
        {
            //Getting Grid Length
            var gridData = window.FindAllDescendants(cf => cf.ByClassName("GridViewRow"));
            var gridLength = gridData.Length;
            // Getting Grid Data 
            for (int z = 0; z <= gridLength - 1; z++)
            {
                string rowX = "Row_" + z;
                var getRowX = window.FindFirstDescendant(cf => cf.ByAutomationId(rowX));
                getRowX.DrawHighlight();
                int getCellLength = getRowX.FindAllChildren(cf => cf.ByClassName("GridViewCell")).Length;

                for (int j = 2; j <= getCellLength - 1; j++)
                {
                    string cellX = "CellElement_" + z + "_" + j;
                    var getCellX = window.FindFirstDescendant(cf => cf.ByAutomationId(cellX)).AsGridCell();
                    this.excelWorksheet.Cells[z + 2, j - 1].Value = getCellX.Value;
                }
            }
            // Getting total order value 
            for (int z = 0; z <= gridLength - 1; z++)
            {
                string rowS = "Row_" + z;
                var getRowS = window.FindFirstDescendant(cf => cf.ByAutomationId(rowS));
                int getCellLength = getRowS.FindAllChildren(cf => cf.ByClassName("GridViewCell")).Length;
                var expandTab = getRowS.FindFirstDescendant(cf => cf.ByAutomationId("CellElement_" + z + "_" + 0)).AsButton();
                expandTab.Click();
                var ordersTab = getRowS.FindFirstDescendant(cf => cf.ByName("Orders"));
                ordersTab.AsButton().Click();
                var ordersGrid = ordersTab.FindAllDescendants(cf => cf.ByClassName("GridViewRow").And(cf.ByName("Telerik.Windows.Examples.Order")));
                var ordersGridLength = ordersGrid.Length;
                decimal totalOrderValue = 0;
                for (var y = 0; y <= ordersGridLength - 1; y++)
                {
                    var row = "Row_" + y;
                    var getRow = ordersTab.FindFirstDescendant(cf => cf.ByAutomationId(row));
                    var ordersCell = ordersTab.FindFirstDescendant(cf => cf.ByAutomationId("CellElement_" + y + "_" + 4)).AsGridCell();
                    var orderValue = ordersCell.Value.Trim('$');
                    var order = decimal.Parse(orderValue);
                    totalOrderValue += order;
                }
                this.excelWorksheet.Cells[z + 2, 7].Value = "$" + totalOrderValue;
                var detailsTab = getRowS.FindFirstDescendant(cf => cf.ByName("Details")).AsButton();
                detailsTab.Click();
                expandTab.Click();
            }
            return this;
        }
        public HomePage SaveExcel()
        {
            FileInfo excelFile = new FileInfo(excelFilepath);
            this.excel.SaveAs(excelFile);
            return this;
        }
        public void CloseApplication()
        {
            this.application.Close();
        }
    }
}
