using TechTalk.SpecFlow;
using SpecFlowProjectVisionRt.Pages;


namespace SpecFlowProjectVisionRt.Steps
{
    [Binding]
    public class TelerikStepDefinitions 
    {
        
        private HomePage homePage;
        

        [Given(@"I launch the application")]
        public void GivenILaunchTheApplication()
        {
           this.homePage = new HomePage();
           this.homePage.LaunchApplication();
        }

        [Given(@"I click on Gridview Button")]
        public void GivenIClickOnGridviewButton()
        {
            this.homePage.ClickGridViewButton();
        }

        [Given(@"I get Header Row of employee table and adding it to excel worksheet")]
        public void GivenIGetHeaderRowOfEmployeeTable()
        {
            this.homePage.GetHeaderRowIntoExcelFile();
        }

        [Given(@"I click first tab to access employee data")]
        public void GivenIMinimizeFirstTabToAccessAllEmployeeGridData()
        {
            this.homePage.ClickFirstTab();
        }

        [Given(@"I get all employee data and thier total order value")]
        public void GivenIGetAllEmployeeDataAndThierTotalOrderValue()
        {
            this.homePage.GetEmployeeData();
        }

        [Given(@"I save the employee data to excel file")]
        public void ThenISaveTheEmployeeDataToExcelFile()
        {
            this.homePage.SaveExcel();
        }

        [Then(@"I close the application")]
        public void ThenICloseTheApplication()
        {
            this.homePage.CloseApplication();
        }
    }
}
