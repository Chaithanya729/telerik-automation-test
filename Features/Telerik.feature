Feature: To store employee data and the total order value in an excel sheet 



@Functionality
Scenario: Storing Employee Data into Excel
	Given I launch the application 
	And I click on Gridview Button
	And I get Header Row of employee table and adding it to excel worksheet 
	And I click first tab to access employee data
	And I get all employee data and thier total order value 
	And I save the employee data to excel file
	Then I close the application


	