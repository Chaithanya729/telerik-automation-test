# VisionRT Telerik Automation coding test

# Description

A simple automation test to read employee data from Telerik WPF application and storing it into a excel file using 
C#, FlaUI, Specflow using BDD framework and POM approach.  

### Set Up

Clone and install:

```bash
git clone https://github.com/Chaithanya729/telerik-automation-test.git
open telerik-automation-test folder in your Visual Studio IDE 
install NuGet Packages
```

### Running the Application

Change the application file path and excel file path in HomePage Class.
- HomePage.cs : Line 16 applicationPath = "Your_Application_Path"
- HomePage.cs : Line 17 excelFilepath = "Your_Folder_Path_To_store_Excel"

Build your application and run test.

### A typical Project Folder layout

    .
    ├── Features                 # Feature files
    ├── Pages                    # All pages 
    └── Steps                    # Step Definitions

## SpecFlowProjectVisionRt

#### Features

`Telerik.feature` Contains the storing employee data scenario and its steps in gherkin language. 

#### Pages

`HomePage.cs` HomePage class which contains all application elements and methods to perform automation to navigate to grid view and store employee data into excel. 

#### Steps

`TelerikStepDefinitons.cs` Contains all the steps defined in the feature file.

### Future Enhancements

- This is a big learning curve for me to understand FlaUi and to implement the automation test for desktop applications,
given more time , i would like to go implement the scroll feature for getting the total order value which are 
out of the visible grid, by looking into the FlaUI repo, a possible solution is we need to turn off UI virtualization in source code.











