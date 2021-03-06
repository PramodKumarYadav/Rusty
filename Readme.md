# ⚓ : Rusty 

![GitHub contributors](https://img.shields.io/github/contributors/pramodkumaryadav/rusty)
![GitHub last commit](https://img.shields.io/github/last-commit/pramodkumaryadav/rusty)

## Reason of existence
At the time of making this repository, I found no github repositories that can provide a basic framework to test oracle forms. This project started with a goal to fill in that gap for new users.

## Why Rusty?
It seems, UFT has not been able to catch up with changing times, specially when it comes to working with new technologies and way of working (specially git). This makes it difficult to collobarate as a team working on UFT. Apart from that, the tool create zombie tests when you delete a test. The debugger doesnt work with load libraries (need a manual load). Script dont still autosave (meaning, you may overwrite good with bad and unintentioal over intentional). Most of the artifacts UFT uses are incompatible for merging/compare/conflict resolving. For exampe, UFT heavily relies on excel - (not a data format but an application  that creates binaries with which you cannot do merge/compare/resolve conflicts) or formats that are not workable with git (such as object repositories, properties, environments). 

## Solution
In this project, we try to make the best use of what UFT has to offer (its object spy abilities and debugging - which by the way also needs improvement) and using alternative test artifcats that are both UFT and GIT compatible. In the process we provide a framework that addresses most of these challenges,

Result is an end to end framework, that uses none of these non-compatible artifacts but provides substitutes that can be version controlled in GIT and thus better collaborated in a team. I am sure you will love the end to end design which decouples all test areas (such as Test Data, Test scenarios, Test Suites, functions, objects and actions) and thus allowing high level of scalability with minimum maintenace in case of changes.

# A quick compare

A quick compare between Rusty vs Traditional UFT on KPIs and KSFs is  given below. 

| Key Success Factor (KSF)        | Tool/Tech           | Rusty           | Traditional UFT use  | Edge | 
| ------------- |:-------------:|:-------------:| -----:|-----:|
| Version Control     | GIT | By using Standard data/file formats | By using binaries (excels) & incompatible data formats (object-repositories,properties etc) |Rusty| 
| Code collaboration    | Working in Teams | Git makes it easy to colloborate | Without GIT, its all manual, time consuming and error prone |Rusty| 
| Decoupled design     | Design      |Test fns can use, whatever data format is best suited for job (csvs, db tables, xmls)      |   Functions+Data are tightly coupled and refferred in excel sheets; test scenarios & fns are also tightly coupled in excels |Rusty | 
| Project Size | Performance      | in KBs ~500 KBs     |    in MBs ~500 MBs |Rusty |
| Execution Speed | Performance      | Faster by X4     |    Slower by X4 |Rusty |
| Maintenance | Efforts      | low due to de-coupled, no duplication  |  high due to tightly coupled, excel duplication |Rusty |
| Refactoring | Code optimisation      | Possible (code lives in code)    |    Very little (Code lives partly in excels and partly in code) |Rusty |
| Code duplication | Code optimisation      | De-coupled architecture, zero/low code duplication  |    Due to code used from excels, it requires a set template and thus high duplication |Rusty |

# Scope
    - Oracle Forms
    - Web browsers
    - Files

# Workflow

## Workflow (Manual trigger of tests)

| main        | Run TestSuite           | Run TestScenario(s) one by one  | TestData | TestResults |
| ------------- |:-------------:| -----:| -----:| -----:|
| Trigger Tests |  main calls TestScenarios| each scenario that lays in TestScenario dir is triggered| Test data is used for domain specific tests. This is stored in csv files in TestData dir| In the end results should be stored for each Test Scenario ran in TestSuite (to be done)
| In UFT (actions)     | from select-tests-to-run.csv      |   This calls functions (both generic, domain specific) | Config files provide env related data | Should provide details at TestSuite, Test Scenario and Test Steps level.
| test-entrypoint->main | iterate all of them and run if they are selected as Yes      |    For domain specific, we call them with test data functions. general fns are mostly called using test-data-config.xml file | Test Env selection and Root dir location are stroed in System env variables | Results should be stored in a central location (not something to be version controlled though)

## Workflow (automated trigger of tests) - say scheduled via Jenkins

| Trigger        | main        | Run TestSuite           | Run TestScenario(s) one by one  | TestData | TestResults |
| ------------- | ------------- |:-------------:| -----:| -----:| -----:|
| Schedule tests in jenkins to run at a given time | Trigger Tests |  main calls TestScenarios| each scenario that lays in TestScenario dir is triggered| Test data is used for domain specific tests. This is stored in csv files in TestData dir| In the end results should be stored for each Test Scenario ran in TestSuite (to be done)
| This should launch UFT and trigger tests at main | In UFT (actions)     | from select-tests-to-run.csv      |   This calls functions (both generic, domain specific) | Config files provide env related data | Should provide details at TestSuite, Test Scenario and Test Steps level.
| At completion, results status should be available either via email or in Jenkins | test-entrypoint->main | iterate all of them and run if they are selected as Yes (in scheduled tests, we would ideally want to run all of them)     |    For domain specific, we call them with test data functions. general fns are mostly called using test-data-config.xml file | Test Env selection and Root dir location are stroed in System env variables | Results should be stored in a central location (not something to be version controlled though)

# Design

## High Level Diagram
![high-level-design](./Images/high-level-design.jpg)
## Detailed Diagrams
### [entry-point-main](./Images/entry-point-main.jpg)
### [select-test-sceanarios](./Images/select-test-sceanarios.jpg)
### [actual-tests](./Images/actual-tests.jpg)
### [get-test-configuration](./Images/get-test-configuration.jpg)
### [get-browser-functions](./Images/get-browser-functions.jpg)
### [get-forms-and-test-data](./Images/get-forms-and-test-data.jpg)
### [tear-down](./Images/tear-down.jpg)


# Data and abstraction layer
## Objects
- Each oracle form object type is stored in a file called [oracle-form-objects](./FunctionLibrary/oracle-forms-objects.vbs).
    - [Oracle has in total 18 type of objects](https://admhelp.microfocus.com/uft/en/14.50-14.53/UFT_Help/Subsystems/FunctionReference/Subsystems/OMRHelp/Content/Oracle/ORACLEPACKAGELib_P.html?TocPath=Object%20Model%20Reference%20for%20GUI%20Testing|Oracle|_____0) (thats all)
- Each browser object type is stored in a file called [browser-objects](./FunctionLibrary/browser-objects.vbs).
- These objects are logical objects which are created on runtime.
- Once you have created a fn for an object type, you dont need to ever create another one of same type. 
    - `(Note: In uft, if you use object-repository to store your objects, then you make multiple physical copies of same object. In no time you will have a bluky and slow project with tons of duplications to run).`
- An example of an oracle form object is as below.
    ```
    ' Example Recording: OracleFormWindow("Navigator").OracleTabbedRegion("Functions")
    Function GetOracleFormWindow(title)
        
        'Set object based on the parent object and property title
        Dim objOracleFormWindow: Set objOracleFormWindow = OracleFormWindow("title:="&title)

        'Check and Continue only if the object exists and is enabled
        CheckIfObjectExistsAndIsEnabled objOracleFormWindow, title, "OracleFormWindow" 
        
        'Assign this object to function
        Set GetOracleFormWindow = objOracleFormWindow
        
        'Now release this object memory
        Set objOracleFormWindow = Nothing
        
    End Function
    ```
- An example of a browser object is as below.
    ```
    Function GetPageObject(name, title)

        'Set object based on the property browser name and page title
        Dim objPage: Set objPage = Browser("name:="&name).Page("title:="&title)	

        'Check and Continue only if the object exists and is visible
        CheckIfObjectExistsAndIsVisible objPage, title, "Page" 

        'Assign this object to function
        Set GetPageObject = objPage

        'Now release this object memory
        Set objPage = Nothing
        
    End Function
    ```
## Actions
- Each oracle form action is stored in a file called [oracle-forms-actions](./FunctionLibrary/oracle-forms-actions.vbs).
    - The number of actions is limited for each object type. You can find all the actions availble for a object [using object Spy from UFT](./Images/oracle-objects-and-methods.png).
    - For most objects, these actions/methods are limited to somewhere around 20. And most of these actions are common across objects, so expect a small set of actions in total.
- Each browser action is stored in a file called [browser-actions](./FunctionLibrary/browser-actions.vbs).
- We use actions to work on objects (passed as a parameter) to action functions.
- Apart from doing an action on objects, the idea is to also sepereate the intent from implementation. This keeps code clean and readable.
- An example of an oracle action is as below.
    ```
    ' Example Recording: OracleNotification("Caution").Approve
    Function ApproveOracleNotification(object)	
        
        ' Approve the notification window
        object.Approve
        
    End Function
    ```
- An example of a browser action is as below.
    ```
    'Navigate to the URL 
    Function NavigateToURL(objBrowser, URL)
        
        objBrowser.navigate(URL)
        
    End Function
    ```

## test-env-config.xml 
- [test-env-config.xml](./test-env-config-template.xml) is the place where you store your different test environment(s) configuration.
- any setting which is generic should land here. 

## select-tests-to-run.csv
- [select-tests-to-run.csv](./select-tests-to-run.csv) is the place where you should specify all the scenarios that you have and want to run in UFT.
- Mark the ones you want to run as **_Yes_**. Leave the rest. 

## set system environment variables
- For user to select a test env and to give project root directory (seems vbscript wscript method is unreliable getting this simple thing accurately. So have to rely on powershell.)
- TODO: Create a powershell script that sets the test env and root project directory as system variable. Script needs to be allsigned so that it runs on all machines. 
- System environment names:
    - RUSTY_TEST_ENV
    - RUSTY_HOME

## Test Data
- Give test data as csv files in the directory for test data. A [Sample csv test data file](./TestData/InvoiceNrs.csv) is here for your reference.
- Idea is that you keep data seperate from functions. This way, you will have no duplicate instances and would need to change a data in only one location. 

# Function layer
## General functions
- This is where you create general functions for say [browsers](./FunctionLibrary/browser-functions.vbs) and oracle forms.
- Functions to deal with [test data](./FunctionLibrary/test-data-functions.vbs), [test environement](./FunctionLibrary/test-env-functions.vbs), [files](./FunctionLibrary/file-operations.vbs) and databases come here.

## Domain functions
- This is the place where you write your domain specific functions. I cannot add a sample here for obvious reasons but they are build exactly similar to say you would build a common browser function as shown above. 
- You dont have to go to extra length to parameterize your application objects parameters. Your application object parameters will not change in different test environements (dev, test, uat) and there is no point parameterizing them. 
- Also to support above point, if there is only value for your application object attributes, its not a good candidate for parameterisation.
- Going by the above logic, parameterise only the input values. Say if you want to fill a few fields to create an invoice, those field inputs should be parameterised, stored in seperate files in test data, and should be passed to functions, during creating a test (a step in top layer)

# Test Scnearios Layer
- This is your glue layer. This is the place, where everything is glued together to make tests (still logical),that will run with physical test data, when user (later) selects a test environement and runs them.
- General functions use config data to set up environment related areas.
- Domain functions use test data from .csv files from TestData directory to say FindInvoices.
- This is also the place where you know what can be iterated and what cannot. You control, the iterations from here with two options "*"- All, or a number say "4" to get first four records only. In case if the number is higher, its no problem, sql statement will take care of it.

# Test Suite 
- This is the place where you specifiy all your test scenario names and if you want to run them (Yes/No)
- Tests marked Yes will be picked for execution.

# Test Runner (main) 
- This is the place from where you enter and trigger your tests.
- When executing tests manually, you have to go to the action, from where this main is called and manually trigger it.

# Test Driver (UFT)
- When running tests via a scheduler, you will trigger them using another script (to be created), which will trigger the action in UFT for you.

# Install powershell 7
To launch powershell scripts. 
- We use one to set up the root location of project. 
- Later we will use powershell functions to create test reports using standard powershell fns. We can create report in any format of our choice, csv, json, html, table.

# Oracle database connection
- It took me 5 days to get this correctly working. To save you from this frustration and also since test data is very crucial for setting up any meaningful test framework, this section is dedicated to making this properly set up for you. Hopefully you will not have to face the same frustration that I went through in dealing with this!
- To make oracle database connection work in UFT, download a 32 bit driver for your oracle database.
- Go to this page (if your database is 12c - if another database, go to that page), https://www.oracle.com/database/technologies/oracle12c-windows-downloads.html 
	- Scroll down and look for "Oracle Database 12c Release 2 Client (12.2.0.1.0) for Microsoft Windows (32-bit)" - If you are working on oracle 12C
	- Download the zip file "win32_12201_client.zip"
	- Extract the file and run "setup.exe" from the folder : Downloads\win32_12201_client\client32
	- While selecting "What type of installation do you want", 
        - select -> Administrator (1.5GB)
	- Use windows built in account (3rd option)
	- Keep default locations for 
		- oracle base -> C:\app\client\yourUserID (no space in between allowed) 
		- software location -> C:\app\client\yourUserID\product\12.2.0\client_2
		- Save response file for your future reference
		- Install the product
- To test if everything went okay or not, open the ODBC connection for 32 bit
- Check in the drivers section, if a driver with name "Oracle in OraClient12Home1_32bit" is installed or not.
- Now you can run the vbscript for testing database connection by running it in 32 bit command line mode. To do this on windows, 
	- Type %windir%\SysWoW64\cmd.exe in Start Search box.
	- Change directory to your script location (say cd c:\Users\yourUserID\UFT\Rusty\FunctionLibrary) - If the script is in dir FunctionLibrary
	- Now you can run the script by writing cscript in front of it -> cscript database-functions.vbs
	- It may take some time, even a few mins, but then you should see values from msgbox statements pop up from test script. 
- If you can add the script and test it from UFT in a debug mode, then you should be able to navigate step by step and see
expected values from fetched recordset.

- Connection string format:
    - https://www.connectionstrings.com/oracle-in-oraclient11g_home1/ (use standard format from here)
    - Dim connString: connString = "DRIVER={Oracle in OraClient12Home1_32bit};DBQ=yourHostURL:portNr/DBname;Trusted_Connection=Yes;UID=MyUser;Password=myPassword"
    - Ex (with dummy values): connString = "DRIVER={Oracle in OraClient12Home1_32bit};DBQ=ab12.mycompany.com:1521/ORAB;Trusted_Connection=Yes;UID=pramod;Password=myPassword"


# [Naming conventions](https://medium.com/better-programming/string-case-styles-camel-pascal-snake-and-kebab-case-981407998841)
* Naming directories and files
    * Directory names: PascalCase
    * File names: kebab-case
    * No spaces allowed between Directory and File names (So "Unified Functional Tester" and "Assert function library.txt" are bad directory and file names)
* [Naming functions, parameters and variables](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/program-structure/naming-conventions)
    * Function Names: PascalCase 
    * Parameters: camelCase
    * Variables: camelCase
    * Constants: SNAKE_CASE
    * Database fields: snake_case
    * urls: kebab-case
# References
 - [Readme Markdown Cheatsheet](https://github.com/adam-p/markdown-here/wiki/Markdown-Cheatsheet#tables)

# Appendix

# The Problem
I call it Rusty, since UFT in my experience is outdated and little rusty when it comes to working with new technologies  and way of working. To name a few:
- There isnt a good GIT integration. i.e. not only the integration is very basic, if I commit the code from say VS code for a function library or action, UFT wouldnt see it as committed.
- UFT still relies heavily on excels for data storage and most of its user interfacing. Excel being an application will go as a binary in GIT. You cannot collabrate with others with a binary file. 
- UFT uses object repositories for storing objects and properties for saving properties. Both these formats are non readable formats, that cannot be in any sensible way version controlled in git. If multiple people are working on it, you will only see that 'something changed' but you will have no clue 'what changed'. UFT should have come up with substitute artifacts that could be version controlled in GIT and thus make collaboration possible.
- Still no good ways to assert stuf. Yes there are checkpoints but their abilites are pretty limited. There are tons of libraries out there but its addins support seems pretty limited.
- VBScript is an old and ancient language. Although easy to learn, its very limited in todays context. 
- Still no support with parsers for parsing data types (JSON, XML, CSV, YAML). Most programming languages come with parsers to do these jobs.In UFT you have to make these yourselves.
- No standard functions for database connections. You have to make one yourself with ADODB connection and it seems many are still stuck between 32 bit and 64 bit issues. With 64 bit systems now a norm from many years, its a bit of surprise that we still have these issues.
- IT is too bulky as a tool for quick automation. It limits performance and execution speed. 
- If you rename a action, it leaves 'zomby actions' that are just 'hanging there'. 

There are probably more things that I could think of but lets leave it to that.

Nonetheless, there are some applications, which are build in ancient technologies and are still used in companies. No new tools are there around these old technologies and you may be stuck with UFT. If you are, then hopefully rusty can give you a good start with building automation frameworks.

# Rusty's Solution
Now I cannot take away some of the core limitations of UFT, but with Rusty, I have tried to take away a few of the problems to give you a better integration with GIT and thus a better chance of collaborating with other team members.
- Instead of storing objecs in object repositores (that you cant manage in GIT), we create a function library that helps identifying differnt types of objects on run time. This gives a clean and consistant way to deal with objects without adding bulky and duplicate objects. This code (or so to say virtual object repository) is git readable.
- Actions are stored in another function library. This makes the seperation of concerns (intent and implementation) possible. Giving us smaller functions that are git manageable.
- You can now combine the objects and Actions in every two lines of codes to achieve what you want to achieve. With a consistant way to build your tests, it gives a easy way to write tests (instead of writing them in excel as a keyword based approach -which git cannot work with) to a format, that is similar to the excel one but also git maintainable).See example files to understand what I mean here.
- Excels, although not used at the time of writing, will be replaced with either XMLs, or CSVs, both these artifacts being git friendly.
- In general, all artifacts are created keeping a clean design, git maintanability, collaboration and long term maintanece in mind. 

# Common Troubleshooting in UFT
Common Troubleshooting errors and their root cause while working with UFT:

- If you try to set an object and see that it is coming as false, it could be because there are more than one objects with the same property value. In that case, you may have to create another function, with another parameter that can uniqely help you identfiy that object (example: GetOracleTextFieldUsingDescAndToolTip needed to be created because in a page, there were two text fields with same description)

- If you are debugging, you would need to add library manually (bug in UFT that doesnt load libraries using script). This is a common mistake that happens that we call the function without adding the library.

- Save the function from VSCODE (autosave). Close opened function in UFT and reopen again to see changed values. UFT doesnt autorefresh with changed values. 

- Forgetting to load all function libraries before testing a action that needs function libraries.

- If you load a library via code, all the changes in libraries that you made after you opened UFT will not be loaded (specially true if you have saved files but not commited them). 
  To be able to see your changes, close and open UFT again. Rusty- Right!

- Debug doesnt work when you load libraries via code (even though UFT claims that it should). 
  As I have been told by a friend, UFT has this bug since version 14 (somewhere early in year 2019). Close to one year gone, still exists!

- If you load a library manually via "add fn library from UFT", you would still need to close and open it again in UFT for it to show
  latest contents. Thre is no auto refresh, or auto save for that matter. Not good, but hey, I dont call it Rusty without a reason! 
- Best option is:
	- For debugging, add all function libraries manually. 
	- For running, the script will add them via code.
	
- Changing parameters in functions call or in functions defination but not setting it in other place.

- If you rename a function and call it in UFT, it will not tell that the function is missing. Instead, it will say
  "Type mismatch: 'GetBrowsersEXENameDummy'". Thats an incorrect message for a missing function. Puts on off in wrong direction and wastes time.Rusty!
  
- While calling a object, call using set. While calling a string function, call without Set. 
	Dim objXML1: Set objXML1 = GetTestEnvConfigurationValue(pathXML, testEnv, Key)
	Dim BrowserName: BrowserName = GetXMLChildNodeValue(objXML1, Key)

- Mixing this up will result in unexpected errors. For example:
	- Not setting a "object" using set command. Example.
		Instead of doing this: Dim envNode: Set envNode = objXML.SelectNodes(xPath)
		Doing this: Dim envNode: envNode = objXML.SelectNodes(xPath)
		
		Another example while returning a function type: Function from test-env-functions (for xml)
		Correct setting: Set GetTestEnvConfigurationValue = envNode.item(0)
		Incorrect setting: GetTestEnvConfigurationValue = envNode.item(0)
		
	- Setting a non-object using set command. Example (for a function that returns string).
		Instead of doing this: Dim BrowserName: BrowserName = GetXMLChildNodeValue(objXML1, Key)
		Doing this: Dim BrowserName: Set BrowserName = GetXMLChildNodeValue(objXML1, Key)
		
		Another example while returning a function type: Function from test-env-functions(for xml)
		Correct setting: GetXMLChildNodeValue = childXMLNode.item(0).nodeTypedValue
		Incorrect setting: Set GetXMLChildNodeValue = childXMLNode.item(0).nodeTypedValue
        
