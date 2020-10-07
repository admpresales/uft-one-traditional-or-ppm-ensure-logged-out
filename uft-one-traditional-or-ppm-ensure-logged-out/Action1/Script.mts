'===========================================================
'20201007 - Initial creation
'===========================================================

Dim BrowserExecutable

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the PPM login page
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Logout if option is available to logout
'===========================================================================================
If Browser("PPM").Page("PPM").WebElement("Logon button").Exist(0) Then
	Reporter.ReportEvent micPass, "Login Button Exists", "The Logon button exists, no user is logged in."
Else
	Browser("PPM").Page("Main Page").WebElement("menuUserIcon").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	Browser("PPM").Page("PPM").Link("Sign Out Link").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	Reporter.ReportEvent micPass, "Logged Off", "The script logged off the user that was logged in."
End If

AppContext.Close																			'Close the application at the end of your script

