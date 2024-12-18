'Script Name: LOI_Submission_RA_LOI_And_Application_001
'Option Explicit
'Variable Declarations
Dim strBrowser, strURL
Dim strPcoriOnline, strFound, strVerifyPcoriOnline, strHelpInfo, strUserName, strPassword

strBrowserType = DataTable.Value("strBrowserType", dtGlobalSheet)
strURL = DataTable.Value("strURL", dtGlobalSheet)
strPcoriOnline = DataTable.Value("strPcoriOnline", dtGlobalSheet)
strUserName = DataTable.Value("strUserName", dtGlobalSheet)
strPassword = DataTable.Value("strPassword", dtGlobalSheet)

'################################################################### Step 1 #####################################################################################################
'Step 1
'Pre-condition: PCORI Online portal should be available. All Test Data required for this test case will be created during test execution.  
'EMAIL ON Required Make sure to have a RA Campaign to execute this project

'################################################################### Step 2 #####################################################################################################
'Step 2: Navigate to PCORI Portal Login Page
Open_Pcori_Application strBrowserType, strURL

'################################################################### Step 3 #####################################################################################################
'Step 3: Verify on the right side of "User Name" and "Password" text present: "PCORI Online is now open for:"   
'At top of page: 'pcori' text, logo and 'Patient-Centered Outcomes Research Institute' text. 
'Below is a 'User Name' field with help text bubble upon hover: Your user name is your email address 'Password' field above a Log In button 
'with the text links: 'Forgot your password? I New User?  and Click here to visit the PCORI website

Set strBrowser = Browser("name:=Login.*").page("title:=Login.*")
strVerifyPcoriOnline = strBrowser.WebElement("innertext:=PCORI Online is now open for.*", "visible:=True", "html tag:=SPAN").GetAllROProperties("innertext")
If Instr(1, strVerifyPcoriOnline , "PCORI Online is now open for:") > 0 Then
	Reporter.ReportEvent micPass, "Verify PCORI Online is now open for:", "PCORI Online is now open for:" & " text is present on the right side of the screen" 
Else
	Reporter.ReportEvent micFail, "Verify PCORI Online is now open for:", "PCORI Online is now open for:" & " text is not present on the right side of the screen" 
End If

'Verify Logo and Text
If strBrowser.Image("file name:=PCORI-Banner.*", "name:=Image").Exist(2) Then
	Reporter.ReportEvent micPass, "Verify Logo and Text", "Logo and Text is present at the top of the page"
Else
	Reporter.ReportEvent micFail, "Verify Logo and Text", "Logo and Text is not present at the top of the page"
End If

'Verify hover help
strFound = 0
Do 
	strBrowser.WebElement("html tag:=LIGHTNING-PRIMITIVE-ICON", "innerhtml:=<svg focusable.*").FireEvent "OnClick"
	If strBrowser.WebElement("innertext:=Your user name is your email address.*", "class:=slds-popover__body").Exist(.2) Then
		strHelpInfo = strBrowser.WebElement("innertext:=Your user name is your email address.*", "class:=slds-popover__body").GetAllROProperties ("innertext")
		strFound = 1
	End If
Loop Until strFound = 1

If strHelpInfo = "Your user name is your email address."  Then
	Reporter.ReportEvent micPass, "Verify Login Help Text", "Your user name is your email address. is displayed while hover the mouse"
Else
	Reporter.ReportEvent micFail, "Verify Login Help Text", "Your user name is your email address. is not displayed while hover the mouse"
End If

'Verify User Name, Password, Click here, Privacy Policy, Click here to visit the PCORI website., Forgot your password?, New User? and Log In fleids
VerifyEditBox strBrowser, "cmntyusrnm", "User Name"
VerifyEditBox strBrowser, "cmntypswd", "Password"
VerifyLink strBrowser, "Click here ", "Click here"
VerifyLink strBrowser,"Privacy Policy", "Privacy Policy"
VerifyText strBrowser,"Click here to visit the PCORI website.", "Click here to visit the PCORI website."
VerifyButton strBrowser,"Forgot your password?", "Forgot your password?"
VerifyButton strBrowser,"New User?", "New User?"
VerifyButton strBrowser,"Forgot your password?", "Forgot your password?"

'Enter values to user name and password field to verify Log In field
Set_WebEdit strBrowser, "cmntyusrnm", strUserName, "User Name"
Set_WebEdit strBrowser, "cmntypswd", strPassword, "Password"
VerifyButton strBrowser,"Log In", "Log In"

'################################################################### Start of Creating New User #####################################################################################################

'################################################################### Step 4 #####################################################################################################
'Click 'New User?' to create a new account
Set strUIAObject = UIAObject("name:=Login - Google Chrome.*").UIAObject("name:=Login")
Click_UIWebButton strUIAObject, "New User?", "New User"

'################################################################### Step 5 #####################################################################################################
'Step5: Fill out all required fields and 
'check the "In voluntarily providing this information, you agree to abide by our website privacy policy and terms of use."  Checkbox and 
'click "Join PCORI Online"
Set strBrowser = Browser("name:=LightningSelfRegisterPage.*").page("title:=LightningSelfRegisterPage.*")
VerifyEditBox strBrowser, "cmntyfrstnm", "First Name"
strVerifyTextbox = Environment.Value("VerifyTextbox")
If strVerifyTextbox = "Success" Then
	Reporter.ReportEvent micPass, "Verify New User page is displayed", "New User page is displayed, Please enter all the required field and click on Join PCORI Online"
Else
	Reporter.ReportEvent micFail, "Verify New User page is displayed", "New User page is not displayed, execution terminated"
	ExitTest
End If

strFirstName = DataTable.Value("strFirstName", dtGlobalSheet)
strLastName = DataTable.Value("strLastName", dtGlobalSheet)
strEmail = DataTable.Value("strEmail", dtGlobalSheet)
strConfirmEmail = DataTable.Value("strConfirmEmail", dtGlobalSheet)
strNewPassword = DataTable.Value("strNewPassword", dtGlobalSheet)
strConfirmNewPassword = DataTable.Value("strConfirmNewPassword", dtGlobalSheet)

'Enver Values for all mandatory fields FirstName, LastName, Email, Confirm Email, Password, Confirm Password
Set_WebEdit strBrowser, "cmntyfrstnm", strFirstName, "First Name"
Set_WebEdit strBrowser, "cmntylastnm", strLastName, "Last Name"
Set_WebEdit strBrowser, "cmntyemail", strEmail, "EMail ID"
Set_WebEdit strBrowser, "cmntycnfrmemail", strConfirmEmail, "Confirm Email ID"
SetSecure_WebEdit strBrowser, "cmntypswd", strNewPassword, "New Password"
SetSecure_WebEdit strBrowser, "cmntycnfrmpswd", strConfirmNewPassword, "Confirm New Password"

'Select Voluntarily Check box
Click_CheckBoxElement strBrowser, "slds-checkbox_faux", "voluntarily check box"

'Click on I am not Robot Text Box
Set strBrowser = Browser("name:=LightningSelfRegisterPage.*").page("title:=LightningSelfRegisterPage.*").Frame("title:=reCAPTCHA", "html tag:=IFRAME", "name:=a.*")
Click_CheckBox strBrowser, "I'm not a robot", "I am not a robot"

'Click on Join PCORI Online button
Set strUIAObject = UIAObject("name:=LightningSelfRegisterPage.*").UIAObject("name:=LightningSelfRegisterPage")
Click_UIWebButton strUIAObject, "Join PCORI Online", "Join PCORI Online"
'DataTable.ImportSheet ("","","")
'################################################################### Step 6 #####################################################################################################
'Step 6: Verify below "Contact Information" text present:  
'"Welcome to the PCORI Online. Please provide some basic information about yourself before proceeding to the Online homepage."
strApplyingFor = DataTable.Value("strApplyingFor", dtGlobalSheet)
strSalutation = DataTable.Value("strSalutation", dtGlobalSheet)
strGender = DataTable.Value("strGender", dtGlobalSheet)
StrHispanicOrLatino = DataTable.Value("StrHispanicOrLatino", dtGlobalSheet)
strRace = DataTable.Value("strRace", dtGlobalSheet)
strYearOfBirth = DataTable.Value("strYearOfBirth", dtGlobalSheet)
strCommunities = DataTable.Value("strCommunities", dtGlobalSheet)
strInvolvedPCORI = DataTable.Value("strInvolvedPCORI", dtGlobalSheet)
strFederalEmployee = DataTable.Value("strFederalEmployee", dtGlobalSheet)
strPositionTitle = DataTable.Value("strPositionTitle", dtGlobalSheet)
strDepartment = DataTable.Value("strDepartment", dtGlobalSheet)
strEmployerName = DataTable.Value("strEmployerName", dtGlobalSheet)
strEmployerFound = DataTable.Value("strEmployerFound", dtGlobalSheet)

strPhone = DataTable.Value("strPhone", dtGlobalSheet)
strStreet = DataTable.Value("strStreet", dtGlobalSheet)
strCity = DataTable.Value("strCity", dtGlobalSheet)
srtStateProvince = DataTable.Value("srtStateProvince", dtGlobalSheet)
strCountry = DataTable.Value("strCountry", dtGlobalSheet)
srtZipPostalCode = DataTable.Value("srtZipPostalCode", dtGlobalSheet)

Set strBrowser = Browser("name:=ContactInformationPage.*").Page("title:=ContactInformationPage")

Click_RadioBtnElement strBrowser, strApplyingFor, strApplyingFor
'Mailing Address
Set_WebEdit strBrowser, "XXX-XXX-XXXX", strPhone, "Phone"
Set_WebEdit strBrowser, "Street", strStreet, "Street"
Set_WebEdit strBrowser, "City", strCity, "City"
Select_WebList strBrowser, "USA;.*", strCountry, "Country"
Select_WebList strBrowser, "Alaska;.*", srtStateProvince, "State OR Province"
Set_WebEdit strBrowser, "Zip Code", srtZipPostalCode, "Postal Code"
Click_RadioBtnElement strBrowser, strSalutation, strSalutation
'Demographic Information
Click_RadioBtnElement strBrowser, strGender, strGender
Click_RadioBtnElement strBrowser, StrHispanicOrLatino, "Hispanic or Latino"
Click_RadioBtnElement strBrowser, strRace, strRace
Select_WebList strBrowser, "1900;.*", strYearOfBirth, "Date of Year"
Click_RadioBtnElement strBrowser, strCommunities, strCommunities
Click_CheckBox strBrowser, strInvolvedPCORI, strInvolvedPCORI
'Employer Information
Click_RadioBtnElement strBrowser, strFederalEmployee, "Federal Employee"
'Add Position, title and department, Employee lookup Search
Set_WebEdit strBrowser, "Position/Title", strPositionTitle, "Position"
Set_WebEdit strBrowser, "Department", strDepartment, "Department"
Set_WebEdit strBrowser, "Type to Search", strEmployerName, "Lookup Employee Name Search"

Select_WebRadioGroup strBrowser, "input-8", "Employee LookUp"


'''Browser("name:=ContactInformationPage.*").Page("title:=ContactInformationPage")
'''Browser("ContactInformationPage_2").Page("ContactInformationPage").ChildObjects ("WebList").
'''
'''
'''strBPStudies = "Salon||Merry||merry123@gmail.com||Agila Somasundaram||Abigail Keatts||Aleksandra Modrow||Anjana||One_Smoke Test Account_DB||Washington DC||Automation"
'''Set strBrowser = Browser("name:=Sfdc Page.*").Page("title:=Sfdc Page")
'''Set Desc = Description.Create()
'''Desc("micclass").value = "WebEdit"
'''Desc("visible").value = true 
'''Set obj = strBrowser.ChildObjects(Desc)
'''Msgbox obj.count
'''If Instr(1, strBPStudies, "||") > 0 Then
'''	strBPStudies = Split(strBPStudies, "||")
'''	strBPSCount = UBound(strBPStudies)
'''End If
'''for i = 0 To strBPSCount - 1
'''	obj(i).Set strBPStudies(i)
'''	items = obj(i).GetRoProperty("default value")
'''	print items
'''Next
''' @@ script infofile_;_ZIP::ssf15.xml_;_

'###########################################################################################################################################


'''GetCurrentDate
'''CurrDate (Time)
'''CreateResultFolder
'''InitResultFile
'''WriteResult "Step 1", "Verify Login Page is Displayed", "Login Page is loaded successfully", "Pass", "Yes"
'''WriteResult "Step 2", "Verify Login Page is Displayed", "Login Page is loaded successfully", "Pass", "Yes"
'''WriteResult "Step 3", "Verify Login Page is Displayed", "Login Page is loaded successfully", "Fail", "Yes"
