

Public const url = "https://test.salesforce.com/"                                                           ''''''''''"https://login.salesforce.com/?startURL=/"                                            ''''''''''' "https://test.salesforce.com/"
Public const urlprod = "https://login.salesforce.com/?startURL=/"
Public const Reviewerurl = "https://pcori--pqt.sandbox.my.site.com/engagement"             '''''''''''''''"https://pcori.force.com/engagement/s/"           ''''"https://pcori--pqt.sandbox.my.site.com/engagement"  ''''''''''	https://pcori--uat.sandbox.my.site.com/engagement"""""


''''''''''''''''''''''''Misc buttons'''''''''''
Public const strOportunities = "Opportunities"
Public const reportTab = "Reports"
Public const ProjectTab = "Projects"
Public const leadsTab = "Leads"
Public const contactTab = "Contacts"
Public const opportunitiesTab = "Opportunities"
Public const runReportNow = "Run Report Now"
Public const editBtn = "Edit"
Public const saveBtn = "Save"
Public const Convertbtn = "Convert"
Public const deleteBtn = "Delete"
Public const keepCurrentAddressBtn = "Keep Current Address"
Public const searchBtn = "Search"
Public const updateBtn = "Update Address"
Public const servicReqBtn = "Service Requests"
Public const contactsTab = "Accounts"
Public const mergeAccountBtn = "Merge Accounts" 
Public const newTaskBtn = "New Task"
Public const continueBtn = "Continue"
Public const newEventbtn = "New Event"
Public const newBtn = "New"
Public const goBtn = "Go!"
Public const resetBtn = "Reset" 
Public const cancelBtn = "Cancel" 
Public const Nextbtn = "Next -->" 
Public const InPersonReview = "In-Person Review"
Public const saveAndNew = "save & New"
Public const logInButton = "Log In to Sandbox"
Public const logInButton_Prod = "Log In"
Public Const MR_Tile_Button = "ACCESS THE MERIT REVIEWER DASHBOARD"
Public Const portalLoginbtn = "Log in"
Public Const reviewSubmitBtn = "Review/Submit"
Public Const submitBtn = "Submit"


''----- Variables used for MR Automation process:
''''''''acc_name of the weblist for MR phase'''' which updates for each salesforce release'''''''' and will need to update before execution of the script''''

'''''''''''below are for summer 24 sf release'''PQT
Public Const Reviewer_Category_field_PQT = "Reviewer Category"

''''''' Below are for Winter 24 release''''''Prod
Public Const Reviewer_Category_field_Prod = "Reviewer Category"


''''''''''''Admin Id And Password PQT''''''''''''
Public const UserAdminPQT = "sahad@pcori.org.pqt"           		''AdminPassword
Public const AdminPassword = "washington2"

''''''''''''Admin Id And Password PROD''''''''''''
Public const UserAdmin2 = "sahad@pcori.org"           		''AdminPassword
Public const AdminPassword2 = "washington4"

'''''''''Acc_name property value for PFA Level COI field'''''in PQT
Public const PFACOI_Pqt = "PFA Level COI"
'''''''''Acc_name property value for PFA Level COI field'''''in Prod
Public const PFACOI_Prod = "PFA Level COI"

'''''''''Acc_name property value for Application Level COI field'''''in PQT
Public const ApplicationCOI_Pqt = "Application Level COI"
'''''''''Acc_name property value for Application Level COI field'''''in PROD
Public const ApplicationCOI_Prod = "Application Level COI"

'''''''''Acc_name property value for SF Internal all Criterions Score, Human Subject Protection, Overall Comment And MRo Tracking field  for Merit Review '''''in PQT
Public const MRO_Tracking_Pqt = "MRO Tracking Status"
Public const Score1_Pqt = "Criterion 1 Score"
Public const Score2_Pqt = "Criterion 2 Score"
Public const Score3_Pqt = "Criterion 3 Score"
Public const Score4_Pqt = "Criterion 4 Score"
Public const Score5_Pqt = "Criterion 5 Score"
Public const Score6_Pqt = "Criterion 6 Score"
Public const OverallScore_Pqt = "Overall Score"
Public const Human_Subject_Pqt = "Protection of Human Subjects"

'''''''''Acc_name property value for SF Internal all Criterions Score field  for Merit Review '''''in Production
Public const Score1_Prod = "Criterion 1 Score"
Public const Score2_Prod = "Criterion 2 Score"
Public const Score3_Prod = "Criterion 3 Score"
Public const Score4_Prod = "Criterion 4 Score"
Public const Score5_Prod = "Criterion 5 Score"
Public const Score6_prod = "Criterion 6 Score"
Public const OverallScore_Prod = "Overall Score"
Public const Human_Subject_Prod = "Protection of Human Subjects"
Public const MRO_Tracking_Prod = "MRO Tracking Status"

''''''''''''App Page Funding Slate Section Weblist Acc Name'''PQT
Public const Funding_Slate_Stage_accname_Pqt = "Funding Slate Stage"
Public const Funding_Slate_Communication_accname_Pqt = "Funding Slate Communications"

''''''''''''App Page Funding Slate Section Weblist Acc Name'''Prod
Public const Funding_Slate_Stage_accname_Prod = "Funding Slate Stage"
Public const Funding_Slate_Communication_accname_Prod = "Funding Slate Communications"

'''''''''Global User name''''
Public const ProgramUser1 = "Katie Hughes"
Public const ProgramUser2 = "Katie Hughes"
Public const CMAAdmin = "CMA Admin \(Test User\)"
Public const CMAReadOnly = "CMA Coordinator \(Test User\)"
Public const AssgnCMAAdmin = "CMA Admin (Test User)"
Public const MRO = "Carolyn Mohan"
Public const EngOfficer = "Chinenye Anyanwu"

'Gloabl Users Password
Public const passWord = "test123456"
Public const Reviewer1_Pass = "test12345"
Public const passWord2 = "test12345"
Public const passWordMR = "somerstest1234"
Public const passWordCMA = "cmatest234"
Public const somersPassword = "somers12345"
Public const geetaPassword = "geetatest234"
Public const KatiePassword = "katietest234"
Public const carolynPassword = "carolyntest12345"

'''''''Global user user id'''''''

Public const mrManagementUser = "cmohan@pcori.org.pqt"
Public const DMOSUser = "bsomers@pcori.org.pqt"
Public const cmaOperationsuser = "cdmin@gmail.com.sit"
Public const rpCommunityManager = "jamescmaopr@pcori.org.fc4build"
Public const scienceoperationuser = "khughes@pcori.org.pqt"

''''''''Browser''''
Public const IEBrowser    = "IEXPLORE.EXE"
Public const chromeBrowser    = "chrome.exe"
'''''''Merit Reviewer''''''''''
Public const userReviewer1 = "pcori.reg+maryreviewer1@gmail.com.pqt"   '''test12345'''passWord | RAPID Smoke Test 1
Public const userReviewer2 = "pcori.reg+robin2@gmail.com"   '''test1234''''passWord
Public const userReviewer3 = "pcori.reg+bob3@gmail.com" ''''''''test1234''''passWord
Public const userReviewer4 = "pcori.reg+mason4@gmail.com"  '''test1234''''passWord
Public const userReviewer5 = "pcori.reg+mia5@gmail.com"   ''''test1234''''passWord
Public const userReviewer6 = "pcori.reg+randytest6@gmail.com"   
Public Const reviewer1 = "Mary Reviewer 1"
Public Const reviewer2 = "Mary Reviewer 2"
Public Const reviewer3 = "Mary Reviewer 3"
Public Const reviewer4 = "Mary Reviewer 4"
Public Const reviewer5 = "Mary Reviewer 5"
Public Const reviewer6 = "Mary Reviewer 6"

''''''''Portal user name'''''

Public const AOExternal = "PI Test User"
Public const PIExternal = "PI Test User"
Public const FinanceCont = "Financial Contact Smoke Test"
'Public const RAPIEx = "RAPI External Portal"
Public const RAPIDEx = "PID test user"
Public const RAAOEx = "AO Test User"


''''''''Varible using for merit review' cycle, DIpanel and DIcampaign''and DI application name'''''''
Public Const NewCycle = "Reg test Automation cycle FC"
Public const stPanelDI = "Reg test D&I Panel FC"
Public const AwdInstitue = "RTP - Test Account "
Public const DICampgn = "D&I 19C1"
Public const AppDI1 = "Application DI 1 QA Reg FC"
Public const attachmentuploadpath = "C:\Users\sahad\Desktop\blankresumedocument.docx"
'Public const Reviewtype = "Online Review - DI"
Public const Reviewtype = "Ready to Review"

'''''Attachment upload file path'''''''
Public const uploadresumepath = "C:\Users\sahad\Desktop\blankresumedocument.docx"
Public const Miscvariable = "C:\QTP\MiscVariable.txt"
'Set MyBrowser = Browser("micclass:=Browser").Page("micclass:=Page")

''''''''''''''''''Campaign file path'''''''''''''
Public const TextfilePathForCampaign_Broad = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - AD.txt"
Public const TextfilePathForCampaign_LC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - LC.txt"
Public const TextfilePathForCampaign_IMRI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - IMRI.txt"
Public const TextfilePathForCampaign_SDM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - SDM.txt"
Public const TextfilePathForCampaign_HSII = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - HSII.txt"
Public const TextfilePathForCampaign_PLACER = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - PLACER.txt"
Public const TextfilePathForCampaign_PASTdue = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\Autocreated\Campaign - Past Due.txt"
Public const TextfilePathForCampaign_Methods = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - Methods.txt"
Public const TextfilePathForCampaign_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - SOE.txt"
Public const TextfilePathForCampaign_SleepHealth = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - SleepHealth.txt"
Public const TextfilePathForCampaign_MMM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - MMM.txt"
Public const TextfilePathForCampaign_IDD = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - IDD.txt"

''''''''''''Project Name file Path'''''''''''''
Public const TextfilePath_LOI_Project_name_BPS = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - AD.txt"
Public const TextfilePath_LOI_Project_name_LC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - LC.txt"
Public const TextfilePath_LOI_Project_name_IMRI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - IMRI.txt"
Public const TextfilePath_LOI_Project_name_SDM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - SDM.txt"
Public const TextfilePath_LOI_Project_name_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - SOE.txt" 
Public const TextfilePath_LOI_Project_name_Methods = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - Methods.txt"
Public const TextfilePath_LOI_Project_name_PLACER = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - PLACER.txt"
Public const TextfilePath_LOI_Project_name_SleepHealth = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - SleepHealth.txt"
Public const TextfilePath_LOI_Project_name_MMM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - MMM.txt"
Public const TextfilePath_LOI_Project_name_IDD = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - IDD.txt"
Public const TextfilePath_LOI_Project_name_HSII = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - HSII.txt"


''''''''''''Panel Name file Path'''''''''''''
Public const TextfilePath_Panel_name_BPS = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel BPS.txt"
Public const TextfilePath_Panel_name_Methods = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel Methods.txt"
Public const TextfilePath_Panel_name_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel SOE.txt"
Public const TextfilePath_Panel_name_PLACER = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel PLACER.txt"
Public const TextfilePath_Panel_name_HSII = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel HSII.txt"
Public const TextfilePath_Panel_name_LC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel LC.txt"
Public const TextfilePath_Panel_name_IMRI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel IMRI.txt"
Public const TextfilePath_Panel_name_SDM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel SDM.txt"
Public const TextfilePath_Panel_name_CC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel CC.txt"
Public const TextfilePath_Panel_name_CRN = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel CRN.txt"
Public const TextfilePath_Panel_name_CER = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel CER.txt"
Public const TextfilePath_Panel_name_OS = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Panel OS.txt"


''''''''''''''''Test Data Criterions Strength and weaknesses File Path'''''''''''''''''''''
Public const TextfilePath_R1_C1_Strength = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 1 Strength.txt"
Public const TextfilePath_R1_C1_Weaknesses = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 1 Weaknesses.txt"
Public const TextfilePath_R1_C2_Strength = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 2 Strength.txt"
Public const TextfilePath_R1_C2_Weaknesses = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 2 Weaknesses.txt"
Public const TextfilePath_R1_C3_Strength = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 3 Strength.txt"
Public const TextfilePath_R1_C3_Weaknesses = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 3 Weaknesses.txt"
Public const TextfilePath_R1_C4_Strength = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 4 Strength.txt"
Public const TextfilePath_R1_C4_Weaknesses = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 4 Weaknesses.txt"
Public const TextfilePath_R1_C5_Strength = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 5 Strength.txt"
Public const TextfilePath_R1_C5_Weaknesses = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 5 Weaknesses.txt"
Public const TextfilePath_R1_C6_Strength = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 6 Strength.txt"
Public const TextfilePath_R1_C6_Weaknesses = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Criterions 6 Weaknesses.txt"
Public const TextfilePath_R1_Comments = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Comments.txt"
Public const TextfilePath_R1_Human_Subjects_Comments = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\MR Cycle Test\Criterions_Test_Data\Reviewer 1 Human Subjects Comments.txt"

'''''''''''''''''Function Begins'''''''''''''''''''''''''''''''''



Public Function Open_SalesForce_Application()
Err.Clear
On Error Resume Next
Wait 3

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn = DeskTop.ChildObjects(btncalc)			 
	'SystemUtil.CloseProcessByName IEBrowser
	'SystemUtil.Run IEBrowser , url,,,3
	SystemUtil.CloseProcessByName ChromeBrowser
	SystemUtil.Run ChromeBrowser, url,,,3 
	wait 3
	refresh_Chrome_browser()

				If err <> 0 Then    	
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"Opening the Salesforce Application" & "Failed", "Opening the Salesforce Application"& "-" & "Failed to open"
				Else 
					LogReport 0,"Opening the Salesforce Application" & "Success", "Opening the Salesforce Application"& "-" & "Was successful"
				End If  

End Function
Public Function Open_SalesForce_Application_Prod()
Err.Clear
On Error Resume Next
Wait 3

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn = DeskTop.ChildObjects(btncalc)			 
	'SystemUtil.CloseProcessByName IEBrowser
	'SystemUtil.Run IEBrowser , url,,,3
	SystemUtil.CloseProcessByName ChromeBrowser
	SystemUtil.Run ChromeBrowser, urlprod,,,3 
	wait 3
	refresh_Chrome_browser()

				If err <> 0 Then    	
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"Opening the Salesforce Application" & "Failed", "Opening the Salesforce Application"& "-" & "Failed to open"
				Else 
					LogReport 0,"Opening the Salesforce Application" & "Success", "Opening the Salesforce Application"& "-" & "Was successful"
				End If  

End Function

Public Function Open_SalesForce_Application_PartnerUsers()
On Error Resume Next
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)			 
SystemUtil.CloseProcessByName("IEXPLORE.EXE")
'SystemUtil.CloseProcessByName "chrome.exe"
'SystemUtil.Run "chrome.exe" ,url,,,3 
 SystemUtil.Run "iexplore.exe" , Reviewerurl,,,3
wait 3

If err <> 0 Then    	
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1,"Opening the Sales Force Application" , "-" & "Failed to open"
Else 
LogReport 0,"Opening the Sales Force Application" , "-" & "Was successfull"
End If   
End Function
Public Function Open_ReviewersPortal()
On Error Resume Next
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)			 
'SystemUtil.CloseProcessByName("IEXPLORE.EXE")
SystemUtil.CloseProcessByName "chrome.exe"
SystemUtil.Run "chrome.exe" ,Reviewerurl,,,3 
'SystemUtil.Run "iexplore.exe" , Reviewerurl,,,3
wait 3

If err <> 0 Then    	
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1,"Opening the Sales Force Application" , "-" & "Failed to open"
Else 
LogReport 0,"Opening the Sales Force Application" , "-" & "Was successfull"
End If   
End Function
Public Function clk_Button_usingName(strName)

On error resume next
wait 4
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


print pHwnd

Set editO = Description.Create

editO("micclass").value = "WebButton"
editO("visible").value = true  
editO("name").value = strName 
Set editObject =  getParentObject().ChildObjects(editO) 
print editObject.count
editObject(0).highlight
editObject(0).click

If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
else
LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
End If
End Function
''Function to click on ANY link with specified Innertext as parameter
Function clk_link_Object2(strName)
Err.Clear
On error resume next
Wait 3

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn = DeskTop.ChildObjects(btncalc)
	strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
	print strHwnd
	
	Set pa = Description.Create
	pa("micclass").value = "Page"
	
	Set l = Description.Create
	l("micclass").value = "Link"
	l("innertext").value = strName


Set lo =  getParentObject().ChildObjects(l)
print lo.count
lo(0).Highlight

		    If lo.count = 0 Then
		  		LogReport 1,"clk_link_Object2 - " & strName,"Link with innertext" & "-" & strName & "NOT found - NOT expected"
		  	else
		  		LogReport 0,"clk_link_Object2 - " & strName,"Link with innertext" & "-" & strName & "-" & "is FOUND succesfully"
		  	End If

lo(0).click

End Function


Public Sub navigateToAccoutsPageThrughTopSearchBox(strData)
wait 5

setWebEditBox(strData)
'srch_SearchBox(strData)
wait 5
clk_Button_usingName "Search"

End Sub 
Function setWebEditBox(strData)
On error resume next

wait 3
strArr = split(strData,",")

Set oShell = CreateObject("WScript.Shell") 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"



print pHwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebEdit"
editO("visible").value = true  
'editO("html tag").value = "INPUT" 
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 

for i = 0 To uBound(strArr)
'id =   edObj(i).GetRoProperty("html id")
If Not(strArr(i) = "")Then

edObj(i).set strArr(i)  
oShell.SendKeys strSearchText				            
End If

Next
If err <> 0 Then
err.clear   

End If
End Function

Function selectWeblist1(strData)
On error resume next

wait 4
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strhwnd

Set pa = Description.Create
pa("micclass").value = "Page"



print hwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
'editO("html tag").value = "SELECT" 
'editO("type").value = "ComboBox Select" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count

for i = 0 To uBound(strArr)
items = edObj(i+1).GetRoProperty("all items")
print items
print strArr(i+1)
If Not(strArr(i+1) = "")Then	
edObj(i+1).select trim(strArr(i))        
End If
Next


End Function
Public Function clk_Button_usingName2(strName)

On error resume next
wait 4
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


print pHwnd

Set editO = Description.Create

editO("micclass").value = "WebButton"
editO("visible").value = true 
editO("class").value = "setupSearchButton"
editO("name").value = strName 
Set editObject =  getParentObject().ChildObjects(editO) 
print editObject.count
editObject(1).highlight
editObject(1).click

If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
else
LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
End If
End Function

Public Function verifyIfWebElementDoesnotexist(strName)
On error resume next 

wait 3
Set l = Description.Create
l("micclass").value = "WebElement"            
l("innertext").value = strName
Set lo =  getParentObject().ChildObjects.ChildObjects(l)
If lo.count = 0 Then      
LogReport 0,"verityWebElementDoesNotExist - ", "The Link --" & strName & "- doesnot exist _Expected"
else
LogReport 1,"verityWebElementDoesNotExist - ", "The Link --" & strName & "-Exists - Not Expected"			
End If



End Function

Public Function verifyIfButtonDoesExist1(strName)
On error resume next
Set btncalc = Description.Create() 
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)

''print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"




Set l = Description.Create
l("micclass").value = "WebButton"
'l("html tag").value = "A"
l("name").value = strName
Set lo =  getParentObject().ChildObjects(l)
If lo.count = 0 Then      
LogReport 1,"verityLinkDoesExist - ", "The WebButton --" & strName & "- doesnot exist _Expected"
else
LogReport 0,"verityLinkDoesExist - ", "The WebButton --" & strName & "-Exists - Not Expected"

End If



End Function

Public Function verifyIfButtonDoesnotExist(strName)
On error resume next
Set btncalc = Description.Create() 
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)

''print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"




Set l = Description.Create
l("micclass").value = "WebButton"
'l("html tag").value = "A"
l("name").value = strName
Set lo =  getParentObject().ChildObjects(l)
print lo.count

If lo.count > 0 Then      
LogReport 1,"verifyIfButtonDoesnotExist- ", "The WebButton --" & strName & "- exists _Not Expcted"
else
LogReport 0,"verifyIfButtonDoesnotExist- ", "The WebButton --" & strName & "- doesnot Exists -  Expected"

End If



End Function

Public Function verifyIfWebElementDoesExist1(strName)
On error resume next 

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
'l("html tag").value = "LI"
l("innertext").value = strName
l("visible").value = True 

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

If lo.count = 1 Then      
LogReport 1,"verifyIfWebElementDoesExist2 - " & strName, "The WebElement --" & strName & "- Doesn't Exist - Not Expected"
else
LogReport 0,"verifyIfWebElementDoesExist2 - " & strName, "The WebElement --" & strName & "- Exists - Expected"

End If



End Function
Function waitForButton(strName)
On error resume next

Wait 3


status = 0
'start timer
StartTime = Now

Do While  status =  0
'create the object description

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)

Set editO = Description.Create

editO("micclass").value = "WebButton"
editO("visible").value = true  
editO("name").value = strName 
' Set editObject =  getParentObject().ChildObjects(editO) 

Set editObject = btn(0).Page("micclass:=Page").ChildObjects(editO) 
intcount = editObject.count
print intcount

status =  editObject.count
Print  status
'End Timer
EndTime = Now
'get the elapsed time
TimeDiff = CallTimeSeconds(StartTime,EndTime)                           
Print "Waiting for " & "-" & strName& "-" &"Time elaspsed is " & TimeDiff
'Logger Information: 
'                                                     ' "INFO","WaitForNextSpanObject","Waiting for " & "-" & strWebName & "-" &"Time elaspsed is " &TimeDiff
'exit the loop if the object is not found for more than a minute
If  TimeDiff> 30 Then
'Log the results
Print " the webelement" & "-" & strName & "was not found and exiting"

Exit Do
End If    

If status >  0 Then
'Log the results   
wait 3                
print "was found"                    
status = 1

End If
Loop  
waitForButton =  TimeDiff       
End function




Public Function login_intoSalesForce_Application(userid, password)
On error resume next 
wait 2
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"




print pHwnd

Set editO = Description.Create

editO("micclass").value = "WebEdit"
editO("visible").value = true
editO("html tag").value = "input"
Set editObject =  getParentObject().ChildObjects(editO)  

print editObject.count
editObject(0).set userid
editObject(1).set password


If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1,"Login to Sales Force" & "-" & userid & "-" & password,"failed to log in"
else
LogReport 0,"Login to Sales Force" & "-" & userid & "-" & password,"was succesfully loged in"
End If

End Function

Function getParentObject()
On error resume next
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
intcount = btn.count
print intcount

Select Case intcount
Case 1
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
intc = btn.count - 1
Case 2
strHwnd =  btn(btn.count - 2).GetRoProperty("hwnd")
intc = btn.count -2
End Select



Set editObject =  btn(intc).Page("micclass:=Page")
set getParentObject = editObject
If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
'LogReport 1,"Get Parent Object","The parent object was not found"
else
LogReport 0,"Get Parent Object","The parent object was found succesfully"
End If  
End Function
Public Sub navigateAndLoginToSalesForce(strUserName,strPassword )
'Navigate to Panorama.
'
Open_SalesForce_Application
wait 8
'waitForButton logInButton
'
'Enter valid ISE ID and password.
'
login_intoSalesForce_Application strUserName , strPassword 

'
'Click “Log in” button.

clk_Button_usingName logInButton


End Sub 

Public Function getRandomString(ofLength)

If Not IsNumeric(ofLength) Or IsEmpty(ofLength) Or ofLength = "" Then
getRandomString = ""
Exit Function
End If

Dim retval, i
retval = ""
For i = 1 To ofLength
retval = retval & Chr(Int(26*Rnd+97))
Next
getRandomString = retval
End Function

Function fillForm_eventDetails(strEventType)
On error resume next
strNameEvent = "pEv" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = "Name" Then
a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b
c = l_link(i).GetCellData(x,j)
print c
Select Case c

Case "Name"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set strNameEvent
Case "Event Type"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
oEdit1.select strEventType
Case "*Start"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set s + 50
Case "*End"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set s + 52
Case "Description255 remaining"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set "Test"

End Select                               

Next
Next                             




Exit for

End If
End If
Next
End Function
Function clk_singelCheckbox()
On error resume next
wait 3
Set ck = Description.Create

ck("micclass").value = "WebCheckBox"
ck("visible").value = true  


Set co =   getParentObject().ChildObjects(ck)
print co.count
If co.count > 0 Then
LogReport 0,"verifySendSamServeyCheckBox ", "The Check Box exists and is clikable"
else
LogReport 1,"verifySendSamServeyCheckBox ", "The Check Box doesnot exist and is not clickable"
End If
co(0).click
End Function
Function clk_singelCheckbox_newWindow()
On error resume next
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)

Set ck = Description.Create

ck("micclass").value = "WebCheckBox"
ck("visible").value = true  


Set co =   btn(1).Page("micclass:=Page").ChildObjects(ck)
print co.count
If co.count > 0 Then
LogReport 0,"verifySendSamServeyCheckBox ", "The Check Box exists and is clikable"
else
LogReport 1,"verifySendSamServeyCheckBox ", "The Check Box doesnot exist and is not clickable"
End If
co(0).click
End Function
Public Function clk_Button_usingName_NewWindow(strName)
    
     On error resume next
     wait 3
        Set btncalc = Description.Create()  
        btncalc("micclass").value = "Browser"
        
        Set btn =DeskTop.ChildObjects(btncalc)
    
     Set editO = Description.Create
    
     editO("micclass").value = "WebButton"
     editO("visible").value = true  
     editO("name").value = strName 
    ' Set editObject =  getParentObject().ChildObjects(editO) 
    
    Set editObject = btn(1).Page("micclass:=Page").ChildObjects(editO) 
     editObject(0).click
    
    If err <> 0 Then
                  LogReport 4,"Error", err.number & "-" & err.description
                  LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
                  else
                  LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
    End If
End Function
Function RandomString( ByVal strLen ) 
Dim str
Const LETTERS = "abcdefghijklmnopqrstuvwxyz0123456789" 
For i = 1 to strLen
str = str & Mid( LETTERS, RandomNumber( 1, Len( LETTERS ) ), 1 )
Next
RandomString = str



End Function

Function setWebEditBox_NewWindow(HtmlID,strData)
                  On error resume next
                  
                    wait 3
                                                                    strArr = split(strData,",")
                                                                   
                                                                   Set btncalc = Description.Create()  
                                                                    btncalc("micclass").value = "Browser"
                                                                  
                                                                    Set btn =DeskTop.ChildObjects(btncalc)
                                                                   
                                                                    
                                                                    
                                                                  
                                                                    
                                                                  
                                                                   
                                                                    Set editO = Description.Create
                                                                    
                                                                    editO("micclass").value = "WebEdit"
                                                                    editO("visible").value = true  
                                                                      editO("html tag").value = "INPUT"
                                                                       editO("html id").value = HtmlID
                                                                       'editO("type").value = "text" 
                                                                     Set edObj =  btn(1).Page("micclass:=Page").ChildObjects(editO) 
                                                                    
                                                                     for i = 0 To uBound(strArr)
                                                                        'id =   edObj(i).GetRoProperty("html id")
                                                                        If Not(strArr(i) = "")Then
                                                                                
                                                                            edObj(i).set strArr(i)        
                                                                        End If
                                                                         
                                                                     Next
                If err <> 0 Then
                   err.clear   
                  
     End If
     End Function

Function cloneProject()
On error resume next
strNameEvent = "1 a projAUT" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = "Primary Campaign Source" Then
l_link(i).GetROProperty("rows") 

a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b
c = l_link(i).GetCellData(x,j)
print c
Select Case c

Case "*Short Project Title"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set strNameEvent
Case "Project Name*-off layout"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
oEdit1.select strNameEvent
'                                      Short Project Title
Case "*Full Project Title4981 remaining"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set strNameEvent
'                                       Case "Description255 remaining"
'                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
'                                       oEdit1.set "Test"
'                                                                    
End Select                               

Next
Next                             




Exit for

End If
End If
Next
End Function
Function click_DetailListEventEdit()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"

'        Set l_link=  getParentObject().ChildObjects(odesc)
Set l_link= getParentObject().ChildObjects(odesc)

''print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
''print strName
strArr = split(strName,";")
If Not(strName = "") Then
''print strArr(0)
If trim(strArr(0)) = "Action" Then
If trim(strArr(2)) = "Task" Then
a = l_link(i).GetROProperty("rows") 
b = l_link(i).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(i).GetCellData(x,j)
print c & x & j
' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
'                                      If c = stragree  And c2 = strstat Then
Set oEdit1 = l_link(i).ChildItem(2,1, "Link", 0)
oEdit1.click
'                                     End If

Next
Next                            




Exit for
End if


End If
End If

next
End function
Function TypeTheInternalAndExternalNotes()
On error resume next
strNameEvent = "1 a projAUT" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = "Decision Or Action Item255 remaining" Then
l_link(i).GetROProperty("rows") 

a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b
c = l_link(i).GetCellData(x,j)
print c
Select Case c

Case "Internal Notes255 remaining"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set strNameEvent
'                                       Case "Event Type"
'                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
'                                       oEdit1.select strEventType
'                                      Short Project Title
Case "External Notes255 remaining"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set strNameEvent
Case "Project Contact"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set "Chuck Charles"

'                                       Case "Description255 remaining"
'                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
'                                       oEdit1.set "Test"
'                                                                    
End Select                               

Next
Next                             




Exit for

End If
End If
Next
End Function

Function create_Panel()
			On error resume next
			 strDate =  now + 30
             strA = split(strDate,":")
             strDeadLineDate = strA(0) & ":" & strA(1) & space(1) & "PM"
             
              strDate2 =  now + 30
              strB = split(strDate2,":")
              strformatted = strB(0) & ":" & strB(1) & space(1) & "PM"             
'              wait 2
'              strDate3 =  now + 30
'              strC = split(strDate3,":")
'             strformatted1 = strC(0) & ":" & strC(1) '& space(1) & "PM"  
'              
              
			strNameEvent = "Panel 1" & "_" & RandomString(3) 
			strArrN = split(straddinfo,",")
			s = Date
			Set odesc=Description.Create
			odesc("micclass").value="WebTable"
			
			Set l_link=  getParentObject().ChildObjects(odesc)
			print l_link.count
			For i = 0 to l_link.count - 1
			strName = l_link(i).GetROProperty("column names")
			print strName
			strArr = split(strName,";")
			If Not(strName = "") Then
			print strArr(0)
			If trim(strArr(0)) = "*Panel Name" Then
			l_link(i).GetROProperty("rows") 
			
			a = l_link(i).GetROProperty("rows")  
			b = l_link(i).GetROProperty("cols") 		  	       
			For x  = 1 to a     
			For j  =1 to b
			c = l_link(i).GetCellData(x,j)
			print c
			Select Case c
			Case "*Panel Name"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set strNameEvent
			
			Case "*Cycle"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "Reg test Automation cycle"
			Case "Program"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "Addressing Disparities"
			'Selecting PFA''''''
			Case "PFA"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "Addressing Disparities Cycle 3"
					
			Case "*MRO"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "Benjamin Somers"
			Case "Online Review Deadline"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set strDeadLineDate
			Case "InPerson Review Deadline"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set  strformatted 
'			Case "Panel Due Date"
'			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
'            oEdit1.set  s + 25
'''''''''''''Selecting Panel Due Date''''''''''''''''''''

            HtmlID = "00N39000003i9zt"
            strData = date + 30
            set_Web_Edit HtmlID,strData
            
            wait 2
            clk_singelCheckbox()
            
			
			                                  
			End Select                               
			
			Next
			Next                             
			
			
			
			
			Exit for
			
			End If
			End If
			Next
			
			create_Panel = strNameEvent
End Function
Function click_PanelAssignment()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"

'        Set l_link=  getParentObject().ChildObjects(odesc)
Set l_link= getParentObject().ChildObjects(odesc)

''print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
''print strName
strArr = split(strName,";")
If Not(strName = "") Then
''print strArr(0)
If trim(strArr(0)) = "Action" Then
If trim(strArr(1)) = "Assignment Number" Then
a = l_link(i).GetROProperty("rows") 
b = l_link(i).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(i).GetCellData(x,j)
print c & x & j
' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
'                                      If c = stragree  And c2 = strstat Then
Set oEdit1 = l_link(i).ChildItem(4,2, "Link", 0)
oEdit1.click
'                                     End If

Next
Next                            




Exit for
End if


End If
End If

next
End function
Function click_COIExpertise()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"

'        Set l_link=  getParentObject().ChildObjects(odesc)
Set l_link= getParentObject().ChildObjects(odesc)

''print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
''print strName
strArr = split(strName,";")
If Not(strName = "") Then
''print strArr(0)
If trim(strArr(0)) = "Action" Then
If trim(strArr(1)) = "COI & Expertise Number" Then
a = l_link(i).GetROProperty("rows") 
b = l_link(i).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(i).GetCellData(x,j)
print c & x & j
' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
'                                      If c = stragree  And c2 = strstat Then
Set oEdit1 = l_link(i).ChildItem(4,2, "Link", 0)
oEdit1.click
'                                     End If

Next
Next                            




Exit for
End if


End If
End If

next
End function
Function click_PanelAssignment2()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"

'        Set l_link=  getParentObject().ChildObjects(odesc)
Set l_link= getParentObject().ChildObjects(odesc)

''print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
''print strArr(0)
If trim(strArr(0)) = "Action" Then
If trim(strArr(1)) = "Review ID" Then
a = l_link(i).GetROProperty("rows") 
b = l_link(i).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(i).GetCellData(x,j)
print c & x & j
' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
'                                      If c = stragree  And c2 = strstat Then
'                                                   Set oEdit1 = l_link(i).ChildItem(4,2, "Link", 0)
'                                           oEdit1.click
'                                     End If

Next
Next                            




Exit for
End if


End If
End If

next
End function
Function click_EditToAddAReviewer()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"

'        Set l_link=  getParentObject().ChildObjects(odesc)
Set l_link= getParentObject().ChildObjects(odesc)

''print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
''print strArr(0)
If trim(strArr(2)) = "Action" Then
If trim(strArr(3)) = "Request Number" Then
a = l_link(i).GetROProperty("rows") 
b = l_link(i).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(i).GetCellData(x,j)
print c & x & j
' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
'                                      If c = stragree  And c2 = strstat Then
Set oEdit1 = l_link(i).ChildItem(2,3, "Link", 0)
oEdit1.click
'                                     End If

Next
Next                            




Exit for
End if


End If
End If

next
End function
Function clk_link_Object3(strName)
On error resume next

wait 5
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"




Set l = Description.Create
l("micclass").value = "Link"
l("innertext").value = strName


Set lo =  getParentObject().ChildObjects(l)
print lo.count
lo(1).click

'    If err <> 0 Then
'  	LogReport 4,"Error", err.number & "-" & err.description
'  	LogReport 1,"cliking the link" & strName,"Failed to click the link" & "-" & strName
'  	else
'  	LogReport 0,"cliking the link" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
'  End If



End Function

Function Add_Reviewers()
On error resume next
strNameEvent = "panelNa" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = "Reviewer 1" Then
l_link(i).GetROProperty("rows") 

a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b
c = l_link(i).GetCellData(x,j)
print c
Select Case c
Case "Reviewer 1"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set reviewer1

Case "Reviewer 2"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set "Merit Reviewer TA"
Case "Reviewer 3"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set "Merit Review Tester 1"
'                                      Short Project Title
Case "Reviewer 4"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set "Merit Review Tester 2"
Case "Reviewer 4"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set " Sunitha Maddala"
Case "Create Online Reviews"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebCheckBox", 0)
oEdit1.click


'                                       Case "Description255 remaining"
'                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
'                                       oEdit1.set "Test"
'                                                                    
End Select                               

Next
Next                             




Exit for

End If
End If
Next

create_Panel = strNameEvent
End Function
Function verifyFielsValues()
On error resume next
strNameEvent = "panelNa" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = "Reviewer 1" Then
l_link(i).GetROProperty("rows") 

a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b
c = l_link(i).GetCellData(x,j)
print c
Select Case c
Case "Panel"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
strD =  oEdit1.getRoProperty("default value")
If strD = "CDR ST 1 Panel" Then
LogReport 0,"cliking the link" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
else
LogReport 1,"cliking the link" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
End If


'                                       Case "Description255 remaining"
'                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
'                                       oEdit1.set "Test"
'                                                                    
End Select                               

Next
Next                             




Exit for

End If
End If
Next

create_Panel = strNameEvent
End Function
Public Function set_Web_Edit(HtmlID,strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebEdit"
odesc("visible").value= True
'odesc("html tag").value= "TEXTAREA" '"LABEL"
odesc("html id").value= HtmlID   '"con4"
'odesc("innertext").value= strLabelName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y        

Set od =Description.Create
od("micclass").value="WebEdit"
od("abs_x").value= x 
od("abs_y").value= y 
'           

Set Os =  getParentObject().ChildObjects(od)

print Os.count

Os(0).set strData
End Function
Public Function set_Web_List(strLabelName,strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebList"'"WebElement"
odesc("visible").value= True
odesc("html tag").value= "SELECT"'"LABEL"

odesc("innertext").value= strLabelName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x + 1
y= O(0).getRoProperty("abs_y")
print y + 1        

Set od =Description.Create
od("micclass").value="WebList"
'od("abs_x").value= x + 1
od("abs_y").value= y
'           

Set Os =  getParentObject().ChildObjects(od)
Print Os.count

Os(0).select strData
End Function

Sub fill_Progress_reportForm(strName,strPeriod)
On error resume next
set_Web_Edit "Start Date", Date
set_Web_Edit "Due Date", Date + 300
set_Web_Edit "Progress Report Name", strName
set_Web_List "Progress Report Status","Open"
set_Web_List "Progress Report Time Period",strPeriod
End Sub
Sub deleteReviewersName()
On error resume next
set_Web_Edit "Reviewer 1", ""
set_Web_Edit "Reviewer 2", ""
set_Web_Edit "Reviewer 3", ""
set_Web_Edit "Reviewer 4", ""

End Sub

Sub click_PencilImageToEdit(strstatus)
		On error resume next
		Set odesc=Description.Create
		odesc("micclass").value="WebElement"
		odesc("visible").value= True
		odesc("html tag").value= "SPAN"
		odesc("innertext").value= strstatus
		
		Set O =  getParentObject().ChildObjects(odesc)
		print O.count
		x = O(0).getRoProperty("abs_x")
		print x + 1
		y= O(0).getRoProperty("abs_y")
		print y + 1 
		
		Set od=Description.Create
		od("micclass").value="WebElement"
		od("class").value =  "fa fa-pencil-square-o"
		od("abs_y").value= y
		'od("html tag").value= I
		Set Ob =  getParentObject().ChildObjects(od)
		print Ob.count
		Ob(0).click
End Sub



Function inlineEdit(strTableName,strData)
		On error resume next
		strNameEvent = "panelNa" & "_" & RandomString(3) 
		strArrN = split(straddinfo,",")
		s = Date
		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		For i = 0 to l_link.count - 1
				strName = l_link(i).GetROProperty("column names")
				print strName
				strArr = split(strName,";")
				If Not(strName = "") Then
				print strArr(0)
				If trim(strArr(0)) = strTableName Then
				l_link(i).GetROProperty("rows") 
				
				a = l_link(i).GetROProperty("rows")  
				b = l_link(i).GetROProperty("cols") 		  	       
				For x  = 1 to a     
				For j  =1 to b
				c = l_link(i).GetCellData(x,j)
				print c
				Select Case c
				Case strTableName
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
				
				oEdit1.set strData
				Exit function
				
				'                                       Case "Description255 remaining"
				'                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
				'                                       oEdit1.set "Test"
				'                                                                    
				End Select                               
				
				Next
				Next                             
				
				
				
				
				Exit for
				
				End If
				End If
		Next
		
		
End Function

Public Function set_editlinesForReviewers(strTableName,strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebElement"
odesc("visible").value= True
odesc("html tag").value= "DIV"
odesc("innertext").value= strTableName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y  

Set od=Description.Create
od("micclass").value="WebElement"
'od("class").value =  "x-grid3-cell-inner x-grid3-col-00N39000003LYgD"
od("html tag").value="DIV"
od("abs_x").value= x
Set Ob =  getParentObject().ChildObjects(od)
print Ob.count

For i = 1 to Ob.count - 1
Ob(i).FireEvent "ondblclick"
wait 2
inlineEdit strTableName,strData

clk_Button_usingName "Save"
wait 3
Next

End Function
Public Function select_onlineReviewcheckbox(strTableName)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebElement"
odesc("visible").value= True
odesc("html tag").value= "DIV"
odesc("innertext").value= strTableName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y  

Set od=Description.Create
od("micclass").value="WebElement"
'od("class").value =  "x-grid3-cell-inner x-grid3-col-00N39000003LYgD"
od("html tag").value="DIV"
od("abs_x").value= x 
Set Ob =  getParentObject().ChildObjects(od)
print Ob.count

For i = 1 to Ob.count - 1
wait 3
Ob(i).FireEvent "ondblclick"
wait 3
inline_checkbox "Create Online Reviews"

clk_Button_usingName "Save"
wait 5
Next

End Function
Function inline_checkbox(strTableName)
On error resume next
strNameEvent = "panelNa" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = strTableName Then
l_link(i).GetROProperty("rows") 

a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b
c = l_link(i).GetCellData(x,j)
print c
Select Case c
Case strTableName
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebCheckBox", 0)

oEdit1.click
Exit function

'                                       Case "Description255 remaining"
'                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
'                                       oEdit1.set "Test"
'                                                                    
End Select                               

Next
Next                             




Exit for

End If
End If
Next


End Function

Function selectRadioButton(strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True


Set O =  getParentObject().ChildObjects(odesc)
print O.count

O(0).select strData
End Function

Public Function verifyIfLinkDoesExist1(strName)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)

print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"




Set l = Description.Create
l("micclass").value = "Link"
l("html tag").value = "A"
l("innertext").value = strName
Set lo =  getParentObject().ChildObjects(l)
If lo.count = 1 Then   	
LogReport 0,"verityLinkDoesExist - ", "The Link --" & strName & "- does exist _Expected-passed"
else
LogReport 1,"verityLinkDoesExist - ", "The Link --" & strName & "-does not Exists - Not Expected-failed"

End If



End Function
Function selectRadioButton1(strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True           

Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(1).Select strData
End Function
Function selectRadioButton2(strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True           

Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(2).select strData
End Function

Function click_applicationInListView()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"

'        Set l_link=  getParentObject().ChildObjects(odesc)
Set l_link= getParentObject().ChildObjects(odesc)

''print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
''print strArr(0)
If trim(strArr(2)) = "Action" Then
If trim(strArr(3)) = "Request Number" Then
a = l_link(i).GetROProperty("rows") 
b = l_link(i).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(i).GetCellData(x,j)
print c & x & j
' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
''                                      If c = stragree  And c2 = strstat Then
'                                                  Set oEdit1 = l_link(i).ChildItem(2,3, "Link", 0)
'                                                  oEdit1.click
'                                     End If

Next
Next                            




Exit for
End if


End If
End If

next
End function


Function closeDate()
On error resume next
strNameEvent = "projAUT" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = "*Close Date" Then
l_link(i).GetROProperty("rows") 

a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b
c = l_link(i).GetCellData(x,j)
print c
Select Case c

Case "*Close Date"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
oEdit1.set date
'                                    
'                                                                    
End Select                               

Next
Next                             




Exit for

End If
End If
Next
End Function

Function click_ReviewInPanel()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"

'        Set l_link=  getParentObject().ChildObjects(odesc)
Set l_link= getParentObject().ChildObjects(odesc)

''print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
''print strArr(0)
If trim(strArr(0)) = "Action" Then
If trim(strArr(1)) = "Review ID" Then
a = l_link(i).GetROProperty("rows") 
b = l_link(i).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(i).GetCellData(x,j)
print c & x & j
' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
'                                      If c = stragree  And c2 = strstat Then
Set oEdit1 = l_link(i).ChildItem(2,2, "Link", 0)
oEdit1.click
'                                     End If

Next
Next                            




Exit for
End if


End If
End If

next
End function

Function changeReviewStatus(strData)
On error resume next
strNameEvent = "projAUT" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = "Review ID" Then
l_link(i).GetROProperty("rows") 

a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b
c = l_link(i).GetCellData(x,j)
print c
Select Case c

Case "MRO Tracking Status"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
oEdit1.select strData
'                                    
'                                                                    
End Select                               

Next
Next                             




Exit for

End If
End If
Next
End Function

Function clickApplicationsInAList(strStatus)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"
odesc("cols").value= 9  

'            
Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b
c = l_link(i).GetCellData(x,j)
print c
If trim( c) = strStatus Then
Set oEdit1 = l_link(i).ChildItem(x,j, "Link", 0)
oEdit1.click
print x
print j

Exit function
End If
print c
next
next                           


Next


'
End Function 

Function clk_MagnifyingClass_StationLookUP()

On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)

print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"




Set l = Description.Create
l("micclass").value = "Image"
l("image type").value = "Image Link"
l("visible").value = true


Set lo =  getParentObject().ChildObjects(l)
print lo.count
For i = 0 to lo.count - 1
name = lo(i).GetRoProperty("alt")
If name = "Panel Lookup (New Window)" Then
lo(i).click
Exit for
End If
Next
End Function

Function dblClick_Panelfield()
	On error resume next
	Set odesc=Description.Create
	odesc("micclass").value="WebElement"
	odesc("visible").value= True
	odesc("html tag").value= "DIV"
	odesc("innertext").value= "Panel"
	
	Set O =  getParentObject().ChildObjects(odesc)
	print O.count
	x = O(0).getRoProperty("abs_x")
	print x 
	y= O(0).getRoProperty("abs_y")
	print y  
	
	Set od=Description.Create
	od("micclass").value="WebElement"
	'od("class").value =  "x-grid3-cell-inner x-grid3-col-00N39000003LYer"
	od("html tag").value="DIV"
	'od("html id").value = HtmlID
	od("abs_x").value= x
	Set Ob =  getParentObject().ChildObjects(od)
	print Ob.count
	
	
	Ob(1).FireEvent "ondblclick"
End Function

Function clk_Link_ObjectInApage_NewWindow(strData)
On error resume next

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)

wait 5

Set l = Description.Create
l("micclass").value = "Link"
l("visible").value = True
l("innertext").value = strData
' Browser("hwnd:=" & strHwnd).Page("title:=salesforce.com - Unlimited Edition").highlight

Set lo = btn(1).Page("micclass:=Page").ChildObjects(l)
print lo.count

For i= 0 to lo.count - 1 
ret = lo(i).GetRoProperty("innertext")
If ret = strData Then
lo(i).click
Exit for
End If
Next
End Function

Function LogReport(micPass,strTestStepName,strResultDesc)

Reporter.ReportEvent micPass, strTestStepName, strResultDesc
End Function

Function OpenApplicationFromReviewRecord()
			On error resume next
			strNameEvent = "projAUT" & "_" & RandomString(3) 
			strArrN = split(straddinfo,",")
			s = Date
			Set odesc=Description.Create
			odesc("micclass").value="WebTable"
			
			Set l_link=  getParentObject().ChildObjects(odesc)
			print l_link.count
			For i = 0 to l_link.count - 1
				strName = l_link(i).GetROProperty("column names")
				print strName
				strArr = split(strName,";")
					If Not(strName = "") Then
						print strArr(0)
						If trim(strArr(0)) = "Review ID" Then
						l_link(i).GetROProperty("rows") 
						
						a = l_link(i).GetROProperty("rows")  
						b = l_link(i).GetROProperty("cols") 		  	       
						For x  = 1 to a     
						For j  =1 to b
						c = l_link(i).GetCellData(x,j)
						print c
						Select Case c
						
						Case "Application"
						Set oEdit1 = l_link(i).ChildItem(x,j + 1, "Link", 0)
						oEdit1.click
						'                                    
						'                                                                    
						End Select                               
						
						Next
						Next                             
						
						
						
						
						Exit for
					
					End If
				End If
			Next
End Function
Function CreateNewReview(strStatus)
On error resume next
strNameEvent = "projAUT" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
		print strArr(0)
		If trim(strArr(0)) = "Deadline" Then
		l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 		  	       
		For x  = 1 to a     
				For j  =1 to b
				c = l_link(i).GetCellData(x,j)
				print c
				Select Case c
				
				Case "Reviewer Type"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
				oEdit1.select "In-Person"
				Case "Status"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
				oEdit1.select strStatus
				
				
				
				Case "Reviewer"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
				oEdit1.set reviewer1
				Case "*Application"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
				oEdit1.set "projAUT_hew"
				'                                    
				'                                                                    
				End Select                               
				
				Next
		Next                             
		
		
		
		
		Exit for
		
		End If
		End If
Next
End Function
Function CreateNewReview2()
				On error resume next
				strNameEvent = "projAUT" & "_" & RandomString(3) 
				strArrN = split(straddinfo,",")
				s = Date
				Set odesc=Description.Create
				odesc("micclass").value="WebTable"
				
				Set l_link=  getParentObject().ChildObjects(odesc)
				print l_link.count
				For i = 0 to l_link.count - 1
					strName = l_link(i).GetROProperty("column names")
					print strName
					strArr = split(strName,";")
					If Not(strName = "") Then
					print strArr(0)
					If trim(strArr(0)) = "Deadline" Then
					l_link(i).GetROProperty("rows") 
					
					a = l_link(i).GetROProperty("rows")  
					b = l_link(i).GetROProperty("cols") 		  	       
					For x  = 1 to a     
					For j  =1 to b
					c = l_link(i).GetCellData(x,j)
					print c
					Select Case c
					
					
					
					Case "Deadline"
					Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
					oEdit1.click
					wait 2
					Set odesc=Description.Create
					odesc("micclass").value="WebElement"
					odesc("class").value="Weekday"
					odesc("innertext").value="29"
					Set l_link=  getParentObject().ChildObjects(odesc)
					print l_link(0).click
					Exit function
					
					
					'                                    
					'                                                                    
					End Select                               
					
					Next
					Next                             
					
					
					
					
					Exit for
					
					End If
					End If
				Next
End Function
Function OverAllScore()
On error resume next
strNameEvent = "projAUT" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = "Overall Comments" Then
l_link(i).GetROProperty("rows") 

a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b
c = l_link(i).GetCellData(x,j)
print c
Select Case c

Case "Overall Score"
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
oEdit1.select "4"


'                                    
'                                                                    
End Select                               

Next
Next                             




Exit for

End If
End If
Next
End Function

Function getDateStampOfReview(strStatus)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"
odesc("cols").value= 7  

'            
Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
a = l_link(0).GetROProperty("rows") 
b = l_link(0).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(0).GetCellData(x,j)
If Trim(c) = strStatus Then
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebElement", 0)                                	 

intText =  oEdit1.getRoProperty("innertext")
print intText

Set oEdit2 = l_link(i).ChildItem(x,3, "WebElement", 0)
intText2 =  oEdit2.getRoProperty("innertext")                                      	
getDateStampOfReview =  intText  & "," & intText2
Exit function
End If

next
next    

'
End Function 
Function verifyIfReviewExists(strReviewInDashboard,strReviewafterSumbitted)
On error resume next
If strReviewInDashboard <> strReviewafterSumbitted  Then
LogReport 0,"verity review - ", "The REview doesnot exist"
else
LogReport 1,"verifyReview - ", "The review still exists"

End If
End Function    

Function getReviewFromStatus(strStatus)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"
odesc("cols").value= 6 

'            
Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
a = l_link(0).GetROProperty("rows") 
b = l_link(0).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(0).GetCellData(x,j)
If Trim(c) = strStatus Then
Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebElement", 0)                                	 

intText =  oEdit1.getRoProperty("innertext")
print intText


getReviewFromStatus =  intText  
Exit function
End If

next
next    

'
End Function 

Function verifyIfReviewExistsinClosed(strStatus)
On error resume next
If strStatus = "Review Submitted"  Then
LogReport 0,"verity review - ", "The review does exist"
else
LogReport 1,"verifyReview - ", "The review does not exists"
End if
End Function

Sub click_magnifyinglass(strStatus)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebElement"
odesc("visible").value= True
odesc("html tag").value= "SPAN"
odesc("innertext").value= strStatus

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x + 1
y= O(0).getRoProperty("abs_y")
print y + 1  

Set od=Description.Create
od("micclass").value="WebElement"
od("class").value =  "fa fa-search"
od("abs_y").value= y
Set Ob =  getParentObject().ChildObjects(od)
print Ob.count
Ob(0).click
End Sub
Function testingtable()
On error resume next
strNameEvent = "panel" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = "Add to Discussion Line" Then
l_link(i).GetROProperty("rows") 

a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b

Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebButton", 0)
Print oEdit1.count
'                                                       

Next
Next                             




Exit for

End If
End If
Next
End Function


Public Function select_InPerson(strTableName)
			On error resume next
			Set odesc=Description.Create
			odesc("micclass").value="WebElement"
			odesc("visible").value= True
			odesc("html tag").value= "DIV"
			odesc("innertext").value= strTableName
			
			Set O =  getParentObject().ChildObjects(odesc)
			print O.count
			x = O(0).getRoProperty("abs_x")
			print x 
			y= O(0).getRoProperty("abs_y")
			print y  
			
			Set od=Description.Create
			od("micclass").value="WebElement"
			'od("class").value =  "x-grid3-cell-inner x-grid3-col-00N39000003LYgD"
			od("html tag").value="DIV"
			od("abs_x").value= x
			Set Ob =  getParentObject().ChildObjects(od)
			print Ob.count
			
			For i = 1 to Ob.count - 1
				Ob(i).FireEvent "ondblclick"
				wait 2
				inline_checkbox "Open In-Person"
				
				clk_Button_usingName "Save"
				wait 3
			Next

End Function

Public Function clk_Button_usingName_Review(strName)

On error resume next
wait 4


Set editO = Description.Create

editO("micclass").value = "WebButton"
editO("visible").value = true 
editO("class").value = "btn btn btn-primary"    
editO("name").value = strName 
Set editObject =  getParentObject().ChildObjects(editO)
print editObject.count     
editObject(0).highlight
editObject(0).click

If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
else
LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
End If
End Function
Public Function verifyIfWebElementDoesexit_NewWindow(strName)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
intcount = btn.count
print intcount

Select Case intcount
Case 2
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
intc = btn.count - 1
Case 3
strHwnd =  btn(btn.count - 2).GetRoProperty("hwnd")
intc = btn.count -2
End Select

wait 3
Set l = Description.Create
l("micclass").value = "WebElement"            
l("innertext").value = strName
Set lo =  Browser("hwnd:=" & strHwnd).Page("micclass:=Page").ChildObjects(l)
If lo.count > 0 Then      
LogReport 0,"verityWebElementDoesExist - ", "The Link --" & strName & "-  exists _Expected"
else
LogReport 1,"verityWebElementDoesExist - ", "The Link --" & strName & "- doesn't Exist - Not Expected"			
End If



End Function
Function setWebEdit_createReview()
On error resume next
strNameEvent = "panel" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
print strArr(0)
If trim(strArr(0)) = "Deadline" Then
l_link(i).GetROProperty("rows") 

a = l_link(i).GetROProperty("rows")  
b = l_link(i).GetROProperty("cols") 		  	       
For x  = 1 to a     
For j  =1 to b

Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebButton", 0)
Print oEdit1.count
'                                                       

Next
Next                             




Exit for

End If
End If
Next
End Function

Function getReviewID()
		On error resume next
		
		strArrN = split(straddinfo,",")
		
		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		For i = 0 to l_link.count - 1
				strName = l_link(i).GetROProperty("column names")
				print strName
				strArr = split(strName,";")
				If Not(strName = "") and ubound(strArr) > 1 Then
					print strArr(0)
					If trim(strArr(0)) = "Review ID" Then
						a = l_link(i).GetROProperty("rows")  
						b = l_link(i).GetROProperty("cols")                      
						For x  = 1 to a     
							For j = 1 to b
							c = l_link(i).GetCellData(1,2)
							getReviewRecord =  c
							print c
					          Exit function 
					'                                
					
							Next
						Next                             
						
						
						
						
						Exit for
					
					End If
				End If
		'          
		next


End Function

Function setUP_review_InPerson(strStatus,strPanelName)
		On error resume next
		navigateAndLoginToSalesForce UserAdmin, "yahweEthiopia@12"
		wait 8
		'click RP community manager
		clk_link_Object2 "Merit Reviewer Management"
		wait 8
		navigateToAccoutsPageThrughTopSearchBox strPanelName
		wait 5
		clk_link_Object2 strPanelName
		wait 5
		clk_Button_usingName "New Review"
		wait 5
		'select the record type and click continue
		selectWeblist "In-Person Review"
		clk_Button_usingName "Continue"
		wait 5
		'fill the form of the review 
		CreateNewReview strStatus
		OverAllScore()
		CreateNewReview2 
		
		clk_Button_usingName "Save"
		wait 5
		setUP_review_InPerson = getReviewRecord()
End Function

Function reportEditboxenter(strdata)
On error resume next
Set editO = Description.Create

editO("micclass").value = "WebEdit"
editO("visible").value = true  
editO("class").value = " x-form-text x-form-field" 

Set e =  getParentObject().ChildObjects(editO)
e(0).set strdata		     

End Function

Function reportListBoxenter(strData)
On error resume next
Set editO = Description.Create

editO("micclass").value = "WebEdit"
editO("visible").value = true  

editO("class").value = "  x-form-text x-form-field x-trigger-noedit" 
Set e =  getParentObject().ChildObjects(editO)
e(0).click
wait 2
Set editO = Description.Create

editO("micclass").value = "WebElement"
editO("visible").value = true  

editO("innertext").value = strData
Set f =  getParentObject().ChildObjects(editO)  
f.click			     
End Function

Public Function select_onlineReviewcheckbox2(stinnertext)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebElement"
odesc("visible").value= True
odesc("html tag").value= "DIV"
odesc("innertext").value= stinntertext

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y  

Set od=Description.Create
od("micclass").value="WebElement"
'od("class").value =  "x-grid3-cell-inner x-grid3-col-00N39000003LYgD"
od("html tag").value="DIV"
od("abs_x").value= x + 1
Set Ob =  getParentObject().ChildObjects(od)
print Ob.count

'           For i = 1 to 3
Ob(0).FireEvent "ondblclick"
wait 2
inline_checkbox "Open In-Person"

clk_Button_usingName "Save"
wait 7
'           Next

End Function

Function clickReviewLinks()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"
odesc("cols").value= 10
Set oT =  getParentObject().ChildObjects(odesc)
print oT.Count

Set  o= oT(1).ChildItem(1,4, "Link", 0)
o.click
End Function

	Function getReviewRecord()
			On error resume next
			Set odesc=Description.Create
			odesc("micclass").value="WebTable"
			odesc("cols").value= 10
			Set oT =  getParentObject().ChildObjects(odesc)
			print oT.Count
			
			Set  o= oT(1).ChildItem(1,4, "Link", 0)
			intsig = o.getRoProperty("innertext")
			getReviewRecord = intsig
	End Function
Function click_DownLoad_GeneratedDocu()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"

'        Set l_link=  getParentObject().ChildObjects(odesc)
Set l_link= getParentObject().ChildObjects(odesc)

''print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
''print strName
strArr = split(strName,";")
If Not(strName = "") Then
''print strArr(0)
If trim(strArr(0)) = "Action" Then
If trim(strArr(1)) = "Type" Then
a = l_link(i).GetROProperty("rows") 
b = l_link(i).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(i).GetCellData(x,j)
print c & x & j
' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
'                                      If c = stragree  And c2 = strstat Then
Set o = l_link(i).ChildItem(2,1, "Link", 0)
o.click
'                                           print s
'                                     End If

Next
Next                            




Exit for
End if


End If
End If

next
End function 
Function click_PanelApplication(strName)
			On error resume next
			Set odesc=Description.Create
			odesc("micclass").value="WebTable"
			
			'        Set l_link=  getParentObject().ChildObjects(odesc)
			Set l_link= getParentObject().ChildObjects(odesc)
			
			''print l_link.count
			For i = 0 to l_link.count - 1
						strName = l_link(i).GetROProperty("column names")
						''print strName
						strArr = split(strName,";")
						If Not(strName = "") Then
						''print strArr(0)
						If trim(strArr(0)) = "Action" Then
								If trim(strArr(1)) = "Short Project Title" Then
								a = l_link(i).GetROProperty("rows") 
								b = l_link(i).GetROProperty("cols")                     
								For x  = 1 to a    
								For j  =1 to b
								c =  l_link(i).GetCellData(x,j)
								print c & x & j
								' c2 =  l_link(i).GetCellData(x,j + 2)
								'                                     print c2
								'                                     
								'                                      If c = stragree  And c2 = strstat Then
								Set oEdit1 = l_link(i).ChildItem(2,2, "Link", 0)
								oEdit1.click
								'                                     End If
								
								Next
								Next                            
								
								
								
								
								Exit for
								End if
								
								
						End If
			End If
			
			next
End function

Function getTheDocumentCreated()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"

'        Set l_link=  getParentObject().ChildObjects(odesc)
Set l_link= getParentObject().ChildObjects(odesc)

''print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
''print strName
strArr = split(strName,";")
If Not(strName = "") Then
''print strArr(0)
If trim(strArr(0)) = "Action" Then
If trim(strArr(1)) = "Type" Then
a = l_link(i).GetROProperty("rows") 
b = l_link(i).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(i).GetCellData(x,j)
print c & x & j
' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
'                                      If c = stragree  And c2 = strstat Then
Set o = l_link(i).ChildItem(2,3, "Link", 0)
getTheDocumentCreated =   o.GetROProperty("innertext") 
Exit function
'                                           print s
'                                     End If

Next
Next                            




Exit for
End if


End If
End If

next
End function 

Function verifyIfTheDocumentIsGenerated()
On error resume next
strDoc = getTheDocumentCreated
strTitle = split(strDoc,"_")

If strTitle(0)="Summary Statement - Application" Then
LogReport 0,"verifyTitle Of the document ", "The Document is generated"
else
LogReport 1,"verityWebElementDoesExist - ", "The Document is not generated"
End If
End Function

Function verifyTheDateOnDocument()
On error resume next
strDoc = getTheDocumentCreated
strTitle = split(strDoc,"-")
strParsed =  strTitle(2)
strDate = split(strParsed,".")
strDateOnDocu =  strDate(0)
strArr = split(strDateOnDocu,"_")
strReplaced = strArr(0) & "/" & strArr(1) & "/" & strArr(2)
If strReplaced = date Then
LogReport 0,"verifyTitle Of the document ", "The Document is generated"
else
LogReport 1,"verityWebElementDoesExist - ", "The Document is not generated"
End If

End Function

Function fillOnlineReviewCriteria()
		On error resume next
		strNameEvent = "panelNa" & "_" & RandomString(3) 
		strArrN = split(straddinfo,",")
		s = Date
		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		For i = 0 to l_link.count - 1
		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
		print strArr(0)
		If trim(strArr(0)) = "Criterion 1" Then
		l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 		  	       
		For x  = 1 to a     
		For j  =1 to b
		c = l_link(i).GetCellData(x,j)
		print c
		Select Case c
		Case "Criterion 1"
		Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
		oEdit1.set "Importance of research results in the context of the existing body of evidence"
		
		Case "Criterion 2"
		Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
		oEdit1.set "Readiness of the research results for implementation"
		Case "Criterion 3"
		Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
		oEdit1.set "Technical merit of the proposed implementation project (project design, outcomes, and evaluation)"
		'                                      Short Project Title
		Case "Criterion 4"
		Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
		oEdit1.set "Project personnel and environment"
		Case "Criterion 5"
		Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
		oEdit1.set "Patient-centeredness"
		Case "Criterion 6"
		Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
		oEdit1.set  "Patient and stakeholder engagement"
		
		End Select                               
		
		Next
		Next                             
		
		
		
		
		Exit for
		
		End If
		End If
		Next
		
		create_Panel = strNameEvent
End Function
Function cloneProject2()
		strPro = "1AprojAUT" & "_" & RandomString(3) 
		Set odesc=Description.Create
		odesc("micclass").value="WebEdit"
		odesc("visible").value= True
		odesc("default value").value= "RA test App for Automation"
		
		
		Set O =  getParentObject().ChildObjects(odesc)
		   print O.count
		   O(0).set strPro
		
		   O(1).set strPro
		   O(2).set strPro
		   
		   
          print strPro		   
		Set odesc=Description.Create
		odesc("micclass").value="WebEdit"
		odesc("visible").value= True
		odesc("default value").value= "CDR ST 1 Panel"
'		
'		
		    Set O =  getParentObject().ChildObjects(odesc)
		   print O.count
		   O(0).set ""
		   
		   
		   Set odesc=Description.Create
		   odesc("micclass").value="WebEdit"
		   odesc("visible").value= True
		    odesc("default value").value= "NANCY Foley"
'		
'		
		    Set O =  getParentObject().ChildObjects(odesc)
		   print O.count
		   O(0).set ""
		   Set odesc=Description.Create
		   odesc("micclass").value="WebEdit"
		   odesc("visible").value= True
		    odesc("default value").value= "Eugene Abul Kashem"
'		
'		
		    Set O =  getParentObject().ChildObjects(odesc)
		   print O.count
		   O(0).set ""
		   Set odesc=Description.Create
		   odesc("micclass").value="WebEdit"
		   odesc("visible").value= True
		    odesc("default value").value= "Danielle Ackerman"
'		
'		
		    Set O =  getParentObject().ChildObjects(odesc)
		   print O.count
		   O(0).set ""
		   Set odesc=Description.Create
		   odesc("micclass").value="WebEdit"
		   odesc("visible").value= True
		    odesc("default value").value= "Lorenz Adams"
'		
'		
		    Set O =  getParentObject().ChildObjects(odesc)
		   print O.count
		   O(0).set ""
		   Set odesc=Description.Create
		   odesc("micclass").value="WebEdit"
		   odesc("visible").value= True
		    odesc("default value").value= "Kutluk Abdelaal"
'		
'		
		    Set O =  getParentObject().ChildObjects(odesc)
		   print O.count
		   O(0).set ""
		   
		  
		   
		   cloneProject2 = strPro 



End Function


	Function CloneView(strData)
		On error resume next
		strNameEvent = "panelNa" & "_" & RandomString(3) 
		strArrN = split(straddinfo,",")
		s = Date
		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		For i = 0 to l_link.count - 1
				strName = l_link(i).GetROProperty("column names")
				print strName
				strArr = split(strName,";")
				If Not(strName = "") Then
						print strArr(0)
						If trim(strArr(0)) = "*View Name:" Then
						l_link(i).GetROProperty("rows") 
						
						a = l_link(i).GetROProperty("rows")  
						b = l_link(i).GetROProperty("cols") 		  	       
						For x  = 1 to a     
						For j  =1 to b
						c = l_link(i).GetCellData(x,j)
						print c
						Select Case c
							Case "*View Name:"
								Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
								oEdit1.set strData
							
						
						
						End Select                               
						
						Next
						Next                             
						
						
						
						
						Exit for
						
				End If
		  End If
		Next
		
		create_Panel = strNameEvent
End Function

Function fillCriterialRowPanelName(stPanel)
	On error resume next
	Set odesc=Description.Create
	odesc("micclass").value="WebEdit"
		odesc("visible").value= True
		odesc("default value").value= "panelNa_ipz"
'		
'		
		    Set O =  getParentObject().ChildObjects(odesc)
		   print O.coucloneProject2nt
		   O(0).set stPanel
End Function

Function createApplicationsProjects()
	 On error resume next
	 wait 2
	 
	 navigateToAccoutsPageThrughTopSearchBox "RA test App for Automation"
      wait 5
      clk_link_Object2 "RA test App for Automation"
'	clk_link_Object2 "Projects"
'	 wait 5
'	'click the project to clone
'	clk_link_Object2 "IP ST App 1"
	wait 5
	'click the clone button
	clk_Button_usingName "Clone"
	wait 5
	 cloneProject2()
	wait 2
	'UnselectOnlineReviewCheckbox()
	'print stProName
	clk_Button_usingName "Save"
	wait 2
	stProjectname = captureWebElementText_projectName
	print stProjectname
	wait 2
	writeToAfileProject stProjectname
	
	'createApplicationsProjects = stProName
	
	
End Function

Function dblClick_PanelFieldtest(strName)
		On error resume next
		Set odesc=Description.Create
		odesc("micclass").value="WebElement"
		odesc("visible").value= True
		odesc("html tag").value= "DIV"
		odesc("innertext").value= "Short Project Title"
		
		Set O =  getParentObject().ChildObjects(odesc)
		print O.count
		x = O(0).getRoProperty("abs_x")
		print x 
		y= O(0).getRoProperty("abs_y")
		print y  
		
		Set od=Description.Create
		od("micclass").value="WebElement"
		odesc("html tag").value= "DIV"
		'odesc("innertext").value= "Panel"
		od("abs_x").value= x
		Set Ob =  getParentObject().ChildObjects(od)
		print Ob.count
		For i = 1 To Ob.count - 1
			strr = Ob(i).getroProperty("innertext")
			If trim(strr) = strName Then
				
				intloc = i
				print i
			End If
		Next
		
		
		Set odesc=Description.Create
		odesc("micclass").value="WebElement"
		odesc("visible").value= True
		odesc("html tag").value= "DIV"
		odesc("innertext").value= "Panel"
		
		Set O =  getParentObject().ChildObjects(odesc)
		print O.count
		x = O(0).getRoProperty("abs_x")
		print x 
		y= O(0).getRoProperty("abs_y")
		print y  
		
		Set od=Description.Create
		od("micclass").value="WebElement"
		odesc("html tag").value= "DIV"
		'odesc("innertext").value= "Panel"
		od("abs_x").value= x
		Set Ob =  getParentObject().ChildObjects(od)
		print Ob.count
		
		Ob(intloc).FireEvent "ondblclick"
End Function

Function writeToAfilePanel(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\panel.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function
Function writeToAfileProject(stContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\project.txt",2,true)     'open in write mode
            f.Write (stContent)
            f.Close
            Set f = nothing
End Function

Public Function readFromFilePanel()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\panel.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFilePanel = sContent
End Function
Public Function readFromFileProject()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\project.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFileProject = sContent
End Function

Public Function set_editlinesForReviewers2(strTableName,strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebElement"
odesc("visible").value= True
odesc("html tag").value= "DIV"
odesc("innertext").value= strTableName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y  

Set od=Description.Create
od("micclass").value="WebElement"
'od("class").value =  "x-grid3-cell-inner x-grid3-col-00N39000003LYgD"
od("html tag").value="DIV"
od("abs_x").value= x + 1
Set Ob =  getParentObject().ChildObjects(od)
print Ob.count

For i = 0 to Ob.count - 1
Ob(i).FireEvent "ondblclick"
wait 2
inlineEdit strTableName,strData

clk_Button_usingName "Save"
wait 3
Next

End Function
Function writeToAfileRAMatrixView(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\ramatrix.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function
Public Function readFromRAMatrixView()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\ramatrix.txt",1)
          
          sContent = objTxtFile.Readline
          readFromRAMatrixView = sContent
End Function


'''''''''''''Good one to click on review record based on status'''''''''
Function clickReviewRecordBasedOnStatus(strStatus)
	On error resume next
	 Set odesc=Description.Create
    odesc("micclass").value="WebTable"
    odesc("column names").value=";Choose a RowSelect All;Sort by:Review IDSorted: NoneShow Review ID column actions;Sort by:ReviewerSorted: NoneShow Reviewer column actions;Sort by:Record TypeSorted: NoneShow Record Type column actions;Sort by:ApplicationSorted: NoneShow Application column actions;Sort by:Application ProgramSorted AscendingShow Application Program column actions;Sort by:Primary Reviewer RoleSorted: NoneShow Primary Reviewer Role column actions;Sort by:Reviewer TypeSorted: NoneShow Reviewer Type column actions;Sort by:StatusSorted: NoneShow Status column actions;Sort by:DeadlineSorted: NoneShow Deadline column actions;"
    'odesc("cols").value= 12  
         
'            
    Set l_link=  getParentObject().ChildObjects(odesc)
     print l_link.count
                 a = l_link(0).GetROProperty("rows") 
                 print a
                           b = l_link(0).GetROProperty("cols")   
print b                           
                        For x  = 0 to a    
                             For j  =0 to b
                                      c =  l_link(0).GetCellData(x,j)
                                      print c
                                      If Trim(c) = strStatus Then
                                      	Set o = l_link(1).ChildItem(x,2, "Link", 0)                                  
	
                                      	o.click
                                      	intText =  o.getRoProperty("innertext")
                                      	 Exit function
                                      End If
                                      
                            next
                       next    
                      clickReviewRecordBasedOnStatus =  intText
'
End Function 

Function  getProjectNamefromList(strTocompare)
	On error resume next
	Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
'editO("html tag").value = "INPUT" 
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
x = edObj(0).getRoProperty("all items")
print x
strA = split(x,";")
For i = 0 to Ubound(strA)
	Print strA(i)
	z = Replace(strA(i),Left(strA(i),7),"")
    if Trim(z) = strToCompare then
     getProjectNamefromList = strA(i)
     Exit for
    End  if
	 
Next

End Function
Function CreateNewReview_Online(strPanel,strdeadline)
On error resume next
strNameEvent = "1AprojAUT" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
		print strArr(0)
		If trim(strArr(0)) = "Record Type" Then
		l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 		  	       
		For x  = 1 to a     
				For j  =1 to b
				c = l_link(i).GetCellData(x,j)
				print c
				Select Case c
				
				Case "Reviewer Type"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
				oEdit1.select "6"
				Case "Status"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
				oEdit1.select strStatus
				
				
				
				Case "Reviewer"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
				oEdit1.set reviewer6
				Case "Panel"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
				oEdit1.set strPanel
				Case "Deadline"
					Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
					oEdit1.set strdeadline
				'                                    
				'                                                                    
				End Select                               
				
				Next
		Next                             
		
		
		
		
		Exit for
		
		End If
		End If
Next
End Function
Function CreateNewReview2_online()
				On error resume next
				strNameEvent = "projAUT" & "_" & RandomString(3) 
				strArrN = split(straddinfo,",")
				s = Date
				Set odesc=Description.Create
				odesc("micclass").value="WebTable"
				
				Set l_link=  getParentObject().ChildObjects(odesc)
				print l_link.count
				For i = 0 to l_link.count - 1
					strName = l_link(i).GetROProperty("column names")
					print strName
					strArr = split(strName,";")
					If Not(strName = "") Then
					print strArr(0)
					If trim(strArr(0)) = "Deadline" Then
					l_link(i).GetROProperty("rows") 
					
					a = l_link(i).GetROProperty("rows")  
					b = l_link(i).GetROProperty("cols") 		  	       
					For x  = 1 to a     
					For j  =1 to b
					c = l_link(i).GetCellData(x,j)
					print c
					Select Case c
					
					
					
					Case "Deadline"
					Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
					oEdit1.click
					wait 2
					Set odesc=Description.Create
					odesc("micclass").value="WebElement"
					odesc("class").value="Weekday"
					odesc("innertext").value="15"
					Set l_link=  getParentObject().ChildObjects(odesc)
					print l_link(0).click
					Exit function
					
					
					'                                    
					'                                                                    
					End Select                               
					
					Next
					Next                             
					
					
					
					
					Exit for
					
					End If
					End If
				Next
End Function
Function getDeadLine()
On error resume next
strNameEvent = "projAUT" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
		print strArr(0)
		If trim(strArr(0)) = "Panel Name" Then
		l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 		  	       
		For x  = 1 to a     
				For j  =1 to b
				c = l_link(i).GetCellData(x,j)
				print c
				Select Case c
				
				Case "Online Review Deadline"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebElement", 0)
				strRet = oEdit1.getRoProperty("innertext")
				getDeadLine = strRet 
				Exit for
				                                            
				End Select                               
				
				Next
		Next                             
		
		
		
		
		Exit for
		
		End If
		End If
Next
End Function
Function click_OpporutnityInTable()
			On error resume next
			Set odesc=Description.Create
			odesc("micclass").value="WebTable"
			
			'        Set l_link=  getParentObject().ChildObjects(odesc)
			Set l_link= getParentObject().ChildObjects(odesc)
			
			''print l_link.count
			For i = 0 to l_link.count - 1
							strName = l_link(i).GetROProperty("column names")
							''print strName
							strArr = split(strName,";")
							If Not(strName = "") Then
							''print strArr(0)
							If trim(strArr(0)) = "Action" Then
									If trim(strArr(1)) = "Opportunity Name" Then
									a = l_link(i).GetROProperty("rows") 
									b = l_link(i).GetROProperty("cols")                     
									For x  = 1 to a    
									For j  =1 to b
									c =  l_link(i).GetCellData(x,j)
									print c & x & j
									' c2 =  l_link(i).GetCellData(x,j + 2)
									'                                     print c2
									'                                     
									'                                      If c = stragree  And c2 = strstat Then
									Set oEdit1 = l_link(i).ChildItem(2,2, "Link", 0)
									oEdit1.click
									'                                     End If
									
									Next
									Next                            
									
									
									
									
									Exit for
								End if
							
							
					End If
				End If

next
End function

Function fillUserNameInformation()  '(strFirstName, strLastName,strEmailID,strpassword)
     	
     	On error resume next
			   Randomize
         intnumber =  Int((1000000-1)*Rnd+1)
         
          strFirstName = "UFT MR" & "_" & RandomString(1)
          strLastName = " test User" & "_" & RandomString(1)
          strEmailID = "test" & "_" & RandomString(2)&"@yopmail.com"
          'strFirstName = "RAPI" & "_" & RandomString(3)
				Set editO = Description.Create
				editO("micclass").value = "WebEdit"
				editO("visible").value = true               
               
				
				Set O =  getParentObject().ChildObjects(editO) 
				print O.count
				
				For i = 0 to O.count - 1
				
						strName = O(i).getRoProperty("name")
						
						If trim(strName) = "cmntyfrstnm" Then
						 exit for
						
						End If
				Next
				wait 2
				 'First Name
				 O(i).set strFirstName
				 wait 2
				 'Last Name
				 O(i + 1).set strLastName
				 wait 2
				 'Email id
				 O(i + 2).set strEmailID
				 wait 2
				 'Email Confirmation
				 O(i + 3).set strEmailID
				 wait 2
				 'password
				 O(i + 4).set "test1234" 'strpassword
				 wait 2
				 'Password Confirmation
				 O(i  + 5).set "test1234"   'strpassword
				 
				
				
				
                 
				
     End Function
     
     Function clickNewuserlink()
     	On error resume next
     	Set l = Description.Create
        l("micclass").value = "Link"
		Set lo =  getParentObject().ChildObjects(l)
		print lo.count
		For i=0 to lo.count -1
			x = lo(i).getRoproperty("innertext")
			If Trim(x = "New User?") Then
				lo(i).click
				Exit for
			End If
		Next
     End Function
      Function JoinPicori()
     	On error resume next
     	Set l = Description.Create
        l("micclass").value = "Link"
		Set lo =  getParentObject().ChildObjects(l)
		print lo.count
		For i=0 to lo.count -1
			x = lo(i).getRoproperty("innertext")
			If Trim(x = "Join PCORI Portal") Then
				lo(i).click
				Exit for
			End If
		Next
     End Function
     
     Function fillMailingAddress()
     	
     	On error resume next
			   
				Set editO = Description.Create
				editO("micclass").value = "WebEdit"
				editO("visible").value = true  
				
				Set O =  getParentObject().ChildObjects(editO) 
				print O.count
				
				For i = 0 to O.count - 1
				
						strName = O(i).getRoProperty("name")
						
						If trim(strName) = "j_id0:autoCompleteForm:phonenumber" Then
						 exit for
						
						End If
				Next
				
				 ''''''Phone
				 O(i).set "1345678"
				 ''Street
				 'O(i + 1).set "123 main st"
				 ''City
				 'O(i + 2).set "Alexander City"
				 'State/Province
				 'O(i + 3).set "Alabama"
				 ''Country
				 'O(i + 4).set "Afghanistan"
				 ''Zip/Postal Code
				 'O(i + 5).set "20036"
				 ''''Salutation
				  'O(i + 6).set ""
				 
				
				
				
                 
				
     End Function
     
     Function selectRadioGrouplsinusers (strName,i)
     	On error resume next
     	Set editO = Description.Create
				editO("micclass").value = "WebRadioGroup"
				editO("visible").value = true  
				
				Set O =  getParentObject().ChildObjects(editO) 
				print O.count
				
				
				O(i).select strName
     End Function
     Function getDeadLine2()
			On error resume next
			    Set odesc=Description.Create
		        odesc("micclass").value="WebTable"
		        odesc("column names").value="Panel Name.*"
		
		       Set l_link=  getParentObject().ChildObjects(odesc)
		        print l_link.count
		        a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 
				  	       
					For x  = 1 to a     
							For j  =1 to b
							c = l_link(i).GetCellData(x,j)
							print c
							Select Case c
							
							Case "Online Review Deadline"
							Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebElement", 0)
							strRet = oEdit1.getRoProperty("innertext")
							getDeadLine2 = strRet 
							Exit function
							                                            
							End Select                               
							
							Next
					Next                             
					
					
					
					
					
End Function
 Function getDeadLineInPerson()
			On error resume next
			    Set odesc=Description.Create
		        odesc("micclass").value="WebTable"
		        odesc("column names").value="Review ID.*"
		
		       Set l_link=  getParentObject().ChildObjects(odesc)
		        print l_link.count
		        a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 
				  	       
					For x  = 1 to a     
							For j  =1 to b
							c = l_link(i).GetCellData(x,j)
							
							Select Case c
							
							Case "Deadline"
							Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebElement", 0)
							strRet = oEdit1.getRoProperty("innertext")
							getDeadLineInPerson = strRet 
							Exit function
							                                            
							End Select                               
							
							Next
					Next                             
					
					
					
					
					
End Function
Function goTolistLinkReviews(strName)
	On error resume next
	Set odesc=Description.Create
    odesc("micclass").value="Link"
    odesc("html tag").value= "A"   
    odesc("innertext").value = strName           
    'odesc("abx_y").value= "643"
    Set O =  getParentObject().ChildObjects(odesc)
    O(1).click
End Function


Function PopulatePanelNameInView(strName)
			On error resume next
			    Set odesc=Description.Create
		        odesc("micclass").value="WebTable"
		        odesc("column names").value=".*Field;Operator.*"
		
		       Set l_link=  getParentObject().ChildObjects(odesc)
		        print l_link.count
		        a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 
				  	       
					For x  = 1 to a     
							For j  =1 to b
							c = l_link(i).GetCellData(x,j)
							print c
							
							Set oEdit1 = l_link(i).ChildItem(3,1, "WebList", 0)
							oEdit1.select "Panel"
							wait 2
							Set oEdit1 = l_link(i).ChildItem(3,1, "WebList", 0)
							oEdit1.select "equals"
							wait 2
							Set oEdit1 = l_link(i).ChildItem(3,1, "WebEdut", 0)
							oEdit1.set strName
							Exit for
							                           
							
							Next
					Next                             
					
					
					
					
					
End Function

Public Function select_OpneInPerson(strTableName)
			On error resume next
			Set odesc=Description.Create
			odesc("micclass").value="WebElement"
			odesc("visible").value= True
			odesc("html tag").value= "DIV"
			odesc("innertext").value= strTableName
			
			Set O =  getParentObject().ChildObjects(odesc)
			print O.count
			x = O(0).getRoProperty("abs_x")
			print x 
			y= O(0).getRoProperty("abs_y")
			print y  
			
			Set od=Description.Create
			od("micclass").value="WebElement"
			'od("class").value =  "x-grid3-cell-inner x-grid3-col-00N39000003LYgD"
			od("html tag").value="DIV"
			od("abs_x").value= x + 1
			Set Ob =  getParentObject().ChildObjects(od)
			print Ob.count
			
			For i = 1 to Ob.count - 1
				Ob(i).FireEvent "ondblclick"
				wait 2
				inline_checkbox "Open In-Person"
				
				clk_Button_usingName "Save"
				wait 3
			Next

End Function

Function clickReviewSubmitButton()
	On error resume next
	Set odesc=Description.Create
			odesc("micclass").value="WebButton"
			odesc("visible").value= True
			odesc("html tag").value= "INPUT"
			odesc("Name").value= "Review/Submit"
			odesc("class").value = "btn btn btn-primary"
			
			Set O =  getParentObject().ChildObjects(odesc)
			O(0).click
End Function

Function CreateNewReview_Updated(strStatus,strPanel)

                Set odesc=Description.Create
		        odesc("micclass").value="WebTable"
		        odesc("column names").value="Deadline.*"
		
		       Set l_link=  getParentObject().ChildObjects(odesc)
		        print l_link.count
		        a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 
				  	       
					For x  = 1 to a     
							For j  =1 to b
							c = l_link(i).GetCellData(x,j)
								print c
								Select Case c
								
								Case "Reviewer Type"
									Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
									oEdit1.select "In-Person"
								Case "Status"
									Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
									oEdit1.select strStatus
								
								
								
								Case "Reviewer"
									Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
									oEdit1.set reviewer1
								Case "*Application"
									Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
									oEdit1.set strPanel
								'                                    
								'                                                                    
								End Select                               
											                           
							
							Next
					Next                             
					
					

End Function
Function dblclkInPersonNotes()
     	  
		

		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		'printl_link.count
		For i = 0 to l_link.count - 1
		  strName = l_link(i).GetROProperty("column names")
		  ''printstrName
		  strArr = split(strName,";")
		  If Not(strName = "") Then
		  'printstrArr(0)
		  	If trim(strArr(0)) = "Add to Discussion Line" Then
		  	
		  	       a = l_link(i).GetROProperty("rows")  
                   b = l_link(i).GetROProperty("cols") 		  	       
                        For x  = 1 to a     
                             For j  =1 to b
                               c = l_link(i).GetCellData(x,j)
                                  Select Case c
                                  
                                  	Case "In-Person Discussion Notes"
                                  	   Set o = l_link(i).ChildItem(x,j + 1, "WebElement", 0)
                                       o.FireEvent"ondblclick"
                                       Exit function
                                     End select
                             Next
                         Next                        
		  	 
'		    Set oPartNumberLink = l_link(i).ChildItem(3,2, "WebElement", 0)
'
'                 oPartNumberLink.FireEvent"ondblclick"

		    End If
		  End If
		  
		  
			
		Next
     End Function
      Function populateInlinePerson()
    	On error resume next
    	 Set oShell = CreateObject("WScript.Shell")
 Set odesc=Description.Create
           odesc("micclass").value="WebElement"
           odesc("visible").value= True
           odesc("class").value= "cke_button_icon cke_button__bulletedlist_icon"
           Set O =  getParentObject().ChildObjects(odesc)
           print O.count
           wait 5
           O(0).click
      oShell.SendKeys "test note"
    End Function
    
    Function NavigateToInpersonViewAndGetRecordName()
    	On error resume next
    	'Navigate to the application 
		navigateAndLoginToSalesForce UserAdmin, SystemAdminPassword
		
		'click RP community manager
		clk_link_Object2 "Merit Reviewer Management"
		
		'click Reviews Tab
		clk_link_Object2 "Reviews"
	
		'select  MR Prep In Person List view
		selectWeblist "MR Prep for In-Person"
	
		'get the id of the review from the table
		intReviewId = getReviewRecord()	
		
		'click the checkbox in in Person 
		select_onlineReviewcheckbox2 "Open In-Person"
		NavigateToInpersonViewAndGetRecordName = intReviewId 
    End Function
    
    Function searchForTimeBasedWorkflow()
    	On error resume next
    	' go to the set up
		clk_link_Object2 "Setup"
	
		'click the time based workflow
		clk_link_Object2 "Time-Based Workflow"
	
		'click the search button
		clk_Button_usingName2 "Search"
    End Function
    
    Function NavigateToReviewerPortal()
    	On error resume next
    	'Navigate to the reviewer portal and submit the online review score
		Open_ReviewersPortal()
		
		wait 5
		
		'Login to the portal
		login_intoSalesForce_Application userReviewer,passWord
		clk_link_Object2 "Log in"
'		wait 10
'		clk_link_Object2 "Click here to Access the Merit Reviewer Dashboard"
    End Function
    Function createAPanelTestCaseCAP4_1()
       on error resume next
        navigateAndLoginToSalesForce UserAdmin, SystemAdminPassword
wait 2
		'click RP community manager
		clk_link_Object2 "Merit Reviewer Management"
		
		
		clk_link_Object2 "Panels"
		
		clk_Button_usingName "New"
		
		stPanel = create_Panel()
		writeToAfilePanel stPanel
		print stPanel
		fillOnlineReviewCriteria()
		clk_Button_usingName saveBtn
		createAPanelTestCaseCAP4_1 = stPanel

    	
    End Function
    Function testCaseCAP4TestCase2_SelectRecordTypeOfThePanel(stPanel)
    	    On error resume next
    	    navigateAndLoginToSalesForce UserAdmin, SystemAdminPassword
    	    wait 2

			navigateToAccoutsPageThrughTopSearchBox stPanel
			
			clk_link_Object2 stPanel
			
			
			clk_Button_usingName "Edit"
			wait 2			 
			  selectWeblist4 ", Online Review - AD"
			 'selectWeblist ",Online Review - AD"
			 wait 2
			 clk_Button_usingName saveBtn
    End Function
    Function selectApplicationsToPanelCreatedCAP4TestCase3(stPanel)
    	    On error resume next
    	       'navigateAndLoginToSalesForce UserAdmin, SystemAdminPassword
			'''''create Projects
'			wait 2
'			strProj1 =  createApplicationsProjects()
'			writeToAfileProject strProj1
'			print strProj1
'			
'			'''''''''''''''''''''''Note If you need to create 3 projects uncomment the strProj2 and strProj3''''''
'			wait 2
'			strProj2 = createApplicationsProjects()
'			writeToAfileProject1 strProj2
'			print strProj2
'		    strProj3 = createApplicationsProjects()
'			print strProj3
			
			
			
			'click the projects tab and go to the MRA for review view
			wait 2
			 navigateAndLoginToSalesForce cmaOperationsuser, passWordCMA
			 
			 wait 2
			clk_link_Object2 "Projects"
			
			''''select the RA app 2 ready for MR review
			selectWeblistany "fcf","RA App 3 Ready for MR"
			
			strProj1 = readFromFileProject()
            'navigateToAccoutsPageThrughTopSearchBox strData
			
			wait 5
			
			'''''''''Uncomment below line if want to directly add the panel throught application detail page instead of double clicking''''''''''
			
			'insertPanelIntoProjects(stPanel)

			
			
'			'doble click the panel field
			dblClick_PanelFieldtest strProj1
			inlineEdit "Panel",stPanel
			clk_Button_usingName saveBtn
			wait 2
'			
			strProj2 = readFromFileProject1()
			dblClick_PanelFieldtest strProj2
'			
			inlineEdit "Panel",stPanel
			clk_Button_usingName saveBtn
'			
'			dblClick_PanelFieldtest strProj3
'			inlineEdit "Panel",stPanel
'			clk_Button_usingName saveBtn
			wait 3

    End Function
    Function createPanelAssignmentsCAP04TestCase4(stPanel)
    	On error resume next
'    	navigateAndLoginToSalesForce mrManagementUser , somersPassword
'
'wait 4
'			navigateToAccoutsPageThrughTopSearchBox stPanel
'			wait 2
'			
'			clk_link_Object2 stPanel
'			wait 2
'			
'			clk_Button_usingName "New Panel Assignment"
			
			setWebEditBox "," & reviewer1
			
			clk_Button_usingName saveAndNew
			
			setWebEditBox "," & reviewer2
			
			clk_Button_usingName saveAndNew
			
			setWebEditBox "," & reviewer3
			
			clk_Button_usingName saveAndNew
			
			setWebEditBox "," & reviewer4
			
			clk_Button_usingName saveAndNew
			
			setWebEditBox "," & reviewer5		
			
			clk_Button_usingName saveBtn
	  End Function
	  Function unCheckActivePanelAssignmentCAP4TestCase5(stPanel)
	  	On error resume next
	  	'opne the reviewer3 panel assigment and uncheck the active check box
		'click the panel
		navigateAndLoginToSalesForce mrManagementUser , somersPassword
		
		
		navigateToAccoutsPageThrughTopSearchBox stPanel
		
		clk_link_Object2 stPanel
					
		wait 5
		'click the panel assignment
		'click_PanelAssignment()        ''''Have to fix this function not working as expected'''''''
		clk_link_Object2 "PAN-13971"
		  wait 5
		clk_Button_usingName "Edit"
		wait 5
		 clk_singelCheckbox()
		 clk_Button_usingName saveBtn
	  End Function
	  Function ReviewersPortalCOIAndExpertiseCAP4TestCase6()
	  	On error resume next
	  	'open the reviewers portal
				Open_ReviewersPortal()
				
				wait 5
				
				'Login to the portal
				login_intoSalesForce_Application userReviewer1,passWord1
				clk_link_Object2 "Log in"
				wait 5
				verifyIfLinkDoesExist1 "Home"
				verifyIfLinkDoesExist1 "My Profile"
				
				clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
				'clk_link_Object2 "INDICATE COI AND EXPERTISE"
				wait 10
				'click the COI experitise tab
				clk_link_Object2 "COI & Expertise"
'				wait 2
				verifyIfWebElementDoesExist1 "Please click on the PCORI Funding Announcement(s) below to indicate whether you have PFA-level COI(s)."
				verifyIfWebElementDoesExist1 "PCORI Funding Announcement"
				wait 5
				
				'click the program
				clk_link_Object2 "Addressing Disparities Cycle 3"
				
				wait 5
				
				'''''SPS-5679 ( November 1st deployment) COI submission without selecting value'''
				'click submit button
				clk_Button_usingName "Submit"
				wait 2
				verifyIfWebElementDoesExist2 "Please select one option before you hit submit\."
				wait 2
				'''slect no otion'''
				selectRadioButton "NO"
				'click submit button
				clk_Button_usingName "Submit"
				wait 5
'				Note : there is no any go button
'				click the edit link
				clk_link_Object2 "Edit"
				wait 5
				'Verify the follwoing links exist
				verifyIfLinkDoesExist1 "Information for PCORI Merit Reviewers on Confidentiality, Conflict of Interest, and Rating Expertise"
				verifyIfLinkDoesExist1   "PCORI Conflict of Interest Policy"
				verifyIfLinkDoesExist1   "PCORI Non-Disclosure Agreement \(Reviewer\)"
				
				'verify if application content is present
				
				 verifyIfWebElementDoesExist1 "Application Key Personnel"
				 
				 'select the expertise level to high
				 'slect no
				 selectRadioButton "No" 
wait 2				 
				 selectRadioButton1 "High"
				'click submit button
				wait 2
				 clk_Button_usingName "Submit"
				
				 wait 5
				 'clk_link_Object2 "Edit"
				 'submit another question on the list
				 'selectRadioButton "Personal"
				 'selectRadioButton2 "None or Not applicable"
				 'clk_Button_usingName "Submit"
				 
				 
				 			 
				 
				 
				 
				 
	  End Function
	  Function checkIfROIisUpdatedCAP4TestCase7(stPanel)
	  	On error resume next
	  	'Log in As a Management User and verify the ROI submitted above is updated in the panel
 
		 navigateAndLoginToSalesForce mrManagementUser, SomersPassword
		
		'click RP community manager
		clk_link_Object2 "Merit Reviewer Management"
		
		navigateToAccoutsPageThrughTopSearchBox stPanel
		
		clk_link_Object2 stPanel
		
		'click the panel assigment that is updated in test case 6
		'click_PanelAssignment()
		'click_COIExpertise()
		clk_link_Object2 "Go to list \(18\) »"
		
	  End Function
	  Function finalizeThePanelCAP4TestCase8(stPanel)
	  	On error resume next
	  	navigateAndLoginToSalesForce mrManagementUser, somersPassword
	  	
	  	  wait 5
	  	  navigateToAccoutsPageThrughTopSearchBox stPanel
	  	  wait 5
		 'click the panel and edit the Panel finalized checkbox
		 'click the panel
		 clk_link_Object2 stPanel
		 wait 5
		 'click edit button to edit the panel
		  wait 5
		 clk_Button_usingName "Edit"
		 'click the panel finalized and save the panel
		 wait 5
		 clk_singelCheckboxbasedonindexid()
		 clk_Button_usingName saveBtn
	  End Function
	  Function createRAAssighnmentMatrixViewCAP4TestCase9(stPanel)
	  	On error resume next
	  	navigateAndLoginToSalesForce mrManagementUser, passWord1
		  strViewName = "RA App Assignment Matrix" & space(1) & stPanel
		  'write this name to flat file
		  writeToAfileRAMatrixView strViewName
		  wait 4
		  clk_link_Object2 "Projects"
		  wait 2
		
		  selectWeblist "RA App Assignment Matrix"
		  wait 3
		
		  clk_link_Object2 "Edit"
		
		  'clone the view
		  CloneView strViewName
		  'enter crtieria row for panel name
		  fillCriterialRowPanelName stPanel
		  clk_Button_usingName "Save As"
		  
		  createRAAssighnmentMatrixViewCAP4TestCase9 = strViewName
	  End Function
	  Function populateReviewersIntoTheRecordsCAP4TestCase10(strViewName)
	  	On error resume next
	  	
  	navigateAndLoginToSalesForce mrManagementUser, somersPassword
	  	wait 2
		clk_link_Object2 "Projects"
		wait 5
		strViewName = readFromRAMatrixView()
		selectWeblistany "fcf",strViewName
		
		wait 5
		'populate the reviewers
		'insertReviewersIntoProjects(strViewName)   ''''this is when double clicking doesn't work'''''
		
		''''''Below is good working one for double clicking''''
		set_editlinesForReviewers2 "Reviewer 1",reviewer1
		set_editlinesForReviewers2 "Reviewer 2",reviewer2
		set_editlinesForReviewers "Reviewer 3",reviewer3
		set_editlinesForReviewers "Reviewer 4",reviewer4
		set_editlinesForReviewers "Reviewer 5",reviewer5
	
		'select the create online reviews
		select_onlineReviewcheckbox  "Create Online Reviews" 

	  End Function
	  Function getCountOfReviewsWithCOIConflictInterest()
     	On error resume next
     	 Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("cols").value= 6
		Set oT =  getParentObject().ChildObjects(odesc)
		print oT.Count
		x = oT(0).GetRoProperty ("rows")
		print x
		     For i = 1 To x
		     	 c =  oT(0).GetCellData(i,5)
		     	 print c
                     If c = "No" Then
                     	counter = counter + 1
                     End If
                   print c 
		     Next
                    
		    
	
	
		 getCountOfReviewsWithCOIConflictInterest = counter
     End Function
     Function countActiveCOI()
     	On error resume next
     	 Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("cols").value= 6
		Set oT =  getParentObject().ChildObjects(odesc)
		print oT.Count
		x = oT(0).GetRoProperty ("rows")
		print x
		     For i = 1 To x
		         Set oEdit1 = oT(0).ChildItem(i,6, "Image", 0)
		     	 c = oEdit1.GetRoProperty("alt")
		     	 print c
                     If Trim(c) = "Checked" Then
                     	counter = counter + 1
                     End If
                   print c 
		     Next
                    
		    
	
	
		 countActiveCOI = counter
     End Function
     Function getCountOfReviewsInAPanel()
     	On error resume next
     	 Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("cols").value= 10
		Set oT =  getParentObject().ChildObjects(odesc)
		print oT.Count
		x = oT(1).GetRoProperty ("rows")
		print x
		     For i = 1 To x
		     	 c =  oT(1).GetCellData(i,2)
		     	 print c
		     	 next
		     	 
		     	 '''''''Uncomment this if need to count the in-person preview record'''
                     'If c = "In-Person Preview" Then
                     'If c = "Review ID" Then
                     
                     	'counter = counter - 1
                     	
                     	
                     'End If
                 ' print c 
		     'Next
                    
		    
	
	
		 'getCountOfReviewsInAPanel = counter
     End Function
     
     Function UnselectOnlineReviewCheckbox()
			On error resume next
			    Set odesc=Description.Create
		        odesc("micclass").value="WebTable"
		        odesc("column names").value="Reviewer 1.*"
		
		       Set l_link=  getParentObject().ChildObjects(odesc)
		        print l_link.count
		        a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 
				  	       
					For x  = 1 to a     
							For j  =1 to b
							c = l_link(i).GetCellData(x,j)
							print c
							
							Select Case c
								Case "Create Online Reviews"
									Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebCheckBox", 0)
									oEdit1.click
									Exit function
									
							End Select
							
							
							
							Next
					Next                             
					
					
					
					
					
End Function
Function getReviewIDInReview()
On error resume next
strNameEvent = "projAUT" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
		print strArr(0)
		If trim(strArr(0)) = "Review ID" Then
		l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 		  	       
		For x  = 1 to a     
				For j  =1 to b
				c = l_link(i).GetCellData(x,j)
				print c
				Select Case c
				
				Case "Review ID"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebElement", 0)
				strRet = oEdit1.getRoProperty("innertext")
				getReviewIDInReview = strRet 
				Exit for
				                                            
				End Select                               
				
				Next
		Next                             
		
		
		
		
		Exit for
		
		End If
		End If
Next
End Function
Function ChangeMRO(strstatus)
On error resume next
strNameEvent = "1AprojAUT" & "_" & RandomString(3) 
strArrN = split(straddinfo,",")
s = Date
Set odesc=Description.Create
odesc("micclass").value="WebTable"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count
For i = 0 to l_link.count - 1
		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
		print strArr(0)
		If trim(strArr(0)) = "Review ID" Then
		l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 		  	       
		For x  = 1 to a     
				For j  =1 to b
				c = l_link(i).GetCellData(x,j)
				print c
				Select Case c
				
				Case "MRO Tracking Status"
				Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
				 oEdit1.select strstatus
				 
				Exit for
				                                            
				End Select                               
				
				Next
		Next                             
		
		
		
		
		Exit for
		
		End If
		End If
Next
End Function

Function click_Panel_Edit()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebTable"

'        Set l_link=  getParentObject().ChildObjects(odesc)
Set l_link= getParentObject().ChildObjects(odesc)

''print l_link.count
For i = 0 to l_link.count - 1
strName = l_link(i).GetROProperty("column names")
print strName
strArr = split(strName,";")
If Not(strName = "") Then
''print strArr(0)
If trim(strArr(2)) = "Action" Then
If trim(strArr(3)) = "Short Project Title" Then
a = l_link(i).GetROProperty("rows") 
b = l_link(i).GetROProperty("cols")                     
For x  = 1 to a    
For j  =1 to b
c =  l_link(i).GetCellData(x,j)
print c & x & j
' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
'                                      If c = stragree  And c2 = strstat Then
Set oEdit1 = l_link(i).ChildItem(2,3, "Link", 0)
oEdit1.click
Exit function 
'                                     End If

Next
Next                            




Exit for
End if


End If
End If

next
End function


 Function populateKeyProjectPersonel(strPanel)
				On error resume next
		 Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("column names").value="Reviewer 1.*"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 		  	       
				For x  = 1 to a     
						For j  =1 to b
							c = l_link(0).GetCellData(x,j)
							print c
							Select Case c
							
							
						    	
                                    Case "Panel"
                                 	     Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                         oEdit1.set strPanel
                                         Exit function
                                 	
                                    
							                                                                    
							End Select     
                         						
						
						Next
						
				Next                               
				
	

				 
     End Function 
     
     Function insertPanelIntoProjects(stPanel)
     	On error resume next
     	wait 5
     	clk_link_Object2 "Projects"
     	For i=1 to 3
		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("cols").value= 9
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		Set oEdit1 = l_link(i).ChildItem(1,4, "Link", 0)
		    oEdit1.click
		   wait 5
		   clk_Button_usingName "Edit"
		    
		    wait 5
		     populateKeyProjectPersonel stPanel
		     clk_Button_usingName "Save"
		     wait 5
		     clk_link_Object2 "Projects"
		     
		     wait 3
			
			'select the RA app 2 ready for MR review
			'selectWeblist "RA App 3 Ready for MR"
			
			wait 5
		     
        
         Next
     End Function
       Function insertReviewersIntoProjects(strViewName)
     	On error resume next
     	'For i=1 to 3
     	wait 5
		'Set odesc=Description.Create
		'odesc("micclass").value="WebTable"
		'odesc("cols").value= 12
		
		'Set l_link=  getParentObject().ChildObjects(odesc)
		'print l_link.count
		'Set oEdit1 = l_link(i).ChildItem(1,4, "Link", 0)
		    'oEdit1.click
		    stProject = readFromFileProject()
		    clk_link_Object2 stProject
'		    
'		    
		   clk_Button_usingName "Edit"
'		    
		    wait 5
		    populateKeyProjecReviewers()
'		    
		    wait 3
		    clickOnlineReviews_new()
'		    wait 3
		     clk_Button_usingName "Save"
           wait 5

            click_webElementExternalReviewlistinaproject()
            wait 2
            verifyIfWebElementDoesExist1 "Online Review - AD"

		     wait 5
		     clk_link_Object2 "Projects"
		     
'		     wait 3
			
			'select the RA app 2 ready for MR review
			strViewName = readFromRAMatrixView
			HtmlID = "fcf"
            strData = strViewName
		    selectWeblist5 HtmlID,strData
			
			'selectWeblist strViewName
			
			wait 8
		     
        
         'Next
     End Function
     
     Function populateKeyProjecReviewers()
				On error resume next
		 Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("column names").value="Reviewer 1.*"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 		  	       
				For x  = 1 to a     
						For j  =1 to b
							c = l_link(0).GetCellData(x,j)
							print c
							Select Case c
							
							
						    	
                                         Case "Reviewer 1"
	                                 	     Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
	                                         oEdit1.set reviewer1
                                         Case "Reviewer 2"
	                                 	     Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
	                                         oEdit1.set reviewer2
                                         Case "Reviewer 3"
	                                 	     Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
	                                         oEdit1.set reviewer3
                                         Case "Reviewer 4"
	                                 	     Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
	                                         oEdit1.set reviewer4
                                         Case "Reviewer 5"
	                                 	     Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
	                                         oEdit1.set reviewer5
                                       
                                 	
                                    
							                                                                    
							End Select     
                         						
						
						Next
						
				Next                               
				
	

				 
     End Function 
      Function clickOnlineReviews_new()
				On error resume next
		 Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("column names").value="Reviewer 1.*"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 		  	       
				For x  = 1 to a     
						For j  =1 to b
							c = l_link(0).GetCellData(x,j)
							print c
							Select Case c
							
							
						    	
                                         Case "Create Online Reviews"
	                                 	     Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebCheckBox", 0)
	                                         oEdit1.click
	                                         Exit function
                                        
                                       
                                 	
                                    
							                                                                    
							End Select     
                         						
						
						Next
						
				Next                               
				
	

				 
     End Function 
     
Public Function verifyIfLinkDoesExist2( strName)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)

print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"




Set l = Description.Create
l("micclass").value = "Link"
l("html tag").value = "A"
l("innertext").value = strName
Set lo =  getParentObject().ChildObjects(l)
If lo.count > 1 Then   	
LogReport 1,"verityLinkDoesnotExist - ", "The Link --" & strName & "- does not exist _not Expected"
else
LogReport 0,"verityLinkDoesExist - ", "The Link --" & strName & "-does Exists -  Expected"

End If



End Function
Public Function searchForReports(strData)    

wait 5
setWebEditBox(strData)
'srch_SearchBox(strData)
'wait 5
'clk_Button_usingName "Search"

End Function   

Function click_PrievewInFile()
    On error resume next
    Set odesc=Description.Create
        odesc("micclass").value="WebTable"
       
'        Set l_link=  getParentObject().ChildObjects(odesc)
         Set l_link= getParentObject().ChildObjects(odesc)

        ''print l_link.count
        For i = 0 to l_link.count - 1
          strName = l_link(i).GetROProperty("column names")
          ''print strName
          strArr = split(strName,";")
          If Not(strName = "") Then
          ''print strArr(0)
              If trim(strArr(0)) = "Action" Then
                       If trim(strArr(1)) = "Title" Then
                             a = l_link(i).GetROProperty("rows") 
                           b = l_link(i).GetROProperty("cols")                     
                        For x  = 1 to a    
                             For j  =1 to b
                                      c =  l_link(i).GetCellData(x,j)
                                      print c & x & j
                                    ' c2 =  l_link(i).GetCellData(x,j + 2)
'                                     print c2
'                                     
'                                      If c = stragree  And c2 = strstat Then
                                                   Set o = l_link(i).ChildItem(2,1, "Link", 0)
                                           o.click
'                                           print s
'                                     End If
                             
                             Next
                         Next                            
                        
            
            
                 
                Exit for
                       End if
                    
 
            End If
          End If
         
          next
    End function
    
    Function NavigateToReviewerPortalnewuser()
    	On error resume next
    	'Navigate to the reviewer portal and submit the online review score
		Open_ReviewersPortal()
		
		wait 5
		clickNewuserlink()
		'Login to the portal
		'login_intoSalesForce_Application userReviewer,passWord
		'clk_link_Object2 "Log in"
'		wait 10
'		clk_link_Object2 "Click here to Access the Merit Reviewer Dashboard"
    End Function
    
 Function Clickingbackonbrowser()
         
   On error resume next
   Set oWshShell = CreateObject("WScript.Shell")
   oWshShell.SendKeys "{BS}"  
    
                 
End Function

Function clk_link_Object(stTag,StText)
On error resume next

wait 5
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "Link"
l("html tag").value = stTag
l("innertext").value = StText
l("visible").value = True
Set lo =  getParentObject().ChildObjects(l)
print lo.count
lo(0).Highlight
lo(0).click

    If err <> 0 Then
  	LogReport 4,"Error", err.number & "-" & err.description
  	LogReport 1,"cliking the link" & StText,"Failed to click the link" & "-" & StText
  	else
  	LogReport 0,"cliking the link" & StText,"The Link" & "-" & StText & "-" & "is clicked succesfully"
  	End If
End Function

Function CloseLatestOpenedBrowser()
On error resume next
Dim oDescription
Dim BrowserObjectList
Dim oLatestBrowserIndex

Set oDescription=Description.Create
oDescription("micclass").value="Browser"
Set BrowserObjectList=Desktop.ChildObjects(oDescription)
oLatestBrowserIndex=BrowserObjectList.count-1

Browser("creationtime:="&oLatestBrowserIndex).close

Set oDescription=Nothing
Set BrowserObjectList=Nothing

End Function

Function fillUserNameInformationwithoutpassword()  '(strFirstName, strLastName,strEmailID,strpassword)
     	
     	On error resume next
			   Randomize
         intnumber =  Int((1000000-1)*Rnd+1)
         
          strFirstName = "MR" & "_" & RandomString(3)
          strLastName = "User" & "_" & RandomString(1)
          strEmailID = "test" & "_" & RandomString(2)&"@yopmail.com"
          'strFirstName = "RAPI" & "_" & RandomString(3)
				Set editO = Description.Create
				editO("micclass").value = "WebEdit"
				editO("visible").value = true               
               
				
				Set O =  getParentObject().ChildObjects(editO) 
				print O.count
				
				For i = 0 to O.count - 1
				
						strName = O(i).getRoProperty("name")
						
						'If trim(strName) = "communitiesSelfRegPage:theForm:firstName" Then
						If trim(strName) = "cmntyfrstnm" Then
						 exit for
						
						End If
				Next
				wait 2
				 'First Name
				 O(i).set strFirstName
				 wait 2
				 'Last Name
				 O(i + 1).set strLastName
				 wait 2
				 'Email id
				 O(i + 2).set strEmailID
				 wait 2
				 'Email Confirmation
				 O(i + 3).set strEmailID
				 wait 2
				 'password
				 'O(i + 4).set "test1234" 'strpassword
'				 wait 2
'				 'Password Confirmation
'				 O(i  + 5).set "test1234"   'strpassword
				 
				
				
				
                 
				
     End Function
     
     Function fillUserNameInformationwithwrongconfirmpassword()  '(strFirstName, strLastName,strEmailID,strpassword)
     	
     	On error resume next
			   Randomize
         intnumber =  Int((1000000-1)*Rnd+1)
         
          strFirstName = "RAPI" & "_" & RandomString(3)
          strLastName = "User" & "_" & RandomString(1)
          strEmailID = "test" & "_" & RandomString(2)&"@yopmail.com"
          'strFirstName = "RAPI" & "_" & RandomString(3)
				Set editO = Description.Create
				editO("micclass").value = "WebEdit"
				editO("visible").value = true               
               
				
				Set O =  getParentObject().ChildObjects(editO) 
				print O.count
				
				For i = 0 to O.count - 1
				
						strName = O(i).getRoProperty("name")
						
						'If trim(strName) = "communitiesSelfRegPage:theForm:firstName" Then
						If trim(strName) = "cmntyfrstnm" Then
						 exit for
						
						End If
				Next
				wait 2
				 'First Name
				 O(i).set strFirstName
				 wait 2
				 'Last Name
				 O(i + 1).set strLastName
				 wait 2
				 'Email id
				 O(i + 2).set strEmailID
				 wait 2
				 'Email Confirmation
				 O(i + 3).set strEmailID
				 wait 2
				 'password
				 O(i + 4).set "test1234" 'strpassword
				 wait 2
'				 'Password Confirmation
				 O(i  + 5).set "test12345"   'strpassword
				 
				
				
				
                 
				
     End Function
     
    Function selectewebCheckbox()
			On error resume next
			    Set odesc=Description.Create
		        odesc("micclass").value="WebCheckBox"		        
		
		       Set l_link=  getParentObject().ChildObjects(odesc)
		        print l_link.count
		        
							
			Set oEdit1 = l_link(i).ChildItem( "WebCheckBox", 0)
			oEdit1.click
			Exit function		
							
							
			                     
									
					
					
					
End Function 

Function ClickWebCheckbox()
         
         On error resume next
    Set MyBrowser = Browser("micclass:=Browser").Page("micclass:=Page")
    
'Set oWebckbox = Description.Create
'oWebckbox("micclass").Value = "WebCheckBox"
'oWebckbox("visible").Value = True
'oWebckbox ("name").Value= "communitiesSelfRegPage:theForm:j_id48"
'oWebckbox("type").Value= "checkbox"
'oWebckbox("html tag").Value= "INPUT"'
'MyBrowser.oWebckbox.click
wait 2
MyBrowser.WebCheckBox("html tag:=INPUT","name:=communitiesSelfRegPage:theForm:j_id48", "title:=Acknowledgement required").Click


End Function

Function ClickonApplication()
     	On error resume next
     	For i=1 to 3
     	wait 5
		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("cols").value= 12
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		Set oEdit1 = l_link(i).ChildItem(1,4, "Link", 0)
		    oEdit1.click
		   wait 5
'click review record based on status
strRecord = clickReviewRecordBasedOnStatus ("Ready to Review")
'verify the populated fields
wait 5
'verify the MRO Trucking  is not submited
verifyIfWebElementDoesExist1 "Not Submitted"
'verify the record type
verifyIfWebElementDoesExist1 "Online Review - AD "
'verify the reviewer
verifyIfWebElementDoesExist1 "Merit Review Tester 2"

'click the application record link
OpenApplicationFromReviewRecord()    
  
		 
	    
		    
		    
'		   clk_Button_usingName "Edit"
'		    
'		    wait 5
'		    populateKeyProjecReviewers()
'		    
'		    wait 3
'		    clickOnlineReviews_new()
'		    wait 3
'		     clk_Button_usingName "Save"
		     wait 5
		     clk_link_Object2 "Projects"
		     
		     wait 3
			
			'select the RA app 2 ready for MR review
			selectWeblist strViewName
			
			wait 8
		     
        
         Next
     End Function
'''''''''''''''''''''''''''New Function added by sabiha'''''''''''''''''''''01/22/2017'''''''''''''''''''''' ''''''''''''''''
'''''''''''''''''''''''''This is for external portal selecting Country
 Function SelectWebList()  

On error resume next

wait 4
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strhwnd

Set pa = Description.Create
pa("micclass").value = "Page"
print hwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
'editO("html tag").value = "SELECT" 
'editO("type").value = "ComboBox Select" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count

edObj(i+1).Select "Brazil"

Wait 3

End Function


Function VerifystateProvincefield()  

On error resume next

wait 4
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strhwnd

Set pa = Description.Create
pa("micclass").value = "Page"

print hwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true 
editO("selection").value = "Alabama"
editO("html tag").value = "SELECT" 
editO("select type").value = "ComboBox Select" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
If edObj.count = 1 Then   	
LogReport 0,"verifydefaultvaluedoesexist - ", "Alabama --" & strData & "- does  exist _ Expected"
else
LogReport 1,"verifydefaultvaluedoesnotexist - ", "Alabama --" & strData& "-does not Exists -not  Expected"
End if

Wait 3
End Function

Function fillMailingAddress1()
     	
     	On error resume next
			   
				Set editO = Description.Create
				editO("micclass").value = "WebEdit"
				editO("visible").value = true  
				
				Set O =  getParentObject().ChildObjects(editO) 
				print O.count
				
				For i = 0 to O.count - 1
				
						strName = O(i).getRoProperty("name")
						
						If trim(strName) = "j_id0:autoCompleteForm:phonenumber" Then
						 exit for
						
						End If
				Next
				
				 ''''''Phone
				 O(i).set "134-567-8910"
				 ''Street
				 'O(i + 1).set "123 main st"
				 ''City
				 'O(i + 2).set "Alexander City"
				 'State/Province
				 'O(i + 3).set "Alabama"
				 ''Country
				 'O(i + 4).set "Afghanistan"
				 ''Zip/Postal Code
				 'O(i + 5).set "20036"
				 ''''Salutation
				  'O(i + 6).set ""
				 
				
				
				
                 
				
     End Function
     
Function SelectWebListUSA()  

On error resume next

wait 4
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strhwnd

Set pa = Description.Create
pa("micclass").value = "Page"
print hwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
'editO("html tag").value = "SELECT" 
'editO("type").value = "ComboBox Select" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count

edObj(i+1).Select "USA"

Wait 3

End Function

Function fillMailingAddress2()
     	
     	On error resume next
			   
				Set editO = Description.Create
				editO("micclass").value = "WebEdit"
				editO("visible").value = true  
				
				Set O =  getParentObject().ChildObjects(editO) 
				print O.count
				
				For i = 0 to O.count - 1
				
						strName = O(i).getRoProperty("name")
						
						If trim(strName) = "j_id0:autoCompleteForm:phonenumber" Then
						 exit for
						
						End If
				Next
				
				 ''''''Phone
				 O(i).set "123-456-7890"
				 ''Street
				 O(i + 1).set "123 main st"
				 ''City
				 O(i + 2).set "Alexander City"
				 
				 O(i + 3).set "20036"
				 
				 				
				
     End Function
     
     Function fillMailingAddress4()
     	
     	On error resume next
			   
				Set editO = Description.Create
				editO("micclass").value = "WebEdit"
				editO("visible").value = true  
				
				Set O =  getParentObject().ChildObjects(editO) 
				print O.count
				
				For i = 0 to O.count - 1
				
						strName = O(i).getRoProperty("name")
						
						If trim(strName) = "j_id0:autoCompleteForm:phonenumber" Then
						 exit for
						
						End If
				Next
				
				 ''''''Phone
				 O(i).set "134-567-2578"
				 ''Street
				 O(i + 1).set "123 test st"
				 ''City
				 O(i + 2).set "Washington DC"							 
				 ''Zip/Postal Code
				 O(i + 3).set "20036"
				 '''Country'''
				 'SelectWebListUSA()				 
				 wait 2		
				 
				 '''''State/Province
				 Selectstateprovince()	
				 '''''''''Salutation
				  wait 2
				  'selectRadioButton1 " Mr."
				  wait 2	 		
							
                 
				
     End Function
     
     Function Selectstateprovince()  

On error resume next

wait 4
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strhwnd

Set pa = Description.Create
pa("micclass").value = "Page"
print hwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
'editO("html tag").value = "SELECT" 
'editO("type").value = "ComboBox Select" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count

edObj(i).Select "District of Columbia"

Wait 3

End Function
Function selectRadioButton4(strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True           

Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(2).Select strData
End Function
Public Function verifyIfWebEditDoesExist(strName)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebEdit"
l("html tag").value = "INPUT"
l("innertext").value = strName
l("visible").value = true 

strcount = l.count
print strcount
' Set lo =  getParentObject().ChildObjects(l)
If l.count = 0 Then      
LogReport 1,"verifyIfWebElementDoesnotExist1 - ", "The WebEdit --" & strName & "- doesnot exist - not Expected"
else
LogReport 0,"verifyIfWebElementDoesExist1 - ", "The WebEdit --" & strName & "-Exists -  Expected"

End If


End Function

Function selectRadioButton5(strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True           

Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(3).Select strData
End Function
Function selectRadioButton6(strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True           

Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(4).Select strData
End Function
Function Verifyyearofbirthfield(strData)  

On error resume next

wait 4
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strhwnd

Set pa = Description.Create
pa("micclass").value = "Page"

print hwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true 
'editO("selection").value = "Alabama"
editO("html tag").value = "SELECT" 
editO("select type").value = "ComboBox Select" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
If edObj.count = 1 Then   	
LogReport 1,"verifydefaultvaluedoesexist - ", "WebList --" & strData & "- does  exist _Not Expected"
else
LogReport 0,"verifydefaultvaluedoesexist - ", "WebList--" & strData& "-does Exists -  Expected"
End if

Wait 3
End Function

Function selectRadioButton7(strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True           

Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(5).Select strData
End Function

Function selectRadioButton8(strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True           

Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(6).Select strData
End Function

Function selectRadioButton9(strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True           

Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(7).Select strData
End Function

Function Employerfoundradiobutton()
     	
     	On error resume next
			   
				Set editO = Description.Create
				editO("micclass").value = "WebEdit"
				editO("visible").value = true  
				
				Set O =  getParentObject().ChildObjects(editO) 
				print O.count
				
				For i = 0 to O.count - 1
				
						strName = O(i).getRoProperty("name")
						
						If trim(strName) = "j_id0:autoCompleteForm:Company" Then
						 exit for
						
						End If
				Next
				
				 ''''''Employer Name Lookup
				 O(i).set "PCORI"
				 ''''''Position/Title *''''
				 O(i + 1).set "Tester"
				 ''''''Department''''''' 
				 O(i + 2).set "Information Technology"							 
				 
				 
				
							
                 
				
     End Function
     
     Function ReviewersPortalCOIAndExpertiseMRTC19()
	  	On error resume next
	  	'open the reviewers portal
				Open_ReviewersPortal()
				
				wait 5
				
				'Login to the portal
				login_intoSalesForce_Application userReviewer,passWord1
				clk_link_Object2 "Log in"
				wait 5
				verifyIfLinkDoesExist1 "Home"
				verifyIfLinkDoesExist1 "My Profile"
				
				clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
				'clk_link_Object2 "INDICATE COI AND EXPERTISE"
				wait 10
				'click the COI experitise tab
				clk_link_Object2 "COI & Expertise"
				wait 2
				verifyIfWebElementDoesExist1 "Please click on the PCORI Funding Announcement(s) below to indicate whether you have PFA-level COI(s)."
				verifyIfWebElementDoesExist1 "PCORI Funding Announcement"
				wait 5
				
				''''''''click the program'''''Manually update the program name so it can click the right program
				clk_link_Object2 "Addressing Disparities Cycle 3"
				wait 2
				verifyIfWebElementDoesExist1 "To the best of your knowledge, are you or a close family member listed as key personnel on any applications submitted to this panel?"
			    verifyIfWebElementDoesExist1 "Yes"
				verifyIfWebElementDoesExist1 "No"
				verifyIfWebElementDoesExist1 "Unclear"
				
				wait 5
				'slect no
				selectRadioButton "No"
				'click submit button
				clk_Button_usingName "Submit"
				wait 5
				''''''''''''''Verifying the PCORI Funding Announcement from previous step for the submitted PFA Level COI"""""""
			    verifyIfWebElementDoesExist1 "Addressing Disparities Cycle 3"
'				Note : there is no any go button
'				click the edit link
				clk_link_Object2 "Edit"
				wait 5
				'Verify the follwoing links exist
				verifyIfLinkDoesExist1 "Information for PCORI Merit Reviewers on Confidentiality, Conflict of Interest, and Rating Expertise"
				verifyIfLinkDoesExist1   "PCORI Conflict of Interest Policy"
				verifyIfLinkDoesExist1   "PCORI Non-Disclosure Agreement (Reviewer)"
				
				'verify if application content is present
				
				 verifyIfWebElementDoesExist1 "Application Key Personnel"
				 
				 'select the expertise level to high
				 'slect no
				 selectRadioButton "No"   
				 selectRadioButton2 "High"
				'click submit button
				 clk_Button_usingName "Submit"
				
'				 wait 5
'				 clk_link_Object2 "Edit"
'				 'submit another question on the list
'				 selectRadioButton "Personal"
'				 selectRadioButton2 "None or Not applicable"
'				 clk_Button_usingName "Submit"
				 End Function
				 
Function selectWeblist4(strData)
On error resume next

wait 4
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"



print pHwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
'editO("html tag").value = "INPUT" 
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
for i = 0 To uBound(strArr)
items =   edObj(i).GetRoProperty("all items")
print items
print strArr(i)
If Not(strArr(i) = "")Then	


edObj(i).select trim(strArr(i))        
End If

Next


End Function

Function clk_Web_Elementusingname(strName)

On error resume next
wait 4
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

print pHwnd

Set editO = Description.Create

editO("micclass").value = "WebElement"
editO("visible").value = true  
editO("name").value = strName 
Set editObject =  getParentObject().ChildObjects(editO) 
print editObject.count
editObject(4).highlight
editObject(4).click



If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
else
LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
End If
End Function

Function selectWeblistinternal(strData)
			On error resume next

            wait 4
			   strArr = split(strData,",")
			   
			   Set btncalc = Description.Create()  
			    btncalc("micclass").value = "Browser"
			  
			    Set btn =DeskTop.ChildObjects(btncalc)
			    strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
			    print strHwnd
			  
			    Set pa = Description.Create
			    pa("micclass").value = "Page"
			  
			   
			    
			    print pHwnd
			    Set editObject =  getParentObject()
			    
			    Set editO = Description.Create
			    
			    editO("micclass").value = "WebList"
			    editO("visible").value = true  
			    'editO("name").value = "File Share Settings" 
			    'editO("type").value = "text" 
			     Set edObj =  getParentObject().ChildObjects(editO) 
			     print edObj.count
			     for i = 0 To uBound(strArr)
			        items =   edObj(i).GetRoProperty("all items")
			        print items
			        print strArr(i)
			        If Not(strArr(i) = "")Then	
                        
                                             
			            edObj(i).select trim(strArr(i))        
			        End If
			         
			     Next
			           
         
     End Function
     
     Public Function verifyIfButtonDoesExist(strName)
               On error resume next
               Set btncalc = Description.Create() 
                                btncalc("micclass").value = "Browser"
                               
                                Set btn =DeskTop.ChildObjects(btncalc)
                               
                                ''print strHwnd
                               
                                Set pa = Description.Create
                                pa("micclass").value = "Page"
                               
                                
                                
                                
                                Set l = Description.Create
                                l("micclass").value = "WebButton"
                                l("html tag").value = "BUTTON"
                                l("name").value = strName
                                Set lo =  getParentObject().ChildObjects(l)
                                 If lo.count = 1 Then      
                                       LogReport 0,"verity Button Does  Exist- ", "The WebButton --" & strName & "-Exists -  Expected"
                                       else                                       
                                      LogReport 1,"verity Button Does Not Exist - ", "The WebButton --" & strName & "- doesnot exist _Not Expected"                    
                                             End If
                             
              
                             
     End Function
'''''''''''''''''''''''Good function to fill out form with webedit box and lookup icon present in the field to search for a specific record''''''''     
     Function MentorEvaluationdetail()
                On error resume next

            'strNameEvent = "CampaignPAST_AD_Test_Automation" & "_" & RandomString(3) 
            strArrN = split(straddinfo,",")
            
            'DATE of today + 30 days - GOOD ONE
            
            Set odesc=Description.Create
            odesc("micclass").value="WebTable"
            
            Set l_link=  getParentObject().ChildObjects(odesc)
            print l_link.count
            For i = 0 to l_link.count - 1
            strName = l_link(i).GetROProperty("column names")
            print strName
            strArr = split(strName,";")
            If Not(strName = "") Then
            print strArr(0)
            If trim(strArr(0)) = "Owner" Then
            l_link(i).GetROProperty("rows") 
            
            a = l_link(i).GetROProperty("rows")  
            b = l_link(i).GetROProperty("cols")                      
            For x  = 1 to a     
            For j  = 1 to b
            c = l_link(i).GetCellData(x,j)
            print c
            Select Case c
                Case "*Reviewer Name"
            Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
            oEdit1.set "Robin reviewer 1"
                Case "*Mentor Name"
            Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
            oEdit1.set "Benjamin Somers"
                Case "*Panel"
            Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
            wait 2
            stPanel= readFromFilePanel
            oEdit1.set stPanel

                                  
            End Select                               
            
            Next
            Next                             
                
            Exit for
            
            End If
            End If
            Next
            
           
End Function


'''''''''''Good Function to fill out internal form with Webedit box''''''''''''''''''''''''
Function setWebEditBox_EvaluationForm(Comments1, Comments2, Comments3, Comments4, Comments5, Comments6, Comments7, Comments8, Comments9, Comments10, Comments11)
On error resume next

wait 3
strArr = split(Comments1,",")

Set oShell = CreateObject("WScript.Shell") 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

print pHwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebEdit"
editO("visible").value = true  
'editO("html tag").value = "INPUT" 
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 

for i = 0 To uBound(strArr)
'id =   edObj(i).GetRoProperty("html id")
If Not(strArr(i) = "")Then

edObj(i+4).set Comments1  
edObj(i+5).set Comments2
edObj(i+6).set Comments3
edObj(i+7).set Comments4
edObj(i+8).set Comments5
edObj(i+9).set Comments6
edObj(i+10).set Comments7
edObj(i+11).set Comments8
edObj(i+12).set Comments9
edObj(i+13).set Comments10
edObj(i+14).set Comments11

oShell.SendKeys strSearchText				            
End If

Next
If err <> 0 Then
err.clear   

End If
End Function
''''''''''''''''Good function to fill out form with Weblist''''''''''
Function SelectWebListEV()  

On error resume next

wait 4
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strhwnd

Set pa = Description.Create
pa("micclass").value = "Page"
print hwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
'editO("html tag").value = "SELECT" 
'editO("type").value = "ComboBox Select" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
wait 3

edObj(i).Select "Met Expectations"
edObj(i+1).Select "Met Expectations"
edObj(i+2).Select "Met Expectations"
edObj(i+3).Select "Met Expectations"
edObj(i+4).Select "Met Expectations"
edObj(i+5).Select "Met Expectations"
edObj(i+6).Select "Met Expectations"
edObj(i+7).Select "Met Expectations"
edObj(i+8).Select "Met Expectations"
edObj(i+9).Select "Met Expectations"
edObj(i+10).Select "Yes"
edObj(i+11).Select "Yes"

Wait 3

End Function

Function Verifymyprofilepagelayout()
On error resume next
	  	'''''''''''''''''''''''This is Summary Information Section''''''''''''''''''
				wait 2
				verifyIfWebElementDoesExist1 "Name"
				verifyIfWebElementDoesExist1 "Group"
				verifyIfWebElementDoesExist1 "Federal Employee"
				verifyIfWebElementDoesExist1 "Highest Level Of Education"
				verifyIfWebElementDoesExist1 "Degrees"
				verifyIfWebElementDoesExist1 "Degrees Other"
				verifyIfWebElementDoesExist1 "Current Employer"
				verifyIfWebElementDoesExist1 "Position/Title"
				verifyIfWebElementDoesExist1 "Department"
				verifyIfWebElementDoesExist1 "Email"
				verifyIfWebElementDoesExist1 "Phone"
				verifyIfWebElementDoesExist1 "Mailing Address"
				verifyIfWebElementDoesExist1 "Race"
				verifyIfWebElementDoesExist1 "Hispanic/Latino?"
				verifyIfWebElementDoesExist1 "Gender"
				verifyIfWebElementDoesExist1 "Year of Birth"
				'''''''''''''''''''This is reviewer Summary Information'''''''''
				wait 2
				verifyIfWebElementDoesExist1 "Primary Reviewer Role"
				verifyIfWebElementDoesExist1 "Secondary Reviewer Role"
				verifyIfWebElementDoesExist1 "Special Positions"
				verifyIfWebElementDoesExist1 "Personal Statement"
				verifyIfWebElementDoesExist1 "Stakeholder Communities"
				verifyIfWebElementDoesExist1 "Patient Communities"
				'''''''''''''''''''''''This is Board Expertise Section''''''''''
				wait 2
				verifyIfWebElementDoesExist1 "Broad Disease/Condition Expertise"
				verifyIfWebElementDoesExist1 "Broad Population Expertise"
				verifyIfWebElementDoesExist1 "Broad Healthcare Expertise"
				verifyIfWebElementDoesExist1 "Broad Methodological Expertise"
				verifyIfWebElementDoesExist1 "Specific Disease/Condition Expertise"
				verifyIfWebElementDoesExist1 "Specific Population Expertise"
				verifyIfWebElementDoesExist1 "Specific Healthcare Expertise"
				verifyIfWebElementDoesExist1 "Specific Methodological Expertise"
				
				'''''''''''''''''''''This is Advisory Panel Documentation section				wait 2
				
				verifyIfWebElementDoesExist1 "Accepts Stipend(s)"
				verifyIfWebElementDoesExist1 "Accepts Reimbursements"
				verifyIfWebElementDoesExist1 "Awardee Institution/Organization"
				wait 2
				verifyIfButtonDoesExist "Add Attachment"				
				
							
				
				
				
				 End Function
				 
Function selectRadioButtonCoIexternal(strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True           

Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(0).Select strData
End Function

Function COIpagelayoutverificationexternal()
On error resume next
clk_link_Object2 "Home"
wait 2
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
wait 10
clk_link_Object2 "COI & Expertise"
wait 5
clk_link_Object2 "Addressing Disparities Cycle 3"
wait 2
verifyIfWebElementDoesExist1 "Yes"
verifyIfWebElementDoesExist1 "No"
verifyIfWebElementDoesExist1 "Unclear"
wait 2
selectRadioButton "No"
'click submit button
clk_Button_usingName "Submit"
wait 5
verifyIfWebElementDoesExist1 "Project Title"
verifyIfWebElementDoesExist1 "Program Organization"
verifyIfWebElementDoesExist1 "Principal Investigator"
wait 3
verifyIfWebElementDoesExist1 "Do you have any of the following types of Conflicts of Interest with this application?"
verifyIfWebElementDoesExist1 "Please select all that apply. If you do not have a COI, select “No.” If you are unsure whether your situation qualifies as a COI, select “Unclear” and type a description of the potential COI in the box. Selecting “Unclear” will generate an automatic notification for your Merit Review Officer, who will review your information and provide guidance."
verifyIfWebElementDoesExist1 "Project Title"
verifyIfWebElementDoesExist1 "No-The reviewer has no COI with this application."
verifyIfWebElementDoesExist1 "Yes-The reviewer has a personal COI: The reviewer or his/her close relative has a significant personal relationship with the principal investigator or a key personnel."
verifyIfWebElementDoesExist1 "Yes-The reviewer has a professional COI: The reviewer or his/her close relative has a significant professional relationship (employment related or not) with the principle investigator or key personnel."
verifyIfWebElementDoesExist1 "Yes-The reviewer has an institutional COI: The reviewer or his/her close relative is employed by or seeking employment at the applicant entity, or the reviewer may receive professional gain or advancement as a direct result of the application funding decision."
verifyIfWebElementDoesExist1 "Yes-The reviewer has a financial COI: The reviewer or his/her close relative could receive a financial benefit as a result of the application funding decision."
verifyIfWebElementDoesExist1 "Unclear if COI Exists."
wait 2
verifyIfWebElementDoesExist1 "Expertise Rating"
verifyIfWebElementDoesExist1 "Below are descriptions for each level of expertise. Scientist reviewers are expected to rate their level of expertise. Patients and stakeholders are encouraged to rate their level of expertise or may choose not to designate a level of expertise by selecting “None or Not Applicable.”"
verifyIfWebElementDoesExist1 "High"
verifyIfWebElementDoesExist1 "The Reviewer is able to evaluate the application with little or no need to make use of background material or the relevant literature. The Reviewer has likely published in areas closely related to the science presented in the application."
verifyIfWebElementDoesExist1 "Medium"
verifyIfWebElementDoesExist1 "The Reviewer has most of the knowledge to evaluate the application but will require some review of relevant literature to fill in details or increase familiarity with the system employed. The Reviewer may employ similar methodologies in his or her own work but may need to review the literature for recent data relevant to the application."
verifyIfWebElementDoesExist1 "Low"
verifyIfWebElementDoesExist1 "The Reviewer understands the broad concepts but is unfamiliar with the specific methodology or other details, and reviewing the application would require considerable preparation."
verifyIfWebElementDoesExist1 "None or Not Applicable"
verifyIfWebElementDoesExist1 "The Reviewer has only superficial or no familiarity with the concepts and methodology described in the application, or the Reviewer chooses not to answer the question about his/her expertise."
Wait 2
verifyIfWebElementDoesExist1 "Technical Abstract"
verifyIfWebElementDoesExist1 "Public Abstract"
verifyIfWebElementDoesExist1 "Expertise Level" 
verifyIfWebElementDoesExist1 "None or Not applicable"
verifyIfWebElementDoesExist1 "Low"
verifyIfWebElementDoesExist1 "Medium"
verifyIfWebElementDoesExist1 "High"
wait 2
verifyIfButtonDoesExist "Submit"
wait 2
clk_Button_usingName "Submit"
wait 2
verifyIfWebElementDoesExist1 "Explanation Required"
wait 2
selectRadioButtonCoIexternal "Unclear if COI Exists." 'have to fix later clikcing on radio button not working"
'''''''''''click submit button
wait 2
clk_Button_usingName "Submit"
wait 5
verifyIfWebElementDoesExist1 "Explanation Required"
wait 2
selectRadioButtonCoIexternal "No-The reviewer has no COI with this application."    'have to fix later clikcing on radio button not working"
wait 2
selectRadioButton1 "Low"                    'this is good one working''''
clk_Button_usingName "Submit"
End Function
''''''''''''''''''''''''''''Good One to click All Tab internal salesforce Site'''''''''''''''''''''''''''
Function clk_link_Object4()
On error resume next

wait 5
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"




Set l = Description.Create
l("micclass").value = "Link"
l("innerhtml").value = "&nbsp;<img title=""All Tabs"" class=""allTabsArrow"" alt=""All Tabs"" src=""/img/s\.gif"">&nbsp;"
l("html tag").value = "A"


Set lo =  getParentObject().ChildObjects(l)
print lo.count
lo(i).click

'    If err <> 0 Then
'  	LogReport 4,"Error", err.number & "-" & err.description
'  	LogReport 1,"cliking the link" & strName,"Failed to click the link" & "-" & strName
'  	else
'  	LogReport 0,"cliking the link" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
'  End If


End Function

Function selectWebEditcustomreport(strData)
			On error resume next

            wait 4
			   'strArr = split(strData,",")

			    Set btncalc = Description.Create()  
				    btncalc("micclass").value = "Browser"
				  
				    Set btn =DeskTop.ChildObjects(btncalc)
				    strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
				    print strHwnd
				  
				    Set pa = Description.Create
				    pa("micclass").value = "Page"
				  
				    Set editObject =  getParentObject()
				    
				    Set editO = Description.Create
				    
				    editO("micclass").value = "WebEdit"
				    editO("visible").value = true  
				      editO("name").value = "WebEdit" 
				       
				     Set edObj =  getParentObject().ChildObjects(editO) 
				    
				     for i = 0 To edObj.count-1
				        
				        If (edObj(i).GetROProperty("value") = strData)Then
				        	
				          print edObj(i).GetROProperty("value") 
				         else
				            On error resume next
                           edObj(i).set strData				          
											            
				        End If
				         
				     Next
	
End Function

Function selectWeblist_New_Window(strData)
                  On error resume next
                  
                    wait 3
                                strArr = split(strData,",")
                               
                               Set btncalc = Description.Create()  
                                btncalc("micclass").value = "Browser"
                              
                                Set btn =DeskTop.ChildObjects(btncalc)
                                 
                                Set editO = Description.Create
                                editO("micclass").value = "WebList"
                               editO("visible").value = true  
                                 Set edObj =  btn(1).Page("micclass:=Page").ChildObjects(editO) 
                                
                                 for i = 0 To uBound(strArr)
			        items =   edObj(i).GetRoProperty("all items")
			        print items
			        print strArr(i)
			        If Not(strArr(i) = "")Then	
                                          
			            edObj(i).select trim(strArr(i))        
			        End If
			         
			     Next
                If err <> 0 Then
                   err.clear   
                  
     End If
     End Function
     
     Function clk_link_Object2_newwindow(strName)
       On error resume next
       
       wait 5
     	 Set btncalc = Description.Create()  
		  btncalc("micclass").value = "Browser"
		  
		  Set btn =DeskTop.ChildObjects(btncalc)
		  strHwnd =  btn(1).GetRoProperty("hwnd")
		  print strHwnd
		  
		  Set pa = Description.Create
		  pa("micclass").value = "Page"  		 
		  
		  
		  Set l = Description.Create
		  l("micclass").value = "Link"
		  'l("html tag").value = "A"		  
		  l("innertext").value = strName	  
		  
		 Set lo =  getParentObject().ChildObjects(l)
		print lo.count
		wait 2
		  lo(0).click
  
    If err <> 0 Then
  	LogReport 4,"Error", err.number & "-" & err.description
  	LogReport 1,"cliking the link" & strName,"Failed to click the link" & "-" & strName
  	else
  	LogReport 0,"cliking the link" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
  End If

 
  
     End Function
     
Public Function clk_Image_Link_Portal()
Err.Clear
On error resume next

wait 5
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
'print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "Image"
l("alt").value = "Add"
l("image type").value = "Image Link"
l("html id").value = "00N70000003DVyw_right_arrow"



Set lo =  getParentObject().ChildObjects(l)
'print lo.count
lo(0).highlight
lo(0).click

If err <> 0 Then
    LogReport 4,"Error", err.number & "-" & err.description
    LogReport 1,"cliking the link" & "'+' Tab","Failed to click the link" & "-" & "'+' Tab"
else
      LogReport 0,"cliking the link" & "'+' Tab","The Link" & "-" & "'+' Tab" & "-" & "is clicked successfully"
End If
End Function 

Function clickforgotyourpasswordlink()
     	On error resume next
     	Set l = Description.Create
        l("micclass").value = "Link"
        l("html tag").value = "A"
        l("outertext").value = "Forgot your password\?"
		Set lo =  getParentObject().ChildObjects(l)
		print lo.count
		For i=0 to lo.count -1
			x = lo(i).getRoproperty("innertext")
			If Trim(x = "Forgot your password?") Then
				lo(i).click
				Exit for
			End If
		Next
     End Function
     
 Function selectWeblist5(HtmlID,strData)
On error resume next

wait 4
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

print pHwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
editO("html tag").value = "SELECT" 
editO("html id").value = HtmlID
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
for i = 0 To uBound(strArr)
items =   edObj(i).GetRoProperty("all items")
print items
print strArr(i)
If Not(strArr(i) = "")Then	


edObj(i).select trim(strArr(i))        
End If

Next


End Function

Public Function clk_rightarrow_Link_Portal(HtmlID)
Err.Clear
On error resume next

wait 5
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
'print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "Image"
l("alt").value = "Add"
l("image type").value = "Image Link"
l("html id").value = HtmlID



Set lo =  getParentObject().ChildObjects(l)
'print lo.count
lo(0).highlight
lo(0).click

If err <> 0 Then
    LogReport 4,"Error", err.number & "-" & err.description
    LogReport 1,"cliking the link" & "'+' Tab","Failed to click the link" & "-" & "'+' Tab"
else
      LogReport 0,"cliking the link" & "'+' Tab","The Link" & "-" & "'+' Tab" & "-" & "is clicked successfully"
End If
End Function 

Function clk_multipleCheckbox()
On error resume next
wait 3
Set ck = Description.Create

ck("micclass").value = "WebCheckBox"
ck("visible").value = true  


Set co =   getParentObject().ChildObjects(ck)
print co.count
For i=0 to co.count - 1
'n=  RandomNumber(1)	

If co.count > 0 Then
LogReport 0,"verifySendSamServeyCheckBox ", "The Check Box exists and is clikable"
else
LogReport 1,"verifySendSamServeyCheckBox ", "The Check Box doesnot exist and is not clickable"
End If
co(i).click

next
End Function
Function writeToAfileMRAApplication(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\mrapplicationnumber.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function readFromFileMRApplication()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\mrapplicationnumber.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFileMRApplication = sContent
End Function


Public Function clk_manageextrenaluser()

On error resume next
wait 4
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


print pHwnd

Set editO = Description.Create

editO("micclass").value = "WebElement"
editO("visible").value = true
editO("html tag").value = "SPAN"
editO("html id").value = "workWithPortalLabel"
'editO("name").value = strName 
Set editObject =  getParentObject().ChildObjects(editO) 
print editObject.count
editObject(0).highlight
editObject(0).click

If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
else
LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
End If
End Function
     
Public Function set_Web_List2(HtmlID,strLabelName,strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebList"'"WebElement"
odesc("visible").value= True
odesc("html tag").value= "SELECT"'"LABEL"
odesc("html id").value= HtmlID
odesc("innertext").value= strLabelName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x + 1
y= O(0).getRoProperty("abs_y")
print y + 1        

Set od =Description.Create
od("micclass").value="WebList"
'od("abs_x").value= x + 1
od("abs_y").value= y
'           

Set Os =  getParentObject().ChildObjects(od)
Print Os.count

Os(0).select strData
End Function
 
 'Function to select 8 Radio Buttons when create a New User
Function selectallRadioButtonfornewuser()
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True


Set O =  getParentObject().ChildObjects(odesc)
print O.count


O(0).select "Merit Reviewer"
Wait 2
O(1).select "Mr."
Wait 2
O(2).select "Male"
Wait 2
O(3).select "No"
Wait 2
O(4).select "White"
Wait 2
O(5).select "Patient/Consumer"
Wait 2
O(6).select "No"
Wait 2
O(7).select "Opt out"
Wait 2

End Function

Function fillMailingAddressfornewuser()  '(strPhone, strStreet,strCity,strZip)
         
         On error resume next
               Randomize
         intnumber =  Int((9999-1)*Rnd+1)
         
          strPhone = "525-236-6958"
          strStreet = "123 test street"
          strCity = "Washington DC"
        '  strState = "Virginia"
          strZip = "20036"
          
                Set editO = Description.Create
                editO("micclass").value = "WebEdit"
                editO("visible").value = true               
               
                
                Set O =  getParentObject().ChildObjects(editO) 
                print O.count
                
                For i = 0 to O.count - 1
                
                        strName = O(i).getRoProperty("name")
                        
                        If trim(strName) = "j_id0:autoCompleteForm:phonenumber" Then
                         exit for
                        
                        End If
                Next
                wait 2
                 'Phone
                 O(i).set strPhone
                 wait 2
                 'Street
                 O(i + 1).set strStreet
                 wait 2
                 'City
                 O(i + 2).set strCity
                 wait 2
                 'State
                  'O(i + 3).set strState
                ' Wait 2
                 'Zip
                 O(i + 3).set strZip
                 wait 2
                 
                 
                
                                                      
                
                              
     End Function

 Function clk_singelCheckboxNewUser()
On error resume next
wait 3
Set ck = Description.Create

ck("micclass").value = "WebCheckBox"
ck("visible").value = true  


Set co =   getParentObject().ChildObjects(ck)
print co.count
If co.count > 0 Then
LogReport 0,"verifySendSamServeyCheckBox ", "The Check Box exists and is clickable"
else
LogReport 1,"verifySendSamServeyCheckBox ", "The Check Box doesnot exist and is not clickable"
End If
co(0).click
End Function

Function clk_multipleCheckboxformrapplicationpage3()
On error resume next
wait 3
Set ck = Description.Create

ck("micclass").value = "WebCheckbox"
ck("visible").value = true  

Set co =   getParentObject().ChildObjects(ck)
print co.count

'For i=0 to co.count - 1

'For j = 0 to ro.count - 1

If ro.count > 0 Then
LogReport 0,"verifySendSamServeyCheckBox ", "The Check Box exists and is clikable"
else
LogReport 1,"verifySendSamServeyCheckBox ", "The Check Box doesnot exist and is not clickable"
End If
co(0).click
co(33).click
co(34).click
co(35).click
co(36).click
co(50).click
co(92).click


End Function

Public Function set_Web_Editbyname(strName,strData)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebEdit"
odesc("visible").value= True
odesc("html tag").value= "TEXTAREA" '"LABEL"
odesc("name").value= strName   '"con4"
'odesc("innertext").value= strLabelName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y        

Set od =Description.Create
od("micclass").value="WebEdit"
od("abs_x").value= x 
od("abs_y").value= y 
'           

Set Os =  getParentObject().ChildObjects(od)

print Os.count

Os(0).set strData
End Function

Public Function clk_webfileButton_usingName(HtmlID,strData)

On error resume next
wait 4
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


print pHwnd

Set editO = Description.Create

editO("micclass").value = "WebFile"
editO("visible").value = true 
editO("html id").Value = HtmlID
'editO("name").value = strName 
Set editObject =  getParentObject().ChildObjects(editO) 
print editObject.count
editObject(0).highlight
editObject(0).click

If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
else
LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
End If
End Function

  Function attachFileFromFileSystem(i,strFilePath)
                              Set oShell = CreateObject("WScript.Shell")
                  Set btncalc = Description.Create()  
                                                                    btncalc("micclass").value = "Browser"
                                                                  
                                                                    Set btn =DeskTop.ChildObjects(btncalc)
                                                                    strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
                                                                    print strHwnd
                                                                  
                                                                    Set pa = Description.Create
                                                                    pa("micclass").value = "Page"
                                                                  
                                                                   
                                                                    
                                                                    print pHwnd
                                                                    Set editObject =  getParentObject()
                                                                    
                                                                    Set editO = Description.Create
                                                                    
                                                                    editO("micclass").value = "WebFile"
                                                                    editO("visible").value = true  
                                                                    editO("html tag").value = "INPUT" 
                                                                       'editO("type").value = "text" 
                                                                     Set edObj =  getParentObject().ChildObjects(editO) 
                                                                     
                                                                      edObj(i).set strFilePath
                     'strSearchText = strFilePath
                     'oShell.SendKeys strSearchText 
                                                                     'set 
                                                                     wait 2
                                                                     'Browser("hwnd:=" & strHwnd).Dialog("text:=Choose File to Upload").WinButton("text:=&Open").click

End Function


  

Public Function clk_popupMessageBox()
     
        On error resume next
       
        Browser("micclass:=Browser").HandleDialog micOK
     End Function

'''''''''Function to Capture MRA Number from Merit Reviewer Application
Public Function captureWebElementText_MRANumber()
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("Visible").value = True 
l("html tag").value = "TD"
l("html id").value = "j_id0:j_id1:j_id42:j_id44:j_id45:0:j_id46"
Set lo =  getParentObject().ChildObjects(l)
Print lo.count

MRAnumber = lo(0).getRoProperty("innertext")
Print MRAnumber
'Write captured MRA-Number to Function itself in order to call it later in a script
captureWebElementText_MRANumber = MRAnumber
End Function

Function setWebEdit_DashBoard_Reportsearch(HtmlID,strdata)
   	    On Error resume next
   	    
        Set editO = Description.Create    
        editO("micclass").value = "WebEdit"
        editO("visible").value = true 
        editO("html id").value = HtmlID
        editO("html tag").value = "INPUT"
        editO("type").value = "text"
        
    
        
        Set O = getParentObject().ChildObjects(editO)
        O(1).set strdata
        If err <> 0 Then
  	      err.clear   
  	  
        End If
       
   End Function
Function clk_link_ObjectDashboard()
On error resume next

wait 5
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"




Set l = Description.Create
l("micclass").value = "Image"
'l("innerhtml").value = "&nbsp;<img title=""All Tabs"" class=""allTabsArrow"" alt=""All Tabs"" src=""/img/s\.gif"">&nbsp;"
l("html tag").value = "IMG"
l("html id").value = "x-auto-2"


Set lo =  getParentObject().ChildObjects(l)
print lo.count
lo(i).click
End function

Public Function set_Web_Element(strData)
On error resume next
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("html tag").value = "TD"
l("innertext").value = strName
l("visible").value = true 

strcount = l.count
print strcount

'Set Os =  getParentObject().ChildObjects(od)

'print Os.count

strcount(4).set strData
End Function

Function ReviewersPortalCOIAndExpertiseworkflowMRtestcase()
	  	On error resume next
	  	'open the reviewers portal
				Open_ReviewersPortal()
				
				wait 5
				
				'Login to the portal
				login_intoSalesForce_Application userReviewer,passWord1
				clk_link_Object2 "Log in"
				wait 5
				'verifyIfLinkDoesExist1 "Home"
				'verifyIfLinkDoesExist1 "My Profile"
				
				clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
				'clk_link_Object2 "INDICATE COI AND EXPERTISE"
				wait 10
				'click the COI experitise tab
				clk_link_Object2 "COI & Expertise"
'				wait 2
'				verifyIfWebElementDoesExist1 "Please click on the PCORI Funding Announcement(s) below to indicate whether you have PFA-level COI(s)."
'				verifyIfWebElementDoesExist1 "PCORI Funding Announcement"
				wait 5
				
				'click the program
				clk_link_Object2 "Addressing Disparities Cycle 3"
				
				wait 5
				'slect no
				selectRadioButton "No"
'				wait 2
'				HtmlID = ""
'                strData = "Note for Radio button seletion as Yes"
'                set_Web_Edit HtmlID,strData
				wait 2
				
				'click submit button
				clk_Button_usingName "Submit"
				'wait 5
				'verifyIfWebElementDoesExist1 "Thank you for your submission. Your MRO will reach out to you with more information.If you are assigned to another panel, please click the COI & Expertise tab above to return to the list of panels, and then select your other assigned panels to indicate COI & Expertise."  
				
				
'				Note : there is no any go button
'				click the edit link
				clk_link_Object2 "Edit"
				wait 5
'				'Verify the follwoing links exist
'				verifyIfLinkDoesExist1 "Information for PCORI Merit Reviewers on Confidentiality, Conflict of Interest, and Rating Expertise"
'				verifyIfLinkDoesExist1   "PCORI Conflict of Interest Policy"
'				verifyIfLinkDoesExist1   "PCORI Non-Disclosure Agreement (Reviewer)"
				
				'verify if application content is present
				
'				 verifyIfWebElementDoesExist1 "Application Key Personnel"
'				 
'				 'select the expertise level to high
'				 
				 selectRadioButton "Personal"   
				 selectRadioButton2 "High"
'				'click submit button
				 clk_Button_usingName "Submit"					 
				 		 
				 
				 
				 
				 
	  End Function
	  
	  Public Function set_Web_Editdateexternal(HtmlID)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebEdit"
odesc("visible").value= True
'odesc("html tag").value= "TEXTAREA" '"LABEL"
odesc("html id").value= HtmlID   '"con4"
'odesc("innertext").value= strLabelName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y        

Set od =Description.Create
od("micclass").value="WebEdit"
od("abs_x").value= x 
od("abs_y").value= y 
'           

Set Os =  getParentObject().ChildObjects(od)

print Os.count

Os(0).SetSecure date + 30
End Function

Public Function readFromFilePanel1()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\panelmr.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFilePanel1 = sContent
End Function

Function ReviewersPortalCOIAndExpertiseworkflowMRtestcase1()
	  	On error resume next
	  	'open the reviewers portal
				Open_ReviewersPortal()
				
				wait 5
				
				'Login to the portal
				login_intoSalesForce_Application userReviewer,passWord1
				clk_link_Object2 "Log in"
				wait 5
				'verifyIfLinkDoesExist1 "Home"
				'verifyIfLinkDoesExist1 "My Profile"
				
				clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
				'clk_link_Object2 "INDICATE COI AND EXPERTISE"
				wait 10
				'click the COI experitise tab
				clk_link_Object2 "COI & Expertise"
'				wait 2
'				verifyIfWebElementDoesExist1 "Please click on the PCORI Funding Announcement(s) below to indicate whether you have PFA-level COI(s)."
'				verifyIfWebElementDoesExist1 "PCORI Funding Announcement"
				wait 5
				
				'click the program
				clk_link_Object2 "Assessment of Prevention Diagnosis and Treatment Options Cycle 3"
				
				wait 5
				'slect no
				selectRadioButton "NO"
'				wait 2
'				HtmlID = ""
'                strData = "Note for Radio button seletion as Yes"
'                set_Web_Edit HtmlID,strData
				wait 2
				
				'click submit button
				clk_Button_usingName "Submit"
				wait 5
				verifyIfWebElementDoesExist1 "Thank you for your submission. Your MRO will reach out to you with more information.If you are assigned to another panel, please click the COI & Expertise tab above to return to the list of panels, and then select your other assigned panels to indicate COI & Expertise."  
				
				
'				Note : there is no any go button
'				click the edit link
				clk_link_Object2 "Edit"
'				wait 5
'				'Verify the follwoing links exist
				verifyIfLinkDoesExist1 "Information for PCORI Merit Reviewers on Confidentiality, Conflict of Interest, and Rating Expertise"
				verifyIfLinkDoesExist1   "PCORI Conflict of Interest Policy"
				verifyIfLinkDoesExist1   "PCORI Non-Disclosure Agreement (Reviewer)"
				
				'verify if application content is present
				
'				 verifyIfWebElementDoesExist1 "Application Key Personnel"
'				 
'				 'select the expertise level to high
'				 'slect no
				 selectRadioButton "No"   
				 selectRadioButton2 "High"
'				'click submit button
				 clk_Button_usingName "Submit"
'				
'				 wait 5
'				 clk_link_Object2 "Edit"
'				 'submit another question on the list
'				 selectRadioButton "Personal"
'				 selectRadioButton2 "None or Not applicable"
'				 clk_Button_usingName "Submit"
'				 
				 
				 			 
				 
				 
				 
				 
	  End Function
	  Function writeToAfilePanel1(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\panelmr.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Function clk_singelCheckboxbasedonindexid(i)
On error resume next
wait 3
Set ck = Description.Create

ck("micclass").value = "WebCheckBox"
ck("visible").value = true  


Set co =   getParentObject().ChildObjects(ck)
print co.count
If co.count > 0 Then
LogReport 0,"verifySendSamServeyCheckBox ", "The Check Box exists and is clikable"
else
LogReport 1,"verifySendSamServeyCheckBox ", "The Check Box doesnot exist and is not clickable"
End If
co(i).click
End Function

Function createApanel2forMRCOItestcase()
       on error resume next
        navigateAndLoginToSalesForce UserAdmin, SystemAdminPassword

		'click RP community manager
		clk_link_Object2 "Merit Reviewer Management"
		
		
		clk_link_Object2 "Panels"
		
		clk_Button_usingName "New"
		
		stPanel1 = create_Panel()
		writeToAfilePanel1 stPanel1
		print stPanel
		fillOnlineReviewCriteria()
		clk_Button_usingName saveBtn
		createAPanelTestCaseCAP4_1 = stPanel1
		End Function
		


''''''''''''''''''good function to click on OK in popup message box to submit confirmation''''''''''''
Function click_WinButton (dialog, text)
On error resume next

            Set btncalc = Description.Create()  
            btncalc("micclass").value = "Browser"
          
            Set btn =DeskTop.ChildObjects(btncalc)
            strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
            print strHwnd
            
        Browser("hwnd:="&strHwnd).Dialog("text:="&dialog).WinButton("text:="&text).click    
        Wait 3                                                
                                                                  


End Function

Public Function captureWebElementText_projectName()
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("html tag").value = "DIV"
l("html id").value = "opp3_ileinner"

Set lo =  getParentObject().ChildObjects(l)
Print lo.count

projectname = lo(0).getRoProperty("innertext")
Print projectname

'                                Dim arrCorrectName
'                                arrCorrectName = split(nameNewUser, ". ")
                                'Print arrCorrectName(1)
'USE this name later to search for RAPI internally - when verify as Admin - covers TC1-TC2
'nameRAPI = arrCorrectName(1)
'Print nameRAPI
'Write captured Email Address to Function itself in order to call it later in a script
captureWebElementText_projectName = projectname
End Function

Function selectWeblist6(HtmlID,strData)
On error resume next

wait 4
strArr = split(strData,";")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

print pHwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
editO("html tag").value = "SELECT" 
editO("html id").value = HtmlID
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
for i = 0 To uBound(strArr)
items =   edObj(i).GetRoProperty("all items")
print items
print strArr(i)
If Not(strArr(i) = "")Then	


edObj(i).select trim(strArr(i))        
End If

Next


End Function

Public Function captureWebElementText_MR_email()
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("Visible").value = True 
l("html tag").value = "DIV"
l("html id").value = "con15_ileinner"
Set lo =  getParentObject().ChildObjects(l)
Print lo.count

MR_Email = lo(0).getRoProperty("innertext")
Print MR_Email
'Write captured MRA-Number to Function itself in order to call it later in a script
captureWebElementText_MR_email = MR_Email
End Function

Function writeToAfileMeritReviewer1(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\reviewer1.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function readFromFileMeritReviewer1()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\reviewer1.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFileMeritReviewer1 = sContent
End Function
     
 Function writeToAfileMeritReviewer2(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\reviewer2.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function readFromFileMeritReviewer2()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\reviewer2.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFileMeritReviewer2 = sContent
End Function

Function writeToAfileMeritReviewer3(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\reviewer3.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function readFromFileMeritReviewer3()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\reviewer3.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFileMeritReviewer3 = sContent
End Function

Function writeToAfileMeritReviewer4(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\reviewer4.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function readFromFileMeritReviewer4()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\reviewer4.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFileMeritReviewer4 = sContent
End Function
Function writeToAfileMeritReviewer5(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\reviewer5.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function readFromFileMeritReviewer5()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\reviewer5.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFileMeritReviewer5 = sContent
End Function
     
''''''''''''''''''''For demo'''''''''''''''
Function Creating_ApproveScienceMR1()
	
NavigateToReviewerPortalnewuser()
wait 5
fillUserNameInformation()
wait 3
clk_singelCheckbox()
wait 2
''''Last step Join PCORI Portal
wait 2
clk_link_Object2 "Join PCORI Online"
wait 10
selectallRadioButtonfornewuser()
wait 2
fillMailingAddressfornewuser() 
Wait 5
clk_singelCheckboxNewUser()
wait 2
clk_Button_usingName "Submit"

''''''''Applying to become a scientist reviewer'''''

Wait 10
clk_Button_usingName "START MY APPLICATION"

''''''Beggining of TC021''''''

wait 5
HtmlID = "j_id0:j_id2:j_id4:j_id6:j_id50:j_id51"
strData = "Scientist"
selectWeblist6 HtmlID,strData
wait 5
clk_Button_usingName "Next"
wait 5
HtmlID = "j_id0:other:pg:j_id33:j_id35"
strData = "Professional school degree or doctorate degree (PhD, DPhil, ScD, EdD, MD, DO, PsyD, DNP, etc.)"
selectWeblist6 HtmlID,strData
wait 5
setWebEditBox "PhD"
wait 5
HtmlID = "j_id0:other:pg:j_id33:j_id49_unselected"
strData = "Academic Hospital/Clinic/Healthcare Sys"
selectWeblist6 HtmlID,strData
wait 5
clk_rightarrow_Link_Portal "j_id0:other:pg:j_id33:j_id49_right_arrow"

'''''''''''''''Peer reviewed Co-authored''''''''''
wait 5
HtmlID = "j_id0:other:pg:j_id33:j_id56"
strData = "1-4"
selectWeblist6 HtmlID,strData
'''''''''''''' peer reviewed author'''''''''''''''''''''
wait 5
HtmlID = "j_id0:other:pg:j_id33:j_id59:j_id60:j_id61"
strData = "1-2"
selectWeblist6 HtmlID,strData
''''''''''''''''Contract/grant''''''''''''''''''
wait 5
HtmlID = "j_id0:other:pg:j_id33:j_id63"
strData = "No"
selectWeblist6 HtmlID,strData
'''''''''''''''''Previously participated in any peer-reviewed process''''''''''''''''''
wait 5
HtmlID = "j_id0:other:pg:j_id33:j_id65"
strData = "No"
selectWeblist6 HtmlID,strData
''''''''''''''''Served as a Chair or Co-chair for a scientific review panel''''''''''
wait 5
HtmlID = "j_id0:other:pg:j_id33:j_id67"
strData = "No"
selectWeblist6 HtmlID,strData

wait 5
clk_Button_usingName "Next"
''''''''''''''''User landed on page 3'''''''''''''''
wait 5
HtmlID = "j_id0:mraform:j_id46:Dis"
strData = "Test Scientific Reviewer qs 1"
set_Web_Edit HtmlID,strData
wait 5
clk_multipleCheckboxformrapplicationpage3
wait 5
HtmlID = "j_id0:mraform:j_id121:popran"
strData = "Test Scientific Reviewer qs 2"
set_Web_Edit HtmlID,strData
wait 5
HtmlID = "j_id0:mraform:j_id138:hecare"
strData = "Test Scientific Reviewer qs 3"
set_Web_Edit HtmlID,strData
wait 5
HtmlID = "j_id0:mraform:j_id155:meth"
strData = "Test Scientific Reviewer qs 4"
set_Web_Edit HtmlID,strData
wait 5
strName = "j_id0:mraform:j_id170:j_id174"
strData = "Test Scientific Reviewer qs 5"
set_Web_Editbyname strName,strData
wait 5
selectRadioButton "No preference"
wait 5
selectRadioButton1 "Yes"
wait 5
selectRadioButton2 "Yes"
wait 5
HtmlID = "j_id0:mraform:j_id207:block1:j_id247:file"
strData = "Browse..."
clk_webfileButton_usingName HtmlID,strData
wait 5
attachFileFromFileSystem 0,uploadresumepath
wait 5
HtmlID = "j_id0:mraform:j_id207:block1:j_id245:fileName"
strData = "MR_test_User Resume"
set_Web_Edit HtmlID,strData

clk_Button_usingName "Upload"
wait 2
clk_Button_usingName "Submit Application"
wait 5
verifyIfWebElementDoesExist1 "Thank you for applying to become a PCORI Merit Reviewer."
wait 5
clk_link_Object2 "Home"
wait 5
clk_link_Object2 "SUBMIT A MERIT REVIEWER APPLICATION"
wait 2
captureWebElementText_MRANumber()
wait 5

'''''''Below script capture the MRA app number and User email then save that in a text file to use in later step''''''''''''

MRAappnumber = captureWebElementText_MRANumber()
wait 5
writeToAfileMRAApplication MRAappnumber
wait 5

''''''''''''''''''''Capturing the User email ID'''''''''''''''''''''

clk_link_Object2 "My Profile"
wait 5

MR_Reviewer1 = captureWebElementText_MR_email()
wait 5
writeToAfileMeritReviewer1 MR_Reviewer1
wait 5
'''''''''''''Approving the MR''''''''''''''''''as MRO'''''''''''''
navigateAndLoginToSalesForce UserAdmin,SystemAdminPassword
wait 10
navigateToAccoutsPageThrughTopSearchBox "Carolyn Mohan "
wait 4
clk_link_Object2 "Carolyn Mohan"
wait 4
clk_link_Objectforimpersonate()
wait 2
HtmlID = "USER_DETAIL"
strName = "User Detail"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 4
clk_Button_usingName "Login"
wait 2
strData= readFromFileMRApplication()
navigateToAccoutsPageThrughTopSearchBox strData
wait 4
strData= readFromFileMRApplication()
clk_link_Object2 strData
wait 4
clk_Button_usingName "Submit for Approval"
wait 2
click_WinButton "Message from webpage", "OK"
wait 2
clk_link_Object2 "Approve / Reject"
wait 2
setWebEditBox ", Approval Test for MR"
wait 2
clk_Button_usingName "Approve"
wait 2
HtmlID = "globalHeaderNameMink"
strName = "Carolyn Mohan"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 2
clk_onlogoutwhendone_impersonating "Logout"



End Function
     
Function clk_link_Objectforimpersonate()
       On error resume next
       
       wait 5
     	 Set btncalc = Description.Create()  
		  btncalc("micclass").value = "Browser"
		  
		  Set btn =DeskTop.ChildObjects(btncalc)
		  strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
		  print strHwnd
		  
		  Set pa = Description.Create
		  pa("micclass").value = "Page"
		  
		 
		  
		  
		  Set l = Description.Create
		  l("micclass").value = "Link"
		  l("html id").value = "moderatorMutton"
		  l("html tag").value = "A"
		  l("title").value = "User Action Menu"
		  
		  
		 Set lo =  getParentObject().ChildObjects(l)
		print lo.count
		  lo(0).click
  
    If err <> 0 Then
  	LogReport 4,"Error", err.number & "-" & err.description
  	LogReport 1,"cliking the link" & strName,"Failed to click the link" & "-" & strName
  	else
  	LogReport 0,"cliking the link" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
  End If

 
  
     End Function
     
     Function clk_onlink_forimpersonateinternal(HtmlID,strName)
       On error resume next
       
       wait 5
     	 Set btncalc = Description.Create()  
		  btncalc("micclass").value = "Browser"
		  
		  Set btn =DeskTop.ChildObjects(btncalc)
		  strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
		  print strHwnd
		  
		  Set pa = Description.Create
		  pa("micclass").value = "Page"
		  
		 
		  
		  
		  Set l = Description.Create
		  l("micclass").value = "Link"
		  l("innertext").value = strName
		  l("html id").value = HtmlID
		  l("html tag").value = "A"
		  
		  
		  
		 Set lo =  getParentObject().ChildObjects(l)
		print lo.count
		  lo(0).click
  
    If err <> 0 Then
  	LogReport 4,"Error", err.number & "-" & err.description
  	LogReport 1,"cliking the link" & strName,"Failed to click the link" & "-" & strName
  	else
  	LogReport 0,"cliking the link" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
  End If

 
  
     End Function
     Function clk_onlogoutwhendone_impersonating(strName)
       On error resume next
       
       wait 5
     	 Set btncalc = Description.Create()  
		  btncalc("micclass").value = "Browser"
		  
		  Set btn =DeskTop.ChildObjects(btncalc)
		  strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
		  print strHwnd
		  
		  Set pa = Description.Create
		  pa("micclass").value = "Page"
		  
		 
		  
		  
		  Set l = Description.Create
		  l("micclass").value = "Link"
		  l("innertext").value = strName
		  'l("html id").value = HtmlID
		  l("html tag").value = "A"
		  
		  
		  
		 Set lo =  getParentObject().ChildObjects(l)
		print lo.count
		  lo(0).click
  
    If err <> 0 Then
  	LogReport 4,"Error", err.number & "-" & err.description
  	LogReport 1,"cliking the link" & strName,"Failed to click the link" & "-" & strName
  	else
  	LogReport 0,"cliking the link" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
  End If

 
  
     End Function
     
     Function create_Panelwithonlinereviewrecordtype()
			On error resume next
			 strDate =  now + 30
             strA = split(strDate,":")
             strDeadLineDate = strA(0) & ":" & strA(1) & space(1) & "PM"
             
              strDate2 =  now + 30
              strB = split(strDate2,":")
              strformatted = strB(0) & ":" & strB(1) & space(1) & "PM"             
'              wait 2
'              strDate3 =  now + 30
'              strC = split(strDate3,":")
'             strformatted1 = strC(0) & ":" & strC(1) '& space(1) & "PM"  
'              
              
			strNameEvent = "Panel 1" & "_" & RandomString(3) 
			strArrN = split(straddinfo,",")
			s = Date
			Set odesc=Description.Create
			odesc("micclass").value="WebTable"
			
			Set l_link=  getParentObject().ChildObjects(odesc)
			print l_link.count
			For i = 0 to l_link.count - 1
			strName = l_link(i).GetROProperty("column names")
			print strName
			strArr = split(strName,";")
			If Not(strName = "") Then
			print strArr(0)
			If trim(strArr(0)) = "*Panel Name" Then
			l_link(i).GetROProperty("rows") 
			
			a = l_link(i).GetROProperty("rows")  
			b = l_link(i).GetROProperty("cols") 		  	       
			For x  = 1 to a     
			For j  =1 to b
			c = l_link(i).GetCellData(x,j)
			print c
			Select Case c
			Case "*Panel Name"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set strNameEvent
			
			Case "*Cycle"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set NewCycle
			Case "Program"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "Addressing Disparities"
			'Selecting PFA''''''
			Case "PFA"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "Addressing Disparities Cycle 3"
					
			Case "*MRO"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "Carolyn Mohan"
			Case "Online Review Deadline"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set strDeadLineDate
			Case "InPerson Review Deadline"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set  strformatted 
'			Case "Panel Due Date"
'			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
'            oEdit1.set  s + 25
'''''''''''''Selecting Panel Due Date''''''''''''''''''''

            HtmlID = "00N0x000000LpNd"
            strData = date + 30
            set_Web_Edit HtmlID,strData
            
            wait 2
            clk_singelCheckbox()
            wait 2
            selectWeblist5 "00N39000003LYhh", "Online Review - AD"           
			
			                                  
			End Select                               
			
			Next
			Next                             
			
			
			
			
			Exit for
			
			End If
			End If
			Next
			
			create_Panelwithonlinereviewrecordtype = strNameEvent
End Function
Function Create_cycle_assystemadmin() 
wait 5
clk_link_Object2 "Cycles"
wait 2
clk_Button_usingName "New"
wait 2
HtmlID = "Name"
strData = NewCycle
set_Web_Edit HtmlID,strData
wait 2
HtmlID = "00N70000003DN7w"
strData = date + 30
set_Web_Edit HtmlID,strData
wait 2
clk_Button_usingName "Save"
wait 2
	
End Function
Function Create_Panel_assystemadmin() 
wait 2
	clk_link_Object2 "Panels"
wait 2
clk_Button_usingName "New"
stPanel = create_Panelwithonlinereviewrecordtype()
wait 2
writeToAfilePanel stPanel
print stPanel
wait 2
fillOnlineReviewCriteria()
wait 2
clk_Button_usingName saveBtn
End Function
Function insertPanelIntoProjects1()
     	On error resume next
 navigateAndLoginToSalesForce UserAdmin,SystemAdminPassword  	
    	
     	wait 5
     	stProject = readFromFileProject()
     	wait 2
     	navigateToAccoutsPageThrughTopSearchBox stProject
'     	
    	clk_link_Object2 stProject
     	'For i=1 to 3
'		Set odesc=Description.Create
'		odesc("micclass").value="WebTable"
'		odesc("cols").value= 9
'		
'		Set l_link=  getParentObject().ChildObjects(odesc)
'		print l_link.count
'		Set oEdit1 = l_link(i).ChildItem(1,4, "Link", 0)
'		    oEdit1.click
		   wait 5
		   clk_Button_usingName "Edit"
		    
		    wait 5
		    stPanel = readFromFilePanel()
		    strData = stPanel
		    
		    set_Web_Edit "CF00N39000003LYer",strData
		    
		     'populateKeyProjectPersonel stPanel
		     clk_Button_usingName "Save"
		     wait 5
		     'clk_link_Object2 "Projects"
		     
		    ' wait 3
			
			'select the RA app 2 ready for MR review
			'selectWeblist "RA App 3 Ready for MR"		
			
		     
        
         
     End Function
     
Function CMAgeneratingCombinePDFandfinalizingthePanel()
On error Resume Next

'''''''''''Below commented line can be used if you want Impersonate as CMA
	'wait 10
'navigateToAccoutsPageThrughTopSearchBox "CMA Admin (Test User)"
'wait 4
'clk_link_Object2 "CMA Admin \(Test User\)"
'wait 4
'clk_link_Objectforimpersonate()
'wait 2
'HtmlID = "USER_DETAIL"
'strName = "User Detail"
'clk_onlink_forimpersonateinternal HtmlID,strName
'wait 4
'clk_Button_usingName "Login"
wait 2
navigateAndLoginToSalesForce cmaOperationsuser, passWord
wait 4
stPanel = readFromFilePanel()
navigateToAccoutsPageThrughTopSearchBox stPanel
wait 4
clk_link_Object2 stPanel
wait 3
stProject = readFromFileProject()
clk_link_Object2 stProject
wait 5
clk_link_Object2 "Generate PDF -- RA Applications"
wait 25
CloseLatestOpenedBrowser()
wait 3
clk_Button_usingName "Edit"
wait 2
clk_Button_usingName "Save"
wait 3
click_webElementBasedonfieldname "Application Download - MR","Application Download - MR"
wait 2
clk_link_Object2 "Panels"
wait 2
stPanel = readFromFilePanel()
clk_link_Object2 stPanel
wait 5
clk_Button_usingName "Edit"
wait 2
clk_singelCheckboxbasedonindexid()
wait 2
clk_Button_usingName "Save"


''''This line is when you want to log out as impersonated user'''
'wait 2
'HtmlID = "globalHeaderNameMink"
'strName = "CMA Admin \(Test User\)"
'clk_onlink_forimpersonateinternal HtmlID,strName
'wait 2
'clk_onlogoutwhendone_impersonating "Logout"

End Function

Function MROCreatingPanelAssignment()
	wait 5
'navigateToAccoutsPageThrughTopSearchBox "Ethan Chiang"
'wait 4
'clk_link_Object2 "Ethan Chiang"
'wait 4
'clk_link_Objectforimpersonate()
'wait 2
'HtmlID = "USER_DETAIL"
'strName = "User Detail"
'clk_onlink_forimpersonateinternal HtmlID,strName
'wait 4
'clk_Button_usingName "Login"

navigateAndLoginToSalesForce mrManagementUser, passWord1

wait 2
stPanel = readFromFilePanel()
navigateToAccoutsPageThrughTopSearchBox stPanel
wait 3
clk_link_Object2 stPanel
wait 5
			
clk_Button_usingName "New Panel Assignment"
			
setWebEditBox "," & reviewer1
			
clk_Button_usingName saveAndNew
			
setWebEditBox "," & reviewer2
			
clk_Button_usingName saveAndNew
			
setWebEditBox "," & reviewer3
			
clk_Button_usingName saveAndNew
			
setWebEditBox "," & reviewer4
			
clk_Button_usingName saveAndNew
			
setWebEditBox "," & reviewer5		
			
clk_Button_usingName saveBtn
wait 3
stPanel = readFromFilePanel()
clk_link_Object2 stPanel
wait 4
clk_link_Object2 "Go to list \(10\) »"
wait 5

'HtmlID = "globalHeaderNameMink"
'strName = "Ethan Chiang"
'clk_onlink_forimpersonateinternal HtmlID,strName
'wait 2
'clk_onlogoutwhendone_impersonating "Logout"

End Function

Function ReviewersPortalCOIAndExpertisesubmissionmr1()
	  	On error resume next
	  	'open the reviewers portal
				Open_ReviewersPortal()
				
				wait 5
				
				''Login to the portal
				login_intoSalesForce_Application userReviewer1,passWord1
				clk_link_Object2 "Log in"
wait 10				
				
				'clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
				clk_link_Object2 "INDICATE COI AND EXPERTISE"
				wait 10
				'click the COI experitise tab
				clk_link_Object2 "COI & Expertise"
'				wait 2
				'verifyIfWebElementDoesExist1 "Please click on the PCORI Funding Announcement(s) below to indicate whether you have PFA-level COI(s)."
'				verifyIfWebElementDoesExist1 "PCORI Funding Announcement"
				'wait 5
				
				'click the program
				clk_link_Object2 DICampgn
				
				wait 5
				''''selecting no
				click_webElement_forengagementreport_selecting_byindex "No",0
                clk_Button_usingName "Submit"
				wait 5
'				Note : there is no go button available'''''''
'				click the edit link
				clk_link_Object2 "Edit"
				wait 5
				'Verify the follwoing links exist
				'verifyIfLinkDoesExist1 "Information for PCORI Merit Reviewers on Confidentiality, Conflict of Interest, and Rating Expertise"
				'verifyIfLinkDoesExist1   "PCORI Conflict of Interest Policy"
				'verifyIfLinkDoesExist1   "PCORI Non-Disclosure Agreement (Reviewer)"
'				
'				''''verify if application content is present
'				
				 verifyIfWebElementDoesExist1 "Application Key Personnel"

'				 'slecting no for PFA level COI
                wait 5
              click_webElement_forengagementreport_selecting_byindex "No-The reviewer has no COI with this application\.",0
              click_webElement_forengagementreport_selecting_byindex "High",0
              clk_Button_usingName "Submit"
              '				this below code is for when user have multiple COI'''''
				 wait 5
				 'clk_link_Object2 "Edit"
				 'submit another question on the list
				 'selectRadioButton "Personal"
				 'selectRadioButton2 "None or Not applicable"
				 'clk_Button_usingName "Submit"
	
				 			 
				 
				 
				 
				 
	  End Function
	  
Function selectRadioButtonCoIexternal1()
On error resume next

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strhwnd

Set pa = Description.Create
pa("micclass").value = "Page"
print hwnd


Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("html tag").value= "INPUT"
odesc("html id").value = "j_id0:j_id1:j_id17:j_id20:0"
'odesc("value").value = "No"


'"j_id0:j_id1:j_id17:j_id20:0"
Set O =  getParentObject().ChildObjects(odesc)
print O.count

O(1).Select 
'O(1).Click
Wait 2

End Function

Function SelectRadioButtononPopUp(StrInnerText)

Err.Clear
On Error Resume Next
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
                Set oDesc = Description.Create
oDesc("micclass").value = "WebElement"
oDesc("html tag").value = "TD"
oDesc("innertext").value = StrInnerText
oDesc("Visible").value = "True"
                
                If Brwser.WebElement(oDesc).Exist(3) Then
                                
                                Brwser.WebElement(oDesc).highlight
                                Brwser.WebElement(oDesc).Click
                                
                                LogReport 0," Click on the Radiobutton ", " done successfully "
                Else
                                LogReport 1," Click on the Radiobutton ", " Failed "
                End If
                
                If err <> 0 Then
                
                                                LogReport 4,"Error", err.number & "-" & err.description
                End  If
                
                
Set oDesc = Nothing

End Function

Function clk_Web_ElementusingnameexternalCOI(strData)

On error resume next
wait 4
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

print pHwnd

Set editO = Description.Create

editO("micclass").value = "WebElement"
editO("visible").value = true  
editO("innertext").value = strData
editO("html tag").value = "LABEL"
editO("innerhtml").value = "input name=""j_id0:j_id1:j_id17:j_id20"" id=""j_id0:j_id1:j_id17:j_id20:1"" onclick=""A4J\.AJAX\.Submit\('j_id0:j_id1',event,\{'similarityGroupingId':'j_id0:j_id1:j_id17:j_id21','parameters':\{'j_id0:j_id1:j_id17:j_id21':'j_id0:j_id1:j_id17:j_id21'\} \} \)"" type=""radio"" value=""No""><label for=""j_id0:j_id1:j_id17:j_id20:1""> No</label>"
Set editObject =  getParentObject().ChildObjects(editO) 
print editObject.count
'editObject().highlight
editObject(0).Click



If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
else
LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
End If
End Function

Function selectRadioButtonCOIEX(strData)
On error resume next
wait 4
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
'odesc("visible").value= True
odesc("html id").value= "j_id0:j_id1:j_id17:j_id20:0"
odesc("html tag").value= "INPUT"
'odesc("value").value = strData


Set O =  getParentObject().ChildObjects(odesc)
print O.count

O(0).Select strData
End Function

Function ReviewersPortalCOIAndExpertisesubmissionmr2()
	  	On error resume next
	  	'open the reviewers portal
				Open_ReviewersPortal()
				
				wait 5
				
				''Login to the portal
				login_intoSalesForce_Application userReviewer2,passWord1
				clk_link_Object2 "Log in"
wait 10				
				
				'clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
				clk_link_Object2 "INDICATE COI AND EXPERTISE"
				'wait 10
				'click the COI experitise tab
				clk_link_Object2 "COI & Expertise"
'				wait 2
				'verifyIfWebElementDoesExist1 "Please click on the PCORI Funding Announcement(s) below to indicate whether you have PFA-level COI(s)."
'				verifyIfWebElementDoesExist1 "PCORI Funding Announcement"
				'wait 5
				
				'click the program
				clk_link_Object2 DICampgn
				
				wait 5
				'slect no
				click_webElement_forengagementreport_selecting_byindex "No",0
clk_Button_usingName "Submit"
''''''''click the edit link
clk_link_Object2 "Edit"
				wait 5
'				Note : there is no go button available'''''''
'				click the edit link
				clk_link_Object2 "Edit"
				wait 5
				'Verify the follwoing links exist
				'verifyIfLinkDoesExist1 "Information for PCORI Merit Reviewers on Confidentiality, Conflict of Interest, and Rating Expertise"
				'verifyIfLinkDoesExist1   "PCORI Conflict of Interest Policy"
				'verifyIfLinkDoesExist1   "PCORI Non-Disclosure Agreement \(Reviewer\)"
'				
'				''''verify if application content is present
'				
				 'verifyIfWebElementDoesExist1 "Application Key Personnel"
                 wait 5
                 click_webElement_forengagementreport_selecting_byindex "No-The reviewer has no COI with this application\.",0
                 click_webElement_forengagementreport_selecting_byindex "High",0
                 clk_Button_usingName "Submit"
'				this below code is for when user have multiple COI'''''
				 wait 5
				 'clk_link_Object2 "Edit"
				 'submit another question on the list
				 'selectRadioButton "Personal"
				 'selectRadioButton2 "None or Not applicable"
				 'clk_Button_usingName "Submit"
				 
				 
				 			 
				 
				 
				 
				 
	  End Function


Function ReviewersPortalCOIAndExpertisesubmissionmr3()
	  	On error resume next
	  	'open the reviewers portal
				Open_ReviewersPortal()
				
				wait 5
				
				''Login to the portal
				login_intoSalesForce_Application userReviewer3,passWord1
				clk_link_Object2 "Log in"
wait 10				
				
				'clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
				clk_link_Object2 "INDICATE COI AND EXPERTISE"
				'wait 10
				'click the COI experitise tab
				clk_link_Object2 "COI & Expertise"
'				wait 2
				'verifyIfWebElementDoesExist1 "Please click on the PCORI Funding Announcement(s) below to indicate whether you have PFA-level COI(s)."
'				verifyIfWebElementDoesExist1 "PCORI Funding Announcement"
				'wait 5
				
				'click the program
				clk_link_Object2 DICampgn
				
				wait 5
				'slect no
				click_webElement_forengagementreport_selecting_byindex "No",0
clk_Button_usingName "Submit"
''''''''click the edit link
clk_link_Object2 "Edit"
				wait 5
'				Note : there is no go button available'''''''
'				click the edit link
				clk_link_Object2 "Edit"
				wait 5
				'Verify the follwoing links exist
				'verifyIfLinkDoesExist1 "Information for PCORI Merit Reviewers on Confidentiality, Conflict of Interest, and Rating Expertise"
				'verifyIfLinkDoesExist1   "PCORI Conflict of Interest Policy"
				'verifyIfLinkDoesExist1   "PCORI Non-Disclosure Agreement \(Reviewer\)"
'				
'				''''verify if application content is present
'				
				 'verifyIfWebElementDoesExist1 "Application Key Personnel"
                 wait 5
                 click_webElement_forengagementreport_selecting_byindex "No-The reviewer has no COI with this application\.",0
                 click_webElement_forengagementreport_selecting_byindex "High",0
                 clk_Button_usingName "Submit"
'				this below code is for when user have multiple COI'''''
				 wait 5
				 'clk_link_Object2 "Edit"
				 'submit another question on the list
				 'selectRadioButton "Personal"
				 'selectRadioButton2 "None or Not applicable"
				 'clk_Button_usingName "Submit"
				 
				 
				 
				 			 
				 
				 
				 
				 
	  End Function
Function ReviewersPortalCOIAndExpertisesubmissionmr4()
	  	On error resume next
	  	'open the reviewers portal
				Open_ReviewersPortal()
				
				wait 5
				
				''Login to the portal
				login_intoSalesForce_Application userReviewer4,passWord1
				clk_link_Object2 "Log in"
wait 10				
				
				'clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
				clk_link_Object2 "INDICATE COI AND EXPERTISE"
				'wait 10
				'click the COI experitise tab
				clk_link_Object2 "COI & Expertise"
'				wait 2
				'verifyIfWebElementDoesExist1 "Please click on the PCORI Funding Announcement(s) below to indicate whether you have PFA-level COI(s)."
'				verifyIfWebElementDoesExist1 "PCORI Funding Announcement"
				'wait 5
				
				'click the program
				clk_link_Object2 DICampgn
				
				wait 5
				'slect no
				click_webElement_forengagementreport_selecting_byindex "No",0
clk_Button_usingName "Submit"
''''''''click the edit link
clk_link_Object2 "Edit"
				wait 5
'				Note : there is no go button available'''''''
'				click the edit link
				clk_link_Object2 "Edit"
				wait 5
				'Verify the follwoing links exist
				'verifyIfLinkDoesExist1 "Information for PCORI Merit Reviewers on Confidentiality, Conflict of Interest, and Rating Expertise"
				'verifyIfLinkDoesExist1   "PCORI Conflict of Interest Policy"
				'verifyIfLinkDoesExist1   "PCORI Non-Disclosure Agreement \(Reviewer\)"
'				
'				''''verify if application content is present
'				
				 'verifyIfWebElementDoesExist1 "Application Key Personnel"
                 wait 5
                 click_webElement_forengagementreport_selecting_byindex "No-The reviewer has no COI with this application\.",0
                 click_webElement_forengagementreport_selecting_byindex "High",0
                 clk_Button_usingName "Submit"
'				this below code is for when user have multiple COI'''''
				 wait 5
				 'clk_link_Object2 "Edit"
				 'submit another question on the list
				 'selectRadioButton "Personal"
				 'selectRadioButton2 "None or Not applicable"
				 'clk_Button_usingName "Submit"
				 
				 
				 
				 			 
				 
				 
				 
				 
	  End Function


Function ReviewersPortalCOIAndExpertisesubmissionmr5()
	  	On error resume next
	  	'open the reviewers portal
				Open_ReviewersPortal()
				
				wait 5
				
				''Login to the portal
				login_intoSalesForce_Application userReviewer5,passWord1
				clk_link_Object2 "Log in"
wait 10				
				
				'clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
				clk_link_Object2 "INDICATE COI AND EXPERTISE"
				wait 10
				'click the COI experitise tab
				clk_link_Object2 "COI & Expertise"
'				wait 2
				'verifyIfWebElementDoesExist1 "Please click on the PCORI Funding Announcement(s) below to indicate whether you have PFA-level COI(s)."
'				verifyIfWebElementDoesExist1 "PCORI Funding Announcement"
				'wait 5
				
				'click the program
				clk_link_Object2 DICampgn
				
				wait 5
				'slect no
				click_webElement_forengagementreport_selecting_byindex "No",0
clk_Button_usingName "Submit"
''''''''click the edit link
clk_link_Object2 "Edit"
				wait 5
'				Note : there is no go button available'''''''
'				click the edit link
				clk_link_Object2 "Edit"
				wait 5
				'Verify the follwoing links exist
				'verifyIfLinkDoesExist1 "Information for PCORI Merit Reviewers on Confidentiality, Conflict of Interest, and Rating Expertise"
				'verifyIfLinkDoesExist1   "PCORI Conflict of Interest Policy"
				'verifyIfLinkDoesExist1   "PCORI Non-Disclosure Agreement \(Reviewer\)"
'				
'				''''verify if application content is present
'				
				 'verifyIfWebElementDoesExist1 "Application Key Personnel"
                 wait 5
                 click_webElement_forengagementreport_selecting_byindex "No-The reviewer has no COI with this application\.",0
                 click_webElement_forengagementreport_selecting_byindex "High",0
                 clk_Button_usingName "Submit"
'				this below code is for when user have multiple COI'''''
				 wait 5
				 'clk_link_Object2 "Edit"
				 'submit another question on the list
				 'selectRadioButton "Personal"
				 'selectRadioButton2 "None or Not applicable"
				 'clk_Button_usingName "Submit"
				 
				 
				 			 
				 
				 
				 
				 
	  End Function	

Function createRAAssighnmentMatrixViewasMRO(stPanel)
	  	On error resume next
	  	navigateAndLoginToSalesForce mrManagementUser, somersPassword
	  	wait 2
	  	clk_link_Object2 "Projects"
	  	wait 2
	  	
	  	HtmlID = "fcf"
        strData = "RA App Assignment Matrix"
		selectWeblist5 HtmlID,strData
	  		  	
	  	wait 3		
		  clk_link_Object2 "Edit"
		  wait 2
		  stPanel = readFromFilePanel()
		  HtmlId = "fname"	 
		 strData = "RA App Assignment Matrix" & space(1) & stPanel
		 set_Web_Edit HtmlID,strData
		  
		  strViewName = "RA App Assignment Matrix" & space(1) & stPanel
		  'write this name to flat file
		  writeToAfileRAMatrixView strViewName	 
		wait 3
        HtmlID = "fcol2"
        strData = "Panel"
		 selectWeblist5 HtmlID,strData
		 wait 2
		 
	    HtmlID = "fop2"
        strData = "equals"
		 selectWeblist5 HtmlID,strData	
		 wait 2
		 HtmlID = "fval2"
         stPanel = readFromFilePanel()		 
		 strData = stPanel
		 set_Web_Edit HtmlID,strData
		  'clone the view
		  'CloneView strViewName
		  'enter crtieria row for panel name
		  'fillCriterialRowPanelName stPanel
		  clk_Button_usingName "Save As"
		  
		  createRAAssighnmentMatrixViewasMRO = strViewName
	  End Function
				 
Public Function click_webElementExternalReviewlistinaproject(strName)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
l("class").value = "listTitle"
'l("outerhtml").value = "<span class=""listTitle"">External Reviews<span class=""count"">\[5\]</span></span>"
l("visible").value = True
l("html tag").value = "SPAN"
l("innertext").Value = strName
Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click

End Function

Function SelectWebListApplicationDownloads()  

On error resume next

wait 4
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strhwnd

Set pa = Description.Create
pa("micclass").value = "Page"
print hwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
editO("html tag").value = "SELECT" 
editO("select type").value = "ComboBox Select" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
stProject = readFromFileProject()
edObj(i).Select "4567 :" & space(1) &stProject

Wait 3			 
		End Function	
Function OnlineReviewsubmissionexternalMR1() 
	wait 2


'open the reviewers portal
Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer1,passWord1
clk_link_Object2 "Log in"
wait 10				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
wait 5
''''''Application Review download'''''''''''
clk_link_Object2 "Application Reviews Downloads"
wait 2
'SelectWebListApplicationDownloads()
wait 2
clk_Button_usingName "Click here to download application review materials"
'wait 2
'clk_link_Object2 "Cycle III_4567_External Portal_ FullApplication"
'wait 6
''''CloseLatestOpenedBrowser()
'Browser("PCORI Reviewer").WinObject("WinObject").WinButton("Close Tab (Ctrl+W)").Click

'''''''''''''Going back to Dashboard''''''''''''''''''
wait 4
clk_link_Object2 "Reviewer Dashboard"
'''''''''''''Opening a Online Review record and submitting the Review''''''''''''
wait 2
click_PencilImageToEdit Reviewtype
wait 2
clk_link_Object2 "Online Review"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame").WebElement("WebElement").Object.innerText = "Criterion 1: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_2").WebElement("WebElement").Object.innerText = "Criterion 1: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:4:j_id331"
strData = "1"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_3").WebElement("WebElement").Object.innerText = "Criterion 2: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_4").WebElement("WebElement").Object.innerText = "Criterion 2: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:9:j_id331"
strData = "2"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_5").WebElement("WebElement").Object.innerText = "Criterion 3: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_6").WebElement("WebElement").Object.innerText = "Criterion 3: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:14:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_7").WebElement("WebElement").Object.innerText = "Criterion 4: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_8").WebElement("WebElement").Object.innerText = "Criterion 4: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:19:j_id331"
strData = "4"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_9").WebElement("WebElement").Object.innerText = "Criterion 5: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_10").WebElement("WebElement").Object.innerText = "Criterion 5: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:24:j_id331"
strData = "5"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_12").WebElement("WebElement").Object.innerText = "Criterion 6: Strengths"
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_13").WebElement("WebElement").Object.innerText = "Criterion 6: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:29:j_id331"
strData = "6"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:34:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
populateonlinereview_asMR 0
populateonlinereview_asMR 1
populateonlinereview_asMR 2
populateonlinereview_asMR 3
populateonlinereview_asMR 4
populateonlinereview_asMR 5
populateonlinereview_asMR 6
populateonlinereview_asMR 7
populateonlinereview_asMR 8
populateonlinereview_asMR 9
populateonlinereview_asMR 10
populateonlinereview_asMR 11
populateonlinereview_asMR 12

'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_11").WebElement("WebElement").Object.innerText = "Overall Comments---Outstanding"
wait 2
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_Button_usingName "Submit"
wait 2
clk_Button_usingName "OK"

'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
clk_link_Object2 "Closed Reviews"
wait 2
verifyIfReviewExistsinClosed "Review Submitted"
wait 4
click_magnifyinglass "Review Submitted"
''''''''''''''''Verifying Review record is locked''''''''
wait 2
clk_link_Object2 "Online Review"
wait 2
End Function

Function OnlineReviewsubmissionexternalMR2()
	wait 2


'open the reviewers portal
Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer2,passWord1
clk_link_Object2 "Log in"
wait 10				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
wait 5
''''''Application Review download'''''''''''
'clk_link_Object2 "Application Reviews Downloads"
'wait 2
'SelectWebListApplicationDownloads()
'wait 2
'clk_Button_usingName "Click here to download application review materials"
'wait 2
'clk_link_Object2 "Cycle III_4567_External Portal_ FullApplication"
'wait 6
'''''CloseLatestOpenedBrowser()
'Browser("PCORI Reviewer").WinObject("WinObject").WinButton("Close Tab (Ctrl+W)").Click

'''''''''''''Going back to Dashboard''''''''''''''''''
wait 4
clk_link_Object2 "Reviewer Dashboard"
'''''''''''''Opening a Online Review record and submitting the Review''''''''''''
wait 2
click_PencilImageToEdit Reviewtype
wait 2
clk_link_Object2 "Online Review"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame").WebElement("WebElement").Object.innerText = "Criterion 1: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_2").WebElement("WebElement").Object.innerText = "Criterion 1: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:4:j_id331"
strData = "1"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_3").WebElement("WebElement").Object.innerText = "Criterion 2: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_4").WebElement("WebElement").Object.innerText = "Criterion 2: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:9:j_id331"
strData = "2"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_5").WebElement("WebElement").Object.innerText = "Criterion 3: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_6").WebElement("WebElement").Object.innerText = "Criterion 3: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:14:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_7").WebElement("WebElement").Object.innerText = "Criterion 4: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_8").WebElement("WebElement").Object.innerText = "Criterion 4: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:19:j_id331"
strData = "4"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_9").WebElement("WebElement").Object.innerText = "Criterion 5: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_10").WebElement("WebElement").Object.innerText = "Criterion 5: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:24:j_id331"
strData = "5"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_12").WebElement("WebElement").Object.innerText = "Criterion 6: Strengths"
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_13").WebElement("WebElement").Object.innerText = "Criterion 6: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:29:j_id331"
strData = "6"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:34:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
'Browser("Article Management - Console").Page("Dashboard ~ PCORI Reviewer").Frame("Frame").WebElement("WebElement").Object.innerText = "Overall Comments---Outstanding"
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_11").WebElement("WebElement").Object.innerText = "Overall Comments---Outstanding"
populateonlinereview_asMR 0
populateonlinereview_asMR 1
populateonlinereview_asMR 2
populateonlinereview_asMR 3
populateonlinereview_asMR 4
populateonlinereview_asMR 5
populateonlinereview_asMR 6
populateonlinereview_asMR 7
populateonlinereview_asMR 8
populateonlinereview_asMR 9
populateonlinereview_asMR 10
populateonlinereview_asMR 11
populateonlinereview_asMR 12
wait 2
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_Button_usingName "Submit"
wait 2
clk_Button_usingName "OK"

'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
'clk_link_Object2 "Closed Reviews"
'wait 2
'verifyIfReviewExistsinClosed "Review Submitted"
'wait 4
'click_magnifyinglass "Review Submitted"
'''''''''''''''''Verifying Review record is locked''''''''
'wait 2
'clk_link_Object2 "Online Review"
'wait 2
End Function
Function OnlineReviewsubmissionexternalMR3()
	wait 2


'open the reviewers portal
Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer3,passWord1
clk_link_Object2 "Log in"
wait 10				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
wait 5
''''''Application Review download'''''''''''
'clk_link_Object2 "Application Reviews Downloads"
'wait 2
'SelectWebListApplicationDownloads()
'wait 2
'clk_Button_usingName "Click here to download application review materials"
'wait 2
'clk_link_Object2 "Cycle III_4567_External Portal_ FullApplication"
'wait 6
'''''CloseLatestOpenedBrowser()
'Browser("PCORI Reviewer").WinObject("WinObject").WinButton("Close Tab (Ctrl+W)").Click

'''''''''''''Going back to Dashboard''''''''''''''''''
wait 4
clk_link_Object2 "Reviewer Dashboard"
'''''''''''''Opening a Online Review record and submitting the Review''''''''''''
wait 2
click_PencilImageToEdit Reviewtype
wait 2
clk_link_Object2 "Online Review"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame").WebElement("WebElement").Object.innerText = "Criterion 1: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_2").WebElement("WebElement").Object.innerText = "Criterion 1: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:4:j_id331"
strData = "1"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_3").WebElement("WebElement").Object.innerText = "Criterion 2: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_4").WebElement("WebElement").Object.innerText = "Criterion 2: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:9:j_id331"
strData = "2"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_5").WebElement("WebElement").Object.innerText = "Criterion 3: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_6").WebElement("WebElement").Object.innerText = "Criterion 3: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:14:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_7").WebElement("WebElement").Object.innerText = "Criterion 4: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_8").WebElement("WebElement").Object.innerText = "Criterion 4: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:19:j_id331"
strData = "4"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_9").WebElement("WebElement").Object.innerText = "Criterion 5: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_10").WebElement("WebElement").Object.innerText = "Criterion 5: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:24:j_id331"
strData = "5"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_12").WebElement("WebElement").Object.innerText = "Criterion 6: Strengths"
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_13").WebElement("WebElement").Object.innerText = "Criterion 6: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:29:j_id331"
strData = "6"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:34:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
'Browser("Article Management - Console").Page("Dashboard ~ PCORI Reviewer").Frame("Frame").WebElement("WebElement").Object.innerText = "Overall Comments---Outstanding"
populateonlinereview_asMR 0
populateonlinereview_asMR 1
populateonlinereview_asMR 2
populateonlinereview_asMR 3
populateonlinereview_asMR 4
populateonlinereview_asMR 5
populateonlinereview_asMR 6
populateonlinereview_asMR 7
populateonlinereview_asMR 8
populateonlinereview_asMR 9
populateonlinereview_asMR 10
populateonlinereview_asMR 11
populateonlinereview_asMR 12
wait 2
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_Button_usingName "Submit"
wait 2
clk_Button_usingName "OK"


'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
'clk_link_Object2 "Closed Reviews"
'wait 2
'verifyIfReviewExistsinClosed "Review Submitted"
'wait 4
'click_magnifyinglass "Review Submitted"
'''''''''''''''''Verifying Review record is locked''''''''
'wait 2
'clk_link_Object2 "Online Review"
'wait 2
End Function
Function OnlineReviewsubmissionexternalMR4()
	wait 2

'open the reviewers portal
Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer4,passWord1
clk_link_Object2 "Log in"
wait 10				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
wait 5
''''''Application Review download'''''''''''
'clk_link_Object2 "Application Reviews Downloads"
'wait 2
'SelectWebListApplicationDownloads()
'wait 2
'clk_Button_usingName "Click here to download application review materials"
'wait 2
'clk_link_Object2 "Cycle III_4567_External Portal_ FullApplication"
'wait 6
'''''CloseLatestOpenedBrowser()
'Browser("PCORI Reviewer").WinObject("WinObject").WinButton("Close Tab (Ctrl+W)").Click

'''''''''''''Going back to Dashboard''''''''''''''''''
wait 4
clk_link_Object2 "Reviewer Dashboard"
'''''''''''''Opening a Online Review record and submitting the Review''''''''''''
wait 2
click_PencilImageToEdit Reviewtype
wait 2
clk_link_Object2 "Online Review"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame").WebElement("WebElement").Object.innerText = "Criterion 1: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_2").WebElement("WebElement").Object.innerText = "Criterion 1: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:4:j_id331"
strData = "1"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_3").WebElement("WebElement").Object.innerText = "Criterion 2: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_4").WebElement("WebElement").Object.innerText = "Criterion 2: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:9:j_id331"
strData = "2"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_5").WebElement("WebElement").Object.innerText = "Criterion 3: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_6").WebElement("WebElement").Object.innerText = "Criterion 3: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:14:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_7").WebElement("WebElement").Object.innerText = "Criterion 4: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_8").WebElement("WebElement").Object.innerText = "Criterion 4: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:19:j_id331"
strData = "4"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_9").WebElement("WebElement").Object.innerText = "Criterion 5: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_10").WebElement("WebElement").Object.innerText = "Criterion 5: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:24:j_id331"
strData = "5"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_12").WebElement("WebElement").Object.innerText = "Criterion 6: Strengths"
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_13").WebElement("WebElement").Object.innerText = "Criterion 6: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:29:j_id331"
strData = "6"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:34:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
'Browser("Article Management - Console").Page("Dashboard ~ PCORI Reviewer_2").Frame("Frame").WebElement("WebElement").Object.innerText = "Overall Comments---Outstanding"
populateonlinereview_asMR 0
populateonlinereview_asMR 1
populateonlinereview_asMR 2
populateonlinereview_asMR 3
populateonlinereview_asMR 4
populateonlinereview_asMR 5
populateonlinereview_asMR 6
populateonlinereview_asMR 7
populateonlinereview_asMR 8
populateonlinereview_asMR 9
populateonlinereview_asMR 10
populateonlinereview_asMR 11
populateonlinereview_asMR 12
wait 2
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_Button_usingName "Submit"
wait 2
clk_Button_usingName "OK"

'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
'clk_link_Object2 "Closed Reviews"
'wait 2
'verifyIfReviewExistsinClosed "Review Submitted"
'wait 4
'click_magnifyinglass "Review Submitted"
'''''''''''''''''Verifying Review record is locked''''''''
'wait 2
'clk_link_Object2 "Online Review"
'wait 2



End Function

Function OnlineReviewsubmissionExternalMR5()
	wait 2


'open the reviewers portal
Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer5,passWord1
clk_link_Object2 "Log in"
wait 10				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
wait 5
''''''Application Review download'''''''''''
'clk_link_Object2 "Application Reviews Downloads"
'wait 2
'SelectWebListApplicationDownloads()
'wait 2
'clk_Button_usingName "Click here to download application review materials"
'wait 2
'clk_link_Object2 "Cycle III_4567_External Portal_ FullApplication"
'wait 6
'''''CloseLatestOpenedBrowser()
'Browser("PCORI Reviewer").WinObject("WinObject").WinButton("Close Tab (Ctrl+W)").Click

'''''''''''''Going back to Dashboard''''''''''''''''''
wait 2
clk_link_Object2 "Reviewer Dashboard"
'''''''''''''Opening a Online Review record and submitting the Review''''''''''''
wait 5
click_PencilImageToEdit Reviewtype
wait 2
clk_link_Object2 "Online Review"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame").WebElement("WebElement").Object.innerText = "Criterion 1: Strengths"
wait 2

'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_2").WebElement("WebElement").Object.innerText = "Criterion 1: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:4:j_id331"
strData = "1"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_3").WebElement("WebElement").Object.innerText = "Criterion 2: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_4").WebElement("WebElement").Object.innerText = "Criterion 2: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:9:j_id331"
strData = "2"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_5").WebElement("WebElement").Object.innerText = "Criterion 3: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_6").WebElement("WebElement").Object.innerText = "Criterion 3: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:14:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_7").WebElement("WebElement").Object.innerText = "Criterion 4: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_8").WebElement("WebElement").Object.innerText = "Criterion 4: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:19:j_id331"
strData = "4"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_9").WebElement("WebElement").Object.innerText = "Criterion 5: Strengths"
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_10").WebElement("WebElement").Object.innerText = "Criterion 5: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:24:j_id331"
strData = "5"
selectWeblist5 HtmlID,strData
wait 2
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_12").WebElement("WebElement").Object.innerText = "Criterion 6: Strengths"
'Browser("PCORI Reviewer_2").Page("Dashboard ~ PCORI Reviewer").Frame("Frame_13").WebElement("WebElement").Object.innerText = "Criterion 6: Weaknesses"
wait 2
HtmlID = "j_id0:mainForm:j_id241:29:j_id331"
strData = "6"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:34:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
'Browser("Article Management - Console").Page("Dashboard ~ PCORI Reviewer").Frame("Frame").WebElement("WebElement").Object.innerText = "Overall Comments---Outstanding"
populateonlinereview_asMR 0
populateonlinereview_asMR 1
populateonlinereview_asMR 2
populateonlinereview_asMR 3
populateonlinereview_asMR 4
populateonlinereview_asMR 5
populateonlinereview_asMR 6
populateonlinereview_asMR 7
populateonlinereview_asMR 8
populateonlinereview_asMR 9
populateonlinereview_asMR 10
populateonlinereview_asMR 11
populateonlinereview_asMR 12
wait 2
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_Button_usingName "Submit"
wait 2
clk_Button_usingName "OK"

'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
'clk_link_Object2 "Closed Reviews"
'wait 2
'verifyIfReviewExistsinClosed "Review Submitted"
'wait 4
'click_magnifyinglass "Review Submitted"
'''''''''''''''''Verifying Review record is locked''''''''
'wait 2
'clk_link_Object2 "Online Review"
'wait 3
End Function

Public Function click_webElementReviewlistinPanel()
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
l("class").value = "listTitle"
l("outerhtml").value = "<span class=""listTitle"">Reviews<span class=""count"">\[5\]</span></span>"
l("visible").value = True
l("html tag").value = "SPAN"
l("innertext").Value = "Reviews\[5\]"
Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click

End Function

Function RequestingupdateOnlineReviewsubmissionexternalMR1() 
	wait 2


'open the reviewers portal
Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer1,"test123456"
clk_link_Object2 "Log in"
wait 10				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"


wait 2
click_PencilImageToEdit "Online Review - AD"
wait 2
clk_link_Object2 "Online Review"
wait 2
HtmlID = "j_id0:mainForm:j_id241:4:j_id331"
strData = "2"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:9:j_id331"
strData = "3"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:14:j_id331"
strData = "4"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:19:j_id331"
strData = "5"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:24:j_id331"
strData = "6"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:29:j_id331"
strData = "7"
selectWeblist5 HtmlID,strData
wait 2
HtmlID = "j_id0:mainForm:j_id241:33:j_id331"
strData = "5"
selectWeblist5 HtmlID,strData
wait 2
'Browser("Article Management - Console").Page("Dashboard ~ PCORI Reviewer_3").Frame("Frame").WebElement("WebElement").Object.innerText = "Overall Comments--Okay"
wait 2
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_link_Object2 "Submit"
wait 2
clk_Button_usingName "OK"




'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
clk_link_Object2 "Closed Reviews"
wait 2
verifyIfReviewExistsinClosed "Review Submitted"
wait 4
click_magnifyinglass "Review Submitted"
''''''''''''''''Verifying Review record is locked''''''''
wait 2
clk_link_Object2 "Online Review"
wait 2
End Function

Function clickReviewRecordBasedOnStatusAll()
	On error resume next
	
 ''''''''''''''''1st Review record Update
 wait 2
clickReviewRecordBasedOnStatus "Review Submitted"
 wait 2
 verifyIfWebElementDoesExist1 "Revisions In"
wait 4
clk_Button_usingName "Edit"
 wait 2
 changeReviewStatus "Scrubbing"
 wait 2
clk_Button_usingName "Save"
wait 4
clk_Button_usingName "Edit"
wait 2
changeReviewStatus "Review Finalized"
wait 2
clk_Button_usingName "Save"


''''''''''''''2nd Review record Update''''''''''''
wait 2
clickReviewRecordBasedOnStatus "Review Submitted"
wait 2
clk_Button_usingName "Edit"
wait 2
changeReviewStatus "Review Finalized"
wait 2
clk_Button_usingName "Save"
stPanel = readFromFilePanel()
wait 2
clk_link_Object2 stPanel
'''''''''''''''3rd Review record update
wait 2
clickReviewRecordBasedOnStatus "Review Submitted"
Wait 2
clk_Button_usingName "Edit"
wait 2
changeReviewStatus "Review Finalized"
wait 2
clk_Button_usingName "Save" 
stPanel = readFromFilePanel()
wait 2
clk_link_Object2 stPanel
wait 2
''''''''''''''4th Review Record Update'''''''
wait 2
clickReviewRecordBasedOnStatus "Review Submitted"
Wait 2
clk_Button_usingName "Edit"
wait 2
changeReviewStatus "Review Finalized"
wait 2
clk_Button_usingName "Save" 
stPanel = readFromFilePanel()
wait 2
clk_link_Object2 stPanel
wait 2
''''''''''''''5th Review Record Update'''''''''''

wait 2
clickReviewRecordBasedOnStatus "Review Submitted"
Wait 2
clk_Button_usingName "Edit"
wait 2
changeReviewStatus "Review Finalized"
wait 2
clk_Button_usingName "Save" 
stPanel = readFromFilePanel()
wait 2
clk_link_Object2 stPanel
wait 2

wait 2
clickReviewRecordBasedOnStatus "Review Submitted"
Wait 2
clk_Button_usingName "Edit"
wait 2
changeReviewStatus "Review Finalized"
wait 2
clk_Button_usingName "Save" 
stPanel = readFromFilePanel()
wait 2
clk_link_Object2 stPanel
wait 2


End Function 

Public Function click_webElementBasedonfieldname(Html,strName)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
l("class").value = "labelCol"
'l("outerhtml").value = "<span class=""listTitle"">Reviews<span class=""count"">\[5\]</span></span>"
l("visible").value = True
l("html tag").value = "TD"
l("innertext").Value = strName
l("innerhtml").value = Html
Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click
End Function
Function writeToAfileRADiscussionorderView(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\discussionorder.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function
Public Function readFromRADiscussionorderView()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\discussionorder.txt",1)
          
          sContent = objTxtFile.Readline
          readFromRADiscussionorderView = sContent
End Function

Function createRADiscussionOrderasPO(stPanel)
On error Resume Next
wait 2
navigateAndLoginToSalesForce scienceoperationuser, KatiePassword
wait 2
'clk_link_Object2 "Projects"
wait 2
'HtmlID = "fcf"
'strData = "RA App Discussion Order"
'selectWeblist5 HtmlID,strData

wait 3		
		  'clk_link_Object2 "Edit"
		  wait 2
		  stPanel = readFromFilePanel()
		  HtmlId = "fname"	 
		 strData = "RA App Discussion Order" & space(1) & stPanel
		 set_Web_Edit HtmlID,strData
		  
		  strViewName = "RA App Discussion Order" & space(1) & stPanel
		  'write this name to flat file
		  writeToAfileRADiscussionorderView strViewName	 
		wait 3
        HtmlID = "fcol2"
        strData = "Panel"
		 selectWeblist5 HtmlID,strData
		 wait 2
		 
	    HtmlID = "fop2"
        strData = "equals"
		 selectWeblist5 HtmlID,strData	
		 wait 2
		 HtmlID = "fval2"
         stPanel = readFromFilePanel()		 
		 strData = stPanel
		 set_Web_Edit HtmlID,strData
		 clk_Button_usingName "Save"
		 createRADiscussionOrderasPO = strViewName
		 wait 2	 
		   
				 
End Function

Function clickOnDiscussionodercheckbox_new()
				On error resume next
		 Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("column names").value="Add to Discussion Line;;All Online Reviews Completed\?;Yes"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 		  	       
				For x  = 1 to a     
						For j  =1 to b
							c = l_link(0).GetCellData(x,j)
							print c
							Select Case c						
													    	
                                         Case "Add to Discussion Line"
	                                 	     Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebCheckBox", 0)
	                                         oEdit1.click
	                                         Exit function                                       
                                       
                                 	                                 
							                                                                    
							End Select   
                         						
												Next
						
				Next                             
					 
     End Function 
     
     Function POcheckingdiscussionlinecheckbox() 
     On error resume next
     wait 2
     stProject = readFromFileProject()
clk_link_Object2 stProject
wait 2  
	    
clk_Button_usingName "Edit"
wait 4
clickOnDiscussionodercheckbox_new()
wait 4
clk_Button_usingName "Save"
wait 5
clk_link_Object2 "Projects"
wait 2
strViewName = readFromRADiscussionorderView
HtmlID = "fcf"
strData = strViewName
selectWeblist5 HtmlID,strData
wait 2
     End Function
Function writeToAfileRADiscussionorderViewMRO(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\discussionordermro.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function
Public Function readFromRADiscussionorderViewMRO()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\discussionordermro.txt",1)
          
          sContent = objTxtFile.Readline
          readFromRADiscussionorderViewMRO = sContent
End Function
Function MROsettingdiscussionorderranking() 
On error resume next
	'navigateAndLoginToSalesForce mrManagementUser, somersPassword
wait 2
clk_link_Object2 "Projects"
wait 2
HtmlID = "fcf"
strData = "RA App Discussion Order"
selectWeblist5 HtmlID,strData

wait 3		
		  clk_link_Object2 "Edit"
		  wait 2
		  stPanel = readFromFilePanel()
		  HtmlId = "fname"	 
		 strData = "RA App Discussion Order MRO" & space(1) & stPanel
		 set_Web_Edit HtmlID,strData
		  
		  strViewName = "RA App Discussion Order" & space(1) & stPanel
		  'write this name to flat file
		  writeToAfileRADiscussionorderViewMRO strViewName	 
		wait 3
        HtmlID = "fcol2"
        strData = "Panel"
		 selectWeblist5 HtmlID,strData
		 wait 2
		 
	    HtmlID = "fop2"
        strData = "equals"
		 selectWeblist5 HtmlID,strData	
		 wait 2
		 HtmlID = "fval2"
         stPanel = readFromFilePanel()		 
		 strData = stPanel
		 set_Web_Edit HtmlID,strData
		 clk_Button_usingName "Save As"		
			 
'''''''''''''''''''''MRO setting ranking for discussion order'''''''

wait 2
     stProject = readFromFileProject1()
clk_link_Object2 stProject
wait 2  
	    
clk_Button_usingName "Edit"
wait 4				 
set_Web_Edit "00N39000003LYdQ","1"
wait 4
clk_Button_usingName "Save"
wait 5
clk_link_Object2 "Projects"
wait 2
strViewName = readFromRADiscussionorderViewMRO
HtmlID = "00B0x000000siq9_listSelect"
strData = strViewName
selectWeblist5 HtmlID,strData
wait 2
'clk_Button_usingName "Go!"
End Function

Function ReviewerverifyingPrepforinpersonreview()
	'''''open the reviewers portal
Open_ReviewersPortal()				
wait 5				
''Login to the portal
login_intoSalesForce_Application userReviewer3,passWord1
clk_link_Object2 "Log in"
wait 10				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
wait 5
click_PencilImageToEdit "Prep for In-Person"
wait 2
clk_link_Object2 "Report Conflict"
wait 3
clk_link_Object2 "In-Person Review"
wait 2
verifyIfWebElementDoesExist1 "The In-Person Review is not yet open. In the meantime, please review the application details in preparation for the meeting"
'wait 2
'clk_link_Object2 "Application Reviews Downloads"
'wait 2
'SelectWebListApplicationDownloads()
'wait 2
'clk_Button_usingName "Click here to download application review materials"
wait 2
End Function

Function writeToAfileMRprepforinpersonview(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\mrprepforinperson.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function
Public Function readFromMRprepforinpersonview()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\mrprepforinperson.txt",1)
          
          sContent = objTxtFile.Readline
         readFromMRprepforinpersonview = sContent
End Function

Function MROaccessingPrepforinpersonreviewrecord()
	wait 2
navigateAndLoginToSalesForce mrManagementUser, careyPassword
wait 2
clk_link_Object2 "Merit Reviewer Management"
clk_link_Object2 "Reviews"
wait 2
HtmlID = "fcf"
strData = "MR Prep for In-Person"
selectWeblist5 HtmlID,strData
wait 2
clk_link_Object2 "Edit"
		  wait 2
		  stPanel = readFromFilePanel()
		  HtmlId = "fname"	 
		 strData = "MR Prep for In-Person" & space(1) & stPanel
		 set_Web_Edit HtmlID,strData
		  
		  strViewName = "MR Prep for In-Person" & space(1) & stPanel
		  'write this name to flat file
		  writeToAfileMRprepforinpersonview strViewName	 
		wait 3
        HtmlID = "fcol3"
        strData = "Panel"
		 selectWeblist5 HtmlID,strData
		 wait 2
		 
	    HtmlID = "fop3"
        strData = "equals"
		 selectWeblist5 HtmlID,strData	
		 wait 2
		 HtmlID = "fval3"
         stPanel = readFromFilePanel()		 
		 strData = stPanelDI
		 set_Web_Edit HtmlID,strData
		 clk_Button_usingName "Save As"	
		 wait 2
'''''''''''''''1st Review''''''''''''
'stProject = readFromFileProject1()
'Click_onReviewfromlistviewMRPrep stProject
'wait 4
'clk_Button_usingName "Edit"
'wait 2
'clk_singelCheckbox()
'wait 2
'clk_Button_usingName "Save"
''''''''''''2nd review'''''''''''''
'wait 4
'clk_link_Object2 "Reviews"
'wait 2
'HtmlID = "fcf"
'strData = readFromMRprepforinpersonview()
'selectWeblist5 HtmlID,strData
'wait 2
'clk_Button_usingName "Go!"
'stProject = readFromFileProject1()
'Click_onReviewfromlistviewMRPrep stProject
'wait 2
'clk_Button_usingName "Edit"
'wait 2
'clk_singelCheckbox()
'wait 2
'clk_Button_usingName "Save"
'
'''''''''''''3rd Review''''''''''''
'wait 4
'clk_link_Object2 "Reviews"
'wait 2
'wait 2
'HtmlID = "fcf"
'strData = readFromMRprepforinpersonview()
'selectWeblist5 HtmlID,strData
'wait 2
'clk_Button_usingName "Go!"
'stProject = readFromFileProject1()
'Click_onReviewfromlistviewMRPrep stProject
'wait 2
'clk_Button_usingName "Edit"
'wait 2
'clk_singelCheckbox()
'wait 2
'clk_Button_usingName "Save"
''''''''''''''''''''4th Review''''''''''
'wait 4
'clk_link_Object2 "Reviews"
'wait 2
'wait 2
'HtmlID = "fcf"
'strData = readFromMRprepforinpersonview()
'selectWeblist5 HtmlID,strData
'wait 2
'clk_Button_usingName "Go!"
'stProject = readFromFileProject1()
'Click_onReviewfromlistviewMRPrep stProject
'clk_Button_usingName "Edit"
'wait 2
'clk_singelCheckbox()
'wait 2
'clk_Button_usingName "Save"
'''''''''''''''''''''5th Review'''''''''
'wait 4
'clk_link_Object2 "Reviews"
'wait 2
'wait 2
'HtmlID = "fcf"
'strData = readFromMRprepforinpersonview()
'selectWeblist5 HtmlID,strData
'wait 2
'clk_Button_usingName "Go!"
'stProject = readFromFileProject1()
'Click_onReviewfromlistviewMRPrep stProject
'wait 2
'clk_Button_usingName "Edit"
'wait 2
'clk_singelCheckbox()
'wait 2
'clk_Button_usingName "Save"
'wait 2
'stPanel =  readFromFilePanel()
'navigateToAccoutsPageThrughTopSearchBox stPanel
clk_singelCheckboxbasedonindexid 0

dblClick_OpenInperson_Internal_listview "a0yc0000003Px99_00N39000003LYYV"
clk_singelCheckboxby_htmlId "00N39000003LYYV"
click_webElement_forengagementreport_selecting_byindex "All 5 selected records",0

clk_Button_usingName "Save"

navigateToAccoutsPageThrughTopSearchBox stPanelDI
wait 2
clk_link_Object2 stPanelDI
wait 2
click_webElementExternalReviewlistinaproject "Reviews\[5\+\]"
wait 2
clk_link_Object2 "Go to list \(10\) »"	
verifyIfWebElementDoesExist1 "In-Person Review"
verifyIfWebElementDoesExist1 "Ready to Review"
clk_link_Object2 "Panel: "& stPanelDI
		 
End Function

Function clickprepforinpersonRecordBasedOnStatus(strStatus)
	On error resume next
	 Set odesc=Description.Create
    odesc("micclass").value="WebTable"
    odesc("column names").value=";Choose a RowSelect All;Sort by:Review IDSorted: NoneShow Review ID column actions;Sort by:ReviewerSorted: NoneShow Reviewer column actions;Sort by:Record TypeSorted: NoneShow Record Type column actions;Sort by:ApplicationSorted: NoneShow Application column actions;Sort by:Application ProgramSorted AscendingShow Application Program column actions;Sort by:Primary Reviewer RoleSorted: NoneShow Primary Reviewer Role column actions;Sort by:Reviewer TypeSorted: NoneShow Reviewer Type column actions;Sort by:StatusSorted: NoneShow Status column actions;Sort by:DeadlineSorted: NoneShow Deadline column actions;"
    '''''odesc("cols").value= 10  
      '''odesc("cols").value= 12         
'            
    Set l_link=  getParentObject().ChildObjects(odesc)
     print l_link.count
                 a = l_link(0).GetROProperty("rows") 
                 print a
                           b = l_link(0).GetROProperty("cols")  
print b                           
                        For x  = 0 to a    
                             For j  =0 to b
                                      c =  l_link(0).GetCellData(x,j)
                                   
                                      print c
                                   If Trim(c) = strStatus Then
                                     	'Set o = l_link(2).ChildItem(x,2, "Link", 0) 
                                   Set o = l_link(2).ChildItem(x,2, "WebElement", 0)   	
                                    	o.click
                                      	intText =  o.getRoProperty("innertext")
              	
                                      	 Exit function
                                      End If
'                                      
                           next
                       next    
                'clickReviewRecordBasedOnStatus =  intText
clickprepforinpersonRecordBasedOnStatus =  strStatus
End Function 

'Function to click webelement based on the Link kin the same row
Public Function Click_onReviewfromlistviewMRPrep(StrInnertext)
                
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
                                
Err.Clear
On Error Resume Next   
''''                'Capture the page no value and convert it into Integer
''''PageStr = Brwser.WebElement("class:=right", "html tag:=SPAN").GetRoProperty("innertext")
''''PageStr=Split(PageStr, "Pageof")
''''PageNo = Cint (PageStr(1))
'''''print "Page No is " &PageNo
PageNo = 2
For i = 1 To PageNo
'Create the description WebTable object
                Set oWTable = Description.Create
                                oWTable("micclass").value = "WebTable"
                                oWTable("html tag").value = "TABLE" 
                                oWTable("class").value = "x-grid3-row-table"

                Set Tables = Brwser.ChildObjects(oWTable)
Print "# of Tables = " & Tables.Count



                Set oDesc = Description.Create
                oDesc("micclass").value = "Link"
                oDesc("html tag").value = "A" 
                oDesc("visible").value = "True"
                oDesc("innertext").value = StrInnertext

                For j = 0 To Tables.Count-1
                
                                If Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link(oDesc).Exist(10) Then
                                                Wait 5
                                                'WebTableIndex = j
                                                k=0
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link(oDesc).Highlight
                                                Wait 3
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link("html tag:=A","visible:=True", "location:="&k+2).Highlight
                                                Wait 3
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link("html tag:=A","visible:=True", "location:="&k+2).Click
                                                Wait 3
                                
                                        Print " Click on Webelement according to Email link name done successfully" 
Exit Function  
                                                                
                                        LogReport 0,"Click on the link according to email link name " & StrInnertext," email " & "-" & StrInnertext & "-" & " Selected Successfully"  
                
                                End If
                                                
                Next

Next 
                                     
                                                
                If err <> 0 Then
LogReport 4,"Error ", err.number & "-" & err.description
LogReport 1,"Click on the checkbox according to Project link name " & StrInnertext," Project " & "-" & StrInnertext & "-" & " Not Selected"         
                
                End If

Set oDesc = Nothing
Set oWTable = Nothing
                
End Function

Function InpersonReviewsubmissionMR1()
	Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer1,passWord1
clk_link_Object2 "Log in"
wait 15				
				
clk_link_Object2 "Access the Merit Reviewer Dashboard"
'wait 2
'clk_link_Object2 "Application Reviews Downloads"
'wait 2
'SelectWebListApplicationDownloads()
wait 4
clk_link_Object2 "Reviewer Dashboard"
wait 2
click_PencilImageToEdit "Ready to Review"
wait 2
clk_link_Object2 "Report Conflict"
wait 2
selectWeblist5 "j_id0:mainForm:j_id241:0:j_id331","None"
wait 2
clk_Button_usingName "Save"
wait 2
clk_link_Object2 "In-Person Review"
wait 2
selectWeblist5 "j_id0:mainForm:j_id241:5:j_id331","1"
wait 2
populateonlinereview_asMR 0
wait 2
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_Button_usingName "Submit"
wait 2
clk_Button_usingName "OK"






'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
clk_link_Object2 "Closed Reviews"
wait 2
verifyIfReviewExistsinClosed "Review Submitted"
wait 4
click_magnifyinglass "Review Submitted"
wait 2
clk_link_Object2 "Report Conflict"
wait 2
clk_link_Object2 "In-Person Review"
End Function
Function InpersonReviewsubmissionMR2()
	Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer2,passWord1
clk_link_Object2 "Log in"
wait 15				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
'wait 2
'clk_link_Object2 "Application Reviews Downloads"
'wait 2
'SelectWebListApplicationDownloads()
wait 4
clk_link_Object2 "Reviewer Dashboard"
wait 2
click_PencilImageToEdit "Ready to Review"
wait 2
clk_link_Object2 "Report Conflict"
wait 2
selectWeblist5 "j_id0:mainForm:j_id241:0:j_id331","None"
wait 2
clk_Button_usingName "Save"
wait 2
clk_link_Object2 "In-Person Review"
wait 2
selectWeblist5 "j_id0:mainForm:j_id241:5:j_id331","2"
wait 2
populateonlinereview_asMR 0
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_Button_usingName "Submit"
wait 2
clk_Button_usingName "OK"

'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
clk_link_Object2 "Closed Reviews"
wait 2
verifyIfReviewExistsinClosed "Review Submitted"
wait 4
click_magnifyinglass "Review Submitted"
wait 2
clk_link_Object2 "Report Conflict"
wait 2
clk_link_Object2 "In-Person Review"
End Function
Function InpersonReviewsubmissionMR3()
	Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer3,passWord1
clk_link_Object2 "Log in"
wait 15				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
'wait 2
'clk_link_Object2 "Application Reviews Downloads"
'wait 2
'SelectWebListApplicationDownloads()
wait 4
clk_link_Object2 "Reviewer Dashboard"
wait 2
click_PencilImageToEdit "Ready to Review"
wait 2
clk_link_Object2 "Report Conflict"
wait 2
selectWeblist5 "j_id0:mainForm:j_id241:0:j_id331","None"
wait 2
clk_Button_usingName "Save"
wait 2
clk_link_Object2 "In-Person Review"
wait 2
selectWeblist5 "j_id0:mainForm:j_id241:5:j_id331","3"
wait 2
populateonlinereview_asMR 0
wait 2
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_Button_usingName "Submit"
wait 2
clk_Button_usingName "OK"

'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
clk_link_Object2 "Closed Reviews"
wait 2
verifyIfReviewExistsinClosed "Review Submitted"
wait 4
click_magnifyinglass "Review Submitted"
wait 2
clk_link_Object2 "Report Conflict"
wait 2
clk_link_Object2 "In-Person Review"
End Function
Function InpersonReviewsubmissionMR4()
	Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer4,passWord1
clk_link_Object2 "Log in"
wait 15				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
'wait 2
'clk_link_Object2 "Application Reviews Downloads"
'wait 2
'SelectWebListApplicationDownloads()
wait 4
clk_link_Object2 "Reviewer Dashboard"
wait 2
click_PencilImageToEdit "Ready to Review"
wait 2
clk_link_Object2 "Report Conflict"
wait 2
selectWeblist5 "j_id0:mainForm:j_id241:0:j_id331","None"
wait 2
clk_Button_usingName "Save"
wait 2
clk_link_Object2 "In-Person Review"
wait 2
selectWeblist5 "j_id0:mainForm:j_id241:5:j_id331","4"
wait 2
populateonlinereview_asMR 0
wait 2
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_Button_usingName "Submit"
wait 2
clk_Button_usingName "OK"

'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
clk_link_Object2 "Closed Reviews"
wait 2
verifyIfReviewExistsinClosed "Review Submitted"
wait 4
click_magnifyinglass "Review Submitted"
wait 2
clk_link_Object2 "Report Conflict"
wait 2
clk_link_Object2 "In-Person Review"
End Function
Function InpersonReviewsubmissionMR5()
	Open_ReviewersPortal()
				
wait 5
				
''Login to the portal
login_intoSalesForce_Application userReviewer5,passWord1
clk_link_Object2 "Log in"
wait 15				
				
clk_link_Object2 "ACCESS THE MERIT REVIEWER DASHBOARD"
'wait 2
'clk_link_Object2 "Application Reviews Downloads"
'wait 2
'SelectWebListApplicationDownloads()
wait 4
clk_link_Object2 "Reviewer Dashboard"
wait 2
click_PencilImageToEdit "Ready to Review"
wait 2
clk_link_Object2 "Report Conflict"
wait 2
wait 2
selectWeblist5 "j_id0:mainForm:j_id241:0:j_id331","None"
wait 2
clk_Button_usingName "Save"
wait 2
clk_link_Object2 "In-Person Review"
wait 2
selectWeblist5 "j_id0:mainForm:j_id241:5:j_id331","5"
wait 2
populateonlinereview_asMR 0
wait 2
clk_Button_usingName "Save"
wait 2
clickReviewSubmitButton()
wait 2
clk_Button_usingName "Submit"
wait 2
clk_Button_usingName "OK"

'''''''''''''''Verifying review moved to Closed review after submitting
wait 4
clk_link_Object2 "Closed Reviews"
wait 2
verifyIfReviewExistsinClosed "Review Submitted"
wait 4
click_magnifyinglass "Review Submitted"
wait 2
clk_link_Object2 "Report Conflict"
wait 2
clk_link_Object2 "In-Person Review"
End Function

Public Function click_webElementinternal(strName)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
l("class").value = "numericalColumn zen-deemphasize"
'l("outerhtml").value = "<span class=""listTitle"">Reviews<span class=""count"">\[5\]</span></span>"
l("visible").value = True
l("html tag").value = "TH"
l("innertext").Value = strName
'l("innerhtml").value = Html
'l("html id").value = HtmlID
Set edObj =  getParentObject().ChildObjects(l)
edObj(0).Highlight
edObj(0).click
End Function

Function MROSummarystatementgeneration()
	wait 2
navigateAndLoginToSalesForce mrManagementUser, careyPassword
wait 2
clk_link_Object2 "Merit Reviewer Management"
wait 2
'stPanel =  readFromFilePanel()
navigateToAccoutsPageThrughTopSearchBox stPanelDI
wait 2
clk_link_Object2 stPanelDI
wait 3
click_webElementBasedonfieldname "Average In-Person Score","Average In-Person Score"
wait 2
'stProject = readFromFileProject()
clk_link_Object2 AppDI1
wait 3
click_webElementBasedonfieldname "Average In-Person Score","Average In-Person Score"
wait 2
'stPanel =  readFromFilePanel()
clk_link_Object2 stPanelDI
wait 2
'clk_link_Object2 "In-Person Score Report"
'wait 10
'stProject = readFromFileProject()
clk_link_Object2 AppDI1
wait 3
dblclkInPersonNotes
populateInlinePerson
'clk_Button_usingName "Edit"
'wait 3
'Browser("PCORI Reviewer_2").Page("Project Edit: 1AprojAUT_vl0").Frame("Frame").WebElement("00N39000003LYe3EAG_rta_body").Object.innerText = "Test In Person Discussion Notes"
wait 2
clk_Button_usingName "Save"
wait 4
clk_Button_usingName "Generate Summary Statement"
wait 20
CloseLatestOpenedBrowser()
wait 4
'click_webElementinternalNotesnattachment "<img width=""1"" height=""1"" title="""" class=""minWidth"" alt="""" src=""/img/s\.gif""><img title="""" class=""relatedListIcon"" alt="""" src=""/img/s\.gif""><h3 id=""0060x000003wtx9_RelatedNoteList_title"">Notes &amp; Attachments</h3>","Notes & Attachments"
wait 2
'stPanel =  readFromFilePanel()
clk_link_Object2 stPanelDI
wait 2
clk_Button_usingName "Calculate Quartile"
wait 2
click_webElementBasedonfieldname "Quartile","Quartile"
wait 2
'stProject = readFromFileProject()
clk_link_Object2 AppDI1
wait 2
click_webElementBasedonfieldname "Quartile","Quartile"
wait 4
clk_Button_usingName "Generate Summary Statement"
wait 20
CloseLatestOpenedBrowser()
wait 2
clk_Button_usingName "Edit"
wait 2
clk_Button_usingName "Save"
wait 2
click_webElementBasedonfieldname "Summary Statement Download","click_webElementBasedonfieldname"
wait 4
'click_webElementinternalNotesnattachment "<img width=""1"" height=""1"" title="""" class=""minWidth"" alt="""" src=""/img/s\.gif""><img title="""" class=""relatedListIcon"" alt="""" src=""/img/s\.gif""><h3 id=""0060x000003wtx9_RelatedNoteList_title"">Notes &amp; Attachments</h3>","Notes & Attachments"
wait 2
'stPanel =  readFromFilePanel()
clk_link_Object2 stPanelDI
wait 2
clk_Button_usingName "Edit"
wait 2
selectWeblist6 "00N39000003LYhf","Yes, include quartile"
wait 2
clk_Button_usingName "Save"
wait 2
'stProject = readFromFileProject()
clk_link_Object2 AppDI1
wait 4
clk_Button_usingName "Generate Summary Statement"
wait 20
CloseLatestOpenedBrowser()
'wait 4
'click_webElementinternalNotesnattachment "<img width=""1"" height=""1"" title="""" class=""minWidth"" alt="""" src=""/img/s\.gif""><img title="""" class=""relatedListIcon"" alt="""" src=""/img/s\.gif""><h3 id=""0060x000003wtx9_RelatedNoteList_title"">Notes &amp; Attachments</h3>","Notes & Attachments"
End Function
Function setWebEditBoxany(HtmlID,strData)
     	  On error resume next
  	       	    
Set odesc=Description.Create
odesc("micclass").value="WebEdit"
odesc("visible").value= True

odesc("html id").value= HtmlID 
'odesc("innertext").value= strLabelName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y        

Set od =Description.Create
od("micclass").value="WebEdit"
od("abs_x").value= x 
od("abs_y").value= y 
'           

Set Os =  getParentObject().ChildObjects(od)

print Os.count

Os(0).set strData
     End Function
     Function selectWeblistany(HtmlID,strData)
			On error resume next

            wait 4
			   strArr = split(strData,",")
			   
			   Set btncalc = Description.Create()  
			    btncalc("micclass").value = "Browser"
			  
			    Set btn =DeskTop.ChildObjects(btncalc)
			    strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
			    print strHwnd
			  
			    Set pa = Description.Create
			    pa("micclass").value = "Page"
			  
			   
			    
			    print pHwnd
			    Set editObject =  getParentObject()
			    
			    Set editO = Description.Create
			    
			    editO("micclass").value = "WebList"
			    editO("visible").value = true 
			    editO("html id").value = HtmlID
			    editO("html tag").value = "SELECT"
			    'editO("name").value = "File Share Settings" 
			    'editO("type").value = "text" 
			     Set edObj =  getParentObject().ChildObjects(editO) 
			     print edObj.count
			     for i = 0 To uBound(strArr)
			        items =   edObj(i).GetRoProperty("all items")
			        print items
			        print strArr(i)
			        If Not(strArr(i) = "")Then	
                        
                                             
			            edObj(i).select trim(strArr(i))        
			        End If
			         
			     Next
			           
         
     End Function
     Function selectWeblistformanualshare(HtmlID,strData)
			On error resume next

            wait 4
			   strArr = split(strData,",")
			   
			   Set btncalc = Description.Create()  
			    btncalc("micclass").value = "Browser"
			  
			    Set btn =DeskTop.ChildObjects(btncalc)
			    strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
			    print strHwnd
			  
			    Set pa = Description.Create
			    pa("micclass").value = "Page"
			  
			   
			    
			    print pHwnd
			    Set editObject =  getParentObject()
			    
			    Set editO = Description.Create
			    
			    editO("micclass").value = "WebList"
			    editO("visible").value = true 
			    editO("html id").value = HtmlID
			    editO("html tag").value = "SELECT"
			    'editO("name").value = "File Share Settings" 
			    'editO("type").value = "text" 
			     Set edObj =  getParentObject().ChildObjects(editO) 
			     print edObj.count
			     for i = 0 To uBound(strArr)
			        items =   edObj(i).GetRoProperty("all items")
			        print items			        
			        print strArr(i)
			        If Not(strArr(i) = "")Then	
                        
                                             
			            edObj(i).Select trim(strArr(i))        
			        End If
'			         
			     Next
			           
         
     End Function
     Public Function clk_rightarrow_Link_todomanualshare()
Err.Clear
On error resume next

wait 5
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
'print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "Image"
l("class").value = "rightArrowIcon"
l("alt").value = "Add"
l("image type").value = "Image Link"
l("html tag").value = "IMG"



Set lo =  getParentObject().ChildObjects(l)
'print lo.count
lo(0).highlight
lo(0).click

If err <> 0 Then
    LogReport 4,"Error", err.number & "-" & err.description
    LogReport 1,"cliking the link" & "'+' Tab","Failed to click the link" & "-" & "'+' Tab"
else
      LogReport 0,"cliking the link" & "'+' Tab","The Link" & "-" & "'+' Tab" & "-" & "is clicked successfully"
End If
End Function 

Public Function captureWebElementText_ADPNumber()
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("Visible").value = True 
l("html tag").value = "TD"
l("html id").value = "j_id0:j_id2:j_id3:j_id35:j_id36:0:j_id41"
Set lo =  getParentObject().ChildObjects(l)
Print lo.count

ADPnumber = lo(0).getRoProperty("innertext")
Print ADPnumber
'Write captured MRA-Number to Function itself in order to call it later in a script
captureWebElementText_ADPNumber = ADPnumber
End Function

Public Function click_webElementinternal2()
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
l("class").value = "dataCell  numericalColumn"
'l("outerhtml").value = "<span class=""listTitle"">Reviews<span class=""count"">\[5\]</span></span>"
l("visible").value = True
l("html tag").value = "TD"
'l("innertext").Value = strName
'l("innerhtml").value = Html
'l("html id").value = HtmlID
Set edObj =  getParentObject().ChildObjects(l)
edObj(0).Highlight
edObj(0).click
End Function

Public Function captureWebElementText_onlinereviewscore()
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("class").value = "dataCell  numericalColumn"
l("Visible").value = True 
l("html tag").value = "TD"

'l("html id").value = "j_id0:j_id2:j_id3:j_id35:j_id36:0:j_id41"
Set lo =  getParentObject().ChildObjects(l)
lo(0).Highlight
Print lo.count

score = lo(0).getRoProperty("innertext")
Print score
'Write captured MRA-Number to Function itself in order to call it later in a script
captureWebElementText_onlinereviewscore = score
 If score =  "numeric value" Then
                
                Print "number Captured successfully"
                                LogReport 0," webelement "  & "-" & " numbercaptured_ successfully verified_expected"
                Else  
                                LogReport 1," webbelement " & "-" &  " Failed - number did not captured_ not expected "
                End If
                
                                If err <> 0 Then
                                                LogReport 4,"Error", err.number & "-" & err.description

                                End  IF


End Function

Public Function captureWebElementText_reviewscoreinapplication(HtmlID)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("Visible").value = True 
l("html tag").value = "DIV"
l("html id").value = HtmlID
Set lo =  getParentObject().ChildObjects(l)
lo(0).highlight
Print lo.count

If lo.count = 1 Then

   Print "Innertext captured successfully"
                                LogReport 0," webelement "  & "-" & " numbercaptured_ successfully verified_expected"
                Else  
                                LogReport 1," webbelement " & "-" &  " Failed - number did not captured_ not expected "
                End If
                
                                If err <> 0 Then
                                                LogReport 4,"Error", err.number & "-" & err.description

                                End  IF

	

	


onlinereviewscore = lo(0).getRoProperty("innertext")
Print onlinereviewscore
'Write captured MRA-Number to Function itself in order to call it later in a script
captureWebElementText_reviewscoreinapplication = onlinereviewscore

End Function
Function writeToAfileProject1(stContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\project1.txt",2,true)     'open in write mode
            f.Write (stContent)
            f.Close
            Set f = nothing
End Function
Public Function readFromFileProject1()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\project1.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFileProject1 = sContent
End Function

Function clk_link_Object_insetupinternal(HtmlID,strName)
On error resume next

wait 5
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "Link"
l("html id").value = HtmlID
l("html tag").value = "A"
l("innertext").value = strName


Set lo =  getParentObject().ChildObjects(l)
print lo.count
lo(0).click

  If err <> 0 Then
  	LogReport 4,"Error", err.number & "-" & err.description
  	LogReport 1,"cliking the link" & strName,"Failed to click the link" & "-" & strName
  	else
  	LogReport 0,"cliking the link" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
  End If



End Function

''''''Working as expected when creating panel with DI online review record
Function create_Panelwith_DI_onlinereviewrecordtype()
			On error resume next
			 strDate =  now + 30
             strA = split(strDate,":")
             strDeadLineDate = strA(0) & ":" & strA(1) & space(1) & "PM"
             
              strDate2 =  now + 30
              strB = split(strDate2,":")
              strformatted = strB(0) & ":" & strB(1) & space(1) & "PM"             
'              wait 2
'              strDate3 =  now + 30
'              strC = split(strDate3,":")
'             strformatted1 = strC(0) & ":" & strC(1) '& space(1) & "PM"  
'              
              
			'strNameEvent = "Panel 1" & "_" & RandomString(3) 
			strNameEvent = stPanelDI
			strArrN = split(straddinfo,",")
			s = Date
			Set odesc=Description.Create
			odesc("micclass").value="WebTable"
			
			Set l_link=  getParentObject().ChildObjects(odesc)
			print l_link.count
			For i = 0 to l_link.count - 1
			strName = l_link(i).GetROProperty("column names")
			print strName
			strArr = split(strName,";")
			If Not(strName = "") Then
			print strArr(0)
			If trim(strArr(0)) = "*Panel Name" Then
			l_link(i).GetROProperty("rows") 
			
			a = l_link(i).GetROProperty("rows")  
			b = l_link(i).GetROProperty("cols") 		  	       
			For x  = 1 to a     
			For j  =1 to b
			c = l_link(i).GetCellData(x,j)
			print c
			Select Case c
			Case "*Panel Name"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set strNameEvent
			
			Case "*Cycle"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set NewCycle
			Case "Program"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "Dissemination and Implementation"
			'Selecting PFA''''''
			Case "Primary Campaign Source"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "Dissemination and Implementation"
					
			Case "*MRO"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "Cary Scheiderer"
			Case "Online Review Deadline"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set strDeadLineDate
			Case "InPerson Review Deadline"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set  strformatted 
			Case "*Panel COI Due Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
            oEdit1.set  s + 25
'''''''''''''Selecting Panel Due Date''''''''''''''''''''

'            HtmlID = "00N39000003i9zt"
'            strData = date + 30
'            set_Web_Edit HtmlID,strData
            
            wait 2
            clk_singelCheckbox()
            wait 2
            selectWeblist6 "00N39000003LYhh", "Online Review – DI"           
			
			                                  
			End Select                               
			
			Next
			Next                             
			
			
			
			
			Exit for
			
			End If
			End If
			Next
			
			create_Panelwith_DI_onlinereviewrecordtype = strNameEvent
End Function

'Enter value on Web Edit TEXT AREA by index if many web edit Text Area present
Function setWebEditTextAreaIndex(strData, i)
On error resume next

wait 3

Set oShell = CreateObject("WScript.Shell") 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
'print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

'print pHwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebEdit"
editO("visible").value = true  
editO("html tag").value = "TEXTAREA" 
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 


edObj(i).highlight
edObj(i).set strData 
oShell.SendKeys strSearchText				            
Print " Entering the data " & "-" & strData & "-" & "is successful "

If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1," WebEdit Text Area Index value " & strData,"Failed to enter the data" & "-" & strData
else
LogReport 0,"WebEdit Text Area Index value " & strName,"Entering the data " & "-" & strName & "-" & "is successful"
End If
End Function

Function Create_cycle_assystemadmin() 
wait 5
clk_link_Object2 "Cycles"
wait 2
clk_Button_usingName "New"
wait 2
HtmlID = "Name"
strData = "Reg test Automation cycle" & RandomString(1)
set_Web_Edit HtmlID,strData
wait 2
HtmlID = "00N70000003DN7w"
strData = date + 30
set_Web_Edit HtmlID,strData
wait 2
clk_Button_usingName "Save"
wait 2
	
End Function
Function Create_Panel_assystemadmin() 
wait 2
	clk_link_Object2 "Panels"
wait 2
clk_Button_usingName "New"
stPanel = create_Panelwith_DI_onlinereviewrecordtype()
wait 2

''''''If want to save it in text file then uncomment the below line''' otherwise panel name will be reading from varible on top
'writeToAfilePanel stPanel
print stPanel
wait 2
fillOnlineReviewCriteria()
wait 2
clk_Button_usingName saveBtn
End Function

'Enter value on Web Edit box by index if many web edit box present
Function setWebEditBoxByIndex(strData, i)
On error resume next

wait 3

Set oShell = CreateObject("WScript.Shell") 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
'print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

'print pHwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebEdit"
editO("visible").value = true  
editO("html tag").value = "INPUT" 

'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 


edObj(i).highlight
edObj(i).set strData 
oShell.SendKeys strSearchText
Print " Entering the data " & "-" & strData & "-" & "is successful "				            

If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1," WebEdit box by Index value " & strData," Failed to enter the data" & "-" & strData
else
LogReport 0,"WebEdit box by Index value " & strData," Entering the data " & "-" & strData & "-" & "is successful"
End If
End Function

'Select web list by index 
Function selectWeblist_OMR(strData)
Err.Clear
On error resume next
Wait 3

	strArr = split(strData,",")
	
	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
	print strHwnd
	
	Set pa = Description.Create
	pa("micclass").value = "Page"
	
	print pHwnd
	Set editObject =  getParentObject()
	
	Set editO = Description.Create
	editO("micclass").value = "WebList"
	editO("visible").value = true  

Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count

for i = 0 To uBound(strArr)
items = edObj(i).GetRoProperty("all items")
print items
print strArr(i)
If Not(strArr(i) = "")Then	

edObj(i).select trim(strArr(i))   

End If

Next

		    If err <> 0 Then
		  		LogReport 4,"Error", err.number & "-" & err.description
		  		LogReport 1,"selectWeblist - " & strData,"Failed to select webList as " & "-" & strData
		  	else
		  		LogReport 0,"selectWeblist - " & strData,"Select webList as " & "-" & strData & "-" & "was done succesfully"
		  	End If


End Function
Function attachFileFromFileSystem(i,strFilePath)
                              Set oShell = CreateObject("WScript.Shell")
                  Set btncalc = Description.Create()  
                                                                    btncalc("micclass").value = "Browser"
                                                                  
                                                                    Set btn =DeskTop.ChildObjects(btncalc)
                                                                    strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
                                                                    print strHwnd
                                                                  
                                                                    Set pa = Description.Create
                                                                    pa("micclass").value = "Page"
                                                                  
                                                                   
                                                                    
                                                                    print pHwnd
                                                                    Set editObject =  getParentObject()
                                                                    
                                                                    Set editO = Description.Create
                                                                    
                                                                    editO("micclass").value = "WebFile"
                                                                    editO("visible").value = true  
                                                                    editO("html tag").value = "INPUT" 
                                                                       'editO("type").value = "text" 
                                                                     Set edObj =  getParentObject().ChildObjects(editO) 
                                                                     
                                                                      edObj(i).set strFilePath
                     'strSearchText = strFilePath
                     'oShell.SendKeys strSearchText 
                                                                     'set 
                                                                     wait 2
                                                                     'Browser("hwnd:=" & strHwnd).Dialog("text:=Choose File to Upload").WinButton("text:=&Open").click

End Function
	
	Public Function click_webElement_forengagementreport_selecting_byindex(strName,i)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
'l("class").value = "actionColumn"
'l("outerhtml").value = "<span class=""listTitle"">Reviews<span class=""count"">\[5\]</span></span>"
l("visible").value = True
l("html tag").value = "LABEL"
l("innertext").Value = strName
'l("innerhtml").value = Html
Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click
End Function
        
Function populateonlinereview_asMR(i)
    	On error resume next
    	 Set oShell = CreateObject("WScript.Shell")
 Set odesc=Description.Create
           odesc("micclass").value="WebElement"
           odesc("visible").value= True
           odesc("outerhtml").value= "<p><br></p>"
           odesc("html tag").value= "P"
           odesc("innerhtml").value= "<br>"
           Set O =  getParentObject().ChildObjects(odesc)
           print O.count
           wait 2
           O(i).click
           wait 2
           
           O(i).object.innertext = "test MR comments"
      'oShell.SendKeys "test MR comments"
    End Function
    Function dblClick_OpenInperson_Internal_listview(HtmlId)
	On error resume next
	Set odesc=Description.Create
	odesc("micclass").value="WebElement"
	odesc("visible").value= True
	odesc("html tag").value= "DIV"
	odesc("html id").value= HtmlId
	
	Set O =  getParentObject().ChildObjects(odesc)
	print O.count
	x = O(0).getRoProperty("abs_x")
	print x 
	y= O(0).getRoProperty("abs_y")
	print y  
	
	Set od=Description.Create
	od("micclass").value="WebElement"
	'od("class").value =  "x-grid3-cell-inner x-grid3-col-00N39000003LYer"
	od("html tag").value="DIV"
	od("html id").value = HtmlID
	od("abs_x").value= x
	Set Ob =  getParentObject().ChildObjects(od)
	print Ob.count
	
	
	Ob(0).FireEvent "ondblclick"
End Function
Function clk_singelCheckboxby_htmlId(HtmlId)
On error resume next
wait 3
Set ck = Description.Create

ck("micclass").value = "WebCheckBox"
ck("visible").value = true 
ck("html id").value = HtmlId
ck("html tag").value = "INPUT"


Set co =   getParentObject().ChildObjects(ck)
print co.count
If co.count > 0 Then
LogReport 0,"verifyCheckBox ", "The Check Box exists and is clikable"
else
LogReport 1,"verifyCheckBox ", "The Check Box doesnot exist and is not clickable"
End If
co(0).click
End Function

Public Function verifyIfweblistvalueDoesExistbyhtmlid(HtmlID,stvalues)
             Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
                                
On error resume next

            wait 4
			   			   
			   Set btncalc = Description.Create()  
			    btncalc("micclass").value = "Browser"
			  
			    Set btn =DeskTop.ChildObjects(btncalc)
			    strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
			    print strHwnd
			  
			    Set pa = Description.Create
			    pa("micclass").value = "Page"
			  		   
			    
			   
			    Set editObject =  getParentObject()
			    
			    Set editO = Description.Create
			    
			    editO("micclass").value = "WebList"
			    editO("visible").value = true 			    
			    editO("html id").value = HtmlID
			    editO("html tag").value = "SELECT"
			     Set edObj =  getParentObject().ChildObjects(editO) 
			     print edObj.count
			     for i = 0 To uBound(strArray)
			        items =   edObj(i).GetRoProperty("all items")
			        print items
			        
			        Exit For
			        Next
			         
			        If items = stvalues Then			        
			                     
			        Print " Weblist values verified successfully" 
			         LogReport 0,"Weblist Values verification" & HtmlID," Weblist " & "-" & stvalues & "-" & " verified Successfully"  
			        Else  
			        print " Weblist values verification failed"
 LogReport 1,"Weblist values verification" & HtmlID," Weblist " & "-" & stvalues & "-" & " Not verified"   
                                                                
                                       
                 
                                End  If
                                                
               


 
                                     
                                                
'                If err <> 0 Then
'LogReport 4,"Error ", err.number & "-" & err.description
'     
'                
'                End If
                

Set editO = Nothing

			         
			  
			           
     End Function
     
  Public Function verifyIfWebElementDoesExist2(strName)
On error resume next 

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
'l("html tag").value = "LI"
l("innertext").value = strName
l("visible").value = True 

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

If lo.count = 0 Then      
LogReport 1,"verifyIfWebElementDoesExist2 - " & strName, "The WebElement --" & strName & "- Doesn't Exist - Not Expected"
else
LogReport 0,"verifyIfWebElementDoesExist2 - " & strName, "The WebElement --" & strName & "- Exists - Expected"

End If

End Function

Function CreatenewpatientroleMRuser()
	On error resume next
	'wait 2
NavigateToReviewerPortalnewuser()
wait 2
fillUserNameInformation()
wait 3
clk_singelCheckbox()
wait 2
'Last step Join PCORI Portal
wait 2
clk_link_Object2 "Join PCORI Portal"
wait 2
selectallRadioButtonfornewuser()
'wait 2
fillMailingAddressfornewuser() 
'Wait 5
clk_singelCheckboxNewUser()
'wait 2
clk_Button_usingName "Submit"
wait 2
clk_link_Object2 "My Profile"
wait 2
captureWebElementText_MRuseremail()
MRemailfromprofile = captureWebElementText_MRuseremail()
writeToAfileMRemail MRemailfromprofile
wait 2
clk_link_Object2 "Home"	
	
End Function

Public Function click_webElement__selecting_radiobuttonany(strName)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
l("visible").value = True
l("html tag").value = "LABEL"
l("innertext").Value = strName

Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click
End Function
Function writeToAfileEmployerName(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\employername.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function captureWebElementText_RAPI_Name()
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("html tag").value = "DIV"
l("html id").value = "con2_ileinner"

Set lo =  getParentObject().ChildObjects(l)
Print lo.count

nameNewUser = lo(0).getRoProperty("innertext")
Print nameNewUser
        
captureWebElementText_RAPI_Name = nameNewUser


End Function

Function writeToAfileRAPIname(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\rapiname.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function readFromFileRAPIname()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\rapiname.txt",1)
          
          sContent = objTxtFile.Readline
         readFromFileRAPIname = sContent
End Function

Public Function readFromEmloyerName()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\employername.txt",1)
          
          sContent = objTxtFile.Readline
          readFromEmloyerName = sContent
End Function
Public Function captureWebElementText_newuser_email()
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("Visible").value = True 
l("html tag").value = "DIV"
l("html id").value = "con15_ileinner"
Set lo =  getParentObject().ChildObjects(l)
Print lo.count

User_Email = lo(0).getRoProperty("innertext")
Print User_Email
'Write captured MRA-Number to Function itself in order to call it later in a script
captureWebElementText_newuser_email = User_Email
End Function

Function writeToAfilenewuseremail(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\useremail.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function readFromFileuseremail()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile("C:\QTP\useremail.txt",1)
          
          sContent = objTxtFile.Readline
          readFromFileuseremail = sContent
End Function

Public Function captureWebElementText_RAPIuser_email()
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("Visible").value = True 
l("html tag").value = "DIV"
l("html id").value = "con15_ileinner"
Set lo =  getParentObject().ChildObjects(l)
Print lo.count

User_Email = lo(0).getRoProperty("innertext")
Print User_Email
'Write captured MRA-Number to Function itself in order to call it later in a script
captureWebElementText_RAPIuser_email() = User_Email
End Function

Function writeToAfileRAPIuseremail(strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile("C:\QTP\useremail.txt",2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function verifyIf_Image_LinkDoesExist_Internal(stTitle)
On error resume next 

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"
Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "Image"
l("alt").value = "All Tabs"
l("image type").value = "Image Link"
l("class").value = "allTabsArrow"
l("title").value = stTitle
Set lo =  getParentObject().ChildObjects(l)
print lo.count
lo(0).highlight


If err <> 0 Then
    LogReport 4,"Error", err.number & "-" & err.description
    LogReport 1,"verityLinkDoesnotExist - ", "The Link --" & stTitle & "- does not exist _not Expected"
else
      LogReport 0,"verityLinkDoesExist - ", "The Link --" & stTitle & "-does Exists -  Expected"
End If
End Function 

Public Function verifyIfLinkDoes_notExist1( HtmlTag,strName)
     	On error resume next 
     	'printstrHwnd
     	Set l = Description.Create
		 l("micclass").value = "Link"
		 l("html tag").value = HtmlTag
		 l("innertext").value = strName
		 Set lo =  getParentObject().ChildObjects(l)
		  If lo.count = 0 Then   	
		        LogReport 0,"verifyIfLinkDoesNotExist1 - ", "The Link --" & strName & "- doesnot exist"
		        else
				LogReport 1,"verifyIfLinkDoesNotExist1 - ", "The Link --" & strName & "-Exists "
		  End If
		
     End Function
     
 Public Function Creating_Approve_PatientReviewer1()
     	''''''''Creating new user'''''''''

NavigateToReviewerPortalnewuser()
wait 5
fillUserNameInformation()
wait 3
clk_singelCheckbox()
wait 2
''''Last step Join PCORI Portal
wait 2
clk_link_Object2 "Join PCORI Online"
wait 10
selectallRadioButtonfornewuser()
wait 2
fillMailingAddressfornewuser() 
Wait 5
clk_singelCheckboxNewUser()
wait 2
clk_Button_usingName "Submit"

''''''''Applying to become a Patient reviewer''''''

wait 10
clk_Button_usingName "START MY APPLICATION"

HtmlID = "j_id0:j_id2:j_id4:j_id6:j_id50:j_id51"
strData = "Patient"
selectWeblist6 HtmlID,strData
wait 5
clk_Button_usingName "Next"

wait 5
HtmlID = "j_id0:other:j_id7:j_id34:j_id36"
strData = "Bachelor’s degree (BA, BS, AB, etc.)"
selectWeblist6 HtmlID,strData
wait 5
setWebEditBox "BS"
wait 5
HtmlID = "j_id0:other:j_id7:j_id34:j_id46_unselected"
strData = "Caregiver"
selectWeblist6 HtmlID,strData
wait 5
clk_rightarrow_Link_Portal "j_id0:other:j_id7:j_id34:j_id46_right_arrow"
clk_Button_usingName "Next"

wait 5
HtmlID = "j_id0:mraform:j_id46:Dis"
strData = "Test Scientific Reviewer qs 1"
set_Web_Edit HtmlID,strData
wait 5
clk_multipleCheckboxformrapplicationpage3
wait 5
HtmlID = "j_id0:mraform:j_id121:popran"
strData = "Test Scientific Reviewer qs 2"
set_Web_Edit HtmlID,strData
wait 5
HtmlID = "j_id0:mraform:j_id138:hecare"
strData = "Test Scientific Reviewer qs 3"
set_Web_Edit HtmlID,strData
wait 5
HtmlID = "j_id0:mraform:j_id155:meth"
strData = "Test Scientific Reviewer qs 4"
set_Web_Edit HtmlID,strData
wait 5
strName = "j_id0:mraform:j_id170:j_id174"
strData = "Test Scientific Reviewer qs 5"
set_Web_Editbyname strName,strData
wait 5
selectRadioButton "No preference"
wait 5
selectRadioButton1 "Yes"
wait 5
selectRadioButton2 "Yes"
wait 5
clk_Button_usingName "Submit Application"
wait 5
clk_link_Object2 "Home"
wait 5
clk_link_Object2 "Submit a Merit Reviewer Application"
wait 2
MRAappnumber = captureWebElementText_MRANumber()
wait 5
writeToAfileMRAApplication MRAappnumber
wait 5
clk_link_Object2 "My Profile"
wait 5

MR_Reviewer2 = captureWebElementText_MR_email()
wait 5
writeToAfileMeritReviewer2 MR_Reviewer2
wait 5
clk_link_Object2 "Logout"

'''''''''Internal User ( MRO Admin approving the MRA Application)'''''''''

navigateAndLoginToSalesForce UserAdmin,SystemAdminPassword
wait 10
navigateToAccoutsPageThrughTopSearchBox "Carolyn Mohan "
wait 4
clk_link_Object2 "Carolyn Mohan"
wait 4
clk_link_Objectforimpersonate()
wait 2
HtmlID = "USER_DETAIL"
strName = "User Detail"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 4
clk_Button_usingName "Login"
wait 2
strData= readFromFileMRApplication()
navigateToAccoutsPageThrughTopSearchBox strData
wait 4
strData= readFromFileMRApplication()
clk_link_Object2 strData
wait 4
clk_Button_usingName "Submit for Approval"
wait 2
click_WinButton "Message from webpage", "OK"
wait 2
clk_link_Object2 "Approve / Reject"
wait 2
setWebEditBox ", Approval Test for MR"
wait 2
clk_Button_usingName "Approve"
wait 2
HtmlID = "globalHeaderNameMink"
strName = "Carolyn Mohan"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 2
clk_onlogoutwhendone_impersonating "Logout"
		
     End Function
     
     Public Function Creating_Approve_StakeholderReviewer1()
     	''''''''Creating new user'''''''''

NavigateToReviewerPortalnewuser()
wait 5
fillUserNameInformation()
wait 3
clk_singelCheckbox()
wait 2
''''Last step Join PCORI Portal'''''''''''''
wait 2
clk_link_Object2 "Join PCORI Online"
wait 10
selectallRadioButtonfornewuser()
wait 2
fillMailingAddressfornewuser() 
Wait 5
clk_singelCheckboxNewUser()
wait 2
clk_Button_usingName "Submit"

''''''''Applying to become a Stakeholder reviewer''''''

wait 10
clk_Button_usingName "START MY APPLICATION"
wait 5


HtmlID = "j_id0:j_id2:j_id4:j_id6:j_id50:j_id51"
strData = "Stakeholder"
selectWeblist6 HtmlID,strData
wait 5
clk_Button_usingName "Next"

wait 5
HtmlID = "j_id0:other:j_id7:j_id34:j_id36"
strData = "Bachelor’s degree (BA, BS, AB, etc.)"
selectWeblist6 HtmlID,strData
wait 5
setWebEditBox "BS"
wait 5
HtmlID = "j_id0:other:j_id7:j_id34:j_id46_unselected"
strData = "Purchaser"
selectWeblist5 HtmlID,strData
wait 5
clk_rightarrow_Link_Portal "j_id0:other:j_id7:j_id34:j_id46_right_arrow"
clk_Button_usingName "Next"
''''''''''''''''''''''''''''''user landed on Page 3''''''''''''''''
wait 5
HtmlID = "j_id0:mraform:j_id46:Dis"
strData = "Test Scientific Reviewer qs 1"
set_Web_Edit HtmlID,strData
wait 5
clk_multipleCheckboxformrapplicationpage3
wait 5
HtmlID = "j_id0:mraform:j_id121:popran"
strData = "Test Scientific Reviewer qs 2"
set_Web_Edit HtmlID,strData
wait 5
HtmlID = "j_id0:mraform:j_id138:hecare"
strData = "Test Scientific Reviewer qs 3"
set_Web_Edit HtmlID,strData
wait 5
HtmlID = "j_id0:mraform:j_id155:meth"
strData = "Test Scientific Reviewer qs 4"
set_Web_Edit HtmlID,strData
wait 5
strName = "j_id0:mraform:j_id170:j_id174"
strData = "Test Scientific Reviewer qs 5"
set_Web_Editbyname strName,strData
wait 5
selectRadioButton "No preference"
wait 5
selectRadioButton1 "Yes"
wait 5
selectRadioButton2 "Yes"
wait 5
wait 5
HtmlID = "j_id0:mraform:j_id207:block1:j_id247:file"
strData = "Browse..."
clk_webfileButton_usingName HtmlID,strData
wait 5
attachFileFromFileSystem 0,uploadresumepath
wait 5
HtmlID = "j_id0:mraform:j_id207:block1:j_id245:fileName"
strData = "MR_test_User Resume"
set_Web_Edit HtmlID,strData
wait 5
clk_Button_usingName "Upload"

clk_Button_usingName "Submit Application"
wait 5
verifyIfWebElementDoesExist1 "Thank you for applying to become a PCORI Merit Reviewer."
wait 5
clk_link_Object2 "Home"
wait 5
wait 5
clk_link_Object2 "Submit a Merit Reviewer Application"
wait 2
MRAappnumber = captureWebElementText_MRANumber()
wait 5
writeToAfileMRAApplication MRAappnumber
wait 5
clk_link_Object2 "My Profile"
wait 5

MR_Reviewer3 = captureWebElementText_MR_email()
wait 5
writeToAfileMeritReviewer3 MR_Reviewer3
wait 5
clk_link_Object2 "Logout"

navigateAndLoginToSalesForce UserAdmin,SystemAdminPassword
wait 10
navigateToAccoutsPageThrughTopSearchBox "Carolyn Mohan "
wait 4
clk_link_Object2 "Carolyn Mohan"
wait 4
clk_link_Objectforimpersonate()
wait 2
HtmlID = "USER_DETAIL"
strName = "User Detail"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 4
clk_Button_usingName "Login"
wait 2
strData= readFromFileMRApplication()
navigateToAccoutsPageThrughTopSearchBox strData
wait 4
strData= readFromFileMRApplication()
clk_link_Object2 strData
wait 4
clk_Button_usingName "Submit for Approval"
wait 2
click_WinButton "Message from webpage", "OK"
wait 2
clk_link_Object2 "Approve / Reject"
wait 2
setWebEditBox ", Approval Test for MR"
wait 2
clk_Button_usingName "Approve"
wait 2
HtmlID = "globalHeaderNameMink"
strName = "Carolyn Mohan"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 2
clk_onlogoutwhendone_impersonating "Logout"
		
     End Function
     
       Public Function Creating_Approve_4thReviewer()
       ''''''''Creating 4th Reviewer'''''

''''''''Creating new user'''''''''

NavigateToReviewerPortalnewuser()
wait 5
fillUserNameInformation()
wait 3
clk_singelCheckbox()
wait 2
''''Last step Join PCORI Portal'''''''''''''
wait 2
clk_link_Object2 "Join PCORI Online"
wait 10
selectallRadioButtonfornewuser()
wait 2
fillMailingAddressfornewuser() 
Wait 5
clk_singelCheckboxNewUser()
wait 2
clk_Button_usingName "Submit"

''''''''Applying to become a Stakeholder reviewer''''''

wait 10
clk_Button_usingName "START MY APPLICATION"
wait 5


HtmlID = "j_id0:j_id2:j_id4:j_id6:j_id50:j_id51"
strData = "Stakeholder"
selectWeblist6 HtmlID,strData
wait 5
clk_Button_usingName "Next"

wait 5
HtmlID = "j_id0:other:j_id7:j_id34:j_id36"
strData = "Bachelor’s degree (BA, BS, AB, etc.)"
selectWeblist6 HtmlID,strData
wait 5
setWebEditBox "BS"
wait 5
HtmlID = "j_id0:other:j_id7:j_id34:j_id46_unselected"
strData = "Purchaser"
selectWeblist5 HtmlID,strData
wait 5
clk_rightarrow_Link_Portal "j_id0:other:j_id7:j_id34:j_id46_right_arrow"
clk_Button_usingName "Next"
''''''''''''''''''''''''''''''user landed on Page 3''''''''''''''''
wait 5
HtmlID = "j_id0:mraform:j_id46:Dis"
strData = "Test Scientific Reviewer qs 1"
set_Web_Edit HtmlID,strData
wait 5
clk_multipleCheckboxformrapplicationpage3
wait 5
HtmlID = "j_id0:mraform:j_id121:popran"
strData = "Test Scientific Reviewer qs 2"
set_Web_Edit HtmlID,strData
wait 5
HtmlID = "j_id0:mraform:j_id138:hecare"
strData = "Test Scientific Reviewer qs 3"
set_Web_Edit HtmlID,strData
wait 5
HtmlID = "j_id0:mraform:j_id155:meth"
strData = "Test Scientific Reviewer qs 4"
set_Web_Edit HtmlID,strData
wait 5
strName = "j_id0:mraform:j_id170:j_id174"
strData = "Test Scientific Reviewer qs 5"
set_Web_Editbyname strName,strData
wait 5
selectRadioButton "No preference"
wait 5
selectRadioButton1 "Yes"
wait 5
selectRadioButton2 "Yes"
wait 5
wait 5
HtmlID = "j_id0:mraform:j_id207:block1:j_id247:file"
strData = "Browse..."
clk_webfileButton_usingName HtmlID,strData
wait 5
attachFileFromFileSystem 0,uploadresumepath
wait 5
HtmlID = "j_id0:mraform:j_id207:block1:j_id245:fileName"
strData = "MR_test_User Resume"
set_Web_Edit HtmlID,strData
wait 5
clk_Button_usingName "Upload"

clk_Button_usingName "Submit Application"
wait 5
verifyIfWebElementDoesExist1 "Thank you for applying to become a PCORI Merit Reviewer."
wait 5
clk_link_Object2 "Home"
wait 5
wait 5
clk_link_Object2 "Submit a Merit Reviewer Application"
wait 2
MRAappnumber = captureWebElementText_MRANumber()
wait 5
writeToAfileMRAApplication MRAappnumber
wait 5
clk_link_Object2 "My Profile"
wait 5

MR_Reviewer4 = captureWebElementText_MR_email()
wait 5
writeToAfileMeritReviewer4 MR_Reviewer4
wait 5
clk_link_Object2 "Logout"

navigateAndLoginToSalesForce UserAdmin,SystemAdminPassword
wait 10
navigateToAccoutsPageThrughTopSearchBox "Carolyn Mohan "
wait 4
clk_link_Object2 "Carolyn Mohan"
wait 4
clk_link_Objectforimpersonate()
wait 2
HtmlID = "USER_DETAIL"
strName = "User Detail"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 4
clk_Button_usingName "Login"
wait 2
strData= readFromFileMRApplication()
navigateToAccoutsPageThrughTopSearchBox strData
wait 4
strData= readFromFileMRApplication()
clk_link_Object2 strData
wait 4
clk_Button_usingName "Submit for Approval"
wait 2
click_WinButton "Message from webpage", "OK"
wait 2
clk_link_Object2 "Approve / Reject"
wait 2
setWebEditBox ", Approval Test for MR"
wait 2
clk_Button_usingName "Approve"
wait 2
HtmlID = "globalHeaderNameMink"
strName = "Carolyn Mohan"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 2
clk_onlogoutwhendone_impersonating "Logout"
       
     End Function
     
     Public Function Creating_Approve_5thReviewer()
     ''''''''Creating new user'''''''''

NavigateToReviewerPortalnewuser()
wait 5
fillUserNameInformation()
wait 3
clk_singelCheckbox()
wait 2
''''Last step Join PCORI Portal'''''''''''''
wait 2
clk_link_Object2 "Join PCORI Online"
wait 10
selectallRadioButtonfornewuser()
wait 2
fillMailingAddressfornewuser() 
Wait 5
clk_singelCheckboxNewUser()
wait 2
clk_Button_usingName "Submit"

''''''''Applying to become a Stakeholder reviewer''''''

wait 10
clk_Button_usingName "START MY APPLICATION"
wait 5

HtmlID = "j_id0:j_id2:j_id4:j_id6:j_id50:j_id51"
strData = "Stakeholder"
selectWeblist6 HtmlID,strData
wait 5
clk_Button_usingName "Next"

wait 5
HtmlID = "j_id0:other:j_id7:j_id34:j_id36"
strData = "Bachelor’s degree (BA, BS, AB, etc.)"
selectWeblist6 HtmlID,strData
wait 5
setWebEditBox "BS"
wait 5
HtmlID = "j_id0:other:j_id7:j_id34:j_id46_unselected"
strData = "Purchaser"
selectWeblist5 HtmlID,strData
wait 5
clk_rightarrow_Link_Portal "j_id0:other:j_id7:j_id34:j_id46_right_arrow"
clk_Button_usingName "Next"
''''''''''''''''''''''''''''''user landed on Page 3''''''''''''''''
wait 5
HtmlID = "j_id0:mraform:j_id46:Dis"
strData = "Test Scientific Reviewer qs 1"
set_Web_Edit HtmlID,strData
wait 5
clk_multipleCheckboxformrapplicationpage3
wait 5
HtmlID = "j_id0:mraform:j_id121:popran"
strData = "Test Scientific Reviewer qs 2"
set_Web_Edit HtmlID,strData
wait 5
HtmlID = "j_id0:mraform:j_id138:hecare"
strData = "Test Scientific Reviewer qs 3"
set_Web_Edit HtmlID,strData
wait 5
HtmlID = "j_id0:mraform:j_id155:meth"
strData = "Test Scientific Reviewer qs 4"
set_Web_Edit HtmlID,strData
wait 5
strName = "j_id0:mraform:j_id170:j_id174"
strData = "Test Scientific Reviewer qs 5"
set_Web_Editbyname strName,strData
wait 5
selectRadioButton "No preference"
wait 5
selectRadioButton1 "Yes"

selectRadioButton2 "Yes"

wait 5
HtmlID = "j_id0:mraform:j_id207:block1:j_id247:file"
strData = "Browse..."
clk_webfileButton_usingName HtmlID,strData
wait 5
attachFileFromFileSystem 0,uploadresumepath
wait 5
HtmlID = "j_id0:mraform:j_id207:block1:j_id245:fileName"
strData = "MR_test_User Resume"
set_Web_Edit HtmlID,strData
wait 5
clk_Button_usingName "Upload"

clk_Button_usingName "Submit Application"
wait 5
verifyIfWebElementDoesExist1 "Thank you for applying to become a PCORI Merit Reviewer."
wait 5
clk_link_Object2 "Home"
wait 5
wait 5
clk_link_Object2 "Submit a Merit Reviewer Application"
wait 2
MRAappnumber = captureWebElementText_MRANumber()
wait 5
writeToAfileMRAApplication MRAappnumber
wait 5
clk_link_Object2 "My Profile"
wait 5

MR_Reviewer5 = captureWebElementText_MR_email()
wait 5
writeToAfileMeritReviewer5 MR_Reviewer5
wait 5
clk_link_Object2 "Logout"

navigateAndLoginToSalesForce UserAdmin,SystemAdminPassword
wait 10
navigateToAccoutsPageThrughTopSearchBox "Carolyn Mohan "
wait 4
clk_link_Object2 "Carolyn Mohan"
wait 4
clk_link_Objectforimpersonate()
wait 2
HtmlID = "USER_DETAIL"
strName = "User Detail"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 4
clk_Button_usingName "Login"
wait 2
strData= readFromFileMRApplication()
navigateToAccoutsPageThrughTopSearchBox strData
wait 4
strData= readFromFileMRApplication()
clk_link_Object2 strData
wait 4
clk_Button_usingName "Submit for Approval"
wait 2
click_WinButton "Message from webpage", "OK"
wait 2
clk_link_Object2 "Approve / Reject"
wait 2
setWebEditBox ", Approval Test for MR"
wait 2
clk_Button_usingName "Approve"
wait 2
HtmlID = "globalHeaderNameMink"
strName = "Carolyn Mohan"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 2
clk_onlogoutwhendone_impersonating "Logout"
     
     End Function
     
     Public Function verifyIfLinkDoesExist3(strName,StRef,i)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)

print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"




Set l = Description.Create
l("micclass").value = "Link"
l("html tag").value = "A"
l("innertext").value = strName
l("href").value = StRef

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(i).Highlight
If lo(i).count = 1 Then 

LogReport 0,"verityLinkDoesExist - ", "The Link --" & strName & "- does exist _Expected-passed"
else
LogReport 1,"verityLinkDoesExist - ", "The Link --" & strName & "-does not Exists - Not Expected-failed"

End If



End Function

Function impersonate_as(userName)
Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
navigateToAccoutsPageThrughTopSearchBox userName
wait 2
'''''clk_link_Object2 userName
Brwser.WebTable("column names:=Name;Following", "html tag:=TABLE", "class:=list").Link("html tag:=A", "innertext:="&userName, "index:=0").Click
wait 2
clk_link_Objectforimpersonate
wait 2
HtmlID = "USER_DETAIL"
strName = "User Detail"
clk_onlink_forimpersonateinternal HtmlID,strName
wait 2
clk_Button_usingName "Login"
	
End Function
     
     Function logOut_asUser(userName)
	HtmlID = "globalHeaderNameMink"
clk_onlink_forimpersonateinternal HtmlID,userName
wait 2
clk_onlogoutwhendone_impersonating "Logout"
wait 2
		
End Function

Function writeToAfile_Misc(stData)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile(Miscvariable,2,true)     'open in write mode
            f.Write (stData)
            f.Close
            Set f = nothing
End Function
'Read from text file - IHS Campaign Name
Public Function readFromFile_Misc()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile(Miscvariable,1)
          
          stData1 = objTxtFile.Readline
          readFromFile_Misc = stData1
          
End Function

''''''''''''''''''''''New Functions for Lightning UI''''''''''''''''''''''''''''

Public Function verifyIf_Image_DoesExist_External(stFileName)
On error resume next 

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"
Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "Image"
l("html tag").value = "IMG"
l("image type").value = "Plain Image"
'l("class").value = "allTabsArrow"
l("file name").value = stFileName
Set lo =  getParentObject().ChildObjects(l)
print lo.count
lo(0).highlight


If err <> 0 Then
    LogReport 4,"Error", err.number & "-" & err.description
    LogReport 1,"verityLinkDoesnotExist - ", "The Link --" & stTitle & "- does not exist _not Expected"
else
      LogReport 0,"verityLinkDoesExist - ", "The Link --" & stTitle & "-does Exists -  Expected"
End If
End Function 
Function VerifywebElementnotclickable(stInnerText,ststatus)
On error resume next
wait 3

Dim Brwser
Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
Set oDesc = Description.Create  
oDesc("micclass").value = "WebElement"
oDesc("html tag").value = "LIGHTNING-LAYOUT-ITEM"
oDesc("visible").value = "True"
oDesc("innertext").value = stInnerText
'oDesc("class").value = "checkImg"
'oDesc("index").value = index

                'Brwser.Image(oDesc).Highlight
                'buttonstatus = Brwser.GetRoProperty("disabled")
                
                If ststatus =  "disabled" Then
                
                Print "WebElement verified successfully"
                                LogReport 0," webelement status check of " & strCheck," webelement status check of  " & "-" & strCheck & "-" & " not editable- successfully verified"
                Else  
                                LogReport 1," webelement status check of " & strCheck," webelement status check of  " & "-" & strCheck & "-" & " Failed - WebElement clickable not expected "
                End If
                
                                If err <> 0 Then
                                                LogReport 4,"Error", err.number & "-" & err.description

                                End  IF
End Function
Function Verifywebbuttonclickable(stvalue,ststatus)
On error resume next
wait 3

Dim Brwser
Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
Set oDesc = Description.Create  
oDesc("micclass").value = "WebButton"
oDesc("html tag").value = "BUTTON"
oDesc("visible").value = "True"
oDesc("Value").value = stvalue
'oDesc("class").value = "checkImg"
'oDesc("index").value = index

                Brwser.Image(oDesc).Highlight
                buttonstatus = Brwser.GetRoProperty("disabled")
                
                If ststatus =  "enabled" Then
                
                Print "Button verified successfully"
                                LogReport 0," webbutton status check of " & strCheck," webbutton status check of  " & "-" & strCheck & "-" & " editable_ successfully verified_expected"
                Else  
                                LogReport 1," webbutton status check of " & strCheck," webbutton status check of  " & "-" & strCheck & "-" & " Failed - button not_editable not expected "
                End If
                
                                If err <> 0 Then
                                                LogReport 4,"Error", err.number & "-" & err.description

                                End  IF
End Function
Public Function click_webElementBasedonXpath(stPath)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
'l("class").value = "labelCol"
'l("outerhtml").value = "<span class=""listTitle"">Reviews<span class=""count"">\[5\]</span></span>"
l("visible").value = True
l("html tag").value = "LIGHTNING-PRIMITIVE-ICON"
l("xpath").Value = stPath
'l("innerhtml").value = Html
Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click
End Function

Public Function read_FromFile_anyFilePath (anyFilePath)
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile(anyFilePath,1)
          
          sContent = objTxtFile.Readline
         read_FromFile_anyFilePath = sContent
End Function

Function set_WebEditBox_by_Name(Stname,strData)
 On error resume next
wait 3
strArr = split(strData,",")

Set oShell = CreateObject("WScript.Shell") 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"
print pHwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebEdit"                                     '''''''"WebEdit"
editO("visible").value = True 
editO("name").value = Stname
''editO("html tag").value = "INPUT" 
''editO("disabled").value = "0" 
Set edObj =  getParentObject().ChildObjects(editO) 

for i = 0 To uBound(strArr)
'id =   edObj(i).GetRoProperty("html id")
If Not(strArr(i) = "")Then

edObj(i).set strArr(i)  
oShell.SendKeys strSearchText				            
End If

Next
If err <> 0 Then
err.clear   

End If
     End Function       

''Function to Click WebElement
Public Function click_webElement_top_email_personnel(Innertext, HtmlTag)
Err.Clear
On error resume next
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("visible").value = True

l("html tag").value = HtmlTag
l("innertext").value = Innertext

Set edObj =  getParentObject().ChildObjects(l)
edObj(i).click

			If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"click_webElement_top_email_personnel - " & Innertext,"Failed to click the webElement" & "-" & Innertext
			else
			  	LogReport 0,"click_webElement_top_email_personnel - " & Innertext,"The webElement" & "-" & Innertext & "-" & "is clicked successfully"
			End If
			
End Function

Function selectWeblist_byClass(stClass,sTselectionname)
On error resume next

wait 4
strArr = split(sTdata,";")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

print pHwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
editO("html tag").value = "DIV" 
editO("select type").value = "ComboBox Select" 
'editO("html id").value = HtmlID
editO("class").value = stClass 
editO("selection").value = sTselectionname 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
for i = 0 To uBound(strArr)
items =   edObj(i).GetRoProperty("all items")
print items
print strArr(i)
If Not(strArr(i) = "")Then	


edObj(i).select trim(strArr(i))        
End If

Next


End Function

Public Function clk_weblist_using_accname(stAccname)
    
     On error resume next
     wait 4
     Set btncalc = Description.Create()  
     btncalc("micclass").value = "Browser"
  
     Set btn =DeskTop.ChildObjects(btncalc)
     strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
     print strHwnd
  
     Set pa = Description.Create
     pa("micclass").value = "Page"
  
          
     print pHwnd
    
     Set editO = Description.Create
     
     editO("micclass").value = "WebList"
     editO("visible").value = true  
     editO("acc_name").value = stAccname
     Set editObject =  getParentObject().ChildObjects(editO) 
     editObject(0).highlight
     editObject(0).click
    
    If err <> 0 Then
  	  LogReport 4,"Error", err.number & "-" & err.description
  	  LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
  	  else
  	  LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
    End If
End Function

Public Function click_webEdit_box_By_placeholder(stHolder)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebEdit"
l("visible").value = True
l("html tag").value = "INPUT"
l("placeholder").value = stHolder
Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click
End Function

Public Function click_webElement_by_indexand_innertext (i,stText,stTag)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")

print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebElement"
'l("class").value = "fa fa-search"
l("html tag").value = stTag
l("visible").value = True
l("innertext").value = stText

Set edObj =  getParentObject().ChildObjects(l)
edObj.count
edObj(i).Highlight
edObj(i).click

  If err <> 0 Then
  	  LogReport 4,"Error", err.number & "-" & err.description
  	  LogReport 1,"cliking the Webelement" & stText,"Failed to click the Webelement" & "-" & stText
  	  else
  	  LogReport 0,"cliking the Webelement" & stText,"The Webelement" & "-" & stText & "-" & "is clicked succesfully"
    End If
End Function
'Function to Capture innertext of the Link based on Table Column and Index
''k -  starts with 2 Always, it COI & Expetise record Numbr line
''m - is Column in COI & ExpertiseTable based on Sort by:COI Due DateSorted: NoneShow COI Due Date column actions = 0,  "Minus" m = Column # we want

Public Function captureLinkText_COI_Number(k,m)
On error resume next 

Set odesc=Description.Create
odesc("micclass").value="WebTable"
odesc("column names").value =";Choose a RowSelect All;Sort by:COI & Expertise NumberSorted: NoneShow COI & Expertise Number column actions;Sort by:Research ApplicationSorted: NoneShow Research Application column actions;Sort by:Reviewer NameSorted AscendingShow Reviewer Name column actions;Sort by:Application Level COISorted: NoneShow Application Level COI column actions;Sort by:PFA Level COISorted: NoneShow PFA Level COI column actions;Sort by:Expertise LevelSorted: NoneShow Expertise Level column actions;Sort by:ActiveSorted: NoneShow Active column actions;Sort by:COI Due DateSorted: NoneShow COI Due Date column actions;"
odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count

		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
			print strArr(2)
        		If trim(strArr(2)) = "Sort by:COI & Expertise NumberSorted: NoneShow COI & Expertise Number column actions" Then
			l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows") 
              print a		
		b = l_link(i).GetROProperty("cols") 		
              print b		
		For x  = 1 to k     
				For j  = 1 to b
				c = l_link(i).GetCellData(x,b-m)
				print c
				Next
		Next
			End if
		End If
captureLinkText_COI_Number = c

End Function

Public Function click_webElement_COI_Number_in_Related_List(StText, HtmlTag)
Err.Clear
On error resume next
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebElement"
l("visible").value = True
l("html tag").value = HtmlTag
l("innertext").value = StText

Set edObj =  getParentObject().ChildObjects(l)
print edObj.count
edObj(0).Highlight

'			If err <> 0 Then
'				LogReport 4,"Error", err.number & "-" & err.description
'				LogReport 1,"click_webElement_COI_Number_in Related List - " & StText,"Failed to click the webElement" & "-" & StText
'			else
'			  	LogReport 0,"click_webElement_COI_Number_in Related List - " & StText,"The webElement" & "-" & StText & "-" & "is clicked successfully"
'			End If
	edObj(0).click		
End Function

Function click_Checkbox_by_Name(stName)
On error resume next
wait 3
Set ck = Description.Create

ck("micclass").value = "WebCheckBox"
ck("visible").value = true  
ck("name").value = stName

Set co =   getParentObject().ChildObjects(ck)
print co.count

				If co.count > 0 Then
					LogReport 0,"click_Checkbox_stName "&stName, "The Check Box exists"&stName&"- Expected"
				else
					LogReport 1,"click_Checkbox_stName "&stName, "The Check Box does NOT exist "&stName&"- NOT expected"
				End If

co(0).click

End Function
Function dblClick_webElement_byOutertext (stText)

	On error resume next
	
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

	Set odesc=Description.Create
	odesc("micclass").value="WebElement"
	odesc("visible").value= True
	odesc("html tag").value= "DIV"
	odesc("outertext").value= stText
	
	Set O =  getParentObject().ChildObjects(odesc)
	print O.count

	
	O(0).DoubleClick
	Wait 5
	
End Function
''Function to verify if Web Element with specified Outertext as parameter Exists
Public Function verify_IfWebElementDoesExist_byOutertext(outerText)
Err.Clear
On error resume next
Wait 3

	Set l = Description.Create
	l("micclass").value = "WebElement"
	l("outertext").value = outerText&".*"
	l("html tag").value = "DIV"
	l("visible").value = True 
	
	'strcount = l.count
	'print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

				If lo.count = 0 Then      
				LogReport 1,"verify_IfWebElementDoesExist_byOutertext - " & outerText, "WebElement -" & outerText & "- Doesn't Exist - Not Expected"
				else
				LogReport 0,"verify_IfWebElementDoesExist_byOutertext - " & outerText, "WebElement -" & outerText & "- Exists - Expected"
				End If

End Function
Public Function click_webEdit_box_By_placeholder_and_Index(stHolder,i)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

Set l = Description.Create
l("micclass").value = "WebEdit"
l("visible").value = True
l("html tag").value = "INPUT"
l("placeholder").value = stHolder
Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click
End Function
Public Function clk_SFLButton_usingName(strName)
    
     On error resume next
     wait 4
     Set btncalc = Description.Create()  
     btncalc("micclass").value = "Browser"
  
     Set btn =DeskTop.ChildObjects(btncalc)
     strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
     print strHwnd
  
     Set pa = Description.Create
     pa("micclass").value = "Page"
  
          
     print pHwnd
    
     Set editO = Description.Create
     
     editO("micclass").value = "SFLButton"
     editO("visible").value = True
     editO("type").value = "button"
     editO("html tag").value = "BUTTON"
     editO("name").value = strName 
     Set editObject =  getParentObject().ChildObjects(editO) 
     editObject(0).highlight
     editObject(0).click
    
    If err <> 0 Then
  	  LogReport 4,"Error", err.number & "-" & err.description
  	  LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
  	  else
  	  LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
    End If
End Function
Public Function verifyWebElement_By_Outertext(Stoutertext)
Err.Clear
On error resume next
Wait 3 

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
	print strHwnd
	
	Set pa = Description.Create
	pa("micclass").value = "Page"
	
	Set l = Description.Create
	l("micclass").value = "WebElement"
	l("outertext").value = Stoutertext
	l("visible").value = True 
	
Set lo = getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

				If lo.count = 0 Then      
					LogReport 1,"verifyWebElement_By_Outertext - " & Stoutertext, "WebElement --" & Stoutertext & "- Doesn't Exist - Not Expected"
				else
					LogReport 0,"verifyWebElement_By_Outertext - " & Stoutertext, "WebElement --" & Stoutertext & "- Exists - Expected"
				End If

End Function

Public Function verifyWebElement_By_OuterHTML(stTag,Stouterhtml)
Err.Clear
On error resume next
Wait 3 

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
	print strHwnd
	
	Set pa = Description.Create
	pa("micclass").value = "Page"
	
	Set l = Description.Create
	l("micclass").value = "WebElement"
	l("outerhtml").value = Stouterhtml
	l("visible").value = True 
	l("html tag").value = stTag
	
Set lo = getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

				If lo.count = 0 Then      
					LogReport 1,"verifyWebElement_By_OuterHtml - " & Stouterhtml, "WebElement --" & Stouterhtml & "- Doesn't Exist - Not Expected"
				else
					LogReport 0,"verifyWebElement_By_OuterHtml " & Stouterhtml, "WebElement --" & Stouterhtml & "- Exists - Expected"
				End If

End Function

Public Function clk_weblist_using_name(stname)
    
     On error resume next
     wait 4
     Set btncalc = Description.Create()  
     btncalc("micclass").value = "Browser"
  
     Set btn =DeskTop.ChildObjects(btncalc)
     strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
     print strHwnd
  
     Set pa = Description.Create
     pa("micclass").value = "Page"
  
          
     print pHwnd
    
     Set editO = Description.Create
     
     editO("micclass").value = "WebList"
     editO("visible").value = true  
     editO("name").value = stname
     Set editObject =  getParentObject().ChildObjects(editO) 
     editObject(0).highlight
     editObject(0).click
    
    If err <> 0 Then
  	  LogReport 4,"Error", err.number & "-" & err.description
  	  LogReport 1,"cliking the weblist" & stname,"Failed to click the weblist" & "-" & stname
  	  else
  	  LogReport 0,"cliking the weblist" & stname,"The weblist" & "-" & stname & "-" & "is clicked succesfully"
    End If
End Function

Public Function captureLinkText_Review_Number(k,m)
On error resume next 

Set odesc=Description.Create
odesc("micclass").value="WebTable"
odesc("column names").value =";Choose a RowSelect All;Sort by:Review IDSorted: NoneShow Review ID column actions;Sort by:ReviewerSorted: NoneShow Reviewer column actions;Sort by:Record TypeSorted: NoneShow Record Type column actions;Sort by:ApplicationSorted: NoneShow Application column actions;Sort by:Application ProgramSorted AscendingShow Application Program column actions;Sort by:Primary Reviewer RoleSorted: NoneShow Primary Reviewer Role column actions;Sort by:Reviewer TypeSorted: NoneShow Reviewer Type column actions;Sort by:StatusSorted: NoneShow Status column actions;Sort by:DeadlineSorted: NoneShow Deadline column actions;"
odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count

		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
			print strArr(2)
        		If trim(strArr(2)) = "Sort by:Review IDSorted: NoneShow Review ID column actions" Then
			l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows") 
              print a		
		b = l_link(i).GetROProperty("cols") 		
              print b		
		For x  = 1 to k     
				For j  = 1 to b
				c = l_link(i).GetCellData(x,b-m)
				print c
				Next
		Next
			End if
		End If
captureLinkText_Review_Number = c

End Function

Function setWebEditBoxByIndex(strData,stTag, i)
On error resume next

wait 3

Set oShell = CreateObject("WScript.Shell") 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
'print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

'print pHwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebEdit"
editO("visible").value = true  
editO("html tag").value = StTag
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 


edObj(i).highlight
edObj(i).set strData 
oShell.SendKeys strSearchText
Print " Entering the data " & "-" & strData & "-" & "is successful "				            

If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1," WebEdit box by Index value " & strData," Failed to enter the data" & "-" & strData
else
LogReport 0,"WebEdit box by Index value " & strData," Entering the data " & "-" & strData & "-" & "is successful"
End If
End Function

 Function setWebEditBoxany_by_acc_name(staccname ,strData)
     	 On error resume next
wait 3

	Set oShell = CreateObject("WScript.Shell") 
	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
	print strHwnd
	
	Set pa = Description.Create
	pa("micclass").value = "Page"
	
	print pHwnd
	Set editObject =  getParentObject()
	
	Set editO = Description.Create
	editO("micclass").value = "WebEdit"
	editO("acc_name").value = staccname
	editO("visible").value = true  
	editO("html tag").value = "DIV" 

	Set edObj =  getParentObject().ChildObjects(editO) 


edObj(i).highlight
edObj(i).set strData 
oShell.SendKeys strSearchText                                                              

				If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"setWebEditBox_ByIndex" & i,"Failed to enter the data -" & strData
				else
				LogReport 0,"setWebEditBox_ByIndex" & i,"Entering the data -" & strData & "- is successful"
				End If

     End Function

Function SelectWebListOnline_Review()  

On error resume next

wait 4
strArr = split(strData,";")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strhwnd

Set pa = Description.Create
pa("micclass").value = "Page"
print hwnd
Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("visible").value = true  
''editO("html tag").value = "SELECT" 
''editO("type").value = "ComboBox Select" 
Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
wait 3

edObj(i).Select "1"
edObj(i+1).Select "1"
edObj(i+2).Select "1"
edObj(i+3).Select "1"
edObj(i+4).Select "1"
edObj(i+5).Select "1"
edObj(i+6).Select "1"
edObj(i+7).Select "Yes"


Wait 3

End Function
