
'Public const runOn = "Methods" '''"Methods"   '''"Regular" ''' "Methods" '' ''

Public const url = "https://test.salesforce.com/"                                                           ''''''''''"https://login.salesforce.com/?startURL=/"                                            ''''''''''' "https://test.salesforce.com/"
Public const urlprod = "https://login.salesforce.com/?startURL=/"
Public const Reviewerurl = "https://pcori--pqt.sandbox.my.site.com/engagement"             '''''''''''''''"https://pcori.force.com/engagement"           ''''"https://pcori--pqt.sandbox.my.site.com/engagementt"  ''''''''''	https://pcori--uat.sandbox.my.site.com/engagement"""""

'''''''''''''''''''PI And AO ID and Password
Public const AOEmail = "pcori.reg+raaosmoketest@gmail.com"        '''''''''''''''''PQT Email = pcori.reg+raaosmoketest@gmail.com.pqt
Public const PIEmail = "pcori.reg+4@gmail.com"                        '''''''''''''''''PQT email = pcori.reg+4@gmail.com.pqt  ''' Prod email = pcori.reg+4@gmail.com
Public const passWord1 = "test123456"                            '''''''''''' | PQT PWD = test123456 | prod - test123456

''''''''''''Admin Id And Password PQT''''''''''''
Public const UserAdminPQT = "sahad@pcori.org.pqt"           		''AdminPassword
Public const AdminPassword = "washington2"

''''''''''''Admin Id And Password PROD''''''''''''
Public const UserAdmin2 = "sahad@pcori.org"           		''AdminPassword
Public const AdminPassword2 = "washington4"

'''''''''''''EA Reviewer ID and password''''User - EA Reviewer Smoke Test''''
Public const EA_Reviewer = "pcori.reg+eareviewer@gmail.com.pqt"     ''''PQT User ID = pcori.reg+eareviewer@gmail.com.pqt'''  | Prod User Id = pcori.reg+eareviewer@gmail.com
Public const EA_Reviewer_Pass = "test1234"


''----- Variables used for LOI and Application Automation process:
''----- LOI ---------------------------------------------------------------------------
''''''''acc_name of the weblist for LOI phase'''' which updates for each salesforce release'''''''' and will need to update before execution of the script''''

'''''''''''below are for Winter 25 sf release'''PQT 
Public Const statusEALOI_accName = "LOI Status"
Public Const LOIConversion_app_recordType_accName = "Record Type"
Public Const StatusApp_accName = "Status"

''''''' Below are for Summer 24 release''''''Prod
Public Const statusEALOI_accName_Prod = "LOI Status"
Public Const LOIConversion_app_recordType_accName_Prod = "Record Type"
Public Const StatusApp_accName_Prod = "Status"

''''''***********************************************************************************
Public const TestCycleR1 = "Test Cycle R1 - kw7"
'''Public const Current_Cycle = "Cycle 3 2018"
Public const AD_program = "Addressing Disparities"
Public const LOI_form = "18C3-LOI Form-Broads"
Public const APP_form = "18C3-App-Broads"
Public const AllTab = "All Tabs"
Public const CyclTab = "Cycles"
Public const InstitOrg = "Opt Out"

Public const testAppNumber = "TEST 123456 kkk"
Public const Program_Methods= "Improving Methods for Conducting PCOR"
Public const Convertbtn = "Convert"
Public const editBtn = "Edit"
Public const saveBtn = "Save"
Public Const campaignsTab = "Campaigns"
Public Const submitBtn = "Submit"
Public Const statusDraft = "Draft"
Public Const blueRoundRAPreawards = "Research Awards Pre Award"
Public Const listviewLOIstatusDraft = "RA LOI 01. Draft with External"
Public Const listviewLOIstatusDraftMethods = "RA LOI 01. Draft with External-Methods"
Public Const listviewLOIstatusSubmitted = "RA LOI 02. Submitted"
Public Const listviewLOIstatusPendingReview = "RA LOI 04. Pending Review"
Public Const listviewLOIstatusPendingReviewMethods = "RA LOI 04. Pending Review Methods"
Public Const tabLOIReviews = "LOI Reviews"
Public Const externalLOIstatusHtmlID = "00N39000003LYaZ_ileinner"
Public Const internalLOIstatusHtmlID = "lea13_ileinner"
Public Const statusLOIclosed = "Closed"
Public Const statusNewLOI = "New LOI"
Public Const statusUnderReview = "Under Review"
Public Const statusSubmittedLOI = "Submitted"
Public Const statusSubmittedtoAO = "Submitted to AO"
Public Const statusPendingLOIReview = "Pending LOI Review"

Public const cancelBtn = "Cancel"
Public const deleteBtn = "Delete"
Public const Nextbtn = "Next -->"
Public const newBtn = "New"
Public const continueBtn = "Continue"
Public const innertextChangeLOIReviewOwner = "\[Change\]"
Public const listviewMyPendingLOIReview = "Sabiha My Pending Reviews"
Public const listviewMyPendingLOIReview_uniqueN = "Sabiha_My_Pending_Reviews"

Public const listviewMyPendingLOIReview_uniqueN_methods = "Sabiha_My_Pending_Reviews_Methods"
Public const listviewMyPendingLOIReviewMethods = "Sabiha Methods Review"

Public const list_view_RA_App_Methods_Program_Review = "APP Methods my Cycle Under Review"
Public const list_view_RA_App_Methods_Program_Review_u = "APP_Methods_my_Cycle_Under_Review"

Public const list_view_RA_App_Program_Review = "APP my Cycle Under Review"
Public const list_view_RA_App_Program_Review_u = "APP_my_Cycle_Under_Review"

Public const htmlIDLOIreviewcheckBox = "00N70000003DUv3"
Public const newLOIreviewBtn = "New LOI Review"
Public const htmlIDemailProofField = "00N39000003LYbZ_ileinner"
Public const listviewDoNotInvite = "RA LOI 07. Do Not Invite"
Public const statusNotInvited = "Not Invited"
Public const changeStatusBtn = "Change Status"
Public const finalCommentsScreenedCheckBoxHtmlID = "00N39000003LYaf_chkbox"
Public const invitedToApplyinPreviewEmail = "Invited to Apply"
Public const htmlIDWebListPreviewEmailCommunic = "00N39000003LYbZ"
Public const listviewInvitedToApply = "RA LOI 06. Invited to Apply"
Public const htmlIDwebElementProjectNameonLOI = "00N70000003CCW8_ileinner"
Public const convertLOIBtn = "Convert to Selected"
Public const webclassSortLOIlistviewBYname = "x-grid3-hd-inner x-grid3-hd-FULL_NAME"
Public const htmlIDofEmailProofwhenEdit = "00N39000003LYaY"
Public const webElement_NewUserinfoPage = "Welcome to the PCORI Online\. Please provide some basic information about yourself before proceeding to the Online homepage\."
Public const LOIparametersMethods = "RecordTypeId:01239000000Hr4z"
Public const generatePDF_RAloiMethods = "Generate Pdf RA LOI Methods"
Public const generatePDF_RAloi = "Generate Pdf RA LOI"
Public const MethodsLOIreviewRecord = "Methods LOI Review"
Public const sharing_btn = "Sharing"
Public const project_Detail = "Project Detail"

''----- Application ------------------------------------------------------------------------
Public const tabProjects = "Projects"
Public const listviewRAappDraft = "RA App 1. Draft"
Public Const internalRAappstatusHtmlID = "opp11_ileinner"
Public Const externalRAappstatusHtmlID = "00N39000003LYdb_ileinner"
Public Const PIapprovalField = "00N39000003LqON_ileinner"
Public Const RAAOapprovalField = "00N39000003LqOA_ileinner"
Public const listviewPendingAOapproval = "RA App 2. Pending AO Approval"
Public Const statusPendingAOApproval = "Pending AO Approval"
Public Const statusPendingPIapplication = "Pending PI Application"
Public Const statusAdminWithdrawn = "Administratively Withdrawn"
Public const listviewAdminWithdraw = "RA App 5. Admin Withdraw"
Public const linkApplicationsPortal = "Applications"
Public const listviewRAappUnderReview = "RA App 3.Under Review Broad"
Public Const statusProgrammaticReview = "Programmatic Review"
Public const linkClearedforReviewPhase = "cleared\.for@review\.phase"
Public const webElementClearedForReview = "cleared for review phase"
Public const listviewRAappUnderReviewBroad = "RA App 3.Under Review Broad"
Public const listviewRAappUnderReviewMethods = "RA App 3.Under Review Methods"
Public const assignContractAdminRAappHtmlID = "00N39000003APwx"
Public const HtmlIDofContractAdminonRAapp = "lookup00539000006JIUB00N39000003APwx"
Public const statusApplicationClosed = "Application Closed"
Public const winBtnMessageFromWebpage = "Message from webpage"
Public const listviewRAAOWithdraw = "RA App 5.Applicant Withdraw"
Public const statusApplicationAwarded = "Application Awarded"
Public const PP_email_address = "zztestemail997@yopmail.com"

Public const valid_rule_msg_opport_SPS_4802 = "Error: Invalid Data. Review all error messages below to correct your data.You do not have necessary access to this fields."
Public const valid_rule_msg_lead_SPS_4802 = "Error: Invalid Data. Review all error messages below to correct your data.You do not have necessary permission to edit this fields."
										
''----- PORTAL variables --------------------------------------------------------------------
Public Const ResearchAwardsbtn = "Research Awards"
Public Const portalLoginbtn = "Log in"
Public Const updatedInfoNewUser = "PCORIteststreet"
Public Const joinPortal = "Join PCORI Online"
Public Const myProfilelink = "My Profile"
Public Const logoutBtn = "Logout"
Public Const verifyUseronPortal = "Advisory Panels"
Public Const linkFundingOpports = "Funding Opportunities"
Public Const linkMyLoisandApps = "My LOIs and Applications"
Public Const linkEADashboard = "Engagement Awards Dashboard"
Public Const tabLOIs = "LOIs"
Public Const tabApplications = "Applications"
Public Const tabOpenItems = "Open Items"
Public Const saveAndNextBtn = "Save & Next"
Public Const linkHome = "Home"
Public Const tabClosedItems = "Closed Items"
Public Const reviewSubmitBtn = "Review/Submit"
Public const resubmissionAppfield = "44488"
Public Const EngagementAwardsbtn = "Engagement Award Program"
Public const webElement_onPortal_LogInPage = "PCORI Online is now open for: Cycle 2 2019 Applications \(Broad; Pragmatic Clinical Studies; Pediatric Anxiety; Dissemination & Implementation\)"

''----- Attachments File Path For Application Generic''''''''''''''
Public Const attachmentResubmission = "C:\Users\eferdous\OneDrive - PCORI\Automation\Attachments for TEST\Resubmission.pdf"
Public Const attachmentAutomation = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\Automation-TEST.docx"
Public Const attachmentPDF = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PDF_File_TEST.pdf"
Public const PeopleandPlaces = "C:\Users\eferdous\OneDrive - PCORI\Automation\Attachments for TEST\PeopleandPlaces.pdf"
Public const BudgetJustification = "C:\Users\eferdous\OneDrive - PCORI\Automation\Attachments for TEST\BudgetJustification.pdf"
Public const attachmentLettersofSupport = "C:\Users\eferdous\OneDrive - PCORI\Automation\Attachments for TEST\LettersOfSupport.pdf"
Public const attachmentResearchPlan = "C:\Users\eferdous\OneDrive - PCORI\Automation\Attachments for TEST\ResearchPlan.pdf"
Public const MethodologyStandards = "C:\Users\eferdous\OneDrive - PCORI\Automation\Attachments for TEST\Methodology-Standards-Checklist.xlsx"
Public const MilestonesTemplate = "C:\Users\eferdous\OneDrive - PCORI\Automation\Attachments for TEST\Milestones-Template.xlsx"
Public const SubcontractorDetailedBudget = "C:\Users\eferdous\OneDrive - PCORI\Automation\Attachments for TEST\Subcontractor-Detailed-Budget.xlsx"
Public Const Prior_Summary_Statement_Totest = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\7697_PriorSummaryStatement.pdf"

''''''----- Project Personnel Tab (on LOI and Application)
Public const projectPersonnel_firstRecord = "TestPPFN,TestLN,Opt Out,Test Other role,Other degree test,573-569-8585,testemail465@yopmail.com"
Public const projectPersonnel_firstRecord2 = "Test PP FN,Test PP LM,Opt Out,Test Other role,573-569-8585,testemail465@yopmail.com"
Public const projectPersonnel_fillOutwebLists = "Patient,Patient Partner,Industry,BA"
Public const projectPersonnel_fillOutwebLists2 = "Stakeholder,Stakeholder Partner,Payer"
Public const htmlIDrightArrowOnPPNewRecord = "j_id0:j_id2:j_id3:mainForm:j_id835:j_id854:10:inputFieldId_unselected"
Public const editedPPNrecordRAapp = "NewNameFirst,NewNameLast,Org,Patient,333-444-5555,testemail78465@yopmail.com"
Public const newPPNrecordRAapp = "AATestFirstName,AATestLastName,Opt Out,Test Other role- Patient,333-444-8585,zztestemail997@yopmail.com"

''''''----- Project Personnel Tab EA (Application) DI''''''''
Public const projectPersonnel_RecordEAApp_DI = "TestPPFN,TestLN,Opt Out,Other degree test,573-569-8585,testemail465@yopmail.com,Testemail78465@yopmail.com"
Public const projectPersonnel_fillOutwebLists_EAApp_DI = "Yes,Patient,Clinician,Patient Partner,,BCH"
Public const projectPersonnel_Relevant_ExperienceEAApp_DI_allitems = "--None--;0–2 years;3–5 years;6–9 years;10–15 years;16 years \+"

''''''----- LOI specific ----------------
Public Const htmlIDofLoiStatusInOpenItemsTab = "j_id0:mainForm:j_id282:iInquirySection:j_id283:0:j_id301:6:j_id312"
Public Const htmlIDofLoiPFAInClosedItemsTab = "j_id0:mainForm:j_id317:iInquirySection:j_id318:0:j_id322:1:j_id333"
Public Const htmlIDofLoiStatusInClosedItemsTab = "j_id0:mainForm:j_id317:iInquirySection:j_id318:0:j_id322:6:j_id333"
Public Const htmlIDpaperPencilIconPortal = "slds-truncate sorting_1"
Public Const statusLOIWithdrawn = "Withdrawn"

Public const htmlID_foredit_LOI_preScreen_MRO_webEdit = "CF00N39000003LYb5"
Public const htmlID_foredit_LOI_preScreen_MRO_webEditcomments = "00N39000003LYb4"
Public const htmlID_foredit_LOI_preScreen_MRO_webList = "00N39000003LYb3"

Public const htmlID_forView_LOIPreScreen_MRO_webEdit = "CF00N39000003LYb5_ileinner"
Public const htmlID_forView_LOIPreScreen_MRO_webEditComments = "00N39000003LYb4_ileinner"
Public const htmlID_forView_LOIPreScreen_MRO_webList = "00N39000003LYb3_ileinner"

''----- Application specific ---------------
Public Const htmlIDofRAappNameInOpenItemsTab = "j_id0:mainForm:j_id160:iProposalSection:j_id161:0:j_id195:0:j_id206"
Public Const htmlIDofRAappStatusInOpenItemsTab = "j_id0:mainForm:j_id160:iProposalSection:j_id161:0:j_id195:3:j_id206"
Public Const htmlIDofRAappStatusInClosedItemsTab = "j_id0:mainForm:j_id240:iProposalSection:j_id241:0:j_id251:3:j_id262"
Public Const htmlIDofRAappProjectNameinClosedItemsTab = "j_id0:mainForm:j_id240:iProposalSection:j_id241:0:j_id251:0:j_id262"
Public Const htmlIDofRAappNumberinClosedItemsTab = "j_id0:mainForm:j_id240:iProposalSection:j_id241:0:j_id251:1:j_id262"


'''''''''''''''''''MISC'''''''
Public const strOportunities = "Opportunities"
Public const logInButton = "Log In to Sandbox"
Public const logInButton_Prod = "Log In"
Public const reportTab = "Reports"
Public const leadsTab = "Leads"
Public const contactTab = "Contacts"
Public const opportunitiesTab = "Opportunities"
Public const runReportNow = "Run Report Now"
Public const searchBtn = "Search"
Public const updateBtn = "Update Address"
Public const servicReqBtn = "Service Requests"
Public const contactsTab = "Accounts"
Public const mergeAccountBtn = "Merge Accounts" 
Public const newTaskBtn = "New Task"
Public const newEventbtn = "New Event"
Public const goBtn = "Go!"
Public const resetBtn = "Reset" 


''----- Gloabl USERS and passwords
Public const DMOS_user1 = "Charles Arnott"

Public const mrManagementUser = "cmohan@pcori.org.uat2"  		''passwordBenjamin
Public const passwordBenjamin = "carolyntest1234"
Public const MROname = "Carolyn Mohan"
Public const MR_management_user = "Ashley Pidal"

Public const cmaOperationsuser = "cdmin@gmail.com.uat2"    		''passwordCMA
Public const passwordCMA = "cmatest234"

Public const scienceoperationuser = "pcori.reg+gbhat@gmail.com"  		''passWord2
Public const geetaPassword = "geetatest234"

Public const scienceAdmin = "khughes@pcori.org.uat2"        		''passWord2
Public const katiePassword = "katietest234"

Public const scienceLimitedAccess = "sbashir@pcori.org.uat2"  		''passWord1
Public const limitedPO_password = "bashirtest1234"
Public const science_LimitedAccess_name = "Surair Bashir"

Public const EOFieldaccess = "acole@pcori.org.pqt"     ''''User - Alana Cole''''
Public const AlanaPassword = "alanatest123"


Public const IEBrowser = "IEXPLORE.EXE"
Public const ChromeBrowser = "chrome.exe"


'' ----- Text Files Pathes when Write to Text File and Read From Text File functions used:
'' AUTO created:
'' RA LOI:
Public const TextfilePathForRAPI_email = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\RAPI Email.txt"
Public const TextfilePathForRAPI_name = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\RAPI Name.txt"

Public const TextfilePathForRAAO_email = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\RAAO Email.txt"
Public const TextfilePathForRAAO_name = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\RAAO Name.txt"

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
Public const TextfilePathForCampaign_EADI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - EADI.txt"
Public const TextfilePathForCampaign_EACB = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - EACB.txt"
Public const TextfilePathForCampaign_EASCS = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - EASCS.txt"
'Public const TextfilePathForCampaign_HSII = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - HSII 2.txt"
Public const TextfilePathForCampaign_Mental_Health = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - Mental Health.txt"
Public const TextfilePathForCampaign_Managing_Pain = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - Managing Pain.txt"
Public const TextfilePathForCampaign_Violence_Trauma = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Campaign - Violence Trauma.txt"



''''''''''''''''''''''LOI Number File path'''''''''''''''''''
Public const TextfilePathForLOInumber_Broad = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - AD.txt"
Public const TextfilePathForLOInumber_LC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - LC.txt"
Public const TextfilePathForLOInumber_IMRI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - IMRI.txt"
Public const TextfilePathForLOInumber_SDM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - SDM.txt"
Public const TextfilePathForLOInumber_HSII = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - HSII.txt"
Public const TextfilePathForLOInumber_PLACER = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - PLACER.txt"
Public const TextfilePathForLOInumber_Methods = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - Methods.txt"
Public const TextfilePathForLOInumber_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - SOE.txt"
Public const TextfilePathForLOInumber_SleepHealth = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - SleepHealth.txt"
Public const TextfilePathForLOInumber_MMM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - MMM.txt"
Public const TextfilePathForLOInumber_IDD = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - IDD.txt"
Public const TextfilePathForLOInumber_EADI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - EADI.txt"
Public const TextfilePathForLOInumber_EACB = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - EACB.txt"
Public const TextfilePathForLOInumber_EASCS = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - EASCS.txt"
Public const TextfilePathForLOInumber_Mental_Health = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - Mental Health.txt"
Public const TextfilePathForLOInumber_Managing_Pain = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - Managing Pain.txt"
Public const TextfilePathForLOInumber_Violence_Trauma = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\LOI Number - Violence Trauma.txt"

''''''''''''''''''LOI review number file path''''''''''''''''
Public const TextfilePathForReviewer1_Number = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Reviewer1 Number.txt"
Public const TextfilePathForReviewer2_Number = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Reviewer2 Number.txt"
Public const TextfilePathForReviewer3_Number = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Reviewer3 Number.txt"
Public const TextfilePathForReviewer4_Number = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Reviewer4 Number.txt"
Public const TextfilePathForReviewer1_Number_M = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\M Reviewer1 Number.txt"
Public const TextfilePathForReviewer2_Number_M = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\M Reviewer2 Number.txt"
Public const TextfilePathForReviewer1_Number_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\SOE Reviewer1 Number.txt"
Public const TextfilePathForReviewer2_Number_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\SOE Reviewer2 Number.txt"


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
Public const TextfilePath_LOI_Project_name_EADI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - EADI.txt"
Public const TextfilePath_LOI_Project_name_EACB = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - EACB.txt"
Public const TextfilePath_LOI_Project_name_EASCS = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - EASCS.txt"
Public const TextfilePath_LOI_Project_name_HSII = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - HSII.txt"
Public const TextfilePath_LOI_Project_name_Mental_Health = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - Mental Health.txt"
Public const TextfilePath_LOI_Project_name_Managing_Pain = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - Managing Pain.txt"
Public const TextfilePath_LOI_Project_name_Violence_Trauma = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Name - Violence Trauma.txt"

''''''''''RA APP:''''''''''''
Public const PP_number_fromAPP_AD = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\TEXT Files\Autocreated\Project Personnel - Number - AD.txt"

'' RA LOI METHODS:
Public const TextfilePathForRAPI_email_M = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\Autocreated\M RAPI Email.txt"
Public const TextfilePathForRAPI_name_M = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\Autocreated\M RAPI Name.txt"

Public const TextfilePathForRAAO_email_M = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\Autocreated\M RAAO Email.txt"
Public const TextfilePathForRAAO_name_M = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\Autocreated\M RAAO Name.txt"

Public const TextfilePathForReviewer3_Number_M = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\Autocreated\M Reviewer3 Number.txt"
Public const TextfilePathForReviewer4_Number_M = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\Autocreated\M Reviewer4 Number.txt"


'' RA APP Methods:
Public const PP_number_fromAPP_Methods = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\Autocreated\Project Personnel - Number - Methods.txt"

''' ***

Public const TextfilePathFor_attachmentsAPP = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\NumberOfAttachments.txt"
Public const TextfilePathFor_attachmentsAPP_LOItemplate = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\LOITemplate - attachmentNumber.txt"
Public const TextfilePathFor_attachmentsAPP_Lettersofsupport = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\LettersofSupport - attachmentNumber.txt"
Public const file_PFA_TC43 = "C:\Users\eferdous\OneDrive - PCORI\Automation\TEXT Files\Autocreated\Test PFA for TC 43.txt"

'''''''''''**************LOI Description Page External***********'''''''''''''''''
Public const LoiDescriptionApplyPage_Broad = "Description This Broad PCORI Funding Announcement seeks comparative clinical effectiveness research applications that address four of PCORI’s National Priorities for Health: Increase Evidence for Existing Interventions and Emerging Innovations in Health; Accelerate Progress Toward an Integrated Learning Health System; Achieve Health Equity; and Advance the Science of Dissemination, Implementation, and Health Communication\. Applicants to the Cycle 1 2024 BPS PFA may select from PCORI's topic themes, as applicable, that speak to everyday health issues facing large numbers of Americans: promoting health for children, youth, and older adults; addressing violence and trauma, or substance use; improving mental and behavioral health; and clinical conditions such as cardiovascular disease, sleep disturbance, and pain management\. Instructions Please click on the Apply button to create your LOI for the Broad Pragmatic Studies\.   PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_IMRI = "Description This PCORI Funding Announcement \(PFA\) seeks to fund implementation projects that promote the uptake of peer-reviewed findings from specific, high-priority, PCORI-funded research in the context of the body of related evidence\. PCORI has identified four areas of eligible evidence, each of which is the focus of important PCORI-funded research: Obesity treatment in primary care settings; Nonsurgical treatment options can improve or eliminate symptoms for women with urinary incontinence \(UI\); Several kinds of therapy and medicines can reduce or stop symptoms for people with posttraumatic stress disorder \(PTSD\); The use of narrow-spectrum versus broad-spectrum antibiotics to treat children’s acute respiratory tract infections \(ARTIs\)\. Instructions Please click on the Apply button to create your LOI for the Open Competition PFA: Implementation of Findings from PCORI Research Investments\.   PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_LC = "Description The intent of this limited competition PCORI Funding Announcement \(PFA\) is to move evidence developed with PCORI funded research toward practical use in improving health care and health outcomes\. This funding opportunity gives PCORI awardees the chance, following the generation of results from their PCORI-funded research award, to propose next steps to move their findings into practice, drawing on the knowledge and experience they gained during their PCORI-funded research award\. PCORI will fund projects that aim to implement patient-centered comparative clinical effectiveness research \(CER\) results obtained from PCORI-funded research in real-world practice settings, and in selected cases, projects that focus on the dissemination of these findings\. Instructions Please click on the Apply button to create your LOI for the Limited Competition PFA: Implementation Awards\.   PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_SDM = "Description This PCORI Funding Announcement \(PFA\) promotes the targeted implementation and systematic uptake of shared decision making \(SDM\) in healthcare settings, in line with PCORI’s goal of supporting patients in making informed decisions about their care\. This PFA seeks projects that propose active, multi-component approaches to implementing effective shared decision making \(SDM\) strategies that address existing barriers and obstacles to uptake and maintenance so that these interventions are effectively and sustainably integrated into practice\. Applicants may take either of two approaches: Propose to implement a SDM strategy that was formally tested and demonstrated to be effective in the context of a PCORI-funded research award; or Propose an implementation project that will incorporate new PCORI-funded clinical comparative clinical effectiveness research evidence from PCORI-funded research into an existing, tested SDM strategy, and then implement the updated SDM strategy\. Instructions Please click on the Apply button to create your LOI for the Implementation Effective Shared Decision Making Approaches in Practice Settings\. PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_Methods = "Description This PCORI Funding Announcement \(PFA\) for Improving Methods for Conducting Patient-Centered Outcomes Research \(PCOR\), also referred to as the “Methods PFA,” aims to fund studies that address high-priority methodological gaps in PCOR and comparative clinical effectiveness research \(CER\)\. PCORI seeks to fund projects that address important methodological gaps and lead to improvements in the strength and quality of evidence generated by PCOR/CER studies\. Priorities: Methods related to ethical and human subject protections \(HSP\) issues in PCOR/CER; methods to improve study design; methods to support data research networks; methods to improve use of artificial intelligence and machine learning in clinical research Instructions Please click on the Apply button to create your LOI for the Improving Methods for Conducting Patient-Centered Outcomes Research\. PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_SOE = "Description This PCORI Funding Announcement \(PFA\) seeks to fund studies that that build an evidence base on engagement in research, including: Measures to capture structure/context, process and outcomes of engagement in research; Techniques that lead to effective engagement in research; How effective engagement techniques should be modified and resourced for different contexts, settings and communities to ensure equity in engagement and research\. Applications should focus on the development and validation of measures and/or the development and testing of engagement techniques to generate evidence on the most effective engagement approaches, particularly for underrepresented populations\. Instructions Please click on the Apply button to create your LOI for the Science of Engagement \(SoE\)\.   PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_SleepHealth = "Description This PCORI Funding Announcement \(PFA\) seeks to fund high-quality, comparative clinical effectiveness research \(CER\) projects that focus on sleep health\. Applications may propose CER studies of screening, diagnostic and treatment approaches for sleep disorders; interventions promoting sleep health; or system-level strategies delivered in hospital, clinic or community settings to improve patient-centered sleep outcomes\. PCORI is particularly interested in submissions that address the following Special Areas of Emphasis \(SAEs\): Promoting Sleep Health Equity, Chronic Conditions Co-occurring with Sleep Disturbances, and Focus on Sleep Health Beyond Diagnosed Sleep Disorders\. Instructions Please click on the Apply button to create your LOI for the Sleep Health Topical   PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_PLACER = "Description This PCORI funding announcement \(PFA\) seeks to fund large, randomized trials in patient-centered clinical comparative effectivness research \(CER\) that are structured into two well-integrated phases: an initial feasibility phase followed by a second phase in which conduct of the full-scale clinical trial proceeds once specific milestones and deliverables are achieved\. The proposed trials should address critical decisional dilemmas that require important new evidence about the comparative clinical effectiveness of available interventions\. This PFA seeks applications addressing three of PCORI’s National Priorities for Health: increase evidence for existing interventions and emerging innovations in health; accelerate progress toward an integrated learning health system; achieve health equity\. To be considered responsive, applications must propose research meeting the distinctive requirements of this PFA and address at least one National Priority for Health\. Applicants have the option to choose up to three of PCORI’s topic themes\. Instructions Please click on the Apply button to create your LOI for the Phased Large Awards for Comparative Effectiveness Research\. PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_MMM = "Description This Targeted PCORI Funding Announcement \(Targeted PFA\) seeks to fund large randomized controlled trials \(RCTs\) and/or well-designed observational studies comparing multicomponent strategies to improve early detection of, and timely care for, complications up to six weeks postpartum for groups more often underserved or experiencing the greatest disparities in health outcomes, including Black, American Indian/Alaska Native \(AI/AN\), Hispanic, rural, and low socioeconomic status \(SES\) populations\. Instructions Please click on the Apply button to create your LOI for the Improving Postpartum Maternal Outcomes for Populations Experiencing Disparities\. PCORI funding opportunities are here: http://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_HSII = "Description This Call for Proposals provides an opportunity for HSII Participants to propose projects that promote the uptake of specific evidence from PCORI-funded comparative clinical effectiveness research within their healthcare delivery settings\. Applicants may propose either an HSII Implementation Project or HSII Pilot Project focused on one of the following topics: - Intensive Lifestyle Treatment for Weight Loss in Primary Care Settings - Appropriate Antibiotic Prescribing for Children With Acute Respiratory Tract Infections \(ARTIs\) - Addressing Hypertension in Adults  - Monitoring Electronic Patient-Reported Outcomes During Cancer Treatment  Instructions Please click on the Apply button to create your LOI for the PCORI Health Systems Implementation Initiative \(HSII\)--Implementation Projects\.   PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_Mental_Health = "Description The Improving Mental and Behavioral Health Topical PCORI funding announcement \(PFA\) seeks to fund patient-centered comparative clinical effectiveness research \(CER\) projects that focus on improving mental and behavioral health\. PCORI has identified three Special Areas of Emphasis for this PFA: mental and behavioral health of children and youth; suicide prevention and crisis response; and strategies to improve mental health care access and delivery\. Instructions Please click on the Apply button to create your LOI for the Improving Mental And Behavioral Health Topical\.   PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_Managing_Pain = "Description The Managing Pain Topical PCORI funding announcement \(PFA\) seeks to fund patient-centered comparative clinical effectiveness research \(CER\) projects that focus on interventions to improve patient-centered outcomes for individuals living with acute and/or chronic pain\. PCORI has identified four Special Areas of Emphasis for this PFA: urogynecological and pelvic pain; pain in individuals living with limitations in cognitive functioning; pain in individuals living with sickle cell disease; and neuropathic pain\. Instructions Please click on the Apply button to create your LOI for the Managing Pain Topical\.   PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"
Public const LoiDescriptionApplyPage_Violence_Truama = "Description The Addressing Violence and Trauma Topical PCORI funding announcement \(PFA\) seeks to fund patient-centered comparative clinical effectiveness research \(CER\) projects that focus on addressing violence and trauma\. PCORI has identified four Special Areas of Emphasis for this PFA: Intentional trauma; unintentional trauma; and substance use and trauma\. Instructions Please click on the Apply button to create your LOI for Addressing Violence and Trauma Topical\.   PCORI funding opportunities are here: https://www\.pcori\.org/funding-opportunities"


''''''''''Contact Information Tab'''''''''''''''
Public const Instruction_Text1 = "For instructions on using our system, click here to access our PCORI Portal User Guide\."
Public const Instruction_Text2 = "To view additional information about the PCORI submission process, click here to view our FAQ page\."
Public const Instruction_Text3 = "Click 'Save & Next' to continue to the next tab\. Otherwise you could receive an error message\."
Public const Instruction_Text4 = "- A Principal Investigator \(PI\) and an Administrative Official \(AO\) with a valid email address are required to submit your LOI to PCORI\.  - To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking “New User\.” - User login information from the previous PCORI Online were not migrated to the new PCORI Online\. - The AO and the PI cannot be the same individual\.  - Individuals assigned at the “Contact Information” tab will have access to the LOI and Application\. - To find your Congressional District please click here\.  - Fields marked with \(\*\) are required\."
Public const Principal_Investigator_Contact = "\* Principal Investigator \(Contact\) "
Public const DualPI_Name = "Dual Principal Investigator Name"
Public const DualPI_Email = "Dual Principal Investigator Email"
Public const Administrative_Official = "\* Administrative Official "
Public const PI_Designee_1 = "PI Designee 1"
Public const PI_Designee_2 = "PI Designee 2"
Public const Financial_Officer = "Financial Officer"
Public const Organization_LOI = "\* Organization "
Public const Congressional_District = "\* Congressional District "
Public const Department = "\* Department "
Public const Link1URL = "https://www\.pcori\.org/sites/default/files/PCORI-Online-Pre-Award-User-Guide\.pdf"
Public const Link2URL = "https://www\.pcori\.org/funding-opportunities/applicant-and-awardee-resources/frequently-asked-questions-faqs"
Public const Link3URL = "https://pcori\.force\.com/engagement"                           '''''''''''''''for Prod = https://pcori\.my\.site\.com/engagement"""""" For PQT = https://pcori\.force\.com/engagement"
Public const Link4URL = "http://www\.house\.gov/representatives/find/"

'''''''''''''''''''Contact Information tab LOI LC'''''''''''''''
Public const Instruction_Text4_LC = "- A Principal Investigator \(PI\) and an Administrative Official \(AO\) with a valid email address are required to submit your LOI to PCORI\.  - To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking “New User\.” - User login information from the previous PCORI Online were not migrated to the new PCORI Online\. - The AO and the PI cannot be the same individual\.  - Individuals assigned at the “Contact Information” tab will have access to the LOI and Application\. - To find your Congressional District please click here\.  - Fields marked with \(\*\) are required\."

''''''''''''''SDM Contact Information tab''''''''''''
Public const Instruction_Text4_SDM = "- A Principal Investigator \(PI\) and an Administrative Official \(AO\) with a valid email address are required to submit your LOI to PCORI\.  - To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking ""New User\."" - User login information from the previous PCORI Online were not migrated to the new PCORI Online\. - The AO and the PI cannot be the same individual\.  - Individuals assigned at the ""Contact Information"" tab will have access to the LOI and Application\. - To find your Congressional District please click here\.  - Fields marked with \(\*\) are required\."

'''''''''''''HSII Contact Information tab''''''''''''
Public const Instruction_Text1_HSII = "For instructions on using our system, click here to access our HSII User Guide\."
Public const Instruction_Text2_HSII = "To view additional information about the PCORI application process, click here to access our Applicant FAQs page\."
Public const Instruction_Text3_HSII = "To view HSII FAQs, click here\."
Public const Instruction_Text5_HSII = "- A Principal Investigator \(PI\) and an Administrative Official \(AO\) with a valid email address are required to submit your LOI to PCORI\. - To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking “New User\.” - The AO and the PI cannot be the same individual\. - Individuals assigned at the “Contact Information” tab will have access to the LOI and Application\. - To find your Congressional District please click here\. - Fields marked with \(\*\) are required"
Public const DualPI_Title = "Dual Principal Investigator Position Title 0 of 255 Characters "
Public const HSII_Participant = "\* HSII Participant Health System Name 0 of 255 Characters "
Public const Link1URL_HSII = "https://www\.pcori\.org/sites/default/files/PCORI-Online-HSII-User-Guide\.pdf"
Public const Link3URL_HSII = "https://www\.pcori\.org/impact/putting-evidence-work/health-systems-implementation-initiative/health-systems-implementation-initiative-faqs"
Public const Link4URL_HSII = "https://www\.pcori\.org/impact/putting-evidence-work/health-systems-implementation-initiative/health-systems-implementation-initiative-faqs"

'''''''''''''''Contact Information tab application ''''''''''
Public const Instruction_Textapp1 = "Contact Information"
Public const Instruction_Textapp2 = "NOTE: PLEASE CONFIRM THAT YOU HAVE ADDED AN ADMINISTRATIVE OFFICIAL \(AO\) AND PRINCIPAL INVESTIGATOR \(PI\) AND THAT YOUR AO AND PI ARE NOT THE SAME PERSON\."
Public const Location_Satellite_field = "Location/Satellite"
Public const Organization_field_IMRI = "\* Organization If organization not provided, please email: PFA@pcori\.org "
Public const Html_ID_for_Location_Satellite_field = "j_id0:j_id2:j_id3:mainForm:j_id713:12:inputFieldId"
''''Public const Instruction_Text2_App = "To view additional information about the PCORI submission process, click here to access our Applicant FAQs page\."



''''''''''''SDM App Contact Information tab''''''''''''
Public const Financial_Officer_SDM = "\* Financial Officer "

'

''''''''''''''''''''''Pre-Screen Questionnaire''''''''''''''Tab
Public const PreScQIns1 = "Do any of the specific aims of your research propose:"
Public const Decision_Aid = "\* Creation of a decision aid or tool --None-- Yes No "
Public const New_Intervention = "\* Development of a novel clinical intervention that has not yet been tested \(Answer ""no"" if your application is proposing to tailor or adapt a tested intervention to a specific population or condition\) --None-- Yes No "
Public const Practice_Guidelines = "\* Creation of practice guidelines   --None-- Yes No "
Public const Cost_Effective_Analysis = "\* Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care   --None-- Yes No "
Public const PreScQIns2 = "If the answer to any of the previous questions is ""yes"" your Letter of Intent will not progress past the review stage\."
Public const foreign_organization = "\* Does your proposed project involve a foreign organization as the prime applicant or foreign organization\(s\) as sub-contractor\(s\) or project performance site\(s\)\? --None-- Yes No If yes, please make sure that you review PCORI's award eligibility requirements\. "
Public const PreScQLinkURL1 = "https://help\.pcori\.org/hc/en-us/sections/200564830-Who-Can-Apply"
Public const PreScQLinkURL2 = "https://help\.pcori\.org/hc/en-us/articles/202901800-Are-foreign-organizations-eligible-for-funding-"

''''''''''''''''''Pre-Screen Questionnaire Methods''''''''''''''
Public const clinical_prediction_model_LOI = "\* Disease or condition-specific clinical prediction model --None-- Yes No "

''''''''''''''''''Pre-Screen Questionnaire SOE'''''''''''''
Public const Cost_Effective_Analysis_SOE = "\* Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care --None-- Yes No "
Public const Cost_Effective_Analysis_PostTextOuterHtml_SOE = "<b>If the answer to any of the previous questions is ""yes"" your Letter of Intent will not progress past the review stage\.</b>"

''''''''''''''''''Pre-Screen Questionnaire PLACER'''''''''''''
Public const Cost_Effective_Analysis_PostTextOuterHtml_PLACER = "<b>If the answer to any of the previous questions is ""yes"" your Letter of Intent will be judged non-responsive and you will not be invited to submit a full application\.</b>"

''''''''''''''''''''''Pre-Screen Questionnaire DI LOI''''''''''''''Tab
Public const Pre_ScQIns_DI_1 = "Do any of the specific aims of your project propose:"
Public const Practice_Guidelines_DI = "\* Creation of practice guidelines --None-- Yes No "
Public const Cost_Effective_Analysis_DI_LOI = "\* Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care   --None-- Yes No "
Public const PreScQIns2_DI_PostText_Outerhtml = "<b>If the answer to any of the previous questions is ""yes"" your Letter of Intent will not progress past the review stage\.</b>"
Public const Foreign_organization_DI = "\* Does your proposed project involve a foreign organization as the prime applicant or foreign organization\(s\) as sub-contractor\(s\) or project performance site\(s\)\? --None-- Yes No If yes, please make sure that you review PCORI's award eligibility requirements and the guidance detailed in this FAQ\. "
Public const Passive_dissemination = "\* Passive Dissemination --None-- Yes No "
Public const Dissemination_Focused_Project = "\* Are you proposing a dissemination or an implementation planning project\? --None-- Yes No "
Public const Cost_Effective_Analysis_LC_LOI = "\* Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care --None-- Yes No "

''''''''''''''''''''''Pre-Screen Questionnaire SDM LOI''''''''''''''Tab
Public const Efficacy_Effectiveness_SDM_Strategies_LOI = "\* To establish efficacy or effectiveness of SDM strategies or to study the comparative effectiveness of multiple SDM strategies --None-- Yes No "
Public const Disseminate_funded_CER_Methods_study_LOI = "\* To implement findings that are not associated with a PCORI-funded CER or methods study --None-- Yes No "
Public const Translate_adapt_shared_decision_making_LOI = "\* To translate or adapt a shared decision making approach without the primary purpose of actively implementing it --None-- Yes No "
Public const Develop_new_tool = "\* To develop or validate a new tool or system for patients or clinicians without the primary purpose of implementing evidence --None-- Yes No "
Public const Cost_Effective_Analysis_SDM_LOI = "\* Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care --None-- Yes No "

''''''''''''''''''''''Pre-Screen Questionnaire HSII LOI''''''''''''''Tab
Public const Executed_HSII_MFA = "\* Please confirm that the applicant organization has a fully executed Master Funding Agreement for HSII with PCORI\. --None-- Yes No "
Public const HSII_PL = "\* Please confirm that at least one of the proposed Principal Investigators \(contact or Dual\) is eligible to submit an HSII project proposal \(i\.e\., is listed as a Project Lead in the HSII Master Funding Agreement\)\. If you answer ‘No’, please do not proceed as you are not eligible to submit an HSII project proposal\. --None-- Yes No "
Public const PreScQIns2_HSII_redtext_Outerhtml = "<b style=""color: Red;"">If you answer ‘No’, please do not proceed as you are not eligible to submit an HSII project proposal\.</b>"

''''''''''''''''Pre Screen Questionnaire Tab External D&I Application '''''
Public const DI_Prescreen_Inst1 = "Do any of the specific aims of your project propose:"
Public const Dissemination = "\* Dissemination --None-- Yes No "
Public const Cost_Effective_Analysis_DI = "\* Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care   --None-- Yes No If the answer to any of the previous questions is ""yes"" your Letter of Intent will not progress past the review stage\. "
Public const Aims_develop_validate_new_tool_system = "\* To develop and/or validate a new tool or system for patients and/or clinicians without the primary purpose of implementing evidence --None-- Yes No "
Public const App_Prescreen_Foreign_Organization = "\* Does your proposed project involve a foreign organization as the prime applicant or foreign organization\(s\) as sub-contractor\(s\) or project performance site\(s\)\? --None-- Yes No "
Public const App_Prescreen_Foreign_Organization_postText = "If yes, please make sure that you review PCORI's award eligibility requirements\."
Public const Html_ID_for_Dissemination_field = "j_id0:j_id2:j_id3:mainForm:j_id713:2:inputFieldId"
Public const Html_ID_for_Aims_develop_validate_new_tool_system_field = "j_id0:j_id2:j_id3:mainForm:j_id713:3:inputFieldId"
Public const Html_ID_for_Creation_of_practice_guidelines_field = "j_id0:j_id2:j_id3:mainForm:j_id713:4:inputFieldId"
Public const Html_ID_for_cost_effectiveness_analysis_field = "j_id0:j_id2:j_id3:mainForm:j_id713:5:inputFieldId"
Public const Html_ID_for_Foreign_Organization_field = "j_id0:j_id2:j_id3:mainForm:j_id713:6:inputFieldId"
Public const PreScQLinkURL2_DI = "https://www\.pcori\.org/funding-opportunities/applicant-and-awardee-resources/frequently-asked-questions/who-can-apply-faqs"

''''''''''''''''Error Message for Pre Screen Tab DI if no fields r filled out''''''''''
Public const PScQ_Tab_err1_DI = "Warning This is a required field\. Dissemination"
Public const PScQ_Tab_err2_DI =  "Warning This is a required field\. To develop and/or validate a new tool or system for patients and/or clinicians without the primary purpose of implementing evidence"
Public const PScQ_Tab_err3_DI = "Warning This is a required field\. Creation of practice guidelines"

''''''''''''''''Error Message for Pre Screen Tab if no fields r filled out''''''''''
Public const PScQ_Tab_err1 = "Warning This is a required field\. Creation of a decision aid or tool"
Public const PScQ_Tab_err2 = "Warning This is a required field\. Development of a novel clinical intervention that has not yet been tested \(Answer ""no"" if your application is proposing to tailor or adapt a tested intervention to a specific population or condition\)"
Public const PScQ_Tab_err3 =  "Warning This is a required field\. Creation of practice guidelines  "
Public const PScQ_Tab_err4 = "Warning This is a required field\. Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care  "
Public const PScQ_Tab_err5 = "Warning This is a required field\. Does your proposed project involve a foreign organization as the prime applicant or foreign organization\(s\) as sub-contractor\(s\) or project performance site\(s\)\?"

'''''''''''''''''''Pre Screen Questionnaire Tab SOE Application''''''''''''''''''''

Public const Decision_Aid_SOE_App = "\* Creation of a decision aid or tool   For all category 1 \(measurement\) proposals, the creation of a decision aid or tool will be deemed non-responsive to this funding mechanism\. The creation of a decision aid or tool may be allowable under category 2 \(engagement methods\) in certain circumstances\. --None-- Yes No "
Public const New_Intervention_SOE_App = "\* Development of a novel clinical intervention that has not yet been tested \(Answer ""no"" if your application is proposing to tailor or adapt a tested intervention to a specific population or condition\) For all category 1 \(measurement\) proposals, the development of a clinical intervention will be deemed non-responsive to this funding mechanism\. The development of a clinical intervention may be allowable under category 2 \(engagement methods\) in certain circumstances\. --None-- Yes No "
Public const Cost_Effective_Analysis_SOE_App = "\* Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care   Formal cost effectiveness analysis is not allowable within PCORI projects --None-- Yes No "
Public const Practice_Guidelines_SOE_App = "\* Creation of practice guidelines   Creation of practice guidelines is not responsive to this funding mechanism --None-- Yes No "
Public const Foreign_Organization_SOE_App = "\* Does your proposed project involve a foreign organization as the prime applicant or foreign organization\(s\) as sub-contractor\(s\) or project performance site\(s\)\? Please consult our Applicant FAQs regarding whether foreign organizations can apply for funding and whether foreign sites/subcontractors are allowed --None-- Yes No "

'''''''''''''''''''Pre Screen Questionnaire Tab Methods Application''''''''''''''''''''
Public const Cost_Effective_Analysis_Methods_App = "\* Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care   --None-- Yes No If the answer to any of the previous questions is ""Yes,"" your Letter of Intent will not progress past the review stage\. "
Public const Cost_Effective_Analysis_PostTextOuterHtml_Methods_App = "<b>If the answer to any of the previous questions is ""Yes,"" your Letter of Intent will not progress past the review stage\.</b>"

'''''''''''''''Placer Acknowledgement of PLACER Requirements Tab'''''''''''''''''''''
Public const Inst_PLACER_APR_1 = "Answering yes to each item below acknowledges the applicant's awareness of the requirements for PLACER research awards:"
Public const PLACER_Participants_randomized = "\* Projects must involve randomization of individuals or clusters as the primary investigative focus\. --None-- Yes No "
Public const Requirements_for_DCC = "\* All projects will be expected to have a Data Coordinating Center \(DCC\) with an independent scientific leadership role in study decision making about the analytical, statistical, and data management aspects of both study phases\. For applicants to this PFA, a Letter of Intent \(LOI\) may be submitted without a DCC being in place—with the proviso that invited applicants must include a DCC as part of their application submission\. PCORI requires that the PI of the DCC and the PI of the CCC be named and serve as dual-PIs to promote parity, scientific independence, and autonomy of clinical, data management, and analytical input into study decisions\. If other arrangements are proposed, applicants must seek prior approval from PCORI to ensure the co-equal and independent voices of the CCC and DCC in trial design, conduct, and oversight\. --None-- Yes No "

''''''''''''''''''''SDM Application Prescreen Questionaaire Tab''''''''
Public const Translate_shared_decision_making = "\* To translate or adapt a shared decision making approach without actively implementing it --None-- Yes No "
Public const Aims_develop_validate_new_tool_system_SDM = "\* To develop or validate a new tool or system for patients or clinicians without the primary purpose of actively implementing evidence --None-- Yes No "
Public const Cost_Effective_Analysis_SDM_App = "\* Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care --None-- Yes No If the answer to any of the previous questions is ""yes"" your Letter of Intent will not progress past the review stage\. "
Public const Html_ID_for_Efficacy_SDM_Strategies = "j_id0:j_id2:j_id3:mainForm:j_id713:2:inputFieldId"
Public const Html_ID_for_Disseminate_funded_study = "j_id0:j_id2:j_id3:mainForm:j_id713:3:inputFieldId"
Public const Html_ID_for_Translate_shared_decision_making = "j_id0:j_id2:j_id3:mainForm:j_id713:4:inputFieldId"
Public const Html_ID_for_Aims_develop_validate_new_tool_system_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:5:inputFieldId"
Public const Html_ID_for_Creation_of_practice_guidelines_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:6:inputFieldId"
Public const Html_ID_for_cost_effectiveness_analysis_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:7:inputFieldId"

'''''''''''''''''''Pre Screen Questionnaire Tab BPS Application''''''''''''''''''''

Public const Decision_Aid_BPS = "\* Creation of a decision aid or tool   --None-- Yes No "
Public const Practice_Guidelines_BPS = "\* Creation of practice guidelines --None-- Yes No "
Public const Cost_Effective_Analysis_BPS_App = "\* Conducting a formal cost effectiveness analysis or directly comparing the costs of care between two or more alternative approaches to providing care  --None-- Yes No If the answer to any of the previous questions is ""Yes,"" your Letter of Intent will not progress past the review stage\. "

'''''''''''''''''''Pre Screen Questionnaire Tab PLACER Application''''''''''''''''''''
Public const Cost_Effective_Analysis_PLACER_App = "\* Conducting a formal cost effectiveness analysis   --None-- Yes No If the answer to any of the previous questions is ""yes"" your Letter of Intent will be judged non-responsive and you will not be invited to submit a full application\. "


''''''''''''''********Resubmission Tab Broad''''''''''''''''''''''
Public const ResubIns1 = "LOI, Application Resubmission Questions"
Public const ResubIns2 = "Listed below are questions regarding the submission of an LOI and/or Application to a previous PCORI funding announcement\. If you require additional assistance locating a previous application title and/or ID number, please contact pfa@pcori\.org"
Public const LOI_Resubmission = "\* Have you submitted this project to PCORI before as an LOI\? --None-- Yes No "
Public const Previous_LOI_Invited = "\* Was that LOI invited for a full application\? --None-- Yes No "
Public const App_Resubmission = "\* Does this project meet PCORI’s requirements to be considered a resubmission \(The same Principal Investigator is resubmitting to the same PCORI Funding Announcement and the Principal Investigator received a Merit Review Summary Statement\)\? Note: A prior Pragmatic Clinical Studies application \(post Cycle 1 2021\) may be resubmitted to the Broad Pragmatic Studies \(BPS\) PFA\. Any prior submission to a Broad PFA/BPS PFA may resubmit to the 2023 Broad Pragmatic Studies PFA\. --None-- Yes No "
Public const Previous_LOI_Bypass_Review = "\* After the previous Merit Review cycle, did you receive an invitation from PCORI inviting you to bypass the LOI process for this project\? --None-- Yes No "
Public const Previous_Application = "If you answered yes to any of the above questions, please enter the LOI/Application number of your prior project referenced above\. If you answered No to all questions above, please leave this field blank\. 0 of 255 Characters "

'''''''''''''********Resubmission Tab Methods''''''''''''''''''''''
Public const Previous_Application_Methods = "Please enter the LOI/Application number of your prior project referenced above\. If you answered No to all questions above, please leave this field blank\. 0 of 255 Characters "
Public const App_Resubmission_Methods = "\* Does this project meet PCORI’s requirements to be considered a resubmission \(The same Principal Investigator is resubmitting to the same PCORI Funding Announcement and the Principal Investigator received a Merit Review Summary Statement\)\? --None-- Yes No "

'''''''''''''********Resubmission Tab SOE''''''''''''''''''''''
Public const Previous_LOI_Bypass_Review_SOE = "\* After the previous Merit Review cycle, did you receive an invitation from PCORI inviting you to bypass the LOI process for this project\? --None-- Yes No If answered ""yes"", please attach the invitation letter in the ""Templates & Uploads"" section\. "
Public const Previous_Application_SOE = "Please enter the LOI/Application number of your prior project referenced above\. 0 of 255 Characters "
Public const LOI_Resubmission_SOE = "\* Have you submitted this project to PCORI before as an LOI\? --None-- Yes No If you answered 'Yes' to this question, please answer the following three questions regarding your LOI Resubmission\. If you do not see values in the drop-down lists for the fields below, click ""Save"" at the bottom of the page to respond to the additional questions about your resubmission\. "

''''''''''''''*****Resubmission tab Error Broad''''''''''''''''''''
Public const Resub_Tab_err1 = "Warning This is a required field\. Have you submitted this project as an LOI to any prior PCORI funding opportunities\?"
Public const Resub_Tab_err2 = "Warning This is a required field\. Was that LOI invited for a full application\?"
Public const Resub_Tab_err3 = "Warning This is a required field\. Does this project meet PCORI’s requirements to be considered a resubmission \(The same Principal Investigator is resubmitting to the same PCORI Funding Announcement and the Principal Investigator received a Merit Review Summary Statement\)\? Note: A prior Pragmatic Clinical Studies application \(post Cycle 1 2021\) may be resubmitted to the Broad Pragmatic Studies \(BPS\) PFA\. Any prior submission to a Broad PFA/BPS PFA may resubmit to the 2023 Broad Pragmatic Studies PFA\."
Public const Resub_Tab_err4 = "Warning This is a required field\. After the previous Merit Review cycle, did you receive an invitation from PCORI inviting you to bypass the LOI process for this project\?"

''''''''''''''*****Resubmission tab Error Methods'''''''''''''''''''
Public const Resub_Tab_Methods_err3 = "Warning This is a required field\. Does this project meet PCORI’s requirements to be considered a resubmission \(The same Principal Investigator is resubmitting to the same PCORI Funding Announcement and the Principal Investigator received a Merit Review Summary Statement\)\?"

''''''''''''''********Resubmission Tab DI LOI'''''''''''''''
Public const LOI_Resubmission_DI_in_LOIPhase = "\* Have you previously submitted this project as an LOI to any PCORI PFA\? --None-- Yes No If you answered 'Yes' to this question, please answer the following questions regarding your LOI Resubmission\. If you do not see values in the drop-down lists for the fields below, click “Save” at the bottom of the page to respond to the additional questions about your resubmission\. "
Public const Was_LOI_Invited_DI_in_LOIPhase = "Have you submitted this project to PCORI before as a full application\? --None-- "
Public const Number_of_submissions_DI_in_LOIPhase =  "How many times have you submitted this project to PCORI before as a full application\? --None-- "
Public const Project_Resubmitted_DI_in_LOIPhase = "Have you submitted this project to PCORI before as a full application\? --None-- "
Public const Summary_Statement_received_DI_in_LOIPhase  = "Have you received a summary statement from your latest full application submission\? --None-- "
Public const Previous_Application_DI_in_LOIPhase = "Please enter the ID\(s\) of your prior application\(s\) 0 of 255 Characters "
Public const Previous_LOI_Invited_LC_LOIPhase = "Was that LOI invited for a full application\? --None-- "


''''''''''''''*****Resubmission tab Error DI LOI''''''''''''''''''''
Public const Resub_Tab_DI_LOI_err1 = "Error: Since you answered Yes to 'Have you submitted this project to PCORI before as an LOI\?' question you must answer all the questions in the Resubmission tab\. "
Public const Resub_Tab_DI_LOI_err2 = "Error: Since you answered Yes to 'Have you submitted this project to PCORI before as a full Application\?' question you must answer all the questions in the Resubmission tab\. "

''''''''''''''''''Resubmission Tab SDM LOI''''''''''''''''
Public const LOI_Resubmission_SDM_in_LOIPhase = "\* Have you submitted this project to the SDM PFA before as an LOI\? --None-- Yes No If you answered 'Yes' to this question, please answer the following questions regarding your LOI Resubmission\. If you do not see values in the drop-down lists for the fields below, click “Save” at the bottom of the page to respond to the additional questions about your resubmission\. "
Public const Was_LOI_Invited_SDM_in_LOIPhase = "Was that LOI invited for a full application\? --None-- "

''''''''''''''''''Resubmission Tab PLACER LOI''''''''''''''''

Public const LOI_Resubmission_PLACER = "\* Have you submitted this project as an LOI to any prior PCORI funding opportunities\? --None-- Yes No "

'''''''''''Resubmission Tab for Mental Health and Managing Pain'''''''''''
Public const LOI_Resubmission_newpfa_LOI_Phase = "\* Have you submitted this project as an LOI to any prior PCORI funding opportunities\? --None-- Yes No "
Public const Previous_Application_newpfa_LOI_Phase = "If you answered yes to any of the above questions, please enter the LOI/Application number of your prior project referenced above\. If you answered No to all questions above, please leave this field blank\. 0 of 255 Characters "

''''''''''''''''''Application Phase Resubmission Tab' External D&I'''''''''''
Public const AppDI_resubmission_Inst1 = "LOI / Application Resubmission Questions"
Public const AppDI_resubmission_Inst2 = "Listed below are questions regarding the submission of an LOI and/or Application to a previous PCORI funding announcement\. If you require additional assistance locating a previous application title and/or ID number, please contact pfa@pcori\.org\. For PCORI’s resubmission policy, please visit the Applicant FAQs\."
Public const ResubURLapp1 = "https://www\.pcori\.org/funding-opportunities/applicant-and-awardee-resources/frequently-asked-questions/who-can-apply-faqs"
Public const LOI_Resubmission_DI = "\* Have you previously submitted this project as an LOI to any PCORI PFA\? --None-- Yes No If you do not see values in the fields below, click “Save” at the bottom of the page to respond to the additional questions about your resubmission\. "
Public const Resub_App_post_text1_OuterHtml = "<font color=""purple"">If you do not see values in the fields below, click “Save” at the bottom of the page to respond to the additional questions about your resubmission\.</font>"
Public const LOI_Resubmission_IMRI = "\* Have you previously submitted this project as an LOI to any PCORI PFA\? --None-- Yes No If you do not see values in the fields below, click ""Save"" at the bottom of the page to respond to the additional questions about your resubmission\. "
Public const Resub_IMRI_App_post_text1_OuterHtml = "<font color=""purple"">If you do not see values in the fields below, click ""Save"" at the bottom of the page to respond to the additional questions about your resubmission\.</font>"
Public const Was_LOI_Invited_DI = "Was that LOI invited for a full application\? --None-- Yes No "
Public const Project_Resubmitted_DI = "Have you submitted this project to PCORI before as a full application\?"
Public const Number_of_submissions = "How many times has this project been submitted to PCORI as a full application\? --None-- 1 2 3 NA "
Public const Summary_Statement_received = "Have you received a summary statement from your latest full application submission\? --None-- Yes No "
Public const Previous_Application_innertext = "Please enter the ID of your prior application\(s\)"
Public const Html_ID_for_LOI_Resubmission_DI_field = "j_id0:j_id2:j_id3:mainForm:j_id713:2:inputFieldId"
Public const Html_ID_for_LOI_Was_LOI_Invited_DI_field = "j_id0:j_id2:j_id3:mainForm:j_id713:3:inputFieldId"
Public const Html_ID_for_Project_Resubmitted_DI_field = "j_id0:j_id2:j_id3:mainForm:j_id713:4:inputFieldId"
Public const Html_ID_for_Number_of_submissions_field = "j_id0:j_id2:j_id3:mainForm:j_id713:5:inputFieldId"
Public const Html_ID_for_Summary_Statement_received_field = "j_id0:j_id2:j_id3:mainForm:j_id713:6:inputFieldId"
Public const Html_ID_for_Previous_Application_field = "j_id0:j_id2:j_id3:mainForm:j_id713:7:inputFieldId"

''''''''''''''*****Resubmission tab Error DI -LC''''''''''''''''''''
Public const Resub_Tab_LCerr1 = "Warning This is a required field\. Have you previously submitted this project as an LOI to any PCORI PFA\?"
Public const Resub_Tab_FCDefault_err1 = "Complete the following fields to save changes on this tab "
Public const Resub_Tab_LCerr2 = "Error: If answered 'Yes' to previously submitting a project to PCORI as an LOI, must complete the following questions on the Resubmission tab "
Public const Resub_Tab_LCerr3 = "Error: Since you answered 'Yes' to 'Have you submitted this project to PCORI before as a full application\?' you must answer the number of submission and summary statement received questions on the Resubmission tab\. "
Public const Resub_Tab_LCerr4 = "Error: Since you answered 'Yes' to 'Have you submitted this project to PCORI before as a full application\?' you must answer the ID of your prior application\(s\) "

''''''''''''''*****Resubmission tab Error DI -SDM''''''''''''''''''''
Public const Resub_Tab_SDMerr1 = "Warning This is a required field\. Have you submitted this project to the SDM PFA before as an LOI\?"

''''''''''''''''''''Resubmission Tab SOE Application'''''''''''''
Public const Previous_Application_SOE_App  = "Please enter the LOI/Application number of your prior project referenced above\. 5 of 255 Characters "
Public const Html_ID_for_LOI_Resubmission_SOE_field = "j_id0:j_id2:j_id3:mainForm:j_id713:2:inputFieldId"

''''''''''''''Resubmission Tab SDM Application''''''''''''''''''''
Public const LOI_Resubmission_DI_SDM = "\* Have you submitted this project to PCORI before as an LOI\? --None-- Yes No If applicant answered 'Yes' to this question, please answer the following three questions regarding your LOI Resubmission\. If you do not see values in the fields below, click ""Save"" at the bottom of the page to respond to the additional questions about your resubmission\. "
Public const Resub_App_post_text1_OuterHtml_SDM = "<font color=""purple"">If applicant answered 'Yes' to this question, please answer the following three questions regarding your LOI Resubmission\. If you do not see values in the fields below, click ""Save"" at the bottom of the page to respond to the additional questions about your resubmission\.</font>"
Public const Html_ID_for_LOI_Resubmission_DI_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:2:inputFieldId"
Public const Html_ID_for_LOI_Was_LOI_Invited_DI_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:3:inputFieldId"
Public const Html_ID_for_Project_Resubmitted_DI_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:4:inputFieldId"
Public const Html_ID_for_Number_of_submissions_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:5:inputFieldId"
Public const Html_ID_for_Summary_Statement_received_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:6:inputFieldId"

'''''''''''''Resub Tab Error SDM Application'''''''
Public const Resub_Tab_LCerr1_SDM = "Warning This is a required field\. Have you submitted this project to PCORI before as an LOI\?"

''''''''''''''********Resubmission Tab PLACER Application''''''''''''''''''''''
Public const Previous_LOI_Invited_PLACER_App = "Was that LOI invited for a full application\? --None-- Yes No "
Public const App_Resubmission_PLACER_App = "Does this project meet PCORI’s requirements to be considered a resubmission \(The same Principal Investigator is resubmitting to the same PCORI Funding Announcement and the Principal Investigator received a Merit Review Summary Statement\?\) --None-- Yes No "
Public const Previous_LOI_Bypass_Review_PLACER_App = "After previous Merit Review cycle, did you receive an invitation from PCORI inviting you to bypass the LOI process for this project\? --None-- Yes No If you answered ""yes,"" please attach the invitation letter in the ""Templates & Uploads"" section "
Public const Previous_LOI_Bypass_Review_PLACER_BoldedInsText_App = "<b>If you answered ""yes,"" please attach the invitation letter in the ""Templates &amp; Uploads"" section</b>"
Public const LOI_Resubmission_PLACER_App = "\* Have you submitted this project to PCORI before as an LOI\? --None-- Yes No If applicant answered 'Yes' to this question, please answer the following three questions regarding your LOI Resubmission\. If you do not see values in the fields below, click “Save” at the bottom of the page to respond to the additional questions about your resubmission\. "
Public const Resub_App_post_text1_OuterHtml_PLACER = "<font color=""purple"">If applicant answered 'Yes' to this question, please answer the following three questions regarding your LOI Resubmission\. If you do not see values in the fields below, click “Save” at the bottom of the page to respond to the additional questions about your resubmission\.</font>"

''''''''''''''********Resubmission Tab Broad Application''''''''''''''''''''''
Public const Previous_Application_BPS = "Please enter the LOI/Application number of your prior project referenced above\."
Public const Previous_Application_Inst = "If you answered No to all questions above, please leave this field blank\."
Public const Html_ID_for_LOI_Resubmission_BPS_field = "j_id0:j_id2:j_id3:mainForm:j_id713:2:inputFieldId"

'''''''''''*******PI Information Tab Broad''''''''''''''''''''
Public const PI_Work_Telephone = "\* PI Work Telephone Number "
Public const Primary_Group_Identification = "\* For the purpose of this project, with which group does the PI or project lead identify primarily\?--None-- Patient/Consumer Caregiver/Family member of patient Patient/Caregiver Advocacy Organization Clinician Clinic/Hospital/Health System Purchaser Payer Industry Research Policy Maker Training Institution "
Public const Previous_involvement_PCORI = "\* Have you interacted with PCORI in the past in the following ways\? \(Select all that apply\) Joined a PCORI email list Visited PCORI’s website Participated in applicant training Watched a PCORI webinar Attended PCORI sponsored event in-person Attended event where PCORI was featured Met with PCORI staff Met with a PCORI Ambassador Applied to review PCORI funding app Applied for PCORI funding Received PCORI funding Served as a PCORI Merit reviewer Participated in a PCORI Advisory Panel Other \(please specify\) None of the above "
Public const Position_Title = "\* Position Title "
Public const Project_Lead_Degree = "\* Degree AAS AB APRN BA BC BCH BCHIR BM BMBC BMEDS BPHAR BS BSC BSN CHB CHM DA DBA DC DCH DDS DES DM DMD DMH DMS DNSC DO DPH DPHIL DMP DSC DVM EDD HS JD LPN MA MB MBA MBBC MBCHB MCHIR MD MED MIS MLIS MN MPA MPH MPHIL MS MSN MSURG MSW ND NN NP PA PHD PHRMD PTA RN SB SCD Other \(please specify\) "
Public const Project_Lead_Degree_Other = "Please describe ""Other"" degree"
Public const Relevant_Exp_Terminal_Degree = "\* How many years of research experience do you have after attaining your terminal degree\? --None-- 0-4 years 5-9 years 10\+ years "
Public const Relevant_experience = "\* How many years of research experience do you have related to this field\? --None-- 0–2 years 3–5 years 6–9 years 10–15 years 16 years \+ "
Public const Grants_Funded_as_PI = "\* As the PI or project lead, approximately how many grants or contracts have you had funded\? --None-- 0 1-5 6-10 11-15 16-20 21-25 26 or greater "
Public const Contract_Fund = "\* Total dollar amount \(direct cost\) for largest grants or contract for which you were the PI --None-- N/A Less than \$500,000 \$500,000 - 1 million \$1\.1 - 5 million \$5\.1 - 10 million Greater than \$10 million "
Public const Previous_Contracts = "\* Have you received grants or contracts from: \(Choose all that apply\) PCORI AHRQ CDC NIH RWJF Other \(please specify\) None of the above "

'''''''''''*******PI Information Tab Methods''''''''''''''''''''
Public const Grants_Funded_as_PI_Methods = "\* As the PI or project lead, approximately how many grants or contracts have you had funded\? --None-- 0 1-5 6-10 11-15 16-20 21-25 26 or greater "

'''''''''''*******PI Information Tab Placer''''''''''''''''''''
Public const Grants_Funded_as_PI_PLACER = "\* As the PI or project lead, approximately how many grants or contracts have you had funded --None-- 0 1-5 6-10 11-15 16-20 21-25 26 or greater "

'''''''''''Error on PI Information Tab'''''''''''''
Public const PIinfo_Tab_err1 = "Warning This is a required field\. PI Work Telephone Number"
Public const PIinfo_Tab_err2 = "Warning This is a required field\. For the purpose of this project, with which group does the PI or project lead identify"
Public const PIinfo_Tab_err3 = "Warning This is a required field\. Have you interacted with PCORI in the past in the following ways\? \(Select all that apply\)"
Public const PIinfo_Tab_err4 = "Warning This is a required field\. Position Title"
Public const PIinfo_Tab_err5 = "Warning This is a required field\. Degree"
Public const PIinfo_Tab_err6 = "Warning This is a required field\. How many years of research experience do you have after attaining your terminal degree\?"
Public const PIinfo_Tab_err7 = "Warning This is a required field\. How many years of research experience do you have related to this field\?"
Public const PIinfo_Tab_err8 = "Warning This is a required field\. As the PI or project lead, approximately how many grants or contracts have you had funded"
Public const PIinfo_Tab_err9 = "Warning This is a required field\. Total dollar amount \(direct costs\) for largest grants/contract for which you were the PI"
Public const PIinfo_Tab_err10 = "Warning This is a required field\. Have you received grants or contracts from: \(Choose all that apply\)"

'''''''''PI Information Tab SOE Application'''''''''
Public const Previous_Contracts_SOE_App = "\* Have you received grants or contracts from: \(Choose all that apply\) PCORI CDC NIH RWJF Other \(please specify\) None of the above AHRQ "


'''''''''''''''''''''''''PI Information DI''''''''''''''''''''''''''''
Public const Primary_Group_Identification_DI_LOI = "\* For the purpose of this project, with which group does the PI or project lead identify primarily\? --None-- Patient/Consumer Caregiver/Family member of patient Patient/Caregiver Advocacy Organization Clinician Clinic/Hospital/Health System Purchaser Payer Industry Research Policy Maker Training Institution "
Public const Relevant_experience_DI_LOI = "\* How many years of relevant experience do you have related to this field\? --None-- 0–2 years 3–5 years 6–9 years 10–15 years 16 years \+ "
Public const Grants_Funded_as_PI_DI_LOI = "\* As the PI or project lead, approximately how many grants or contracts have you had funded\? --None-- 0 1-5 6-10 11-15 16-20 21-25 26 or greater "
Public const Contract_Fund_DI = "\* Total dollar amount \(direct cost\) for largest grant/contract for which you were the PI --None-- N/A Less than \$500,000 \$500,000 - 1 million \$1\.1 - 5 million \$5\.1 - 10 million Greater than \$10 million "

'''''''''''''''''''''''''PI Information HSII'''''''''''''''''''''''''''
Public const Contact_PI_listed_HSII_MFA = "\* Is the Principal Investigator \(contact\) a Project Lead for the HSII Participant organization \(i\.e\., listed as a Project Lead in the HSII Master Funding Agreement\)\? \(Note: the contact PI is not required to be an HSII Project Lead\.\) --None-- Yes No "
Public const Relevant_experience_HSII = "\* How many years of relevant healthcare experience do you have\? --None-- 0–2 years 3–5 years 6–9 years 10–15 years 16 years \+ "
Public const Grants_Funded_as_PI_HSII = "\* As the PI or project lead, approximately how many grants/contracts have you had funded\? --None-- 0 1-5 6-10 11-15 16-20 21-25 26 or greater "

'''''''''''''''''''''''''PI Information Managing Pain''''''''''''''''''''''''''
Public const Contract_Fund_ManagingPain_LOI = "\* Total dollar amount \(direct costs\) for largest grants/contract for which you were the PI --None-- N/A Less than \$500,000 \$500,000 - 1 million \$1\.1 - 5 million \$5\.1 - 10 million Greater than \$10 million "


''''''''''''''PI Information tab DI Application'''''''''''''
Public const PI_Tab_InstQs_1_DI = "The following questions pertain to the PI:"
Public const Primary_Group_IdentificationDI = "\* For the purpose of this project, with which group does the PI or project lead identify primarily\? --None-- Industry Research Caregiver/Family member of patient Clinic/Hospital/Health System Clinician Patient/Caregiver Advocacy Organization Patient/Consumer Payer Policy Maker Purchaser Training Institution "
Public const Previous_involvement_PCORI_DI = "Have you interacted with PCORI in the past in the following ways\? \(Select all that apply\)"
Public const Degree_DI = "Degree"
Public const Degree_Other_DI = "Please describe ""Other"" degree "
Public const Relevant_Exp_Terminal_Degree_DI = "\* How many years of research experience do you have after attaining your terminal degree\? --None-- 0-4 years 5-9 years 10\+ years "
Public const Years_of_relevant_experience_DI = "\* How many years of relevant experience do you have related to this field\? --None-- 0–2 years 3–5 years 6–9 years 10–15 years 16 years \+ "
Public const Grants_Funded_as_PI_DI = "\* As the PI or project lead, approximately how many grant or contracts have you had funded\? --None-- 0 1-5 6-10 11-15 16-20 21-25 26 or greater "
Public const Previous_Grants_Contracts_DI = "Have you received grant or contracts from: \(Choose all that apply\)"
Public const HtmlID_for_PI_Work_Telephone_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:2:inputFieldId"
Public const HtmlID_for_Primary_Group_Identification_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:3:inputFieldId"
Public const HtmlID_for_Previous_involvement_PCORI_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:4:inputFieldId_selected"
Public const HtmlID_for_Previous_involvement_PCORI_Other_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:5:inputFieldId"
Public const HtmlID_for_Position_Title_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:6:inputFieldId"
Public const HtmlID_for_Degree_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:7:inputFieldId_selected"
Public const HtmlID_for_Degree_Other_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:8:inputFieldId"
Public const HtmlID_for_Relevant_Exp_Terminal_Degree_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:9:inputFieldId"
Public const HtmlID_for_years_of_relevant_experience_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:10:inputFieldId"
Public const HtmlID_for_Grants_Funded_as_PI_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:11:inputFieldId"
Public const HtmlID_for_Largest_Previous_Grant_Fund_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:12:inputFieldId"
Public const HtmlID_for_Previous_Grants_Contracts_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:13:inputFieldId_selected"
Public const HtmlID_for_Other_Organization_specify_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:14:inputFieldId"

''''''''''*******PI Information Tab ''''''''Application '''''''' Broad''''''''''''''''''''

Public const Previous_Contracts_App_BPS = "\* Have you received grants or contracts from: \(Choose all that apply\) PCORI CDC NIH RWJF Other \(please specify\) None of the above AHRQ "
Public const Previous_involvement_PCORI_App_BPS = "\* Have you interacted with PCORI in the past in the following ways\? \(Select all that apply\) Joined a PCORI email list Participated in applicant training Watched a PCORI webinar Attended PCORI sponsored event in-person Attended event where PCORI was featured Met with PCORI staff Met with a PCORI Ambassador Applied to review PCORI funding app Applied for PCORI funding Received PCORI funding Served as a PCORI Merit reviewer Participated in a PCORI Advisory Panel Other \(please specify\) None of the above Visited PCORI’s website "
Public const Degree_App_BPS = "\* Degree AAS AB APRN BA BC BCHIR BM BMBC BMEDS BPHAR BS BSC BSN CHB CHM DA DBA DC DCH DDS DES DM DMD DMH DMP DMS DNSC DO DPH DPHIL DSC DVM EDD HS JD LPN MA MB MBA MBBC MBCHB MCHIR MD MED MIS MLIS MN MPA MPH MPHIL MS MSN MSURG MSW ND NN NP PA PHD PHRMD PTA RN SB SCD Other \(please specify\) BCH "

''''''''''''''''''''''''''''**********Project Information Tab Broad''''''''''''''''''''''

Public const Project_Name = "\* Project Title "
Public const Rare_Disease_Focus = "\* Is the primary focus of your study on a rare disease\? --None-- Yes No "
Public const BPS_National_Priority_Instext = "For the following questions that pertain to PCORI's National Priorities, please review the goal of each National Priority here\."
Public const BPS_National_Priority_URL = "https://www\.pcori\.org/about/about-pcori/pcori-strategic-plan/pcori-strategic-plan-national-priorities-health"
Public const National_Priorities_primary_BPS = "\* Please identify the primary PCORI National Priority that pertains to your proposal \(Required\) --None-- Increase Evidence for Existing Interventions and Emerging Innovations in Health Accelerate Progress Toward an Integrated Learning Health System Achieve Health Equity Advance the Science of Dissemination, Implementation, and Health Communication "
Public const National_Priorities_secondary_BPS = "Please identify the secondary PCORI National Priority that pertains to your proposal \(If applicable\) --None-- Increase Evidence for Existing Interventions and Emerging Innovations in Health Accelerate Progress Toward an Integrated Learning Health System Achieve Health Equity Advance the Science of Dissemination, Implementation, and Health Communication "
Public const National_Priorities_tertiary_BPS = "Please identify the tertiary PCORI National Priority that pertains to your proposal \(If applicable\) --None-- Increase Evidence for Existing Interventions and Emerging Innovations in Health Accelerate Progress Toward an Integrated Learning Health System Achieve Health Equity Advance the Science of Dissemination, Implementation, and Health Communication "
Public const Topic_Themes_InsText = "For the following questions that pertain to PCORI's Topic Themes, please review the text of each topic here\."
Public const Topic_Themes_URL = "https://www\.pcori\.org/funding-opportunities/what-who-we-fund/research-project-agenda-topic-themes-inform-focused-funding-opportunities/2025-research-project-agenda-topic-themes"
Public const Topic_Themes_Primary = "\* Please identify the primary topic theme that pertains to your proposal \(required\) --None-- Improving outcomes for people with intellectual or developmental disabilities \(IDD\) Addressing substance use Addressing violence and trauma Addressing COVID-19 Addressing rare diseases Improving cardiovascular health Improving mental and behavioral health Managing pain Preventing maternal morbidity and mortality \(MMM\) Promoting sleep health Metabolic and Endocrine Health Cancer Sensory Impairment and Disability Other "
Public const Topic_Themes_Secondary = "Please identify the secondary topic theme that pertains to your proposal \(if applicable\) --None-- Improving outcomes for people with intellectual or developmental disabilities \(IDD\) Addressing substance use Addressing violence and trauma Addressing COVID-19 Addressing rare diseases Improving cardiovascular health Improving mental and behavioral health Managing pain Preventing maternal morbidity and mortality \(MMM\) Promoting sleep health Metabolic and Endocrine Health Cancer Sensory Impairment and Disability Other "
Public const Topic_Themes_Tertiary = "Please identify the tertiary topic theme that pertains to your proposal \(if applicable\) --None-- Improving outcomes for people with intellectual or developmental disabilities \(IDD\) Addressing substance use Addressing violence and trauma Addressing COVID-19 Addressing rare diseases Improving cardiovascular health Improving mental and behavioral health Managing pain Preventing maternal morbidity and mortality \(MMM\) Promoting sleep health Metabolic and Endocrine Health Cancer Sensory Impairment and Disability Other "
Public const BPS_Categories = "\* Does your LOI pertain to a Category 1, Category 2, or Category 3\? --None-- Category 1 \(≤5 million direct costs\) Category 2 \(>5 million direct costs\) Category 3 \(PCORnet® Study\) "
Public const Total_Direct_Costs = "\* Total direct cost "
Public const Total_Indirect_Costs = "\* Total indirect cost "
Public const LOI_Amount_Requested_from_PCORI = "\* Total amount requested "
Public const Patient_Care_Costs = "\* Are you requesting PCORI coverage of patient care costs\? --None-- Yes No "
Public const Total_Patient_Care_Costs = "Total amount of patient care costs requested in LOI"
Public const Project_Duration = "\* Please select your estimated project length \(in months\) --None-- 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73 74 75 76 77 78 "
Public const ProjIns1 = "Which disease or condition is the primary focus of your proposal\? \(Select the best option\)"
Public const Primary_Disease_Condition = "\* Primary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down list for the field below, click “Save” at the bottom of the page to respond to the additional questions\. "
Public const Primary_Disease_Condition_Focus = "\* Primary disease or condition focus --None-- "
Public const Primary_Disease_Condition_Other = "Please describe ""Other"" primary disease or condition 0 of 255 Characters "
Public const ProjIns2 = "Which disease or condition is the secondary focus of your proposal\? \(Select the best option\)"
Public const Secondary_Disease_Condition = "\* Secondary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down list for the field below, click “Save” at the bottom of the page to respond to the additional questions\. "
Public const Secondary_Disease_Focus =  "\* Secondary disease or condition focus --None-- "
Public const Secondary_Disease_Other = "Please describe ""Other"" secondary disease or condition 0 of 255 Characters "
Public const Population_Focus = "\* Does your proposal focus on any of the following populations\? \(Select all that apply\) N/A - this proposal does not focus on a population Children 0-12 Children 13-18 Children 18-21 Adults 21-64 Adults >65 Disabled persons Racial or ethnic minorities: Residents of rural areas Residents of urban areas Veterans Women LGBTQ Low income groups Patients with low health literacy/numeracy and/or limited English proficiency Individuals with multiple chronic conditions Individuals with rare or genetic disease Other \(please specify\) "
Public const Population_Focus_Other = "Please describe ""Other"" population focus 0 of 255 Characters "
Public const Racial_Minority_Focus = "\* Racial or ethnic minorities American Indian or Alaska Native Asian Black or African American Hispanic/Latino Native Hawaiian or Pacific Islander White Two or more races Other \(please specify\) "
Public const Racial_Minority_Other =  "Please describe ""Other"" racial or ethnic minority focus "
Public const Healthcare_Primary_Focus = "\* Which of the following healthcare topics is a primary focus of your proposal\? \(Select the best choice\) --None-- Complementary and Alternative Medicine: Other Complementary and Alternative Medicine: Mindfulness Complementary and Alternative Medicine: Acupuncture Complementary and Alternative Medicine: Supplements Chronic Disease Management: Other Chronic Disease Management: Patient Adherence Chronic Disease Management: Patient Activation Chronic Disease Management: Provider Adherence Chronic Disease Management: Functional Impairments Chronic Disease Management: Access and Quality Chronic Disease Management: Medical Treatment Disparities: Other Disparities: Ethnic and Racial Disparities: Literacy Disparities: Sociological Disparities: Economic Disparities: Demographic Health Care Delivery Systems: Other Health Care Delivery Systems: Care Coordination Health Care Delivery Systems: Patient Participation Health Care Delivery Systems: Primary Care \(Providers\) Health Care Delivery Systems: Quality Improvement Health Care Delivery Systems: Telemedicine Health Care Delivery Systems: Cost Containment Health Care Delivery Systems: Accountable Care Organizations Health Care Delivery Systems: Transitions of Care Health Care Delivery Systems: Integrated Care Medication and Medication Management: Other Medication and Medication Management: Palliative Care Medication and Medication Management: Adherence Medication and Medication Management: Delivery Models Medication and Medication Management: Medication Non- specific Medical Technology: Other Medical Technology: Robotic Tools Medical Technology: Imaging Health Information Technology: Other Health Information Technology: Website Health Information Technology: Patient Portal Health Information Technology: Applications None of the Above: Other Other \(please specify\) "
Public const Healthcare_Primary_Other = "Please describe ""Other"" primary focus healthcare topic 0 of 255 Characters "
Public const Healthcare_Secondary_Focus = "\* Which of the following healthcare topics is a secondary focus of your proposal\? \(Select the best choice\) --None-- Complementary and Alternative Medicine: Other Complementary and Alternative Medicine: Mindfulness Complementary and Alternative Medicine: Acupuncture Complementary and Alternative Medicine: Supplements Chronic Disease Management: Other Chronic Disease Management: Patient Adherence Chronic Disease Management: Patient Activation Chronic Disease Management: Provider Adherence Chronic Disease Management: Functional Impairments Chronic Disease Management: Access and Quality Chronic Disease Management: Medical Treatment Disparities: Other Disparities: Ethnic and Racial Disparities: Literacy Disparities: Sociological Disparities: Economic Disparities: Demographic Health Care Delivery Systems: Other Health Care Delivery Systems: Care Coordination Health Care Delivery Systems: Patient Participation Health Care Delivery Systems: Primary Care \(Providers\) Health Care Delivery Systems: Quality Improvement Health Care Delivery Systems: Telemedicine Health Care Delivery Systems: Cost Containment Health Care Delivery Systems: Accountable Care Organizations Health Care Delivery Systems: Transitions of Care Health Care Delivery Systems: Integrated Care Medication and Medication Management: Other Medication and Medication Management: Palliative Care Medication and Medication Management: Adherence Medication and Medication Management: Delivery Models Medication and Medication Management: Medication Non- specific Medical Technology: Other Medical Technology: Robotic Tools Medical Technology: Imaging Health Information Technology: Other Health Information Technology: Website Health Information Technology: Patient Portal Health Information Technology: Applications None of the Above: Other Other \(please specify\) "
Public const Healthcare_Secondary_Other = "Please describe ""Other"" secondary focus healthcare topic 0 of 255 Characters "
Public const Sample_Size = "\* Targeted sample size for main analysis "
Public const PCORnet_involvement = "\* Does your study proposal include use of a PCORnet® Clinical Research Network or the Coordinating Center for PCORnet\? --None-- Yes No "
Public const PCORnet_Front_Door = "\* Have you contacted the PCORnet® Front Door \(www\.pcornet\.org/front-door\) about this proposal\? \(Note: Required for Category 3 BPS\) --None-- Yes No N/A "
Public const PCORnet_Front_Door_Number = "If you contacted the PCORnet® Front Door about this proposal, please provide the Front Door number the PCORnet® Front Door team gave you\. "
Public const PCORnet_ID_Network = "\* For Category 3: Please identify at least one PCORnet® Clinical Research Network that will submit a Letter of Support on your behalf if you are invited to submit a full application\. ADVANCE GPC INSIGHT OneFlorida PaTH PEDSnet REACHnet STAR Coordinating Center for PCORnet N/A "
Public const Address_SAE = "\* Does your study proposal address a PCORI Special Area of Emphasis \(SAE\)\? --None-- Yes No "
Public const Which_SAE_BPS = "If Yes, which SAE does your study address\? --None-- PTSD Down Syndrome PAD Functional Sensory Impairment "
Public const PCORI_Research_Area_Topic = "\* Does your study proposal address a PCORI Research Priority Area \(IDD/MMM/COVID-19/Rare Disease\)\? 1\) Intellectual and Developmental Disabilities \(IDD\) 2\) Maternal Morbidity and Mortality \(MMM\) 3\) COVID-19 4\) Rare Disease 5\) N/A "
Public const Structured_Mentorship_Activities_BPS_New_field_LOI = "\* Do you intend to request funding for structured mentorship activities\? --None-- Yes No "
Public const Cross_Cutting_Population_Prior_Topic_Theme = "\* Does your study proposal address a cross-cutting population of interest and/or prior PCORI Topic Theme\? Promoting Health for Older Adults Promoting Healthy Children and Youth N/A " 



'''''''''''''''''''Project Information tab Methods'''''''''''''''''''
Public const Methods_Research_Areas_LOI = "\* Please select the research area\(s\) of interest listed in the PFA to which your application is responding \(Select all that apply\) Methods Related to Ethics & Human Subjects Protections Methods to Improve Study Design Methods to Support Data Research Networks Methods to Improve the Use of Artificial Intelligence \(AI\) and Machine Learning \(ML\) Methods - Other "
Public const Methods_Project_Info_Ins_Text_3 = "What methods will be used in conducting the research proposed in the application\? Please identify at least three primary data collection methods \(as applicable\) and/or data analytic methods \(as applicable\) and/or data analytic methods\. \(Examples of data collecting methods: surveys, Delphi technique, focus groups\. Examples of analytic methods: inductive thematic analysis, descriptive statistics, multi-level modeling, propensity scores, Bayesian techniques, machine learning\.\)"
Public const Methods_Used_One_LOI = "\* Methods used- One 0 of 255 Characters "
Public const Methods_Used_Two_LOI =  "\* Methods used- Two 0 of 255 Characters "
Public const Methods_Used_Three_LOI = "\* Methods used- Three 0 of 255 Characters "
Public const Methods_Used_Four_LOI = "Methods used- Four 0 of 255 Characters "
Public const Methods_Used_Five_LOI = "Methods used- Five 0 of 255 Characters "
Public const Methods_Used_Six_LOI = "Methods used- Six 0 of 255 Characters "
Public const Methods_Project_Info_Ins_Text_4 = "What method\(s\) does the application aim to advance\? \(Please identify at least one method\) "
Public const Methods_Advanced_One_LOI = "\* Methods advanced - One 0 of 255 Characters "
Public const Methods_Advanced_Two_LOI = "Methods advanced - Two 0 of 255 Characters "
Public const Methods_Advanced_Three_LOI = "Methods advanced - Three 0 of 255 Characters "
Public const Methods_Advanced_Four_LOI = "Methods advanced - Four 0 of 255 Characters "
Public const PCORnet_Affiliation_LOI = "\* Which PCORnet Network Partners will participate in this study: ADVANCE GPC INSIGHT OneFlorida PaTH PEDSnet REACHnet STAR Coordinating Center for PCORnet N/A "
Public const PCORnet_Network_Partners_LOI = "\* Have you engaged with any of these PCORnet Network Partners\? ADVANCE GPC INSIGHT OneFlorida PaTH PEDSnet REACHnet STAR Coordinating Center for PCORnet \(including Front Door\) N/A: Have not engaged "
Public const Primary_Disease_Condition_Methods_LOI = "Primary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down list for the field below, click ""Save"" at the bottom of the page to respond to the additional questions\. "
Public const Primary_Disease_Condition_Focus_Methods_LOI = "Primary disease or condition focus --None-- "
Public const Secondary_Disease_Condition_Methods_LOI = "Secondary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down list for the field below, click ""Save"" at the bottom of the page to respond to the additional questions\. "
Public const Secondary_Disease_Condition_Focus_Methods_LOI = "Secondary disease or condition focus --None-- "
Public const PCORnet_Front_Door_Number_Methods_LOI = "If you contacted the PCORnet® Front Door about this proposal, please provide the Front Door number you were given by the PCORnet® Front Door team\. "
Public const Primary_Disease_Condition_Focus_htmlId_LOI_Methods = "j_id0:j_id2:j_id3:mainForm:j_id713:13:inputFieldId"
Public const Secondary_Disease_Condition_Focus_htmlId_LOI_Methods = "j_id0:j_id2:j_id3:mainForm:j_id713:17:inputFieldId"
Public const Project_Duration_Methods = "\* Please select your estimated project length \(in months\) --None-- 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 "

''''''''''''''''''''''Project Information Tab SOE""""""""
Public const National_Priorities_primary_SoE = "Please identify the primary PCORI National Priority that pertains to your proposal \(If applicable\) --None-- Increase Evidence for Existing Interventions and Emerging Innovations in Health Accelerate Progress Toward an Integrated Learning Health System Achieve Health Equity Advance the Science of Dissemination, Implementation, and Health Communication " 
Public const Topic_Themes_Primary_SoE = "Please identify the primary topic theme that pertains to your proposal \(If applicable\) --None-- Improving outcomes for people with intellectual or developmental disabilities \(IDD\) Promoting health for older adults Promoting healthy children and youth Addressing substance use Addressing violence and trauma Addressing COVID-19 Addressing rare diseases Improving cardiovascular health Improving mental and behavioral health Managing pain Preventing maternal morbidity and mortality \(MMM\) Promoting sleep health Other "
Public const Primary_Disease_Condition_SoE = "Primary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down list for the field below, click “Save” at the bottom of the page to respond to the additional questions\. "
Public const Primary_Disease_Condition_Focus_htmlId_LOI_SOE = "j_id0:j_id2:j_id3:mainForm:j_id713:17:inputFieldId"
Public const Secondary_Disease_Condition_SoE = "Secondary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down list for the field below, click “Save” at the bottom of the page to respond to the additional questions\. "
Public const Secondary_Disease_Condition_Focus_htmlId_LOI_SOE = "j_id0:j_id2:j_id3:mainForm:j_id713:21:inputFieldId"
Public const Methods_Used_Two_LOI_SOE =  "Methods used- Two 0 of 255 Characters "
Public const Methods_Used_Three_LOI_SOE = "Methods used- Three 0 of 255 Characters "
Public const Methods_Used_Four_LOI_SOE = "Methods used- Four 0 of 255 Characters "
Public const SoE_Project_Info_Ins_Text_3 = "What methods will be used in conducting the research proposed in the application\? Please identify primary data collection methods \(as applicable\) and/or data analytic methods \(as applicable\)\. \(Examples of data collecting methods: surveys, Delphi technique, focus groups\. Examples of analytic methods: inductive thematic analysis, descriptive statistics, multi-level modeling, propensity scores, Bayesian techniques, machine learning\.\)"
Public const SoE_Research_Areas_LOI = "\* Please select the research area\(s\) of interest listed in the PFA to which your application is responding \(Select all that apply\) --None-- Development and/or assessment of validity of measures to capture structure/context, process, and outcomes of engagement in research Development and/or testing of engagement methods to generate evidence on the most effective engagement approaches, particularly for underrepresented populations, and how effectiveness varies by context "

''''''''''''''''''''''Project Information Tab SLeep Health""""""""
Public const Which_PFA_SleepHealth = "PFA Name Please DO NOT change this field --None-- Broad Pragmatic Studies \(Cycle 1 2025\)Improving Mental and Behavioral Health Topical Managing Pain Topical "

'''''''Public const Which_SAE_SleepHealth = "If Yes, which SAE does your study address\? Note: please ensure to select the SAE that corresponds with the PCORI PFA\. --None-- C3 Sleep: Promoting Sleep Health Equity C3 Sleep: Chronic Conditions Co-occurring with sleep disturbances C3 Sleep: Focus on Sleep Health Beyond Diagnosed Sleep Disorders "
Public const Which_SAE_SleepHealth = "If Yes, which SAE does your study address\? --None-- PTSD Down Syndrome PAD Functional Sensory Impairment "

''''''''''''''''''''''Project Information Tab PLACER""""""""
Public const Total_Direct_Costs_Placer = "\* Anticipated total direct costs \(Stage 1\) "
Public const Total_Indirect_Costs_Placer = "\* Anticipated total indirect costs \(Stage 1\) "
Public const Total_Direct_Costs_Stage_2 = "\* Anticipated total direct costs \(Stage 2\) "
Public const Total_Indirect_Costs_Stage_2 = "\* Anticipated total indirect costs \(Stage 2\) "
Public const LOI_Amount_Requested_from_PCORI_Placer = "\* Total cost estimate \(direct \+ indirect costs in Stage 1 and Stage 2\)\. "
Public const Project_Duration_Stage_1 = "\* Please select your estimated Stage 1 duration \(in months\) needed prior to Stage 2\. --None-- 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 "
Public const Project_Duration_Placer = "\* Please select your estimated study duration of Stage 1 and 2 combined \(in months\)\. --None-- 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73 74 75 76 77 78 "
Public const National_Priorities_primary_PLACER = "\* Please identify the primary PCORI National Priority that pertains to your proposal \(Required\) --None-- Increase Evidence for Existing Interventions and Emerging Innovations in Health Accelerate Progress Toward an Integrated Learning Health System Achieve Health Equity "
Public const Primary_Disease_Condition_Focus_htmlId_LOI_PLACER = "j_id0:j_id2:j_id3:mainForm:j_id713:22:inputFieldId"
Public const Secondary_Disease_Condition_Focus_htmlId_LOI_PLACER = "j_id0:j_id2:j_id3:mainForm:j_id713:26:inputFieldId"
Public const National_Priorities_secondary_PLACER = "Please identify the secondary PCORI National Priority that pertains to your proposal \(If applicable\) --None-- Increase Evidence for Existing Interventions and Emerging Innovations in Health Accelerate Progress Toward an Integrated Learning Health System Achieve Health Equity "
Public const National_Priorities_tertiary_PLACER = "Please identify the tertiary PCORI National Priority that pertains to your proposal \(If applicable\) --None-- Increase Evidence for Existing Interventions and Emerging Innovations in Health Accelerate Progress Toward an Integrated Learning Health System Achieve Health Equity "




''''''''''''''''''''''Project Information Tab Mental Health """"""""
Public const Greaterthan_Duaration_InsText = "Applicants who are requesting PCORI review for an increased study duration \(i\.e\., >5 years ≤7 years must complete the separate addendum “Request for Increased Study Duration” of the LOI template\. "
Public const Greaterthan_Duaration_LOI = "\* Request for Increased Study Duration >5 years and ≤7 years\? --None-- Yes No "
Public const Which_SAE_MentalHealth = "If Yes, which SAE does your study address\? --None-- Urogynecological and pelvic pain Pain in individuals living with limitations in cognitive functioning Pain in individuals living with sickle cell disease Neuropathic pain "

''''''''''''''''''''''Project Information Tab MMM""""""""
Public const PCORnet_involvement_MMM = "\* Does any portion of your study proposal include collaborations with existing PCORnet entities \(including CDRNs, PPRNs, or Collaborative Research Groups\)\? --None-- Yes No "
Public const Which_SAE_MMM = "If Yes, which SAE does your study address\? --None-- MMM: Strategies for High Risk Pregnant People  MMM: Populations experiencing disparities  MMM: Observational CER examining variations "
Public const Which_PFA_MMM = "Which PFA\? --None-- Broad Pragmatic Studies Healthy Aging: Optimizing Physical and Mental Functioning Cardiovascular Health Topical Sleep Health Topical Improving Postpartum Maternal Outcomes for Populations Experiencing Disparities Please do NOT remove this value "

'''''''''''''Project Information Tab DI LOI'''''''''''
Public const Eligible_evidence_DI_LOI = "\* Eligible evidence proposed for implementation 0 of 255 Characters "
Public const Total_Direct_Costs_DI = "\* Total direct cost requested for this proposed project "
Public const Total_Indirect_Costs_DI = "\* Total indirect cost requested for this proposed project "
Public const Amount_Requested_from_PCORI_LC = "\* Total amount requested for this proposed project Please note that Dissemination Projects should not exceed \$300,000 in total project costs and Implementation Planning Projects should not exceed \$200,000\. "
Public const Project_Duration_DI = "\* Please select your estimated project length \(in months\) for this proposed project --None-- 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73 74 75 76 77 78 "

Public const Primary_Disease_Condition_LC_LOI = "\* Primary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down list for the field below, click “Save” at the bottom of the page to respond to the additional questions\. "
Public const Primary_Disease_Condition_LC_PostText_HTML = "<font color=""red"">If you do not see values in the drop-down list for the field below, click “Save” at the bottom of the page to respond to the additional questions\.</font>"
Public const Secondary_Disease_Condition_DI_LOI = "\* Secondary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down list for the field below, click ""Save"" at the bottom of the page to respond to the additional questions\. "

Public const Secondary_Disease_Condition_PostText_HTML = "<font color=""red"">If you do not see values in the drop-down list for the field below, click ""Save"" at the bottom of the page to respond to the additional questions\.</font>"
Public const PCORnet_involvement_DI = "\* Does any portion of your project proposal include collaborations with existing PCORnet entities \(including CDRNs, PPRNs, or Collaborative Research Groups\)\? --None-- Yes No "
Public const Dissemination_multiple_PCORI_projects_LOI = "\* Are you proposing a project involving collaboration to implement the results of multiple related PCORI-funded research studies\? --None-- Yes No If you answered 'Yes' to this question, please answer the following two questions regarding your application\. "
Public const Dissemination_multiple_PCORI_projects_PostText_HTML_LOI = "<font color=""purple"">If you answered 'Yes' to this question, please answer the following two questions regarding your application\.</font>"
Public const Primary_Disease_Condition_Focus_htmlId_LOI_LC = "j_id0:j_id2:j_id3:mainForm:j_id713:13:inputFieldId"
Public const Secondary_Disease_Condition_Focus_htmlId_LOI_LC = "j_id0:j_id2:j_id3:mainForm:j_id713:17:inputFieldId"
Public const Population_Focus_htmlId_LOI_LC = "j_id0:j_id2:j_id3:mainForm:j_id713:19:inputFieldId_unselected"
Public const Racial_Ethnic_Minority_Focus_htmlId_LOI_LC = "j_id0:j_id2:j_id3:mainForm:j_id713:21:inputFieldId_unselected"
Public const Amount_Requested_from_PCORI_PostText_HTML = "<font color=""red"">Please note that Dissemination Projects should not exceed \$300,000 in total project costs and Implementation Planning Projects should not exceed \$200,000\.</font>"
Public const HtmlId_for_Contract_Number_DI_LOI = "j_id0:j_id2:j_id3:mainForm:j_id713:6:inputFieldId"
Public const LC_Project_Info_Ins_Text1 = "Please enter the PCORI Contract ID of the original PCORI funded research award\."
'''''''''''''Prior TI-7735 Update Public const Cycle_DI_LOI = "\* Cycle --None-- Cycle 3 2016 Cycle 2 2015 Cycle 3 2015 Cycle I Cycle II Cycle III August 2013 Cycle Fall 2014 Cycle Inaugural Methods PCORnet Phase I PCORnet Phase II Pilot Projects Spring 2014 Cycle Spring 2015 Cycle Winter 2014 Cycle Winter 2015 Cycle Cycle 3 2017 Cycle 1 2017 Cycle 2 2018 "
Public const Cycle_DI_LOI = "\* Cycle "
Public const Contract_Number_DI_LOI = "\* Contract Number "

'''''''''''''''''''''Project Information tab IMRI LOI''''''''''''''''
Public const Amount_Requested_from_PCORI = "\* Total amount requested for this proposed project "
Public const Primary_Disease_Condition_DI_LOI = "\* Primary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down list for the field below, click ""Save"" at the bottom of the page to respond to the additional questions\. "
Public const Primary_Disease_Condition_PostText_HTML = "<font color=""red"">If you do not see values in the drop-down list for the field below, click ""Save"" at the bottom of the page to respond to the additional questions\.</font>"
Public const Primary_Disease_Condition_Focus_htmlId_LOI_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:9:inputFieldId"
Public const Secondary_Disease_Condition_Focus_htmlId_LOI_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:13:inputFieldId"
Public const Population_Focus_htmlId_LOI_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:15:inputFieldId_unselected"
Public const Racial_Ethnic_Minority_Focus_htmlId_LOI_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:17:inputFieldId_unselected"
Public const Total_Direct_Costs_IMRI = "\* Total direct cost requested for this proposed implementation project "
Public const Total_Indirect_Costs_IMRI = "\* Total indirect cost requested for this proposed implementation project "
Public const Project_Duration_IMRI = "\* Please select your estimated project length \(in months\) for this proposed implementation project --None-- 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73 74 75 76 77 78 "

''''''''''''''''''''''''Project Information SDM LOI ''''''''''''''
Public const Primary_Disease_Condition_Focus_htmlId_LOI_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:11:inputFieldId"
Public const Secondary_Disease_Condition_Focus_htmlId_LOI_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:15:inputFieldId"
Public const Population_Focus_htmlId_LOI_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:17:inputFieldId_unselected"
Public const Racial_Ethnic_Minority_Focus_htmlId_LOI_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:19:inputFieldId_unselected"
Public const Total_Direct_Costs_SDM = "\* Total direct cost requested for this proposed SDM implementation project "
Public const Total_Indirect_Costs_SDM= "\* Total indirect cost requested for this proposed SDM implementation project "
Public const ProjectInfo_SDM_Instext1 = "Please enter the title and name of the Principal Investigator of the original PCORI funded research award\. This information is available on the PCORI website\."
Public const Original_funded_title_SDM_LOI = "\* Title of original PCORI funded research award "
Public const Original_PI_SDM_LOI = "\* Principal Investigator of original PCORI funded research award "
Public const Project_Duration_SDM = "\* Please select your estimated project length \(in months\) for this proposed implementation project --None-- 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73 74 75 76 77 78 "
Public const Primary_Disease_Condition_SDM = "\* Primary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down for the field below, click “Save” at the bottom of the page to respond to the additional questions\. "
Public const Primary_Disease_Condition_SDM_PostText_HTML = "<font color=""red"">If you do not see values in the drop-down for the field below, click “Save” at the bottom of the page to respond to the additional questions\.</font>"
Public const Secondary_Disease_Condition_SDM_LOI = "\* Secondary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) If you do not see values in the drop-down for the field below, click “Save” at the bottom of the page to respond to the additional questions\. "
Public const PCORnet_involvement_SDM = "\* Does any portion of your project proposal include collaborations with existing PCORnet entities \(including CRNs, HPRNs, and/or the Coordinating Center\)\? --None-- Yes No "

''''''''''''''''''''''''HSII LOI ''''''''''''''
Public const HSII_Evidence_Topic = "\* Evidence topic selected for this proposed project \(select one\) --None-- Topic 1: Intensive Lifestyle Treatment for Weight Loss in Primary Care Settings Topic 2: Appropriate Antibiotic Prescribing for Children with ARTIs Topic 3: Addressing Hypertension in Adults \(A - Uncontrolled HTN\) Topic 3: Addressing Hypertension in Adults \(B - Undiagnosed HTN\) Topic 3: Addressing Hypertension in Adults \(A and B\) Topic 4: Monitoring Electronic Patient-Reported Outcomes During Cancer Treatment "
Public const HSII_Project_Type = "\* Proposed project type \(select one\) --None-- HSII Implementation Project HSII Pilot Project HSII Implementation and Supplemental Pilot \(Topic 3 only\) "
Public const Amount_Requested_from_PCORI_HSII = "\* Total amount requested for this proposed project "
Public const HSII_Relevant_Care_Delivery_Sites = "\* How many care delivery sites does your health system have, relevant to your selected topic \? "
Public const HSII_Participating_Care_Delivery_Sites = "\* How many care delivery sites are included for participation in the proposed project\? "
Public const HSII_Eligible_Patients = "\* What is the number of patients potentially eligible to receive the program/intervention\? \(Insert the Potentially Eligible number you provided in the LOI Template; 'Target Eligible', if relevant\.\) "
Public const HSII_Intended_Patients = "\* What is the number of patients intended to receive the program/intervention\? \(Insert the Intended Reach number you provided in the LOI Template\.\) "
Public const HSII_Subcontractors = "\* Does your proposed project involve subcontractors\? \(Please see the Call for Proposals for guidance regarding use of subcontractors\.\) --None-- Yes No "
Public const HSII_Subcontractor_List = "If yes, please list the subcontractors and describe the rationale for their involvement in the proposed project\. 0 of 255 Characters "


'''''''''''''''''''''''''Project Information tab DI Application'''''''''''''''''''''''''''''''''''''''
Public const Project_info_Tab_Ins1_DI = "Please note that the information below was prepopulated based on the information previously submitted during the LOI process and should be updated at this time\. Several new questions now appear; please complete these fields before submitting your full application\."
Public const Dissemination_multiple_PCORI_projects = "\* Are you proposing a project involving collaboration to implement the results of multiple related PCORI-funded research studies\? --None-- Yes No If applicant answered 'Yes' to this question, please answer the following two questions regarding your application\. "
Public const Dissemination_multiple_PCORI_projects_Post_text = "<font color=""purple"">If applicant answered 'Yes' to this question, please answer the following two questions regarding your application\.</font>"
Public const Name_of_Pis = "Names of PIs from other PCORI-funded research studies in collaboration on this proposed project "
Public const Contract_Ids = "PCORI Contract IDs of other PCORI-funded research studies whose results you are proposing to implement with this proposed project "
Public const Dissemination_focused = "\* Are you proposing a dissemination-focused project\? --None-- Yes No "
Public const Contract_Start_Date = "\* Projected Start Date "
Public const Contract_End_Date ="\* Projected End Date "
Public const Technical_Abstract = "\* Project Summary Please copy and paste the text from the body of your Project Summary Template\. As a reminder, you must also upload your Project Summary to the Templates & Uploads tab of this application\.                   "
Public const Original_app_PI = "\* What was the name of the PI on the original PCORI funded research award\? "
Public const Eligible_evidenceDI = "\* Eligible evidence proposed for implementation 52 of 255 Characters "
Public const Project_info_Tab_Ins2_DI = "Please enter the PCORI Contract ID of the original PCORI funded research award\. \(Select Cycle from the dropdown list then select the Contract Number\)"
Public const Cycle_DI = "\* Cycle --None-- August 2013 Cycle Cycle 2 2015 Cycle 3 2015 Cycle I Cycle II Cycle III Fall 2014 Cycle Inaugural Methods PCORnet Phase I PCORnet Phase II Pilot Projects Spring 2014 Cycle Spring 2015 Cycle Winter 2014 Cycle Winter 2015 Cycle Cycle 1 2017 Cycle 2 2017 2016: Cycle 3 Cycle 3 2017 Cycle 1 2018 2016: Cycle 2 2015: Cycle 3 Off-Cycle 17C2 Cycle 2 2018 Cycle 3 2016 "
Public const Original_PFA = "\* Original PFA "
Public const Contract_Number_DI = "Contract Number"
Public const Primary_Disease_Condition_DI = "\* Primary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) "
Public const Primary_Disease_Condition_Focus_DI = "\* Primary disease or condition focus --None-- Other Allergies Immune Disorders Lupus "
Public const Primary_Disease_Condition_Other_DI = "Please describe ""Other"" primary disease or condition 65 of 255 Characters "
Public const Secondary_Disease_Condition_DI = "\* Secondary disease or condition --None-- Allergies and Immune Disorders Birth and Developmental Disorders Blood Disorders Cancer Cardiovascular Diseases Dental Care Ear, Nose and Throat Diseases Eye Diseases Functional Limitations and Disabilities Gastrointestinal Disorders Genetic Disorders and Rare Diseases Infectious Diseases Kidney Disease Liver Disease Mental/Behavioral Health Multiple/co-morbid Chronic Conditions Muscular and Skeletal Disorders Neurological Disorders Nutritional and Metabolic Disorders Rare Diseases Reproductive and Perinatal Health Respiratory Diseases Skin Diseases Systemic Disease Toxin Trauma/Injury Urinary Disorders Wellness None of the Above Other \(please specify\) "
Public const Secondary_Disease_Condition_Focus_DI = "=\* Secondary disease or condition focus --None-- Other Allergies Immune Disorders Lupus "
Public const Estimated_Project_Duration_DI = "\* Please select your estimated project length \(in months\) for this proposed project --None-- 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73 74 75 76 77 78 "
Public const Population_Focus_DI = "Does your proposal focus on any of the following populations\? \(Select all that apply\)"
Public const Population_Focus_Other_DI = "Please describe ""Other"" population focus 53 of 255 Characters "
Public const Healthcare_Primary_Other_DI = "Please describe ""Other"" primary focus healthcare topic 67 of 255 Characters "
Public const Healthcare_Secondary_Other_DI = "Please describe ""Other"" secondary focus healthcare topic 69 of 255 Characters "
Public const Racial_Ethnic_Minority_Focus_DI = "Racial or ethnic minorities"
Public const HtmlId_for_Dissemination_multiple_PCORI_projects = "j_id0:j_id2:j_id3:mainForm:j_id713:3:inputFieldId"
Public const HtmlId_for_Dissemination_focused = "j_id0:j_id2:j_id3:mainForm:j_id713:6:inputFieldId"
Public const HtmlId_for_Original_app_PI = "j_id0:j_id2:j_id3:mainForm:j_id713:10:inputFieldId"
Public const HtmlId_for_Cycle_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:12:inputFieldId"
Public const HtmlId_for_Original_PFA = "j_id0:j_id2:j_id3:mainForm:j_id713:13:inputFieldId"
Public const HtmlId_for_Contract_Number_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:14:inputFieldId"
Public const HtmlId_for_Estimated_Project_Duration_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:15:inputFieldId"
Public const HtmlId_for_Estimated_Project_Duration_DI_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:7:inputFieldId"
Public const HtmlId_for_Primary_Disease_Condition_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:17:inputFieldId"
Public const HtmlId_for_Primary_Disease_Condition_DI_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:9:inputFieldId"
Public const HtmlId_for_Primary_Disease_Condition_Focus_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:18:inputFieldId"
Public const HtmlId_for_Primary_Disease_Condition_Focus_DI_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:10:inputFieldId"
Public const HtmlId_for_Secondary_Disease_Condition_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:21:inputFieldId"
Public const HtmlId_for_Secondary_Disease_Condition_DI_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:13:inputFieldId"
Public const HtmlId_for_Secondary_Disease_Condition_Focus_DI = "j_id0:j_id2:j_id3:mainForm:j_id713:22:inputFieldId"
Public const HtmlId_for_Secondary_Disease_Condition_Focus_DI_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:14:inputFieldId"
Public const HtmlId_for_Healthcare_Primary_Focus = "j_id0:j_id2:j_id3:mainForm:j_id713:28:inputFieldId"
Public const HtmlId_for_Healthcare_Primary_Focus_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:20:inputFieldId"
Public const HtmlId_for_Healthcare_Secondary_Focus = "j_id0:j_id2:j_id3:mainForm:j_id713:30:inputFieldId"
Public const HtmlId_for_Healthcare_Secondary_Focus_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:22:inputFieldId"
Public const HtmlId_for_Contract_Start_Date = "j_id0:j_id2:j_id3:mainForm:j_id713:7:inputFieldId"
Public const HtmlId_for_Contract_Start_Date_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:3:inputFieldId"
Public const HtmlId_for_Contract_End_Date = "j_id0:j_id2:j_id3:mainForm:j_id713:8:inputFieldId"
Public const HtmlId_for_Contract_End_Date_IMRI = "j_id0:j_id2:j_id3:mainForm:j_id713:4:inputFieldId"
Public const HtmlId_for_Technical_Abstract = "cke_j_id0:j_id2:j_id3:mainForm:j_id713:9:inputRichTextAreaId"

''''''''''''''''''''***********Error on Project Infromation Tab''''''''''''''''''''''''
Public const Proj_Tab_error1 = "Warning This is a required field\. Project Title"
Public const Proj_Tab_error2 = "Warning This is a required field\. Is the primary focus of your study on a rare disease\?"
Public const Proj_Tab_error3 = "Warning This is a required field\. Please identify the primary PCORI National Priority that pertains to your proposal \(Required\)"
Public const Proj_Tab_error4 = "Warning This is a required field\. Please identify the primary topic theme that pertains to your proposal \(required\)"
Public const Proj_Tab_error5 = "Warning This is a required field\. Does your LOI pertain to a Category 1, Category 2, or Category 3\?"
Public const Proj_Tab_error6 = "Warning This is a required field\. Total direct cost"
Public const Proj_Tab_error7 = "Warning This is a required field\. Total indirect cost"
Public const Proj_Tab_error8 = "Warning This is a required field\. Total amount requested"
Public const Proj_Tab_error9 = "Warning This is a required field\. Are you requesting PCORI coverage of patient care costs\?"
Public const Proj_Tab_error10 = "Warning This is a required field\. Please select your estimated project length \(in months\)"
Public const Proj_Tab_error11 = "Warning This is a required field\. Primary disease or condition"
Public const Proj_Tab_error12 = "Warning This is a required field\. Primary disease or condition focus"
Public const Proj_Tab_error13 = "Warning This is a required field\. Secondary disease or condition"
Public const Proj_Tab_error14 = "Warning This is a required field\. Secondary disease or condition focus"
Public const Proj_Tab_error15 = "Warning This is a required field\. Does your proposal focus on any of the following populations\? \(Select all that apply\)"
Public const Proj_Tab_error16 = "Warning This is a required field\. Racial or ethnic minorities"
Public const Proj_Tab_error17 ="Warning This is a required field\. Which of the following healthcare topics is a primary focus of your proposal\? \(Select the best choice\)"
Public const Proj_Tab_error18 = "Warning This is a required field\. Which of the following healthcare topics is a secondary focus of your proposal\? \(Select the best choice\)"
Public const Proj_Tab_error19 = "Warning This is a required field\. Targeted sample size for main analysis"
Public const Proj_Tab_error20 = "Warning This is a required field\. Does your study proposal include use of a PCORnet® Clinical Research Network or the Coordinating Center for PCORnet\?"
Public const Proj_Tab_error21 = "Warning This is a required field\. Have you contacted the PCORnet® Front Door \(www\.pcornet\.org/front-door\) about this proposal\? \(Note: Required for Category 3 BPS\)"
Public const Proj_Tab_error22 = "Warning This is a required field\. For Category 3: Please identify at least one PCORnet® Clinical Research Network that will submit a Letter of Support on your behalf if you are invited to submit a full application\."
Public const Proj_Tab_error23 = "Warning This is a required field\. Does your study proposal address a PCORI Special Area of Emphasis \(SAE\)\?"
Public const Proj_Tab_error24 = "Warning This is a required field\. Does your study proposal address a PCORI Research Priority Area \(IDD/MMM/COVID-19/Rare Disease\)\?"
Public const Proj_Tab_error25 = "Warning This is a required field\. Projected Start Date"
Public const Proj_Tab_error26 = "Warning This is a required field\. Projected End Date"
Public const Proj_Tab_error27 = "Warning This is a required field\. Project Summary"
Public const Proj_Tab_error28 = "Error: The Projected End Date must be a later date \(greater\) than the Projected Start Date \(Project Information tab\)\. "

'''''''''''''''Project Information Tab SOE Application'''''''''''
Public const Project_info_Tab_Ins1_SOE = "Note that your application’s budget must be within the maximum amounts outlined in the PCORI Funding Announcement\. Applications that exceed these amounts without prior PCORI approval may be deemed administratively non-compliant and removed from consideration\."
Public const Project_info_Tab_Ins1_SOE_outerHtml = "<font color=""red"">Note that your application’s budget must be within the maximum amounts outlined in the PCORI Funding Announcement\. Applications that exceed these amounts without prior PCORI approval may be deemed administratively non-compliant and removed from consideration\.</font>"
Public const Technical_Abstract_SOE = "\* Technical Abstract \(suggested 700-word limit\)"
Public const Public_Abstract_SOE = "\* Public Abstract \(suggested 700-word limit\)"
Public const PCORI_Research_Priority_Area_SOE_App = "\* PCORI Research Priority Area \(IDD/MMM/COVID-19/Rare Disease\) 1\) Intellectual and Developmental Disabilities \(IDD\) 2\) Maternal Morbidity and Mortality \(MMM\) 3\) COVID-19 5\) N/A 4\) Rare Disease "
Public const PCORI_Research_Priority_Area_SOE_App2 = "\* PCORI Research Priority Area \(IDD/MMM/COVID-19/Rare Disease\) 2\) Maternal Morbidity and Mortality \(MMM\) 3\) COVID-19 4\) Rare Disease 5\) N/A 1\) Intellectual and Developmental Disabilities \(IDD\) "
Public const PCORnet_Network_Partners_SOE_App = "\* Have you engaged with any of these PCORnet Network Partners\? GPC INSIGHT OneFlorida PaTH PEDSnet REACHnet STAR Coordinating Center for PCORnet \(including Front Door\) N/A: Have not engaged ADVANCE "
Public const Study_Comparators_SOE_App = "\* Briefly describe the study design and comparators if applicable \(Enter ""N/A"" if not applicable\) 0 of 1000 Characters "
Public const Primary_Outcome_SOE_App = "\* Provide the primary outcome for your proposed study\. \(Identify only one outcome\) 0 of 255 Characters "
Public const Secondary_Outcome_SOE_App = "\* Provide the secondary outcome for your proposed study \(if applicable\) 0 of 255 Characters "
Public const Collecting_cost_data_study_SOE_App = "\* Are you planning on collecting cost data as part of your study\? --None-- Yes No "
Public const Data_Sources_SOE_App = "\* Does your proposed research use any of the following data sources\? \(Select all that apply\) Medicare data Medicaid data Commercially available administrative claims \(SID, HCUP, specific health plan data, etc\.\) Other governmentally funded data sets \(HRS, MEPS, LSOA, etc\.\) Clinical data research networks Patient-Powered research networks Data from disease registries Medical record data Other, please specify none of the above "
Public const Other_Data_Sources_App = "Please describe ""Other"" data sources 0 of 255 Characters "
Public const Measurement_Development_Type_SOEInstext_App = "What methods will be used in conducting the research proposed in the application\? Please identify all primary data collection methods \(as applicable\) and/or data analytic methods\. \(Examples of data collecting methods: surveys, Delphi technique, focus groups\)\. \(Examples of analytic methods: inductive thematic analysis, descriptive statistics, multi-level modeling, propensity scores, Bayesian techniques, machine learning\) "
Public const Measurement_Development_Type_SOE_App = "\* If this is a measurement project, how would you describe the measure development being proposed\? If this is not a measurement project, please select ""Not Applicable\."" --None-- New measure development Previously developed measure being validated for the first time Previously developed and validated measure, being validated in a new population Adaptation and validation of existing measure developed in a related/adjacent field Not Applicable: my project focuses on engagement methods "
Public const Validity_Type_SOE_App = "\* If this is a measurement project, what kind of validity data will be collected\? If this is not a measurement project, please select ""Not Applicable\."" Criterion validity \(e\.g\., concurrent or predictive validity\) Construct validity Not Applicable If you are proposing a measures project but are not including any of the listed types of validity data, your application will be deemed non-responsive to this funding opportunity "
Public const Engagement_Method_Type_SOE_App = "\* If this is an engagement methods project, how would you describe the engagement method being proposed for evaluation\? If this is not an engagement methods project, please select ""Not Applicable"" --None-- Development and evaluation of a new engagement method Evaluation of an existing engagement method Not applicable: this is not an engagement methods project "
Public const Population_Focus_SOE_App = "\* Does your proposal focus on any of the following populations\? \(Select all that apply\) N/A - this proposal does not focus on a population Children 0-12 Children 13-18 Children 18-21 Adults >65 Disabled persons Racial or ethnic minorities Residents of rural areas Residents of urban areas Veterans Women LGBTQ Low income groups Patients with low health literacy/numeracy and/or limited English proficiency Individuals with multiple chronic conditions Individuals with rare or genetic disease Other \(please specify\) Adults 21-64 "
Public const Racial_Ethnic_Minority_Focus_SOE_App = "\* Racial or ethnic minorities American Indian or Alaska Native Black or African American Hispanic/Latino Native Hawaiian or Pacific Islander White Two or more races Other \(please specify\) Asian "
Public const Sample_Size_SOE_App = "\* Targeted sample size for main analysis \(Enter ""0"" if not applicable\) "
Public const SoE_Research_Areas_SOE_App = "\* Please select the research area\(s\) of interest listed in the current Science of Engagements PFA to which your application is responding --None-- Development and/or assessment of validity of measures to capture structure/context, process, and outcomes of engagement in research Development and/or testing of engagement methods to generate evidence on the most effective engagement approaches, particularly for underrepresented populations, and how effectiveness varies by context "
Public const SoE_Research_Areas_SOE_App_Html = "j_id0:j_id2:j_id3:mainForm:j_id713:4:inputFieldId"
Public const PCORnet_involvement_SOE_App_Html = "j_id0:j_id2:j_id3:mainForm:j_id713:16:inputFieldId"
Public const PCORnet_Front_Door_SOE_App_Html = "j_id0:j_id2:j_id3:mainForm:j_id713:19:inputFieldId"
Public const Rare_Disease_Focus_SOE_App_html = "j_id0:j_id2:j_id3:mainForm:j_id713:34:inputFieldId"
Public const Project_Duration_SOE_App_Html = "j_id0:j_id2:j_id3:mainForm:j_id713:35:inputFieldId"
Public const Healthcare_Primary_Focus_SOE_App_Html = "j_id0:j_id2:j_id3:mainForm:j_id713:48:inputFieldId"
Public const Healthcare_Secondary_Focus_SOE_App_Html = "j_id0:j_id2:j_id3:mainForm:j_id713:50:inputFieldId"


'''''''''''''''Project Information Tab Methods Application'''''''''''''
Public const Health_Systems_Factors_Methods_App = "\* If your project is a health systems project, which factors that drive health system change will your proposal test\? \(Select all that apply\) N/A- this proposal is not for a health systems project Personnel \(Additional, re-tasked, or trained to provide intervention\) IT \(New or revamped information technologies\) Incentives \(Changes in reimbursement, out of pocket cost for patients, shared savings agreements, etc\.\) Other \(please specify\) "
Public const Study_Comparators_Methods_App = "\* Name the study comparators \(Enter ""N/A"" if not applicable\) 0 of 1000 Characters "

'''''''''''''''Project Information Tab SDM Application'''''''''''''
Public const Full_Project_Title = "\* Project Title 33 of 5000 Characters "
Public const Project_info_Tab_Ins2_SDM = "Please enter the title and name of the Principal Investigator of the original PCORI funded research award\. This information is available on the PCORI website\."
Public const Primary_Disease_Condition_Focus_SDM = "\* Primary disease or condition focus --None-- Other Allergies Immune Disorders Lupus "
Public const Secondary_Disease_Condition_Focus_SDM = "Please describe ""Other"" secondary disease or condition 51 of 255 Characters "
Public const HtmlId_for_Estimated_Project_Duration_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:12:inputFieldId"
Public const HtmlId_for_Primary_Disease_Condition_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:14:inputFieldId"
Public const HtmlId_for_Primary_Disease_Condition_Focus_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:15:inputFieldId"
Public const HtmlId_for_Secondary_Disease_Condition_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:18:inputFieldId"
Public const HtmlId_for_Secondary_Disease_Condition_Focus_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:19:inputFieldId"
Public const HtmlId_for_Healthcare_Primary_Focus_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:25:inputFieldId"
Public const HtmlId_for_Healthcare_Secondary_Focus_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:27:inputFieldId"
Public const HtmlId_for_Contract_Start_Date_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:6:inputFieldId"
Public const HtmlId_for_Contract_End_Date_SDM = "j_id0:j_id2:j_id3:mainForm:j_id713:7:inputFieldId"

''''''''''''''''''''''Project Information Tab Application Broad"""""""
Public const BPS_Categories_App = "\* Does your application pertain to a Category 1, Category 2, or Category 3\? --None-- Category 1 \(≤5 million direct costs\) Category 2 \(>5 million direct costs\) Category 3 \(PCORnet® Study\) "
Public const National_Priorities_primary_BPS_App = "\* Please identify the primary PCORI National Priority that pertains to your proposal \(Required\) Increase Evidence for Existing Interventions and Emerging Innovations in Health "
Public const National_Priorities_secondary_BPS_App = "Please identify the secondary PCORI National Priority that pertains to your proposal \(If applicable\) Accelerate Progress Toward an Integrated Learning Health System "
Public const National_Priorities_tertiary_BPS_App = "Please identify the tertiary PCORI National Priority that pertains to your proposal \(If applicable\) Achieve Health Equity "
Public const Study_Comparators_BPS_App = "\* Name the study comparators 0 of 1000 Characters "
Public const Requesting_intervention_care_cost = "\* Are you requesting PCORI coverage of intervention/patient care costs\? --None-- Yes No "
Public const Total_direct_intervention_care_cost = "If yes, what are the total direct costs of intervention/patient care requested in application\? "
Public const Population_Focus_BPS_App = "\* Does your proposal focus on any of the following populations\? \(Select all that apply\) N/A - this proposal does not focus on a population Children 0-12 Children 18-21 Adults 21-64 Adults >65 Disabled persons Racial or ethnic minorities Residents of rural areas Residents of urban areas Veterans Women LGBTQ Low income groups Patients with low health literacy/numeracy and/or limited English proficiency Individuals with multiple chronic conditions Individuals with rare or genetic disease Other \(please specify\) Children 13-18 "
Public const Sample_Size_BPS_App = "\* Targeted sample size for main analysis "
Public const Which_PFA_BPS_App = "PFA Name Please DO NOT change this field --None-- Intellectual and Developmental Disabilities Topical Broad Pragmatic Studies \(Cycle 3 2024\) Promoting Healthy Children and Youth Topical "
Public const Structured_Mentorship_Activities_BPS_New_field_App = "\* Do you intend to request funding for structured mentorship activities\? --None-- Yes No Picklist \(Yes/No\) "

''''''''''''''''''''''Project Information Tab Application PLACER""""""
Public const National_Priorities_primary_PLACER_App = "\* Please identify the primary national priority for your proposal \(Required\) Increase Evidence for Existing Interventions and Emerging Innovations in Health "
Public const National_Priorities_secondary_PLACER_App = "Please identify the secondary national priority for your proposal \(if applicable\) Accelerate Progress Toward an Integrated Learning Health System "
Public const National_Priorities_tertiary_PLACER_App = "Please identify the tertiary national priority for your proposal \(if applicable\) Achieve Health Equity "
Public const PLACER_DCC_name = "\* Provide the name of your DCC 0 of 255 Characters "
Public const  PLACER_Feasibility_Stage_Duration = "\* Please select your estimated Feasibility Phase duration \(in months\) needed prior to the Full-Scale Study Phase\. --None-- 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 "
Public const Estimated_Project_Duration_PLACER = "\* Please select your estimated study duration of both the Feasibility and Full-Scale Study Phases combined \(in months\)\. --None-- 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73 74 75 76 77 78 "


''''''''''''''''''''''''''''''Project Personnel Tab Broad'''''''''''''''''''
Public const Pro_Personnel_InText = "For instructions on using our system, click here to access our PCORI Portal User Guide\. To view additional information about the PCORI submission process, click here to view our FAQ page\. Senior or Key Personnel are individuals in addition to the Principal Investigator who contributes to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. At least one key personnel entry is required\. If stakeholder is selected, you may enter ""N/A"" for institution\. PCORI is committed to recognizing the contributions of all of the members of the research team, including patient and stakeholder partners, and makes information about the research projects it funds publicly available through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\. By providing the names of the members of the research team \(including individuals and partnering organizations\), Applicant understands and acknowledges that PCORI may use such information as described above in the event that Applicant is awarded a contract\. Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."
Public const Pro_Personnel_BoldedText = "<strong>Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\.</strong>"
Public const Pro_Personnel_childrecord_Text = "Senior or Key Personnel are individuals in addition to the Principal Investigator who contributes to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. At least one key personnel entry is required\. If stakeholder is selected, you may enter ""N/A"" for institution\."

''''''''''''''''''''''''''''''Project Personnel Tab Methods'''''''''''''''''''
Public const Pro_Personnel_InText_Methods = "For instructions on using our system, click here to access our PCORI Portal User Guide\. To view additional information about the PCORI application process, click here to access our Applicant FAQs page\. Senior or Key Personnel are individuals in addition to the Principal Investigator who contributes to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. At least one key personnel entry is required\. If stakeholder is selected, you may enter ""N/A"" for institution\. PCORI is committed to recognizing the contributions of all of the members of the research team, including patient and stakeholder partners, and makes information about the research projects it funds publicly available through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\. By providing the names of the members of the research team \(including individuals and partnering organizations\), Applicant understands and acknowledges that PCORI may use such information as described above in the event that Applicant is awarded a contract\. Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."

Public const Pro_Personnel_InText_childrecord1_Methods = "Senior or Key Personnel are individuals in addition to the Principal Investigator who contributes to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. At least one key personnel entry is required\. If stakeholder is selected, you may enter ""N/A"" for institution\. "
Public const Pro_Personnel_InText_childrecord2_Methods = "PCORI is committed to recognizing the contributions of all of the members of the research team, including patient and stakeholder partners, and makes information about the research projects it funds publicly available through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\. By providing the names of the members of the research team \(including individuals and partnering organizations\), Applicant understands and acknowledges that PCORI may use such information as described above in the event that Applicant is awarded a contract\. Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."

''''''''''''''''''''''''''''''Project Personnel Tab PLACER'''''''''''''''''''
Public const Pro_Personnel_InText_PLACER = "For instructions on using our system, click here to access our PCORI Portal User Guide\. To view additional information about the PCORI submission process, click here to view our FAQ page\. Senior or Key Personnel are individuals in addition to the Principal Investigator who contributes to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. At least one key personnel entry is required\. If stakeholder is selected, you may enter ""N/A"" for institution\. PCORI is committed to recognizing the contributions of all of the members of the research team, including patient and stakeholder partners, and makes information about the research projects it funds publicly available through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\. By providing the names of the members of the research team \(including individuals and partnering organizations\), Applicant understands and acknowledges that PCORI may use such information as described above in the event that Applicant is awarded a contract\. Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."
Public const Pro_Personnel_childrecord_Text_1_PLACER = "Project Personnel are individuals, in addition to the Principal Investigator, who contribute to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. At least one project personnel entry is required\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\."
Public const Pro_Personnel_childrecord_Text_2_PLACER = "PCORI is committed to recognizing the contributions of all of the members of the research team, including patient and stakeholder partners\. PCORI publicizes information about its funded research projects through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\. By providing the names of the members of the research team \(including individuals and partnering organizations\) below, the applicant understands and acknowledges that PCORI may use such information as described above in the event that the applicant is awarded a contract\. Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."


'''''''''''''''''''DI''''LOI Project Personnel ''''''''''''''''''
Public const Pro_Personnel_InText_DI_LOI = "For instructions on using our system, click here to access our PCORI Portal User Guide\. To view additional information about the PCORI submission process, click here to view our FAQs page\. Senior or Key Personnel are individuals in addition to the Principal Investigator who contributes to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. At least one key personnel entry is required\.  If stakeholder is selected, you may enter ""N/A"" for institution\. PCORI is committed to recognizing the contributions of all of the members of the project team, including patient and stakeholder partners, and makes information about the projects it funds publicly available through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\. By providing the names of the members of the project team \(including individuals and partnering organizations\), Applicant understands and acknowledges that PCORI may use such information as described above in the event that Applicant is awarded a contract\. Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."
Public const Pro_Personnel_BoldedText_DI_LOI = "<p>At least one key personnel entry is required\. <strong>Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\.</strong></p>"
Public const Pro_Personnel_InText_app_DI_childrecord_1 = "Senior or Key Personnel are individuals in addition to the Principal Investigator who contributes to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. At least one key personnel entry is required\. If stakeholder is selected, you may enter ""N/A"" for institution\. "
Public const Pro_Personnel_InText_app_DI_childrecord_2 = "PCORI is committed to recognizing the contributions of all of the members of the project team, including patient and stakeholder partners, and makes information about the projects it funds publicly available through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\. By providing the names of the members of the project team \(including individuals and partnering organizations\), Applicant understands and acknowledges that PCORI may use such information as described above in the event that Applicant is awarded a contract\. Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."
Public const Pro_Personnel_BoldedText_Childrecord_DI_LOI = "<b>Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\.</b>"

'''''''''''''''''''HSII''''LOI Project Personnel ''''''''''''''''''
Public const Pro_Personnel_InText_HSII_LOI = "For instructions on using our system, click here to access our HSII User Guide\. To view additional information about the PCORI application process, click here to access our Applicant FAQs page\.   Senior or Key Personnel are individuals in addition to the Principal Investigator who contributes to the development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\.   At least one key personnel entry is required\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. If stakeholder is selected, you may enter ""N/A"" for institution\.   PCORI is committed to recognizing the contributions of all of the members of the project team, including patient and stakeholder partners\. Applicants should identify patient and stakeholder partners, whether individuals or organizations, if known\. Note that stakeholder partner organizations will be publicly listed on the PCORI website and may be included on public communications\. \(Individual names will not be disclosed\.\) In providing the names of stakeholder organization partners, applicants acknowledge that partners have consented to the disclosure of their names to PCORI and to making their names publicly available\. If a stakeholder partner chooses to remain anonymous, contact pfa@pcori\.org for guidance\.    "
Public const Pro_Personnel_BoldedText_HSII = "<strong>Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\.</strong>"
Public const Pro_Personnel_InText_HSII_childrecord_1 = "Senior or Key Personnel are individuals in addition to the Principal Investigator who contributes to the development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. At least one key personnel entry is required\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. If stakeholder is selected, you may enter ""N/A"" for institution\."
Public const Pro_Personnel_InText_HSII_childrecord_2 = "PCORI is committed to recognizing the contributions of all of the members of the project team, including patient and stakeholder partners\. Applicants should identify patient and stakeholder partners, whether individuals or organizations, if known\. Note that stakeholder partner organizations will be publicly listed on the PCORI website and may be included on public communications\. \(Individual names will not be disclosed\.\) In providing the names of stakeholder organization partners, applicants acknowledge that partners have consented to the disclosure of their names to PCORI and to making their names publicly available\. If a stakeholder partner chooses to remain anonymous, contact pfa@pcori\.org for guidance\."

''''''''''''''''''''''''''''''Project Personnel Tab DI Application'''''''''''''''''''
Public const Pro_Personnel_InText_app_DI = "For instructions on using our system, click here to access our PCORI Portal User Guide\. To view additional information about the PCORI submission process, click here to view our FAQ page\. Project Personnel are individuals, in addition to the Principal Investigator, who contribute to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. At least one project personnel entry is required\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. PCORI is committed to recognizing the contributions of all of the members of the research team, including patient and stakeholder partners\. PCORI publicizes information about its funded research projects through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\. By providing the names of the members of the research team \(including individuals and partnering organizations\) below, the applicant understands and acknowledges that PCORI may use such information as described above in the event that the applicant is awarded a contract\. Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."
Public const Pro_Personnel_BoldedText_app_DI = "<strong>Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\.</strong>"
Public const Pro_Personnel_InText_app_DI_childrecord = "Project Personnel are individuals, in addition to the Principal Investigator, who contribute to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. At least one project personnel entry is required\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. PCORI is committed to recognizing the contributions of all of the members of the research team, including patient and stakeholder partners\. PCORI publicizes information about its funded research projects through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\. By providing the names of the members of the research team \(including individuals and partnering organizations\) below, the applicant understands and acknowledges that PCORI may use such information as described above in the event that the applicant is awarded a contract\. Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."
Public const Pro_Personnel_BoldedText_app_DI_childrecord = "<b>Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\.</b>"
Public const Pro_Personnel_BoldedText_app_SDM = "<strong>Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\.</strong>"

'''''''''''''''''''''''''''''Project Personnel Tab BPS Application'''''''''''''''''''
Public const Pro_Personnel_InText_app_BPS = "For instructions on using our system, click here to access our PCORI Portal User Guide\. To view additional information about the PCORI submission process, click here to view our FAQ page\. Project Personnel are individuals, in addition to the Principal Investigator, who contribute to the scientific development or execution of the project in a substantive and measurable way\. The contribution is independent of financial compensation\. At least one project personnel entry is required\. Note: Please include the PI, or Dual-PIs, as Key Personnel to the Project Personnel list below\. PCORI is committed to recognizing the contributions of all of the members of the research team, including patient and stakeholder partners\. PCORI publicizes information about its funded research projects through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\. By providing the names of the members of the research team \(including individuals and partnering organizations\) below, the applicant understands and acknowledges that PCORI may use such information as described above in the event that the applicant is awarded a contract\. Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."


'''''''''''''''''Templates and Uploads tab'''''''''''
Public const LOI_Template_BPS = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Broad Pragmatic Studies Funding Announcement -- 2024 Standing PFA \(Cycle 2 2024\) Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\. An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\. "
Public const Templates_Tab_Instxt2 = "<p>Note that all uploads must be PDF documents\.</p>"
Public const Templates_Tab_Instxt3 = "<p>LOIs that exceed the PFA-specific word/page limit will not be reviewed\.</p>"
Public const Templates_Tab_Instxt4 = "<p>Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\.</p>"
Public const Templates_Tab_Instxt5 = "<p>An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\.</p>"
Public const Templates_Tab_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Broad-Pragmatic-LOI-Template\.docx"
Public const Templates_Tab_Link = "Broad Pragmatic Studies Funding Announcement -- 2024 Standing PFA \(Cycle 2 2024\)"
Public Const attachmentBroadLOITemplate = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Broad-Pragmatic-LOI-Template.docx"

'''''''''''''''Methods LOI'''''''''''''''
Public const LOI_Template_Methods = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Improving Methods for Conducting Patient-Centered Outcomes Research Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\. An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\. "
Public const Templates_Tab_URL_Methods = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Methods-LOI-Template\.docx"
Public const Templates_Tab_Link_Methods = "Improving Methods for Conducting Patient-Centered Outcomes Research"
Public Const attachment_Methods_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Methods-LOI-Template"
Public Const Templates_Name_Methods = "PCORI-2024-Methods-LOI-Template\.docx"

'''''''''''''''SOE LOI'''''''''''''''
Public const LOI_Template_SOE = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Science Of Engagement \(SoE\) Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\. An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\. " 
Public const Templates_Tab_URL_SOE = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Science-of-Engagement-LOI-Template\.docx"
Public const Templates_Tab_Link_SOE = "Science Of Engagement \(SoE\)"
Public Const attachment_SOE_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-1-Science-of-Engagement-LOI-Template"
Public Const Templates_Name_SOE = "PCORI-2024-Cycle-1-Science-of-Engagement-LOI-Template\.docx"

'''''''''''''''Sleep Health LOI'''''''''''''''
Public const LOI_Template_SleepHealth = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Sleep Health Topical PCORI Funding Announcement -- Cycle 3 2023 Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\. An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\. "
Public const Templates_Tab_URL_SleepHealth = "https://www\.pcori\.org/sites/default/files/PCORI-2023-Cycle-3-Sleep-Health-LOI-Template\.docx"
Public const Templates_Tab_Link_SleepHealth = "Sleep Health Topical PCORI Funding Announcement -- Cycle 3 2023"
Public Const attachment_SleepHealth_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2023-Cycle-3-Sleep-Health-LOI-Template"
Public Const Templates_Name_SleepHealth = "PCORI-2023-Cycle-3-Sleep-Health-LOI-Template\.docx"

'''''''''''''''PLACER LOI'''''''''''''''
Public const LOI_Template_PLACER = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Phased Large Awards for Comparative Effectiveness Research Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\. An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\. "
Public const Templates_Tab_URL_PLACER = "https://www\.pcori\.org/sites/default/files/PCORI-2025-Cycle-1-PLACER-LOI-Template\.docx"
Public const Templates_Tab_Link_PLACER = "Phased Large Awards for Comparative Effectiveness Research"
Public Const attachment_PLACER_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2025-Cycle-1-PLACER-LOI-Template"
Public Const Templates_Name_PLACER = "PCORI-2025-Cycle-1-PLACER-LOI-Template\.docx"


'''''''''''''''MMM LOI'''''''''''''''
Public const LOI_Template_MMM = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Improving Postpartum Maternal Outcomes for Populations Experiencing Disparities Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\. An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\. "
Public const Templates_Tab_URL_MMM = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-1-MMM-LOI-Template\.docx"
Public const Templates_Tab_Link_MMM = "Maternal Morbidity and Mortality Topical PCORI Funding Announcement -- Cycle 1 2024"
Public Const attachment_MMM_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-1-MMM-LOI-Template"
Public Const Templates_Name_MMM = "PCORI-2024-Cycle-1-MMM-LOI-Template\.docx"

'''''''''''''IDD'''''''''''''''''''
Public Const  LOI_Template_IDD = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Intellectual and Developmental Disabilities Topical Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\. An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\. "
Public Const Templates_Tab_Link_IDD = "Intellectual and Developmental Disabilities Topical"
Public Const Templates_Tab_URL_IDD = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-1-IDD-LOI-Template\.docx"
Public Const attachment_IDD_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-1-IDD-LOI-Template"

''''''''''Mental Health''''''''LOI Phase Template & Upload Tab'''''''''
Public Const  LOI_Template_Mental_Health = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Mental and Behavioral Health Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\. An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\. "
Public Const Templates_Tab_Link_Mental_Health = "Improving Mental and Behavioral Health Topical"
Public Const Templates_Tab_URL_Mental_Health = "https://www\.pcori\.org/sites/default/files/PCORI-2025-Cycle-1-Mental-and-Behavioral-Health-LOI-Template\.docx"
Public Const attachment_Mental_Health_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2025-Cycle-1-Mental-and-Behavioral-Health-LOI-Template"


''''''''''Managing Pain''''''''LOI Phase Template & Upload Tab'''''''''
Public Const  LOI_Template_Managing_Pain = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Managing Pain Topical PCORI Funding Announcement -- Cycle 1 2025 Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\. An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\. "
Public Const Templates_Tab_Link_Managing_Pain = "Managing Pain Topical PCORI Funding Announcement -- Cycle 1 2025"
Public Const Templates_Tab_URL_Managing_Pain = "https://www\.pcori\.org/sites/default/files/PCORI-2025-Cycle-1-Managing-Pain-LOI-Template\.docx"
Public Const attachment_Managing_Pain_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2025-Cycle-1-Managing-Pain-LOI-Template"
Public Const Templates_Name_Managing_Pain = "PCORI-2025-Cycle-1-Managing-Pain-LOI-Template\.docx"

''''''''''Violence Trauma''''''''LOI Phase Template & Upload Tab'''''''''
Public Const  LOI_Template_Violence_Trauma = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Mental and Behavioral Health Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen Program staff may invite applicants from previous cycles to resubmit their \(revised\) applications\. If invited, applicants will bypass the LOI review stage\. Instead of completing and uploading an LOI Template, invited applicants are required to upload their Invitation to Resubmit letter and complete the PCORI Online LOI questions by the LOI submission deadline\. Unless the applicant has explicit and documented approval from the program staff to alter the originally submitted study aims of the application, the invited resubmission application’s aims must remain the same as in the original application\. An Invitation to Resubmit is not a guarantee that PCORI will select the application for funding\. Invited applicants must adhere to the updated guidance in the PFA and compete with other invited and new applicants\. "
Public Const Templates_Tab_Link_Violence_Trauma = "Addressing Violence and Trauma Topical"
Public Const Templates_Tab_URL_Violence_Trauma = "https://www\.pcori\.org/sites/default/files/PCORI-2025-Cycle-2-Violence-Trauma-LOI-Template\.docx"
Public Const attachment_Violence_Trauma_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2025-Cycle-2-Violence-Trauma-LOI-Template"


''''''''''''''IMRI'''''''''''''''LOI'''
Public const LOI_Template_IMRI = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Open Competition PFA: Implementation of Findings from PCORI's Research Investments Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Templates_Tab_URL_IMRI = "https://www\.pcori\.org/sites/default/files/PCORI-2025-Cycle-1-Open-Competition-Implementation-LOI-Template\.docx"
Public const Templates_Tab_Link_IMRI = "Open Competition PFA: Implementation of Findings from PCORI's Research Investments"
Public Const attachment_IMRI_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2025-Cycle-1-Open-Competition-Implementation-LOI-Template.docx"
Public Const Templates_Name_IMRI = "PCORI-2025-Cycle-1-Open-Competition-Implementation-LOI-Template\.docx"

''''''''''''''''''LC'''''LOI''''''
Public const LOI_Template_LC = "\* LOI Template A Letter Of Intent is required for new and resubmitted applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Limited Competition PFA: Implementation Awards Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Templates_Tab_URL_LC = "https://www\.pcori\.org/sites/default/files/PCORI-2025-Cycle-1-Limited-Competition-Implementation-LOI-Template\.docx"
Public const Templates_Tab_Link_LC = "Limited Competition PFA: Implementation Awards"
Public Const attachment_LC_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2025-Cycle-1-Limited-Competition-Implementation-LOI-Template"
Public Const Templates_Name_LC  = "PCORI-2025-Cycle-1-Limited-Competition-Implementation-LOI-Template\.docx"

'''''''SDM'''''''''''''''LOI'''
Public const LOI_Template_SDM = "\* LOI Template A Letter Of Intent is required for new applications and must be submitted prior to completion of an application\. Download the PFA-specific LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the PFA-specific word/page limit will not be reviewed\. Implementation of Effective Shared Decision Making Approaches in Practice Settings Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Templates_Tab_URL_SDM = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Shared-Decision-Making-LOI-Template\.docx"
Public const Templates_Tab_Link_SDM = "Implementation of Effective Shared Decision Making Approaches in Practice Settings"
Public Const attachment_SDM_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Shared-Decision-Making-LOI-Template"
Public Const Templates_Name_SDM = "PCORI-2024-Cycle-3-Shared-Decision-Making-LOI-Template\.docx"

'''''''''''''''HSII'''''''''''''''LOI'''
Public const LOI_Template_HSII = "\* LOI Template A Letter Of Intent is required for new applications and must be submitted prior to completion of an application\. Download the LOI template below\. Note that all uploads must be PDF documents\. LOIs that exceed the specific word/page limit will not be reviewed\. Call for Proposals for HSII Implementation Projects - 2024 Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Templates_Tab_Link_HSII = "Call for Proposals for HSII Implementation Projects - 2024"
Public const Templates_Tab_URL_HSII = "https://www\.pcori\.org/sites/default/files/PCORI-2024-HSII-Implementation-LOI-Template\.docx"
Public Const Templates_Name_HSII = "PCORI-2024-HSII-Implementation-LOI-Template\.docx"
Public Const attachment_HSII_LOI_Template = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2023-HSII-Implementation-LOI-Template"


''''''''''''''''SOE Application Templates and uploads Tab'''''''''''''''''''''
Public const Templates_Tab_AppSOE_Instxt1 = "Download PFA-specific application templates below\. Note that all uploads must be PDF documents except the Milestones and Methodology Standards Checklist\."

''''''''''''''''SDM Application Templates and uploads Tab'''''''''''''''''''''
Public const Templates_Tab_AppSDM_Instxt1 = "Download PFA-specific application templates below\. Note that all uploads must be PDF documents, except the Milestones and Subcontractor Detailed Budget\."

''''''''''''''''''''''''''''''Application Tempates and Uploads Tab LC '''''''''''''''''''''
Public const Templates_Tab_AppLC_Instxt1 = "<p>Download PFA-specific application templates below\.</p>"
Public const Templates_Tab_AppLC_Instxt2 = "<p>Note that all uploads must be PDF documents, except the Milestones\.</p>"
Public const Resubmission_Letter_App_LC = "Resubmission Letter Download Implementation Resubmission Letter Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const PI_and_Key_Personnel_App_LC = "\* PI and Key Personnel Download PI and Key Personnel Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Patient_and_Stakeholder_Partner_App_LC  = "Patient and Stakeholder Partner Download Patient and Stakeholder Partner Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Project_Performance_Site_and_Resources_App_LC = "\* Project Performance Site\(s\) and Resources Download Project Performance Site\(s\) and Resources Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Project_Plan_App_LC = "\* Project Plan Download Project Plan Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Project_Summary_App_LC = "\* Project Summary Download Implementation Project Summary Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Milestones_App_LC = "\* Milestones Download Implementation Milestones Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Subcontractor_Detailed_Budget_App_LC = "Subcontractor Detailed Budget Download Subcontractor Detailed Budget Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Budget_Justification_App_LC = "\* Budget Justification Download Budget Justification Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Letters_of_Support_App_LC  = "\* Letters of Support Download Letters of Support Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Resubmission_Letter_LC_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-DI-Resubmission-Letter-Template\.docx"
Public const PI_and_Key_Personnel_LC_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-PI-Key-Personnel-Template\.docx"
Public const Patient_and_Stakeholder_Partner_LC_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Patient-Stakeholder-Partner-Template\.docx"
Public const Project_Performance_Site_and_Resources_LC_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Project-Performance-Sites-Resources-Template\.docx"
Public const Project_Plan_LC_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Limited-Competition-Implementation-Project-Plan-Template\.docx"
Public const Project_Summary_LC_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Limited-Competition-Implementation-Project-Summary-Template\.docx"
Public const Milestones_LC_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Limited-Competition-Implementation-Milestones-Template\.xlsx"
Public const Subcontractor_Detailed_Budget_LC_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Limited-Competition-Implementation-Subcontractor-Detailed-Budget-Template\.xlsx"
Public const Budget_Justification_LC_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Budget-Justification-Template\.docx"
Public const Letters_of_Support_LC_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Limited-Competition-Implementation-Letters-of-Support-Table\.docx"
Public const Leadership_Plan_App = "Leadership Plan Download Leadership Plan Template \(Only required if proposed project is dual-PI\) Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Leadership_Plan_App_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Leadership-Plan-Template\.docx"

''''''''''''''''' Attachments File Path For LC Application '''''''''''''
Public Const attachment_ResubmissionLC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-DI-Resubmission-Letter-Template.docx"
Public Const attachment_PIandKeyPersonnelLC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-PI-Key-Personnel-Template.docx"
Public Const attachmentBudget_Justification_LC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Budget-Justification-Template.docx"
Public const attachmentPatientandStakeholderLC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Patient-Stakeholder-Partner-Template.docx"
Public const attachmentProjectPerformanceSiteLC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Project-Performance-Sites-Resources-Template.docx"
Public const attachmentLettersofSupportLC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Limited-Competition-Implementation-Letters-of-Support-Table.docx"
Public const attachmentProjectPlanLC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Limited-Competition-Implementation-Project-Plan-Template.docx"
Public const attachmentProject_Summary_LC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Limited-Competition-Implementation-Project-Summary-Template.docx"
Public const MilestonesTemplateLC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Limited-Competition-Implementation-Milestones-Template.xlsx"
Public const SubcontractorDetailedBudgetLC = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Limited-Competition-Implementation-Subcontractor-Detailed-Budget-Template.xlsx"

'''''''''''Application IMRI Specific''''''''
Public const Resubmission_Letter_App_IMRI = "Resubmission Letter Download Resubmission Letter Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Resubmission_Letter_IMRI_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-DI-Resubmission-Letter-Template\.docx"
Public const PI_and_Key_Personnel_IMRI_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-PI-Key-Personnel-Template\.docx"
Public const Patient_and_Stakeholder_Partner_IMRI_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Patient-Stakeholder-Partner-Template\.docx"
Public const Project_Plan_IMRI_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Open-Competition-Implementation-Project-Plan-Template\.docx"
Public const Project_Summary_IMRI_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Open-Competition-Implementation-Project-Summary-Template\.docx"
Public const Milestones_IMRI_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Open-Competition-Implementation-Milestones-Template\.xlsx"
Public const Subcontractor_Detailed_Budget_IMRI_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Open-Competition-Implementation-Subcontractor-Detailed-Budget-Template\.xlsx"
Public const Letters_of_Support_IMRI_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Open-Competition-Implementation-Letters-of-Support-Table\.docx"

'''''''''''''''''' Attachments File Path For IMRI Application'''''''''''''''
Public const attachmentLettersofSupportIMRI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Open-Competition-Implementation-Letters-of-Support-Table.docx"
Public const attachmentProjectPlanIMRI= "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Open-Competition-Implementation-Project-Plan-Template.docx"
Public const attachmentProject_Summary_IMRI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Open-Competition-Implementation-Project-Summary-Template.docx"
Public const MilestonesTemplateIMRI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Open-Competition-Implementation-Milestones-Template.xlsx"
Public const SubcontractorDetailedBudgetIMRI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Open-Competition-Implementation-Subcontractor-Detailed-Budget-Template.xlsx"
Public const Leadership_Plan_IMRI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Leadership-Plan-Template.docx"

'''''''''''Application SDM Specific''''''''
Public const Resubmission_Letter_App_SDM = "Resubmission Letter Download Shared Decision-Making Resubmission Letter Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Shared_Decision_Making_Approach_App_SDM = "\* Shared Decision-Making Approach Download Shared Decision-Making Approach Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Project_Performance_Site_and_Resources_App_SDM = "\* Project Performance Site\(s\) and Resources Download Shared Decision-Making Project Performance Site\(s\) and Resources Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Project_Plan_App_SDM = "\* Project Plan Download Shared Decision-Making Project Plan Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Project_Summary_App_SDM = "\* Project Summary Download Shared Decision-Making Project Summary Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Milestones_App_SDM = "\* Milestones Download Shared Decision-Making Milestones Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Subcontractor_Detailed_Budget_App_SDM = "Subcontractor Detailed Budget Download Shared Decision-Making Subcontractor Detailed Budget Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Letters_of_Support_App_SDM = "\* Letters of Support Download Shared Decision-Making Letters of Support Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Budget_Justification_App_SDM = "\* Budget Justification Download Shared Decision-Making Budget Justification Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "

'''''''''''''''''Attachments URL For SDM Application Templates and uploads tab'''''''''''''''
Public const Shared_Decision_Making_Approach_App_SDM_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-1-Shared-Decision-Making-Approach-Template\.docx"
Public const Project_Plan_SDM_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Shared-Decision-Making-Project-Plan-Template\.docx"
Public const Project_Summary_SDM_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Shared-Decision-Making-Project-Summary-Template\.docx"
Public const Milestones_SDM_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Shared-Decision-Making-Milestones-Template\.xlsx"
Public const Subcontractor_Detailed_Budget_SDM_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Shared-Decision-Making-Subcontractor-Detailed-Budget-Template\.xlsx"
Public const Letters_of_Support_SDM_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-3-Shared-Decision-Making-Letters-of-Support-Table\.docx"

'''''''''''''''''' Attachments File Path For SDM Application'''''''''''''''
Public Const attachment_Shared_Decision_Making_Approach_SDM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2023-Cycle-3-Shared-Decision-Making-Approach-Template.docx"
Public Const attachmentProjectPlan_SDM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Shared-Decision-Making-Project-Plan-Template.docx"
Public const attachmentProject_Summary_SDM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Shared-Decision-Making-Project-Summary-Template.docx"
Public const MilestonesTemplate_SDM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Shared-Decision-Making-Milestones-Template.xlsx"
Public const SubcontractorDetailedBudget_SDM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Shared-Decision-Making-Subcontractor-Detailed-Budget-Template.xlsx"
Public const attachmentLettersofSupport_SDM = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-3-Shared-Decision-Making-Letters-of-Support-Table.docx"

'''''''''''''''''' Attachments URL For SOE Application Templates and uploads tab'''''''''''''''
Public const Resubmission_Letter_App_SOE = "Resubmission Letter Download Resubmission Letter Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Resubmission_Letter_SOE_URL = "https://www\.pcori\.org/sites/default/files/PCORI-Resubmission-Letter-Template\.docx"
Public const PI_and_Key_Personnel_App_SOE = "\* Scientific PI and Key Personnel Download Scientific PI and Key Personnel Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const PI_and_Key_Personnel_SOE_URL = "https://www\.pcori\.org/sites/default/files/PCORI-Scientific-PI-Key-Personnel-Template\.docx"
Public const Project_Performance_Site_and_Resources_App_SOE= "\* Project/Performance Site\(s\) and Resources Template Download Project Performance Sites Resources Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Research_Plan_App_SOE_App = "\* Research Plan Download Research Plan Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Research_Plan_SOE_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-2-Science-of-Engagement-Research-Plan-Template\.docx"
Public const Milestones_SOE_App = "\* Milestones Download Milestones Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Milestones_SOE_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-2-Science-of-Engagement-Milestones-Template\.xlsx"
Public const Subcontractor_Detailed_Budget_SOE_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-2-Science-of-Engagement-Subcontractor-Detailed-Budget-Template\.xlsx"
Public const Letters_of_Support_SOE_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-2-Science-of-Engagement-Letters-of-Support-Table\.docx"
Public const Methodology_Standard_Checklist_App_SOE = "\* Methodology Standards Checklist Download Methodology Standards Checklist Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Methodology_Standard_Checklist_App_SOE_URL = "https://www\.pcori\.org/sites/default/files/PCORI-PFA-Methodology-Standards-Checklist\.xlsx"

'''''''''''''''''' Attachments File Path For SOE Application'''''''''''''''
Public Const attachment_Resubmission_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-Resubmission-Letter-Template.docx"
Public const attachmentLettersofSupport_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-2-Science-of-Engagement-Letters-of-Support-Table.docx"
Public const attachmentResearchPlan_SOE= "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-2-Science-of-Engagement-Research-Plan-Template.docx"
Public const attachmentProject_Summary_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-1-Open-Competition-Implementation-Project-Summary-Template.docx"
Public const MilestonesTemplateI_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-2-Science-of-Engagement-Milestones-Template.xlsx"
Public const SubcontractorDetailedBudget_SOE= "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-2-Science-of-Engagement-Subcontractor-Detailed-Budget-Template.xlsx"
Public const Leadership_Plan_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-Leadership-Plan-Template.docx"
Public Const attachment_PIandKeyPersonnel_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-Scientific-PI-Key-Personnel-Template.docx"
Public Const attachment_Methodology_Standard_Checklist_SOE = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-PFA-Methodology-Standards-Checklist.xlsx"

'''''''''''''''''' Attachments URL For Methods Application Templates and uploads tab'''''''''''''''
Public const Research_Plan_Methods_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Methods-Research-Plan-Template\.docx"
Public const Milestones_Methods_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Methods-Milestones-Template\.xlsx"
Public const Subcontractor_Detailed_Budget_Methods_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Methods-Subcontractor-Detailed-Budget-Template\.xlsx"
Public const Letters_of_Support_Methods_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Methods-Letters-of-Support-Table\.docx"

'''''''''''''''''' Attachments File Path For Methods Application'''''''''''''''
Public const attachment_Research_Plan_Methods = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Methods-Research-Plan-Template.docx"
Public const attachment_Milestones_Methods = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Methods-Milestones-Template.xlsx"
Public const attachment_Subcontractor_Detailed_Budget_Methods = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Methods-Subcontractor-Detailed-Budget-Template.xlsx"
Public const attachment_Letters_of_Support_Methods = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Methods-Letters-of-Support-Table.docx"

''''''''''''' Attachments URL For BPS Application Templates and uploads tab'''''''''''''''
Public const Resubmission_Letter_BPS_URL = "https://www\.pcori\.org/sites/default/files/PCORI-Resubmission-Letter-Template\.docx"
Public const Research_Plan_BPS_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Broad-Pragmatic-Research-Plan-Template\.docx"
Public const Project_Performance_Site_and_Resources_App_BPS = "\* Project Performance Site\(s\) and Resources Download Project Performance Site\(s\) and Resources Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Subcontractor_Detailed_Budget_BPS_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Broad-Pragmatic-Subcontractor-Detailed-Budget-Template\.xlsx"
Public const Letters_of_Support_BPS_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Broad-Pragmatic-Letters-of-Support-Table\.docx"
Public const Milestones_BPS_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Broad-Pragmatic-Milestones-Template\.xlsx"
Public const attachment_Subcontractor_Detailed_Budget_BPS = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Broad-Pragmatic-Subcontractor-Detailed-Budget-Template.xlsx"
Public const attachment_Letters_of_Support_BPS = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Broad-Pragmatic-Letters-of-Support-Table.docx"
Public const attachment_Research_Plan_BPS = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Broad-Pragmatic-Research-Plan-Template.docx"
Public const attachment_Milestones_BPS = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Broad-Pragmatic-Milestones-Template.xlsx"

''''''''''''' Attachments URL For PLACER Application Templates and uploads tab'''''''''''''''
Public const Research_Plan_PLACER_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-2-PLACER-Research-Plan-Template\.docx"
Public const PI_and_Key_Personnel_App_PLACER = "\* Scientific PI and Key Personnel Download PI Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Patient_and_Stakeholder_Partner_App_PLACER = "Patient and Stakeholder Partner Download Key Personnel Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Project_Performance_Site_and_Resources_App_PLACER = "\* Project Performance/Site\(s\) and Resources Download Project Performance/Site\(s\) and Resources Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Project_Performance_Site_and_Resources_PLACER_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-1-Project-Performance-Sites-Resources-Template\.docx"
Public const Milestones_PLACER_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-2-PLACER-Milestones-Template\.xlsx"
Public const PCORI_10_Year_Budget_PLACER = "\* PCORI 10-Year Budget PLACER Download PCORI 10-Year Budget Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const PCORI_10_Year_Budget_PLACER_URL = "https://www\.pcori\.org/sites/default/files/PCORI-10-Years-Budget-PLACER-Template\.xlsx"
Public const Letters_of_Support_PLACER_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-2-PLACER-Letters-of-Support-Table\.docx"
Public const attachment_Research_Plan_PLACER = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-2-PLACER-Research-Plan-Template.docx"
Public const attachment_Milestones_PLACER = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-2-PLACER-Milestones-Template.xlsx"
Public const attachment_PCORI_10_Year_Budget_PLACER = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-10-Years-Budget-PLACER-Template.xlsx"
Public const attachment_Letters_of_Support_PLACER = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-2-PLACER-Letters-of-Support-Table.docx"

''''''''''''' Attachments URL For IDD Application Templates and uploads tab'''''''''''''''
Public const Resubmission_Letter_IDD_URL = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-Resubmission-Letter-Template.docx"
Public const Research_Plan_IDD_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-1-IDD-Research-Plan-Template\.docx"
Public const attachment_Research_Plan_IDD = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-1-IDD-Research-Plan-Template.docx"
Public const Milestones_IDD_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-1-IDD-Milestones-Template\.xlsx"
Public const attachment_Milestones_IDD = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-1-IDD-Milestones-Template.xlsx"
Public const Subcontractor_Detailed_Budget_IDD_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-1-IDD-Subcontractor-Detailed-Budget-Template\.xlsx"
Public const attachment_Subcontractor_Detailed_Budget_IDD = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-1-IDD-Subcontractor-Detailed-Budget-Template.xlsx"
Public const Letters_of_Support_IDD_URL = "https://www\.pcori\.org/sites/default/files/PCORI-2024-Cycle-1-IDD-Letters-of-Support-Table\.docx"
Public const attachment_Letters_of_Support_IDD = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-2024-Cycle-1-IDD-Letters-of-Support-Table.docx"

''''''''''''''''''''''Certification Tab Application Captured Outertext'''''''''''''''''''''
Public const LC_Certification_Tab_Ins1 = "By submitting this application, I certify that to the best of my knowledge, the information provided is complete and accurate and that I have obtained all permissions, authorizations, or consents that may be required under applicable law to disclose the enclosed information to PCORI and its contractors \(together “PCORI”\) and for PCORI to use the information as described in this application, the PCORI funding announcement, Submission Instructions and any other application preparation materials\."
Public const SDM_Certification_Tab_Ins1 = "By submitting this application, I certify that to the best of my knowledge, the information provided is complete and accurate and that I have obtained all permissions, authorizations, or consents that may be required under applicable law to disclose the enclosed information to PCORI and its contractors \(together “PCORI”\) and for PCORI to use the information as described in this application, the PCORI funding announcement, Application Guidelines and any other application preparation materials\."
Public const PI_Certification_app = "\* I Agree --None-- Yes No "
Public const AO_Approval_Page_Inst1 = "Only the Administrative Official \(AO\) may access these fields\. Please click the Cancel button if you are not the AO\."
Public const AO_Approval_Page_Inst2 = "\* I hereby certify that, to the best of my knowledge, the information in this PCORI application is true and accurate\. \* I understand that the discovery of false or fictitious information included in the application may result in rejection from the review process or termination of an award\. \* I certify that the funds applied for will be used as outlined in the proposed budget, and in accordance with the contract terms and conditions\. \* To the best of my knowledge, none of the key personnel and /or collaborating institutions are currently banned from receiving federal funds due to debarment or engagement in research misconduct\. \* By submitting this application, I attest that I am recognized by my institution as an official authorized to enter into contractual agreements and commit institution resources\."
Public const AO_Agree_Field = "\* I agree to the statements above \(please select an option\) --None-- Yes No "
Public const AO_Approval_Page_Inst3 = "As the Administrative Official, I approve this application submission to PCORI\."
Public const AO_Approval_Page_Inst4 = "•If you select ‘Reject’, and then select the ‘Save’ button, this application will be sent back to the Project Lead and may be edited\. •Once you select ‘Approved’ and then select the ‘Save’ button, the application cannot be changed without contacting PCORI\."
Public const AO_Approve_Reject_Field  = "--None-- Approve Reject"
Public const AO_Approval_Page_Inst5 = "If you wish to withdraw this application and no longer submit to PCORI, please select ‘Withdraw,’ provide your reasons in the text box below, and select the ‘Save’ button\. Note: This step is permanent\."
Public const AO_Withdraw_Field = "Withdraw Application --None-- Withdraw "
Public const AO_Withdraw_Reason_Field = "Reasons for Withdrawal 0 of 50000 Characters "
Public const AO_Approval_Page_Inst6 = "Select the ‘Save’ button below and scroll to the top of the page to select the ‘Review/Submit’ button\."


''''''''This below text is captured by Outerhtml since the font color is red''''''''''''
Public const AO_Approval_Page_Inst1_redfont = "<p><b><font color=""red"">Only the Administrative Official \(AO\) may access these fields\. Please click the Cancel button if you are not the AO\.</font></b></p>"

''''''''This below text is captured by Outerhtml since the text has paragraph and bulleted text''''''''''''
Public const AO_Approval_Page_Inst2_paragraph = "<div class=""sfdc_richtext"" id=""j_id0:mainForm:j_id252:0:j_id261j_id0:mainForm:j_id252:0:j_id261_00N1O00000BWHi0_div""><p>\* I hereby certify that, to the best of my knowledge, the information in this PCORI application is true and accurate\. </p><br><p>\* I understand that the discovery of false or fictitious information included in the application may result in rejection from the review process or termination of an award\.</p><br><p>\* I certify that the funds applied for will be used as outlined in the proposed budget, and in accordance with the contract terms and conditions\.</p><br><p>\* To the best of my knowledge, none of the key personnel and /or collaborating institutions are currently banned from receiving federal funds due to debarment or engagement in research misconduct\.</p><br><p>\* By submitting this application, I attest that I am recognized by my institution as an official authorized to enter into contractual agreements and commit institution resources\.</p></div>"
Public const AO_Approval_Page_Inst4_paragraph = "<div class=""sfdc_richtext"" id=""j_id0:mainForm:j_id252:2:j_id261j_id0:mainForm:j_id252:2:j_id261_00N1O00000BWHi0_div""><br><br>•If you select ‘Reject’, and then select the ‘Save’ button, this application will be sent back to the Project Lead and may be edited\.<br><br>•Once you select ‘Approved’ and then select the ‘Save’ button, the application cannot be changed without contacting PCORI\.</div>"
Public const AO_Approval_Page_Inst6_boldtext = "<b>Select the ‘Save’ button below and scroll to the top of the page to select the ‘Review/Submit’ button\. </b>"


''''''''''''''''''''''''''''LOI Review'''''''''''''''''''''''''''''
'''''''''''''''''''''BPS/ methods/PLACER/SOE/Sleephealth/MMM/IDD LOI Review Form fields edit mode by outertext'''''''''''
Public const Category_3_Contact_Front_Door = "\*Category 3: Contact Front Door\? Category 3: Contact Front Door\? Help Info --None--"
Public const Special_Area_of_Emphasis_SAE = "\*Special Area of Emphasis \(SAE\) Special Area of Emphasis \(SAE\) Help Info --None--"
Public const Primary_National_Priority_Aims_Aligned = "\*Primary National Priority Aims Aligned\? Primary National Priority Aims Aligned\? Help Info --None-- View all dependencies"
Public const PCORI_National_Priority_Area_Fit = "PCORI National Priority Area Fit PCORI National Priority Area Fit Help Info --None--"
Public const Compelling_research_question = "Compelling research question\? Compelling research question\? Help Info --None--"
Public const Systematic_rev_evidence_synthesis_cited = "Systematic rev/evidence synthesis cited\? Systematic rev/evidence synthesis cited\? Help Info --None--"
Public const Study_Design = "Study Design Study Design Help Info --None--"
Public const Population_appropriate = "\*Population appropriate\? Population appropriate\? Help Info --None--"
Public const Interventions_have_efficacy_or_in_use = "Interventions have efficacy or in use\? Interventions have efficacy or in use\? Help Info --None--"
Public const Usual_Care_Comparator = "Usual Care Comparator\? Usual Care Comparator\? Help Info --None--"
Public const Outcomes = "\*Outcomes Outcomes Help Info --None--"
Public const Sample_size_power_assumption_appropriate = "Sample size/power assumption appropriate Sample size/power assumption appropriate Help Info --None--"
Public const Hypothesized_Effect_Size = "Hypothesized Effect Size Hypothesized Effect Size Help Info --None--"
Public const Appropriate_duration = "Appropriate duration\? Appropriate duration\? Help Info --None--"
Public const Analysis_plan_HTE_appropriate ="Analysis plan/HTE appropriate\? Analysis plan/HTE appropriate\? Help Info --None--"
Public const Patient_stakeholder_community_engagement = "Patient/stakeholder/community engagement Patient/stakeholder/community engagement Help Info --None--"
Public const Does_applicant_have_adequate_experience = "Does applicant have adequate experience\? Does applicant have adequate experience\? Help Info --None--"
Public const PCORI_Research_Priority_Area_Key_Topic =  "\*PCORI Research Priority Area/Key Topic PCORI Research Priority Area/Key Topic Help Info --None--"
Public const PCC_coverage_requested = "\*PCC coverage requested\? PCC coverage requested\? Help Info --None--"
Public const PCC_request_reasonable =  "PCC request reasonable\? PCC request reasonable\? Help Info --None--"
Public const Further_discussion_needed = "Further discussion needed\? --None--"
Public const Type_3_hybrid_study_For_BPS_DIHC_only = "\*Type 3 hybrid study\? \(For BPS DIHC only\) Type 3 hybrid study\? \(For BPS DIHC only\) Help Info --None--"
Public const Rejection_Criteria = "Rejection Criteria Rejection Criteria Help Info --None--"
Public const Responsive_to_Research_Areas_of_Interest = "Responsive to Research Areas of Interest --None--"
Public const Research_Areas_of_Interest_Comments = "Research Areas of Interest Comments Research Areas of Interest Comments Help Info"
Public const Includes_Clinical_Practice_Guidelines = "Includes Clinical Practice Guidelines\? --None--"
Public const Clinical_Practice_Guidelines_Comments = "Clinical Practice Guidelines Comments Clinical Practice Guidelines Comments Help Info"
Public const Does_LOI_include_CER_or_CEA = "Does LOI include CER or CEA\? Does LOI include CER or CEA\? Help Info --None--"
Public const Includes_CER_or_CEA_comments = "Includes CER or CEA comments Includes CER or CEA comments Help Info"
Public const Includes_PCORnet = "Includes PCORnet\? --None--"
Public const Includes_PCORnet_Comments = "Includes PCORnet Comments Includes PCORnet Comments Help Info"
Public const Foreign_Institution = "Foreign Institution\? --None--"
Public const PCORI_Research_Priority_Area = "\*PCORI Research Priority Area PCORI Research Priority Area Help Info --None--"
Public const Greater_Than_Budget_Approval = "\*Greater Than Budget Approval Greater Than Budget Approval Help Info --None--"
Public const Decisional_Dilemma = "Decisional Dilemma Decisional Dilemma Help Info --None--"
Public const Appropriate_setting = "Appropriate setting\? Appropriate setting\? Help Info --None--"
Public const Evidence_of_stakeholder_consultation = "Evidence of stakeholder consultation\? Evidence of stakeholder consultation\? Help Info --None--"
Public const RCT_using_strong_design_methods = "RCT using strong design/methods\? RCT using strong design/methods\? Help Info --None--"
Public const Is_a_DCC_named = "\*Is a DCC named\? Is a DCC named\? Help Info --None--"
Public const Dual_PIs_CCC_DCC = "Dual-PIs CCC/DCC\? Dual-PIs CCC/DCC\? Help Info --None--"
Public const Activities_feasibility_ph_appropriate = "Activities feasibility ph\. appropriate\? Activities feasibility ph\. appropriate\? Help Info --None--"
Public const PI_team_success_trials_of_similar_size = "PI/team success trials of similar size\? PI/team success trials of similar size\? Help Info --None--"
Public const General_proposal_rating = "General proposal rating General proposal rating Help Info --None--"
Public const Compelling_gap_evidence_identified = "Compelling gap in evidence identified\? Compelling gap in evidence identified\? Help Info --None--"
Public const Framework = "Framework Framework Help Info --None--"
Public const Engagement = "Engagement Engagement Help Info --None--"
Public const Participants_appropriate = "Participants appropriate\? --None--"
Public const Cost = "Cost Cost Help Info --None--"
Public const Primary_outcome_measures_include_Sleep = "\*Primary outcome measures include Sleep Primary outcome measures include Sleep Help Info --None--"
Public const Maternal_Health_Primary_Outcome_Measure = "Maternal Health Primary Outcome Measure --None--"
Public const Overlap_with_Funded_Portfolio = "Overlap with Funded Portfolio Overlap with Funded Portfolio Help Info --None--"
Public const LOI_Enthusiasm = "LOI Enthusiasm LOI Enthusiasm Help Info --None--"
Public const Population_appropriate_PLACER = "\*Population appropriate\?"
Public const Outcomes_PLACER = "\*Outcomes"
Public const  PI_team_success_trials_of_similar_size_PLACER = "\*PI/team success trials of similar size\?"
Public const Interventions  = "\*Interventions Interventions Help Info --None--"

''''''''''''''''''''''BPS/Methods/PLACER/SOE/Sleephealth/MMM/IDD LOI Review form tooltip text'''''''''''''''
Public const Category_3_Contact_Front_Door_Tooltip = "For BPS Category 3 I&I Reviewer, did Investigator Team contact Front Door\?"
Public const Special_Area_of_Emphasis_SAE_Tooltip = "Does the LOI address the SAE\?"
Public const Primary_National_Priority_Aims_Aligned_Tooltip = "Do the specific aims align with the \(Primary\) designated PCORI national priority area\?"
Public const PCORI_National_Priority_Area_Fit_Tooltip = "Which PCORI National priority area is a better fit for this LOI\?"
Public const Compelling_research_question_Tooltip = "Does the LOI justify that the research question is of critical importance to relevant decision makers, requiring compelling evidence of the benefits and harms of different treatment options\?"
Public const Systematic_rev_evidence_synthesis_cited_Tooltip = "Does the LOI appropriately cite current systematic reviews or other high-quality evidence syntheses that documents a research need or gap\?"
Public const Study_Design_Tooltip = "Will the proposed study design reasonably answer the proposed research question\?"
Public const Population_appropriate_Tooltip = "Are the study population and prespecified subgroups appropriate to answer the research question\?"
Public const Interventions_have_efficacy_or_in_use_Tooltip =  "Are the proposed interventions appropriate and have documentation of efficacy and/or use\?"
Public const Usual_Care_Comparator_Tooltip = "If usual care is proposed, is justification of its inclusion provided\?"
Public const Outcomes_Tooltip = "Are the primary and secondary outcomes of the study clearly described, and are they relevant to stakeholders and/or patient-centered\?"
Public const Sample_size_power_assumption_appropriate_Tooltip = "Is the sample size appropriate and justified and are the power calculations and assumptions provided\?"
Public const Hypothesized_Effect_Size_Tooltip = "Is the hypothesized effect size realistic\?"
Public const Appropriate_duration_Tooltip = "Is the duration appropriate and justified\?"
Public const Analysis_plan_HTE_appropriate_Tooltip = "Does the statistical analysis plan appropriately correspond to the major aims and outcomes \(HTE where appropriate\?\)"
Public const Patient_stakeholder_community_engagement_Tooltip =  "Has the applicant consulted or planned to consult relevant stakeholders to help determine that the proposed study addresses their need for relevant evidence in decisions\?"
Public const Does_applicant_have_adequate_experience_Tooltip = "Did the LOI provide adequate relevant experience for the PI and/or research team\?"
Public const PCORI_Research_Priority_Area_Key_Topic_Tooltip = "When selecting a value, ensure the LOI aligns broadly with the general research priority area/key topic"
Public const PCC_coverage_requested_Tooltip = "Patient Care Costs: Does the LOI indicate PCC's are included\? E\.g\. costs for the intervention being studied as well as clinical personnel costs\."
Public const PCC_request_reasonable_Tooltip = "If yes to PCCs, are the reasons for including them described, well-supported and justified in the LOI and it is clear they are needed for the success of the study\." 
Public const Type_3_hybrid_study_For_BPS_DIHC_only_Tooltip = "Type 3 hybrid studies assess both CER and implementation, with a focus primarily on implementation\. These studies test an implementation strategy while also observing and gathering information on the clinical intervention’s impact on relevant outcomes\."
Public const Rejection_Criteria_Tooltip = "Indicate primary reason for rejection \(Not CER includes pilot study, guideline dev, coverage rec, instrument dev, single intervention study design, patient characteristics vs clinical options, redundancy portfolio\)"
Public const Research_Areas_of_Interest_Comments_Tooltip = "If responsive to RAI\(s\), please specify"
Public const Clinical_Practice_Guidelines_Comments_Tooltip = "If LOI includes CPG, please explain why responsive"
Public const Does_LOI_include_CER_or_CEA_Tooltip = "Includes CEA\?"
Public const Includes_CER_or_CEA_comments_Tooltip = "Includes Cost Effectiveness \(CER\) or Cost Effectiveness Analysis \(CEA\)\?"
Public const Includes_PCORnet_Comments_Tooltip = "If LOI includes PCORnet, please provide explanation\."
Public const PCORI_Research_Priority_Area_Tooltip = "When selecting a value, ensure the LOI aligns broadly with the general research priority area"
Public const Greater_Than_Budget_Approval_Tooltip = "If the LOI is requesting to exceed the allowable budget, please indicate\:"
Public const Decisional_Dilemma_Tooltip = "Does the LOI convincingly make a case that the specific interventions proposed represent a compelling decisional dilemma\?"
Public const Appropriate_setting_Tooltip = "Is the setting in which the proposed research is slated to take place appropriate to answer the research question\?"
Public const Evidence_of_stakeholder_consultation_Tooltip = "Does the LOI provide evidence of appropriate stakeholder consultation\?"
Public const RCT_using_strong_design_methods_Tooltip = "Are the study design/methods appropriate\? Note: must be an RCT with the unit of randomization at either the individual or cluster level"
Public const Is_a_DCC_named_Tooltip = "Was a DCC named in place at the time of LOI submission\?"
Public const Dual_PIs_CCC_DCC_Tooltip = "Are the CCC and DCC leads named as dual-PIs\?"
Public const Activities_feasibility_ph_appropriate_Tooltip = "Are the activities named in the feasibility phase portion appropriate to support the full-scale study phase research\?"
Public const PI_team_success_trials_of_similar_size_Tooltip = "Do the PI/study team have success in completing trials of similar complexity and scope\?"
Public const General_proposal_rating_Tooltip = "General proposal rating\?"
Public const Compelling_gap_evidence_identified_Tooltip = "Does the LOI document an evidence gap that this proposed research will fill\?"
Public const Framework_Tooltip = "Did the applicant include a conceptual/theoretical framework that will guide the proposed study\? I\.e\. have they outlined what they are doing and how they are doing it\?"
Public const Engagement_Tooltip = "Is there key engagement with relevant patients and stakeholders for the purposes of proposal development, study implementation, and dissemination\?"
Public const Cost_Tooltip = "Do the costs of the study seem reasonable given the scope and duration\?"
Public const Primary_outcome_measures_include_Sleep_Tooltip = "For Sleep Health, at least one primary outcome measure related to sleep\?"
Public const Overlap_with_Funded_Portfolio_Tooltip = "Confirm this LOI does not overlap with previously PCORI-funded studies"
Public const Systematic_rev_evidence_synthesis_cited_accName = "Systematic rev/evidence synthesis cited\?"
Public const LOI_Enthusiasm_Tooltip = "For accepted LOIs, what is your level of enthusiasm\? \(Based on compelling RQ, clarity, rigor, etc\)\?"
Public const Interventions_Tooltip = "Does the study propose continuing or adding interventions\?"

''''''''Need to update here after each SF Releases if value changed''''''''''
'''''''''''''''''''''BPS/Methods/PLACER/SOE/Sleephealth/MMM/IDD  LOI Review form Weblist Acc Name PQT''''''''''
Public const Category_3_Contact_Front_Door_accName = "Category 3: Contact Front Door\?"
Public const Special_Area_of_Emphasis_SAE_accName = "Special Area of Emphasis \(SAE\)"
Public const Responsive_to_Research_Areas_of_Interest_accName = "Responsive to Research Areas of Interest"
Public const Includes_Clinical_Practice_Guidelines_accName = "Includes Clinical Practice Guidelines\?"
Public const Does_LOI_include_CER_or_CEA_accName = "Does LOI include CER or CEA\?"
Public const Includes_PCORnet_accName = "Includes PCORnet\?"
Public const Foreign_Institution_accName = "Foreign Institution\?"
Public const PCORI_Research_Priority_Area_accName = "PCORI Research Priority Area"
Public const Greater_Than_Budget_Approval_accName = "Greater Than Budget Approval"
Public const Decisional_Dilemma_accName = "Decisional Dilemma"
Public const Appropriate_setting_accName = "Appropriate setting\?"
Public const Evidence_of_stakeholder_consultation_accName = "Evidence of stakeholder consultation\?"
Public const RCT_using_strong_design_methods_accName = "RCT using strong design/methods\?"
Public const Is_a_DCC_named_accName = "Is a DCC named\?"
Public const Dual_PIs_CCC_DCC_accName = "Dual-PIs CCC/DCC\?"
Public const Activities_feasibility_ph_appropriate_accName = "Activities feasibility ph\. appropriate\?"
Public const PI_team_success_trials_of_similar_size_accName = "PI/team success trials of similar size\?"
Public const General_proposal_rating_accName = "General proposal rating, --None--"
Public const Compelling_gap_evidence_identified_accName = "Compelling gap in evidence identified\?"
Public const Framework_accName = "Framework"
Public const Engagement_accName = "Engagement"
Public const Participants_appropriate_accName = "Participants appropriate\?"
Public const Outcomes_accName = "Outcomes"
Public const Does_applicant_have_adequate_experience_accName = "Does applicant have adequate experience\?"
Public const Cost_accName = "Cost"
Public const Appropriate_duration_accName = "Appropriate duration\?"
Public const Patient_stakeholder_community_engagement_accName = "Patient/stakeholder/community engagement"
Public const Primary_outcome_measures_include_Sleep_accName ="Primary outcome measures include Sleep"
Public const Maternal_Health_Primary_Outcome_Measure_accName = "Maternal Health Primary Outcome Measure"
Public const Compelling_research_question_accname = "Compelling research question\?"
Public const Population_appropriate_accname = "Population appropriate\?"
Public const Interventions_have_efficacy_or_in_use_accName = "Interventions have efficacy or in use\?"
Public const Study_Design_accName = "Study Design"
Public const Sample_size_power_assumption_appropriate_accName = "Sample size/power assumption appropriate"
Public const PCC_coverage_requested_accName = "PCC coverage requested\?"
Public const PCC_request_reasonable_accName = "PCC request reasonable\?"
Public const Overlap_with_Funded_Portfolio_accName = "Overlap with Funded Portfolio"
Public const Rejection_Criteria_accName = "Rejection Criteria"
Public const Type_3_hybrid_study_For_BPS_DIHC_only_accName = "Type 3 hybrid study\? \(For BPS DIHC only\)"
Public const Further_discussion_needed_accName = "Further discussion needed\?"
Public const PCORI_Research_Priority_Area_Key_Topic_accName = "PCORI Research Priority Area/Key Topic"
Public const Analysis_plan_HTE_appropriate_accName = "Analysis plan/HTE appropriate\?"
Public const Hypothesized_Effect_Size_accName = "Hypothesized Effect Size"
Public const Usual_Care_Comparator_accName = "Usual Care Comparator\?"
Public const LOI_Enthusiasm_accName = "LOI Enthusiasm"
Public const Interventions_accName = "Interventions"

''''''''Need to update here after each SF Releases if value changed''''''''''
''''''''''''''''''''''''''''''''''BPS/Methods/PLACER/SOE/Sleephealth/MMM/IDD  LOI Review form Weblist Acc Name for Prod''''''''''''''''''

Public const Compelling_research_question_accname_Prod = "Compelling research question\?"
Public const Special_Area_of_Emphasis_SAE_accName_Prod = "Special Area of Emphasis \(SAE\)"
Public const Population_appropriate_accname_Prod = "Population appropriate\?"
Public const Interventions_have_efficacy_or_in_use_accName_Prod = "Interventions have efficacy or in use\?"
Public const Outcomes_accName_Prod = "Outcomes"
Public const Study_Design_accName_Prod = "Study Design"
Public const Category_3_Contact_Front_Door_accName_Prod = "Category 3: Contact Front Door\?"
Public const Sample_size_power_assumption_appropriate_accName_Prod = "Sample size/power assumption appropriate"
Public const Patient_stakeholder_community_engagement_accName_Prod = "Patient/stakeholder/community engagement"
Public const PCC_coverage_requested_accName_Prod = "PCC coverage requested\?"
Public const PCC_request_reasonable_accName_Prod = "PCC request reasonable\?"
Public const Overlap_with_Funded_Portfolio_accName_Prod = "Overlap with Funded Portfolio"
Public const Rejection_Criteria_accName_Prod = "Rejection Criteria"
Public const Responsive_to_Research_Areas_of_Interest_accName_Prod = "Responsive to Research Areas of Interest"
Public const Includes_Clinical_Practice_Guidelines_accName_Prod = "Includes Clinical Practice Guidelines\?"
Public const Does_LOI_include_CER_or_CEA_accName_Prod = "Does LOI include CER or CEA\?"
Public const Includes_PCORnet_accName_Prod = "Includes PCORnet\?"
Public const Foreign_Institution_accName_Prod = "Foreign Institution\?"
Public const PCORI_Research_Priority_Area_accName_Prod = "PCORI Research Priority Area"
Public const Greater_Than_Budget_Approval_accName_Prod = "Greater Than Budget Approval"
Public const Decisional_Dilemma_accName_Prod = "Decisional Dilemma"
Public const Appropriate_setting_accName_Prod = "Appropriate setting\?"
Public const Evidence_of_stakeholder_consultation_accName_Prod = "Evidence of stakeholder consultation\?"
Public const RCT_using_strong_design_methods_accName_Prod = "RCT using strong design/methods\?"
Public const Is_a_DCC_named_accName_Prod = "Is a DCC named\?"
Public const Dual_PIs_CCC_DCC_accName_Prod = "Dual-PIs CCC/DCC\?"
Public const Activities_feasibility_ph_appropriate_accName_Prod = "Activities feasibility ph\. appropriate\?"
Public const PI_team_success_trials_of_similar_size_accName_Prod = "PI/team success trials of similar size\?"
Public const General_proposal_rating_accName_Prod = "General proposal rating, --None--"
Public const Compelling_gap_evidence_identified_accName_Prod = "Compelling gap in evidence identified\?"
Public const Framework_accName_Prod = "Framework"
Public const Engagement_accName_Prod = "Engagement"
Public const Participants_appropriate_accName_Prod = "Participants appropriate\?"
Public const Does_applicant_have_adequate_experience_accName_Prod = "Does applicant have adequate experience\?"
Public const Cost_accName_Prod = "Cost"
Public const Appropriate_duration_accName_Prod = "Appropriate duration\?"
Public const Primary_outcome_measures_include_Sleep_accName_Prod ="Primary outcome measures include Sleep"
Public const Maternal_Health_Primary_Outcome_Measure_accName_Prod = "Maternal Health Primary Outcome Measure"
Public const Type_3_hybrid_study_For_BPS_DIHC_only_accName_Prod = "Type 3 hybrid study\? \(For BPS DIHC only\)"
Public const Further_discussion_needed_accName_Prod = "Further discussion needed\?"
Public const PCORI_Research_Priority_Area_Key_Topic_accName_Prod = "PCORI Research Priority Area/Key Topic"
Public const Analysis_plan_HTE_appropriate_accName_Prod = "Analysis plan/HTE appropriate\?"
Public const Hypothesized_Effect_Size_accName_Prod = "Hypothesized Effect Size"
Public const Usual_Care_Comparator_accName_Prod = "Usual Care Comparator\?"
Public const LOI_Enthusiasm_accName_Prod = "LOI Enthusiasm"
Public const Interventions_accName_Prod = "Interventions"




''''''''''''''''''D&I (LC/IMRI/SDM/HSII) LOI Review Form fields edit mode'''''''''''
Public const Project_type = "\*Project type\? Project type\? Help Info --None--"
Public const Will_results_be_available_in_time = "Will results be available in time\? Will results be available in time\? Help Info --None--"
Public const Is_the_original_PCORI_evidence_strong = "Is the original PCORI evidence strong\? --None--"
Public const Does_it_contribute_to_existing_evidence = "Does it contribute to existing evidence\? --None--"
Public const Primary_focus_on_dissemination = "Primary focus on dissemination\? --None--"
Public const Clear_implementation_objective_aims = "Clear project objective/aims\? Clear project objective/aims\? Help Info --None--"
Public const Clear_implementation_approach = "Clear approach/strategies\? Clear approach/strategies\? Help Info --None--"
Public const End_users_described = "End users described\? End users described\? Help Info --None--"
Public const Imp_sites_described_committed = "Sites described & committed\? Sites described & committed\? Help Info --None--"
Public const Appropriate_reach = "Appropriate reach\? --None--"
Public const Unresponsive_LOI = "Unresponsive LOI\? Unresponsive LOI\? Help Info --None--"
Public const Logical_Next_Step = "Logical Next Step\? Logical Next Step\? Help Info --None--"
Public const Eval_plan_include_metrics_of_success = "Eval plan include metrics of success\? Eval plan include metrics of success\? Help Info --None--"
Public const Key_stakeholder_engagement = "Key stakeholder engagement\? Key stakeholder engagement\? Help Info --None--"
Public const Conceptual_theoretical_framework_used = "Conceptual/theoretical framework used\? Conceptual/theoretical framework used\? Help Info --None--" 
Public const Will_project_help_improve_healthcare = "Will project help improve health\(care\)\? Will project help improve health\(care\)\? Help Info --None--"
Public const Implementation_aims_objectives_stated = "Implementation aims/objectives stated\? Implementation aims/objectives stated\? Help Info --None--"
Public const Decision_context_identified = "Decision context identified\? Decision context identified\? Help Info --None--"
Public const Implementation_gap_described = "Implementation gap described\? Implementation gap described\? Help Info --None--"
Public const Results_to_be_implemented_described = "Results to be implemented described\? Results to be implemented described\? Help Info --None--"
Public const Implementation_setting_described = "Implementation setting described\? Implementation setting described\? Help Info --None--"
Public const SDM_implementation_approach_described = "SDM implementation approach described\? SDM implementation approach described\? Help Info --None--"
Public const Will_SDM_approach_provide_useful_info = "Will SDM approach provide useful info\? Will SDM approach provide useful info\? Help Info --None--"
Public const Clear_project_goals_objective = "Clear project goals and objectives\? --None--"
Public const Major_concerns = "Major concerns\? --None--"

''''''''''''''''''''''''''''''''''''D&I (LC/IMRI/SDM) LOI Review Form Tool tip Text '''''''''''
Public const Will_results_be_available_in_time_Tooltip = "Do you think either: 1\) a DFRR for the original PCORI award will be accepted for entry into Peer Review or 2\) a manuscript on the PCORI results to be implemented will be accepted for publication by a peer-reviewed scientific journal by the app due date\?"
Public const End_users_described_Tooltip = "Are the targeted end users of the implementation project clearly described\?"
Public const Unresponsive_LOI_Tooltip = "Does the proposed approach contain unresponsive activities, such as passive dissemination strategies or the development/validation of a new tool/system without the primary purpose of disseminating or implementing evidence\?"
Public const Logical_Next_Step_Tooltip = "In your opinion, is the proposed project a logical and feasible next step for dissemination/implementation of these results\?"
Public const Eval_plan_include_metrics_of_success_Tooltip = "Is the evaluation plan clearly described, including measurable indicators of success\?"
Public const Key_stakeholder_engagement_Tooltip = "Is there key engagement with relevant patients and stakeholders for the purpose of proposal development and project implementation"
Public const Conceptual_theoretical_framework_used_Tooltip = "Did the applicant include a conceptual or theoretical framework to guide the proposed project\?"
Public const Will_project_help_improve_healthcare_Tooltip = "In your opinion, does this project propose to implement results that will improve health/healthcare in the long run, and does this implementation project move us toward that end\?"
Public const Implementation_aims_objectives_stated_Tooltip = "Are the project objectives and specific aims clearly stated\?"
Public const Decision_context_identified_Tooltip = "Did the applicant clearly identify the preference-sensitive decision the proposed SDM strategy addresses and indicate how this new PCORI evidence contributes to patient or provider decision making\?"
Public const Implementation_gap_described_Tooltip = "Did the applicant clearly identify existing barriers to use of the proposed SDM strategy that motivate the proposed implementation project\?"
Public const Results_to_be_implemented_described_Tooltip = "Did the applicant clearly describe the PCORI research findings and related evidence most relevant to their proposed implementation project\?"
Public const Implementation_setting_described_Tooltip = "Is the setting in which implementation will take place clearly described\?"
Public const SDM_implementation_approach_described_Tooltip = "Did the applicant clearly describe the approach for implementing the research findings to end user groups\?"
Public const Will_SDM_approach_provide_useful_info_Tooltip = "In your opinion, is the SDM approach proposed for implementation likely to provide useful information to patients or their caregivers facing health related decisions"
Public const Project_type_Tooltip = "Which LC opportunity is the applicant applying to\?"
Public const Clear_implementation_approach_Tooltip = "Did the applicant clearly describe the approach for implementing the research findings to end user groups\?"

''''''''Need to update here after each SF Releases if value changed''''''''''
'''''''''''''''''''''D&I (LC/IMRI/SDM/HSII) Review form Weblist Acc Name for PQT''''''''''
Public const Project_type_accName = "Project type\?"
Public const Will_results_be_available_in_time_accName = "Will results be available in time\?"
Public const Is_the_original_PCORI_evidence_strong_accName = "Is the original PCORI evidence strong\?"
Public const Does_it_contribute_to_existing_evidence_accName = "Does it contribute to existing evidence\?"
Public const Primary_focus_on_dissemination_accName = "Primary focus on dissemination\?"
Public const Clear_implementation_objective_aims_accName = "Clear project objective/aims\?"
Public const Clear_implementation_approach_accName = "Clear approach/strategies\?"
Public const End_users_described_accName = "End users described\?"
Public const Imp_sites_described_committed_accName = "Sites described & committed\?"
Public const Appropriate_reach_accName = "Appropriate reach\?"
Public const Unresponsive_LOI_accName = "Unresponsive LOI\?"
Public const Logical_Next_Step_accName = "Logical Next Step\?"
Public const Eval_plan_include_metrics_of_success_accName = "Eval plan include metrics of success\?"
Public const Key_stakeholder_engagement_accName = "Key stakeholder engagement\?"
Public const Conceptual_theoretical_framework_used_accName = "Conceptual/theoretical framework used\?"
Public const Will_project_help_improve_healthcare_accName = "Will project help improve health\(care\)\?"
Public const Implementation_aims_objectives_stated_accName = "Implementation aims/objectives stated\?"
Public const Decision_context_identified_accName = "Decision context identified\?"
Public const Implementation_gap_described_accName = "Implementation gap described\?"
Public const Results_to_be_implemented_described_accName = "Results to be implemented described\?"
Public const Implementation_setting_described_accName = "Implementation setting described\?"
Public const SDM_implementation_approach_described_accName = "SDM implementation approach described\?"
Public const Will_SDM_approach_provide_useful_info_accName = "Will SDM approach provide useful info\?"
Public const Clear_project_goals_objective_accName = "Clear project goals and objectives\?"
Public const Major_concerns_accName = "Major concerns\?" 

''''''''Need to update here after each SF Releases if value changed''''''''''
''''''''''''''''''''D&I (LC/IMRI/SDM/HSII) Review form Weblist Acc Name for Prod''''''''''
Public const Project_type_accName_Prod = "Project type\?"
Public const Will_results_be_available_in_time_accName_Prod = "Will results be available in time\?"
Public const Is_the_original_PCORI_evidence_strong_accName_Prod = "Is the original PCORI evidence strong\?"
Public const Does_it_contribute_to_existing_evidence_accName_Prod = "Does it contribute to existing evidence\?"
Public const Primary_focus_on_dissemination_accName_Prod = "Primary focus on dissemination\?"
Public const Clear_implementation_objective_aims_accName_Prod = "Clear project objective/aims\?"
Public const Clear_implementation_approach_accName_Prod = "Clear approach/strategies\?"
Public const End_users_described_accName_Prod = "End users described\?"
Public const Imp_sites_described_committed_accName_Prod = "Sites described & committed\?"
Public const Appropriate_reach_accName_Prod = "Appropriate reach\?"
Public const Unresponsive_LOI_accName_Prod = "Unresponsive LOI\?"
Public const Logical_Next_Step_accName_Prod = "Logical Next Step\?"
Public const Eval_plan_include_metrics_of_success_accName_Prod = "Eval plan include metrics of success\?"
Public const Key_stakeholder_engagement_accName_Prod = "Key stakeholder engagement\?"
Public const Conceptual_theoretical_framework_used_accName_Prod = "Conceptual/theoretical framework used\?"
Public const Will_project_help_improve_healthcare_accName_Prod = "Will project help improve health\(care\)\?"
Public const Implementation_aims_objectives_stated_accName_Prod = "Implementation aims/objectives stated\?"
Public const Decision_context_identified_accName_Prod = "Decision context identified\?"
Public const Implementation_gap_described_accName_Prod = "Implementation gap described\?"
Public const Results_to_be_implemented_described_accName_Prod = "Results to be implemented described\?"
Public const Implementation_setting_described_accName_Prod = "Implementation setting described\?"
Public const SDM_implementation_approach_described_accName_Prod = "SDM implementation approach described\?"
Public const Will_SDM_approach_provide_useful_info_accName_Prod = "Will SDM approach provide useful info\?"
Public const Clear_project_goals_objective_accName_Prod = "Clear project goals and objectives\?"
Public const Major_concerns_accName_Prod = "Major concerns\?" 







'''''''''''''''''For All record Types of RA LOI Review "Review Notes & Programmatic Responsiveness /Alternate PFA Section Pqt'''''''''''
Public const LOI_Accepted_Denied_accName = "LOI Accepted or Denied\?"
Public const Programmatically_Responsive_accName = "Programmatically Responsive\?" 
Public const can_it_move_to_alternate_PFA_accName = "If ""No"" can it move to alternate PFA\?"
Public const If_Yes_then_which_PFA_accName = "If ""Yes"" then which PFA\?"

'''''''''''''''''For All record Types of RA LOI Review "Review Notes Section Prod'''''''''''
Public const LOI_Accepted_Denied_accName_Prod = "LOI Accepted or Denied\?"
Public const Programmatically_Responsive_accName_Prod = "Programmatically Responsive\?" 
Public const can_it_move_to_alternate_PFA_accName_Prod = "If ""No"" can it move to alternate PFA\?"
Public const If_Yes_then_which_PFA_accName_Prod = "If ""Yes"" then which PFA\?"



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Engagement Award realted Variable''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''**************LOI Description Page EA External by outertext***********'''''''''''''''''
Public const LoiDescriptionApplyPage_EADI = "Description The Engagement Award: Dissemination Initiative funding opportunity aims to support projects that help organizations and communities actively communicate pertinent PCORI-funded research findings to their specific audiences, including patients, clinicians, communities, and others, in ways that will command their attention and interest and encourage use of this information in their healthcare decision making\. Instructions Letters of Intent \(LOI\) must be submitted by 5:00 p\.m\. \(ET\) on the deadline date posted in the funding announcement to be reviewed\. Please click on the Apply button to create your LOI for the Engagement Award: Dissemination Initiative funding opportunity\."
Public const LoiDescriptionApplyPage_EACB = "Description The Engagement Award: Capacity Building opportunity funds projects that build communities prepared to participate in patient-centered comparative clinical effectiveness research \(CER\)\. These awards support organizations with strong ties to patients, families, caregivers and the broader health and healthcare community who have a connection to a research focus area and seek to better equip stakeholders to engage as partners in patient-centered CER\. These projects will focus on building the knowledge, competencies and abilities of their community to be meaningful partners with researchers throughout the research process\. Instructions Letters of Intent \(LOI\) must be submitted by 5:00 p\.m\. \(ET\) on the deadline date posted in the funding announcement\. Please click on the Apply button to create your LOI for the Engagement Award: Capacity Building funding opportunity\."
Public const LoiDescriptionApplyPage_EASCS = "Description The Engagement Award: Stakeholder Convening Support funding opportunity provides support to organizations and communities to hold multi-stakeholder convenings, meetings, and conferences that include a combination of patients, caregivers, researchers, clinicians, purchasers, payers, health system leaders, and/or other stakeholders\. These convenings must have a focus on, and commitment to, supporting collaboration around patient-centered outcomes research/comparative clinical effectiveness research \(PCOR/CER\)\. Convenings should be designed with the active collaboration and partnership of patients, community groups, and/or other stakeholder organizations\. Projects should bring together diverse stakeholders around a central focus or shared priority that unifies stakeholders \(e\.g\., geography, health condition, population\) to explore issues related to PCOR/CER or communicate PCORI-funded research findings to targeted end-user audiences\. Instructions Letters of Intent \(LOI\) must be submitted by 5:00 p\.m\. \(ET\) on the deadline date posted in the funding announcement to be reviewed\. Please click on the Apply button to create your LOI for the Engagement Award: Stakeholder Convening Support funding opportunity\."

'''''''''''''''''EA LOI DI by Outertext''''''''''''''''''''''''''''''''
'''''''''''''''''Project Name & Contact Information ''''''''''''''''''''
Public const PNCInfo_DILOI_InsText1 = "You are applying to the Engagement Award: Dissemination Initiative funding opportunity Please utilize the following resources when completing the LOI: PCORI Funding Announcement: Engagement Award: Dissemination Initiative - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Additional Instructions Click ""Save & Next"" to continue to the next tab\. Otherwise you could receive an error message\. To save a draft LOI, you must complete the PL and Organization fields\. A Project Lead \(PL\) and Administrative Official \(AO\) with valid email addresses are required to submit your LOI to PCORI\. To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking “New User\.” Individuals assigned at the ""Contact Information"" tab will have access to the LOI and Full Proposal\. Fields marked with \(\*\) are required\."
Public const Link1_PNCInfo_Name_EALOI_DI = "Engagement Award: Dissemination Initiative - Fall 2024 Cycle"
Public const Link2_PNCInfo_Name_EALOI_DI = "Engagement Award Submission Instructions - Fall 2024 Cycle"
Public const Link3_PNCInfo_Name_EALOI_DI = "Engagement Award Program Resources"
Public const PNCInfo_DILOI_InsText2 = "Required Information The Project Lead 1 \(Contact PL\) and Administrative Official must be an employee or board member of the primary organization/institution on the application\. The Project Lead 1 \(Contact PL\) and Administrative Official cannot be the same individual\."
Public const Project_Name_EADI_LOI = "\* Project Name "
Public const ORG_Name_EADI_LOI = "\* Organization Name If your organization does not appear in the lookup, click ""My Profile"" and then ""Edit Employer Details\."" "
Public const Project_Lead_1_Name_EADI_LOI = "\* Project Lead 1 \(Contact PL\) "
Public const AO_Name_EADI_LOI = "\* Administrative Official "
Public const Project_Lead_2_Name_EADI_LOI = "Project Lead 2 \(co-PL\) "
Public const Project_Lead_Designee_1_Name_EADI_LOI = "Project Lead Designee 1 "
Public const Project_Lead_Designee_2_Name_EADI_LOI = "Project Lead Designee 2 "

'''''''''''''''''Project Name & Contact Information' link for DI'''''''''''''''''''
Public const Link1_PNCInfo_URL_EALOI_DI = "https://www\.pcori\.org/funding-opportunities/announcement/engagement-award-dissemination-initiative-fall-2024-cycle"
Public const Link2_PNCInfo_URL_EALOI_DI = "https://www\.pcori\.org/sites/default/files/PCORI-engagement-award-submission-instructions-fall-2024-cycle\.pdf"
Public const Link3_PNCInfo_URL_EALOI_DI = "https://www\.pcori\.org/funding-opportunities/applicant-and-awardee-resources/applicant-resources/pcori-online-resources-applicants-engagement-award-program"

'''''''''''''''''EA LOI SCS by Outertext''''''''''''''''''''''''''''''''
'''''''''''''''''Project Name & Contact Information''''''''''''''''''''
Public const PNCInfo_SCS_LOI_InsText1 = "You are applying to the Engagement Award: Stakeholder Convening Support funding opportunity Please utilize the following resources when completing the LOI: PCORI Funding Announcement: Engagement Award: Stakeholder Convening Support - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Additional Instructions Click ""Save & Next"" to continue to the next tab\. Otherwise you could receive an error message\. To save a draft LOI, you must complete the PL and Organization fields\. A Project Lead \(PL\) and Administrative Official \(AO\) with valid email addresses are required to submit your LOI to PCORI\. To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking “New User\.” Individuals assigned at the ""Contact Information"" tab will have access to the LOI and Full Proposal\. Fields marked with \(\*\) are required\. "
Public const Link1_PNCInfo_Name_EALOI_SCS = "Engagement Award: Stakeholder Convening Support - Fall 2024 Cycle"

'''''''''''''''''Project Name & Contact Information' link for SCS'''''''''''''''''''
Public const Link1_PNCInfo_URL_EALOI_SCS = "https://www\.pcori\.org/funding-opportunities/announcement/engagement-award-stakeholder-convening-support-april-2024-cycle"

'''''''''''''''''EA LOI CB by Outertext''''''''''''''''''''''''''''''''
'''''''''''''''''Project Name & Contact Information ''''''''''''''''''''
Public const PNCInfo_CBLOI_InsText1 = "You are applying to the Engagement Award: Capacity Building funding opportunity Please utilize the following resources when completing the LOI: PCORI Funding Announcement: Engagement Award: Capacity Building - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Additional Instructions Click ""Save & Next"" to continue to the next tab\. Otherwise you could receive an error message\. To save a draft LOI, you must complete the PL and Organization fields\. A Project Lead \(PL\) and Administrative Official \(AO\) with valid email addresses are required to submit your LOI to PCORI\. To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking “New User\.” Individuals assigned at the ""Contact Information"" tab will have access to the LOI and Full Proposal\. Fields marked with \(\*\) are required\."


''''''''''''''''Project Name & Contact Information' link for CB'''''''''''''''''''
Public const Link1_PNCInfo_URL_EALOI_CB = "https://www\.pcori\.org/funding-opportunities/announcement/engagement-award-capacity-building-fall-2024-cycle"
Public const Link1_PNCInfo_Name_EALOI_CB = "Engagement Award: Capacity Building - Fall 2024 Cycle"
Public const Link2_PNCInfo_Name_EALOI_CB = "Engagement Award Submission Instructions - Fall 2024 Cycle"
Public const Link3_PNCInfo_Name_EALOI_CB = "Engagement Award Program Resources"

'''''''''''''''''EA App DI by Outertext''''''''''''''''''''''''''''''''
'''''''''''''''''Project Name & Contact Information ''''''''''''''''''''
Public const PNCInfo_DIApp_InsText1 = "You are applying to the Engagement Award: Dissemination Initiative funding opportunity Please utilize the following resources when completing the proposal: PCORI Funding Announcement: Engagement Award: Dissemination Initiative - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - April 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Additional Instructions Click ""Save & Next"" to continue to the next tab\. Otherwise you could receive an error message\. To save a draft proposal, you must complete the Project Name, Project Lead 1 \(Contact PL\), Administrative Official, and Organization fields on this tab\. A Project Lead \(PL\) and Administrative Official \(AO\) with valid email addresses are required to submit your proposal to PCORI\. To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking “New User\.” Individuals assigned at the ""Contact Information"" tab will have access to the LOI and Full Proposal\. Fields marked with \(\*\) are required\. "

'''''''''''''''''EA App CB by Outertext''''''''''''''''''''''''''''''''
'''''''''''''''''Project Name & Contact Information ''''''''''''''''''''
Public const PNCInfo_CBApp_InsText1 = "You are applying to the Engagement Award: Capacity Building funding opportunity Please utilize the following resources when completing the proposal: PCORI Funding Announcement: Engagement Award: Capacity Building - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Additional Instructions Click ""Save & Next"" to continue to the next tab\. Otherwise you could receive an error message\. To save a draft proposal, you must complete the Project Name, Project Lead 1 \(Contact PL\), Administrative Official, and Organization fields on this tab\. A Project Lead \(PL\) and Administrative Official \(AO\) with valid email addresses are required to submit your proposal to PCORI\. To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking “New User\.” Individuals assigned at the ""Contact Information"" tab will have access to the LOI and Full Proposal\. Fields marked with \(\*\) are required\."
Public const Pcori_Online_LinkURL_CB = "https://pcori\.my\.site\.com/engagement"

'''''''''''''''''''Pre-Screen Questionnaire Tab DI LOI EA'''''''''
Public const PSQ_DILOI_InsText1 = "Please utilize the following resources when completing the LOI: PCORI Funding Announcement: Engagement Award: Dissemination Initiative - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources"
Public const Non_Responsive_EALOI_DI = "\* Categories of nonresponsiveness The Engagement Award Program does not fund projects that: Do not have a clear focus on patient-centered comparative clinical effectiveness research \(CER\)\. Include a cost-effectiveness analysis of alternative approaches to providing care\. Solely intend to increase patient engagement in health care or healthcare systems rather than healthcare research, particularly patient-centered CER\. Design or test healthcare interventions\. Involve the use of a drug or medical device\. Create clinical practice guidelines, care protocols, or decision support tools\. Create coverage, payment, or policy recommendations or guidelines\. Address only quality measures, quality improvement, or engagement around quality measures\. Recruit and enroll patients for clinical trials\. Involve patients only as subjects\. Are research studies, such as randomized controlled trials, observational studies, pragmatic clinical studies, and systematic reviews\. Create or maintain a registry or recruit people for a registry\. Are designed solely to validate tools or instruments not created through a PCORI-funded project\. Develop funding requests via proposals, grant applications, or other means\. Give grants using PCORI's project funds\. Focus solely on social determinants of health, with no focus on patient-centered CER\. Plan to disseminate research results without including PCORI-funded research or related products\. Implement PCORI-funded findings in a clinical practice setting\. PCORI funds this type of project in a different program\. Aim to influence, directly or indirectly, any federal, state, or local laws, regulations, judicial decisions, or the like, including preparation or planning activities, research, and other background work related to or in contemplation of lobbying activities\. Aim to create an independent corporation \(nonprofit or for-profit\), limited liability company, partnership, or other legal entity\. Does your project contain any of these nonresponsive activities\? --None-- Yes No "

'''''''''''''''''''Pre-Screen Questionnaire Tab SCS LOI EA'''''''''
Public const PSQ_SCS_LOI_InsText1 = "Please utilize the following resources when completing the LOI: PCORI Funding Announcement: Engagement Award: Stakeholder Convening Support - Falll 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - April 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources"

'''''''''''''''''''Pre-Screen Questionnaire Tab CB LOI EA'''''''''
Public const PSQ_CB_LOI_InsText1 = "Please utilize the following resources when completing the LOI: PCORI Funding Announcement: Engagement Award: Capacity Building - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources"

'''''''''''''''''''Pre-Screen Questionnaire Tab DI App EA'''''''''
Public const PSQ_DIApp_InsText1 = "Please utilize the following resources when completing the proposal: PCORI Funding Announcement: Engagement Award: Dissemination Initiative - April 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - April 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources"


'''''''''''''''''''Pre-Screen Questionnaire Tab CB App EA'''''''''
Public const PSQ_CBApp_InsText1 = "Please utilize the following resources when completing the proposal: PCORI Funding Announcement: Engagement Award: Capacity Building - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources"


''''''''''''''''''''''Organization & Project Lead Details Tab for EA DI ''''''''''''''''''
Public const Primary_Group_Identification_EALOI_DI = "\* \(1\) With which group does the Project Lead primarily identify for the purpose of this project\? --None-- Patient/Consumer Caregiver/Family member of patient Patient/Caregiver Advocacy Organization Clinician Clinic/Hospital/Health System Purchaser Payer Industry Research Policy Maker Training Institution "
Public const Project_Lead_Degree_EALOI_DI = "\* \(2\) What degree\(s\) does the Project Lead have\? AAS AB APRN BA BC BCH BCHIR BM BMBC BMEDS BPHAR BS BSC BSN CHB CHM DA DBA DC DCH DDS DES DM DMD DMH DMS DNSC DO DPH DPHIL DMP DSC DVM EDD HS JD LPN MA MB MBA MBBC MBCHB MCHIR MD MED MIS MLIS MN MPA MPH MPHIL MS MSN MSURG MSW ND NN NP PA PHD PHRMD PTA RN SB SCD Other \(please specify\) "
Public const Project_Lead_Degree_Other_EALOI_DI = "\(2a\) If the Project Lead has other degrees not listed, please specify "
Public const Describe_project_lead_team_previous_exp_EALOI_DI = "\* \(3\) Describe the Project Lead and project team's previous experience similar to this project, including their patient-centered comparative clincial effectiveness research \(CER\) experience Share the project lead, team, and organization's experience successfully engaging stakeholders and communities\. Provide examples of successful engagement efforts\. Include examples of the project team's patient-centered CER experience, if any\. \(Applicants may not exceed 2,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Key_Personnel_funded_by_PCORI_EALOI_DI  = "\* \(4\) Has the Project Lead or another member of the project team been previously funded by PCORI\? --None-- Yes No "
Public const Past_PCORI_funding_information_EALOI_DI = "\(4a\) If yes, provide the name of the project team member\(s\) and their previous PCORI contract number\(s\) \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "
Public const Leveraging_Past_PCORI_Funded_Projects_EALOI_DI = "\* \(5\) If this application leverages the project team's previous work funded by PCORI, please describe how If not applicable, enter N/A in the box below\. 0 of 1000 Characters "
Public const Leadership_Plan_EALOI_DI = "\* \(6\) Describe the project team's leadership plan Briefly explain how the project team \(individuals identified on the ""Contact Information"" tab and other key personnel\) will share the responsibilities for this project\. 0 of 1000 Characters "
Public const Organizational_Summary_EALOI_DI = "\* \(7\) Summarize your organization's history, capacity and mission \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 2,050\.\) 0 of 2050 Characters "
Public const Describe_Unique_Capabilities_EALOI_DI = "\* \(8\) Describe the unique capabilities the Project Lead, project team, and organization have to address the problem identified in this project \(Applicants may not exceed 2,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 2,050\.\) 0 of 2050 Characters "
Public const Website_EALOI_DI = "\* \(9\) What is the URL of your organization's website\? Enter ""n/a"" if your organization does not have a website "
Public const Organization_Financial_Status_EALOI_DI = "\* \(10\) Organization's Financial Status --None-- Sole Proprietor/Consultant Partnership Corporation Not-for-profit \(501c3\) Other "
Public const Financial_Status_Other_EALOI_DI = "\(10a\) If other financial status, please specify "
Public const EIN_Number_EALOI_DI = "\* \(11\) Organization's Employer Identification Number \(EIN\) "
Public const Current_Organizational_Budget_Total_EALOI_DI = "\* \(12\) Provide your organization's current annual operating expenses "
Public const Organization_Fiscal_Calendar_EALOI_DI = "\* \(13\) Which best describes your organization's fiscal year calendar\? --None-- Jan - Dec Oct - Sept Other "
Public const Organization_Fiscal_Calendar_Other_EALOI_DI = "\(13a\) If other fiscal calendar, please specify "
Public const Previous_involvement_PCORI_EALOI_DI = "\* \(14\) Have you interacted with PCORI in the past in any of the following ways\? Select all that apply\. Joined a PCORI email list Visited PCORI’s website Participated in applicant training Watched a PCORI webinar Attended PCORI sponsored event in-person Attended event where PCORI was featured Met with PCORI staff Met with a PCORI Ambassador Applied to review PCORI funding app Applied for PCORI funding Received PCORI funding Served as a PCORI Merit reviewer Participated in a PCORI Advisory Panel Other \(please specify\) None of the above "
Public const Previous_involvement_PCORI_Other_EALOI_DI = "\(14a\) If other interaction with PCORI, please specify "
Public const Proj_Lead_Funding_Sources_Past_5_yr_EALOI_DI = "\* \(15\) Have you been the PI/Project Lead for a research grant/contract from the following organizations\? PCORI AHRQ CDC Other Fed Gov agency \(please specify\) Private Foundation \(please specify\) Other \(please specify\) None of the above "
Public const Other_Organizations_Federal_Government_EALOI_DI = "\(15a\) If other federal government organization, please specify "
Public const Other_Organizations_Private_Foundation_EALOI_DI = "\(15b\) If private foundation, please specify "
Public const Other_Organizations_Other_EALOI_DI = "\(15c\) If other organization, please specify "
Public const EIN_Number_Error = "Error: EIN number should be in XXX-XX-XXXX or XX-XXXXXXX numeric format "

'''''''''''''''''''''Organization & Project Lead Details Tab for EA SCS ''''''''''''''''''

Public const Describe_project_lead_team_previous_exp_EALOI_SCS = "\* \(3\) Describe the Project Lead and project team's previous experience similar to this project, including their patient-centered comparative clincial effectiveness research \(CER\) experience Share the project lead, team, and organization's experience successfully planning and facilitating convenings\. Provide examples of successful prior efforts\. Include examples of the project team's patient-centered CER experience, if any\. \(Applicants may not exceed 2,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "

'''''''''''''''''''''Organization & Project Lead Details Tab for EA CB ''''''''''''''''''

Public const Describe_project_lead_team_previous_exp_EALOI_CB ="\* \(3\) Describe the Project Lead and project team's previous experience similar to this project, including their patient-centered comparative clincial effectiveness research \(CER\) experience Share the project lead, team, and organization's experience successfully engaging stakeholders and communities\. Provide examples of successful engagement efforts\. Include examples of the project team's patient-centered CER experience, if any\. \(Applicants may not exceed 2,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "


''''''''''''''''''''''Organization & Project Lead Details Tab for EA DI 'Application'''''''''''''''''
Public const Primary_Group_Identification_EALOI_DI_App = "\* \(1\) With which group does the Project Lead primarily identify for the purpose of this project\? --None-- Industry Research Caregiver/Family member of patient Clinic/Hospital/Health System Clinician Patient/Caregiver Advocacy Organization Patient/Consumer Payer Policy Maker Purchaser Training Institution "
Public const Project_Lead_Degree_EALOI_DI_App = "\* \(2\) What degree\(s\) does the Project Lead have\? AAS AB APRN BC BCH BCHIR BM BMBC BMEDS BPHAR BS BSC BSN CHB CHM DA DBA DC DCH DDS DES DM DMD DMH DMP DMS DNSC DO DPH DPHIL DSC DVM EDD HS JD LPN MA MB MBA MBBC MBCHB MCHIR MD MED MIS MLIS MN MPA MPH MPHIL MS MSN MSURG MSW ND NN NP PA PHD PHRMD PTA RN SB SCD Other \(please specify\) BA "
Public const Describe_project_lead_team_previous_exp_EALOI_DI_App = "\* \(3\) Describe the Project Lead and project team's previous experience similar to this project, including their patient-centered comparative clinical effectiveness research \(CER\) experience Share the project lead, team, and organization's experience successfully engaging stakeholders and communities\. Provide examples of successful engagement efforts\. Include examples of the project team's patient-centered CER experience, if any\. \(Applicants may not exceed 2,000 characters including spaces for this response\. Disregard the automated counter's listed maximum of 3,050\.\) 59 of 3050 Characters "
Public const Past_PCORI_funding_information_EA_DI_App = "\(4a\) If yes, provide the name of the project team member\(s\) and their previous PCORI contract number\(s\) \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 47 of 1000 Characters "
Public const Leveraging_Past_PCORI_Funded_Projects_EALOI_DI_App = "\* \(5\) If this application leverages the project team's previous work funded by PCORI, please describe how If not applicable, enter N/A in the box below\. 57 of 1000 Characters "
Public const Leadership_Plan_EALOI_DI_App = "\* \(6\) Describe the project team's leadership plan Briefly explain how the project team \(individuals identified on the ""Contact Information"" tab and other key personnel\) will share the responsibilities for this project\. 43 of 1000 Characters "
Public const Organizational_Summary_EALOI_DI_App = "\* \(7\) Summarize your organization's history, capacity and mission \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 53 of 3050 Characters "
Public const Describe_Unique_Capabilities_EALOI_DI_App = "\* \(8\) Describe the unique capabilities the Project Lead, project team, and organization have to address the problem identified in this project \(Applicants may not exceed 2,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 2,050\.\) 51 of 2050 Characters "
Public const Previous_involvement_PCORI_EALOI_DI_App = "\* \(14\) Have you interacted with PCORI in the past in any of the following ways\? Select all that apply\. Visited PCORI’s website Participated in applicant training Watched a PCORI webinar Attended PCORI sponsored event in-person Attended event where PCORI was featured Met with PCORI staff Met with a PCORI Ambassador Applied to review PCORI funding app Applied for PCORI funding Received PCORI funding Served as a PCORI Merit reviewer Participated in a PCORI Advisory Panel Other \(please specify\) None of the above Joined a PCORI email list "
Public const Proj_Lead_Funding_Sources_Past_5_yr_EALOI_DI_App = "\* \(15\) Have you been the PI/Project Lead for a research grant/contract from the following organizations\? AHRQ CDC Other Fed Gov agency \(please specify\) Private Foundation \(please specify\) Other \(please specify\) None of the above PCORI "

''''''''''''''''''''''Organization & Project Lead Details Tab for EA CB 'Application'''''''''''''''''
Public const Past_PCORI_funding_information_EA_CB_App = "\(4a\) If yes, provide the name of the project team member\(s\) and their previous PCORI contract number\(s\) \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 47 of 1000 Characters "

'''''''''''''''''Project Summary Tab for EA DI''''''''''''''''''''
Public const PS_Instext1 = "Engagement Award: Dissemination Initiative Funding Track"
Public const EA_DI_Funding_Track_for_LOI = "\* Which Engagement Award: Dissemination Initiative funding track are you applying to\? Applicants for an Engagement Award: Dissemination Initiative are required to focus their project on one of two tracks, as indicated in the funding announcement: Building Capacity for Dissemination or Active Dissemination --None-- Building Capacity for Dissemination Active Dissemination "
Public const Is_this_resubmission_EALOI_DI = "\* Is this a resubmission of a previously submitted Engagement Award LOI or full proposal\? --None-- Yes No "
Public const Resubmission_Details_EALOI_DI = "\* If this is a resubmission, provide the prior PCORI application number\(s\) and describe how this application has changed from the previous submission\(s\)\. If this is not a resubmission, enter N/A in the box below\. 0 of 1000 Characters "
Public const PCORI_Topic_Theme_EALOI_DI = "\* PCORI Topic Theme \(if applicable\) Does your project align with one of PCORI's Topic Themes\? Funding is not limited to these topics\. PCORI welcomes all LOIs that meet its guidelines\. Topics should address issues important to patients, families, caregivers, and the health care community\. Select all that apply\. Not Applicable Improving outcomes for people with intellectual or developmental disabilities \(IDD\) Promoting health for older adults Promoting healthy children and youth Addressing substance use Addressing violence and trauma Addressing COVID-19 Addressing rare diseases Improving cardiovascular health Improving mental and behavioral health Managing pain Preventing maternal morbidity and mortality \(MMM\) Promoting sleep health "
Public const Background_EALOI_DI = "\* Background State the problem or question this project is designed to address\. Describe the opportunity to promote the accessibility, use, and/or uptake of eligible PCORI-funded research findings\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Proposing_Solution_This_Problem_EALOI_DI = "\* Are you proposing a solution to this problem\? --None-- Yes No "
Public const Explain_Proposed_Solution_EALOI_DI = "\* Project Goal What is the specific goal of this project\? Explain what you hope to achieve through the PCORI-funded work you propose in this application\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Objectives_EALOI_DI = "\* Objectives Clearly state the primary aims and objectives of your project that will lead to accomplishing your goal\(s\)\. Clearly state the primary objective of your dissemination effort\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Methods_EALOI_DI = "\* Methods/Activities Provide a concise description of the project methods, activities, and strategies that will be employed, how they will support the overall goal and how they are connected to each other\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Outcomes_projected_EALOI_DI = "\* Projected Outcomes Specify the projected short-term \(during the PCORI-funded project period\), medium-term \(0-2 years post-project period\), and long-term outcomes \(3\+ years post-project period\) and state their significance\. In addition, clearly describe your intended outputs and deliverables for the PCORI-funded project period\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Patient_Stakeholder_Engagement_Plan_EALOI_DI = "\* Patient and Stakeholder Engagement Plan Briefly explain who the patients and stakeholders are, what if any preexisting engagement has taken place, how patients and stakeholders will be engaged, and how often they will be engaged in the project\. Address how these groups can use and benefit from the proposed dissemination efforts\. Describe how stakeholders and the organizations they may represent were involved in developing the proposed project and this LOI, and how they will continue to be involved in executing the proposed project\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Evaluation_Plan_EALOI_DI = "\* Evaluation Plan Referencing the required ""EADI Reporting Tool"" as a basis for evaluation planning, briefly describe any evaluation to be conducted during this project\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "
Public const Sustainability_Plan_EALOI_DI = "\* Sustainability Plan Briefly describe how the project's outcomes contribute to the overall work your organization does and how the outcomes may be adopted and sustained by your organization or project partners\. If you do not intend to sustain the efforts of this project, please justify why this award will accomplish what is needed with respect to dissemination\. Future funding from PCORI should not be assumed\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "
Public const Link1_PS_Tab_EALOI_DI = "PCORI’s Applicant and Awardee FAQs Related to COVID-19 or Other Unexpected Public Health Emergencies"
Public const Link1_PS_Tab_URL_EALOI_DI = "https://www\.pcori\.org/funding-opportunities/applicant-and-awardee-resources/applicant-and-awardee-faqs-related-covid-19-or-other-unexpected-public-health-emergencies"

'''''''''''''''''Project Summary Tab for EA SCS''''''''''''''''''''
Public const PS_Instext1_SCS = "Engagement Award: Stakeholder Convening Support Funding Track"
Public const EA_SCS_Funding_Track_for_LOI = "\* Which Engagement Award: Stakeholder Convening Support funding track are you applying to\? Applicants for an Engagement Award: Stakeholder Convening Support are required to focus their project on one of two tracks, as indicated in the funding announcement: Convening Around Patient-Centered CER or Convening Around Dissemination of PCORI-Funded Research Findings --None-- Convening Around Patient-Centered CER Convening Around Dissemination of PCORI-Funded Research Findings "
Public const Background_EALOI_SCS = "\* Background State the problem or question this project is designed to address\. Describe the opportunity or need to convene the proposed stakeholders around the proposed topic area or focus\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Objectives_EALOI_SCS = "\* Objectives Briefly describe the aims and goals of the project, including the objectives\. Clearly state the primary objective\(s\) of your convening and the expected outcomes and outputs of the convening\. Identify specific steps you will take to accomplish your overall objective\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Patient_Stakeholder_Engagement_Plan_EALOI_SCS = "\* Patient and Stakeholder Engagement Plan Briefly explain who the patients and stakeholders are, what if any preexisting engagement has taken place, how patients and stakeholders will be engaged and how often they will be engaged in the project\. Describe how stakeholders and the organizations they may represent were involved in developing the proposed project and this LOI, and how they will continue to be involved in executing the proposed project\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const EA_SCS_Patient_Centered_CER_Focus_for_LOI = "\* Connection to patient-centered comparative clinical effectiveness research \(CER\) Explain how your project is clearly focused on and supports engagement in patient-centered CER\. Your response should demonstrate an understanding of patient-centered CER\. \(Remember that the Engagement Award program funds research support activities and is NOT a research funding opportunity\.\) \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "
Public const Link1_PS_Tab_EALOI_SCS = "patient-centered CER"
Public const Link1_PS_Tab_URL_EALOI_SCS = "https://www\.pcori\.org/research-results/about-our-research/research-we-support"
Public const Evaluation_Plan_EALOI_SCS = "\* Evaluation Plan In addition to the required ""PCORI Engagement Award Reporting Tool,"" briefly describe any supplementary evaluation to be conducted focused on knowledge sharing and transfer from this project\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "
Public const Sustainability_Plan_EALOI_SCS = "\* Sustainability Plan Briefly describe how the project's outcomes contribute to the overall work your organization does and how the outcomes may be adopted and sustained by your organization or project partners\. Demonstrate a clear link from this project to future opportunities for participation in patient-centered CER\. If the project does not lend itself to sustained activities after the project period concludes, provide justification\. Please note: Future funding from PCORI should not be assumed\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "

'''''''''''''''''Project Summary Tab for EA CB''''''''''''''''''''
Public const Background_EALOI_CB ="\* Background State the problem or question this project is designed to address\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Objectives_EALOI_CB ="\* Objectives Clearly state the primary aims and objectives of your project that will lead to accomplishing your goal\(s\)\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Methods_EALOI_CB = "\* Methods/Activities Provide a concise description of the project methods, activities, and strategies that will be employed, how they will support the overall goal, and how they are connected to each other\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Patient_Stakeholder_Engagement_Plan_EALOI_CB = "\* Patient and Stakeholder Engagement Plan Briefly explain who the patients and stakeholders are, what if any preexisting engagement has taken place, how patients and stakeholders will be engaged and how often they will be engaged in the project\. Describe how stakeholders and the organizations they may represent were involved in developing the proposed project and this LOI, and how they will continue to be involved in executing the proposed project\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const EA_CB_Patient_Centered_CER_Focus_for_LOI = "\* Connection to patient-centered comparative clinical effectiveness research \(CER\) Explain how your project is clearly focused on and supports engagement in patient-centered CER\. Your response should demonstrate an understanding of patient-centered CER\. \(Remember that the Engagement Award program funds research support activities and is NOT a research funding opportunity\.\) \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "
Public const Link1_PS_Tab_EALOI_CB ="patient-centered CER"
Public const Link1_PS_Tab_URL_EALOI_CB = "https://www\.pcori\.org/research-results/about-our-research/research-we-support"
Public const Evaluation_Plan_EALOI_CB = "\* Evaluation Plan In addition to the required ""PCORI Engagement Award Reporting Tool,"" briefly describe any supplementary evaluation to be conducted focused on knowledge sharing and transfer from this project\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "
Public const Sustainability_Plan_EALOI_CB = "\* Sustainability Plan Briefly describe how the project's outcomes contribute to the overall work your organization does and how the outcomes may be adopted and sustained by your organization or project partners\. Demonstrate a clear link from this project to future opportunities for participation in patient-centered CER\. If the project does not lend itself to sustained activities after the project period concludes, provide justification\. Please note: Future funding from PCORI should not be assumed\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "



''''''''''''Project Summary Tab for EA DI App''''''''''''''''
Public const Project_Summary_made_public_EAApp_DI = "\* Project Summary \(to be made public\) Summarize your project for the general public\. This summary may be posted to PCORI\.org or in other PCORI materials\. Please include information about the following items in your abstract\. Remember to demonstrate in your project summary how your project is clearly focused on and supports engagement in patient-centered CER \(3,500 character limit\): Background—Briefly state the problem or question that the project is designed to address\. Proposed Solution to the Problem—Briefly describe the manner in which the problem or question will be resolved, including your project’s location \(i\.e\., city, town, district\) and setting \(i\.e\., clinic, community center, school\)\. Objectives—Briefly describe the aims of the project Activities—Provide a concise description of project activities that will occur throughout the duration of your project\. Outcomes and Outputs \(projected\)—Specify the projected short-term \(during the PCORI-funded project period\), medium-term \(0-2 years post-project period\), and long-term outcomes \(3\+ years post-project period\) and state their significance\. Ensure the outcomes and impact are defined by the time ranges provided above\. In addition, clearly describe your intended outputs and deliverables for the PCORI-funded project period\. Patient and Stakeholder Engagement Plan—Who are the patients and stakeholders involved in and/or impacted by the project, how they will be engaged, and how often will they be engaged in the planning and execution of the proposed project\? Project Collaborators— Which organizations or institutions are helping to lead, subcontract, or support this project in any way\? 0 of 3500 Characters "
Public const Is_this_resubmission_EAApp_DI = "\* Is this a resubmission of a previously submitted Engagement Award LOI or proposal\? --None-- Yes No "
Public const Resubmission_Details_EAApp_DI = "\* If this is a resubmission, provide the prior PCORI application number\(s\) and describe how this application has changed from the previous submission\(s\)\. If this is not a resubmission, enter N/A in the box below\. 26 of 1000 Characters "
Public const Background_EAApp_DI = "\* Background State the problem or question this project is designed to address\. Describe the opportunity to promote the accessibility, use, and/or uptake of eligible PCORI-funded research findings\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 23 of 3050 Characters "
Public const Explain_Proposed_Solution_EAApp_DI = "If yes, explain the proposed solution Explain why it is believed that this solution will work, and be better than previous solutions\. Describe how the solution is achieved \(designed and implemented\) or is at least achievable\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 30 of 3050 Characters "
Public const Objectives_EAApp_DI = "\* Objectives Briefly describe the aims and goals of the project, including the objectives\. Clearly state the primary objective of your dissemination effort\. What are the specific steps you will take to accomplish your overall objectives\? \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 23 of 3050 Characters "
Public const Methods_EAApp_DI = "\* Methods/Activities Provide a concise description of the project methods, activities, and strategies that will be employed, how they will support the overall goal and how they are connected to each other\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) Note: At this time, PCORI recommends that applicants plan for meetings with a virtual component or option, when practicable\. If an applicant opts for a fully in-person meeting, the Workplan and Budget Justification applicant templates in the full proposal must include a plan for activities and costs related to travel and other in-person meeting expenses should issues related to COVID-19 or any unanticipated public health emergency interfere with the meeting\. Travel logistics, accessibility, and health and safety considerations of the participants should be a consideration for any convening, meeting, and/or conference\. If an applicant proposes a meeting with an in-person component, PCORI expects all applicants to implement appropriate safety protocols consistent with applicable public health authorities and federal, state, and local guidance, laws, and regulations\. Please consult PCORI’s Applicant and Awardee FAQs Related to COVID-19 or Other Unexpected Public Health Emergencies to ensure your proposed project adheres to PCORI’s guidance related to applicant pre-award concerns\. 31 of 3050 Characters "
Public const Outcomes_projected_EAApp_DI = "\* Projected Outcomes Specify the projected short-term \(during the PCORI-funded project period\), medium-term \(0-2 years post-project period\), and long-term outcomes \(3\+ years post-project period\) and state their significance\. In addition, clearly describe your intended outputs and deliverables for the PCORI-funded project period\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 31 of 3050 Characters "
Public const Patient_Stakeholder_Engagement_Plan_EAApp_DI = "\* Patient and Stakeholder Engagement Plan Briefly explain who the patients and stakeholders are, what if any preexisting engagement has taken place, how patients and stakeholders will be engaged, and how often they will be engaged in the project\. Address how these groups can use and benefit from the proposed dissemination efforts\. Describe how stakeholders and the organizations they may represent were involved in developing the proposed project and this proposal, and how they will continue to be involved in executing the proposed project\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 52 of 3050 Characters "
Public const Evaluation_Plan_EAApp_DI = "\* Evaluation Plan Referencing the required ""EADI Reporting Tool"" as a basis for evaluation planning, briefly describe any evaluation to be conducted during this project\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 28 of 1000 Characters "
Public const Sustainability_Plan_EAApp_DI = "\* Sustainability Plan Briefly describe how the project's outcomes contribute to the overall work your organization does and how the outcomes may be adopted and sustained by your organization or project partners\. If you do not intend to sustain the efforts of this project, please justify why this award will accomplish what is needed with respect to dissemination\. Future funding from PCORI should not be assumed\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 32 of 1000 Characters "



''''''''''''Project Summary Tab for EA CB App''''''''''''''''
Public const Background_EAApp_CB = "\* Background State the problem or question this project is designed to address\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 23 of 3050 Characters "
Public const Objectives_EAApp_CB = "\* Objectives Clearly state the primary aims and objectives of your project that will lead to accomplishing your goal\(s\)\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 23 of 3050 Characters "
Public const Resubmission_Details_EAApp_CB = "\* If yes, provide the prior application number\(s\) and describe how this application has changed from the previous submission\(s\)\. If this is not a resubmission, enter N/A in the box below\. 26 of 1000 Characters "
Public const Explain_Proposed_Solution_EAApp_CB = "\* Project Goal What is the specific goal of this project\? Explain what you hope to achieve through the PCORI-funded work you propose in this application\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 25 of 3050 Characters "
Public const Methods_EAApp_CB = "\* Methods/Activities Provide a concise description of the project methods, activities, and strategies that will be employed, how they will support the overall goal, and how they are connected to each other\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 31 of 3050 Characters "
Public const Patient_Stakeholder_Engagement_Plan_EAApp_CB = "\* Patient and Stakeholder Engagement Plan Briefly explain who the patients and stakeholders are, what if any preexisting engagement has taken place, how patients and stakeholders will be engaged and how often they will be engaged in the project\. Describe how stakeholders and the organizations they may represent were involved in developing the proposed project and this proposal, and how they will continue to be involved in executing the proposed project\. \(Applicants may not exceed 1,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 3,050\.\) 52 of 3050 Characters "
Public const Evaluation_Plan_EAApp_CB = "\* Evaluation Plan In addition to the required ""PCORI Engagement Award Reporting Tool,"" briefly describe any supplementary evaluation to be conducted focused on knowledge sharing and transfer from this project\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 28 of 1000 Characters "
Public const Sustainability_Plan_EAApp_CB = "\* Sustainability Plan Briefly describe how the project's outcomes contribute to the overall work your organization does and how the outcomes may be adopted and sustained by your organization or project partners\. Demonstrate a clear link from this project to future opportunities for participation in patient-centered CER\. If the project does not lend itself to sustained activities after the project period concludes, provide justification\. Please note: Future funding from PCORI should not be assumed\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 32 of 1000 Characters "
Public const PCORI_Topic_Theme_EAApp_CB = "\* PCORI Topic Theme \(if applicable\) Does your project align with one of PCORI's Topic Themes\? Funding is not limited to these topics\. PCORI welcomes all LOIs that meet its guidelines\. Topics should address issues important to patients, families, caregivers, and the health care community\. Select all that apply\. Not Applicable Improving outcomes for people with intellectual or developmental disabilities \(IDD\) Promoting healthy children and youth Addressing substance use Addressing violence and trauma Addressing COVID-19 Addressing rare diseases Improving cardiovascular health Improving mental and behavioral health Managing pain Preventing maternal morbidity and mortality \(MMM\) Promoting sleep health Promoting health for older adults "
Public const EA_CB_Patient_Centered_CER_Focus_for_App = "\* Connection to patient-centered comparative clinical effectiveness research \(CER\) Explain how your project is clearly focused on and supports engagement in patient-centered CER\. Your response should demonstrate an understanding of patient-centered CER\. \(Remember that the Engagement Award program funds research support activities and is NOT a research funding opportunity\.\) \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 87 of 1000 Characters "
Public const Project_Summary_made_public_EAApp_CB = "\* Project Summary \(to be made public\) Summarize your project for the general public\. This summary may be posted to PCORI\.org or in other PCORI materials\. Please include information about the following items in your abstract\. Remember to demonstrate in your project summary how your project is clearly focused on and supports engagement in patient-centered CER \(3,500 character limit\): Background—Briefly state the problem or question that the project is designed to address\. Project Goal—Briefly describe the specific goal of the project\. Explain what you hope to achieve through the work proposed in this application\. Objectives—Briefly describe the primary aims and objectives of the project that will lead to accomplishing the goal\. Activities—Provide a concise description of project activities that will occur throughout the duration of your project\. Outcomes and Outputs \(projected\)—Specify the projected short-term \(during the PCORI-funded project period\), medium-term \(0-2 years post-project period\), and long-term outcomes \(3\+ years post-project period\) and state their significance\. Ensure the outcomes and impact are defined by the time ranges provided above\. In addition, clearly describe your intended outputs and deliverables for the PCORI-funded project period\. Patient and Stakeholder Engagement Plan—Who are the patients and stakeholders involved in and/or impacted by the project, how they will be engaged, and how often will they be engaged in the planning and execution of the proposed project\? Project Collaborators— Which organizations or institutions are helping to lead, subcontract, or support this project in any way\? 0 of 3500 Characters "

'''''''''''''''''''Additional Project Information Tab for EA DI'''''''''''''''''''''
Public const Amount_Requested_From_PCORI_EALOI_DI = "\* \(1\) Enter the amount of funding requested from PCORI for this application Do not exceed \$300,000 total costs; total costs are all direct costs and indirect costs combined "
Public const Total_Project_Budget_EALOI_DI = "\* \(2\) Enter the project's total budget Include the amount requested from PCORI plus any additional funding from sources other than PCORI that will support the proposed project\. "
Public const Budget_Description_EALOI_DI = "\* \(3\) Provide a brief budget narrative describing how the amount requested from PCORI will be utilized\. Your response must include an estimated dollar amount or percentage of the budget for each category and identify the percentage of indirect costs being required\. The budget categories include: Personnel Costs, Consultant Costs, Supply Costs, Travel Costs, Other Expenses, Subcontractor Costs, and Indirect Costs\. Where possible, identify specific expenses within the budget categories\. Describe if patient/stakeholder partners will receive compensation for their role on the project\. If patient/stakeholder partners will be compensated, include the number of patient/stakeholder partners to be compensated and include an estimated dollar amount or percentage of the total budget allocated to patient/stakeholder compensation\. If compensation of patient/stakeholder partners will occur but details are not yet available, explain how compensation rates will be determined and who will be involved in those decisions\. If patient/stakeholder partners will not be compensated, explain why\. Project budgets should reflect the time and contributions of all partners, including patients and stakeholders\. Fair financial compensation demonstrates that patients, caregivers, and patient/caregiver organizations’ contributions to the project, including related commitments of time and effort, are valuable and valued\. \(See our Financial Compensation of Partners Framework, Budgeting for Engagement Activities, and Cost Principles for more information\. Though these documents discuss compensation in research, the concepts are relevant for Engagement Award projects\.\) Budgets should account for all costs associated with proposed activities and note any in-kind support or external funding\. Applicants should keep personnel costs \(applicant organization/institution staff\) below 50% of the total project budget\. However, higher personnel costs may be considered with strong justification included in application documents\. \(Applicants may not exceed 2,000 characters, including spaces, for this response\. Please disregard the automated counter’s listed maximum of 3,050\.\) 0 of 3050 Characters "
Public const Link1_API_Tab_EALOI_DI = "Financial Compensation of Partners Framework"
Public const Link1_API_Tab_URL_EALOI_DI = "https://www\.pcori\.org/sites/default/files/PCORI-Compensation-Framework-for-Engaged-Research-Partners\.pdf"
Public const Link2_API_Tab_EALOI_DI = "Budgeting for Engagement Activities"
Public const Link2_API_Tab_URL_EALOI_DI = "https://www\.pcori\.org/sites/default/files/PCORI-Budgeting-for-Engagement-Activities\.pdf"
Public const Link3_API_Tab_EALOI_DI = "Cost Principles"
Public const Link3_API_Tab_URL_EALOI_DI = "https://www\.pcori\.org/sites/default/files/PCORI-Cost-Principles-2022\.pdf"
Public const Project_Start_Date_EALOI_DI = "\* \(4\) Enter the project's requested Start Date Projects for the Engagement Award: Dissemination Initiative in the Fall 2024 Cycle may begin between June 1, 2025 and November 1, 2025\. PCORI prefers all contract periods to begin on the first day of the month\. "

Public const Project_End_Date_EALOI_DI = "\* \(5\) Enter the project's requested End Date Projects for the Engagement Award: Dissemination Initiative may be up to 24 months in duration\. The length of the project period must be justified by the level of activity that will occur during the project\. "
Public const Vulnerable_Underserved_Pop_Focus_EALOI_DI = "\* \(6\) Does your project focus on any of the following vulnerable or underserved populations\? The PCORI Engagement Award Program is open to all applicants; however, we would like to know if your LOI focuses on any of the following vulnerable or underserved populations\. \(Please select all that apply\) Children 0-12 Children 13-18 Children 18-21 Adults >65 Disabled persons African-Americans Hispanic/Latino American American Indian/Alaska Native Pacific Islander Asian Residents of rural areas Resident of urban areas Veterans Women LGBT Low income groups Patients w/ low health literacy/numeracy N/A Other \(please specify\) Multiple chronic conditions Rare diseases Genetic make-up affects medical outcomes "
Public const Vulnerable_Underserved_Other_EALOI_DI = "\(6a\) If other vulnerable/underserved population, please specify "
Public const Proposal_Focus_Stakeholder_Community_EALOI_DI = "\* \(7\) Which stakeholder community\(-ies\) does your project focus on\? Select all that apply\. Patient/Consumer Caregiver/Family member of patient Patient/Caregiver Advocacy Organization Clinician Clinic/Hospital/Health System Payer Industry Research Policy Maker Training Institution "
Public const Experience_engaging_population_EALOI_DI = "\* \(8\) What is your organization's experience working with the population\(s\) you seek to engage\? Why is it important to fund dissemination efforts with the population\(s\)\? For what target audience\(s\) and in what ways is your organization a trusted source of information\? Please describe the type of information your organization generally communicates to your target audience\. What approaches does your organization use to communicate information to your target audience\? Please describe the number of people you reach using these approaches\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "
Public const Collaborator_Partner_Orgs_Role_EALOI_DI = "\* \(9\) Provide the name of organizational collaborators or partners and describe the role the organization\(s\) will play in meeting the goals and objectives of the project\. \(Applicants may not exceed 2,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 5,000\.\) 0 of 5000 Characters "
Public const Existing_Project_Funded_by_Others_EALOI_DI = "\* \(10\) Is this a previously existing project that has been funded by others\? --None-- Yes No " 
Public const Describe_funders_explain_EALOI_DI = "\(10a\) If yes, identify the previous funders and explain "
Public const PCORnet_involvement_EALOI_DI = "\* \(11\) Does your project include collaborations with existing PCORnet Network Partners \(i\.e\., Clinical Research Networks, Coordinating Center for PCORnet\)\? --None-- Yes No "

'''''''''''''''''Error message Additional Project Information Tab EA Application''''''''''
Public const Default_Error_Message1 = "Complete the following fields to save changes on this tab "
Public const Application_amount_Error = "Error: The Application Amount limit is \$300,000 "

'''''''''''''''''''Additional Project Information Tab for EA SCS'''''''''''''''''''''
Public const Amount_Requested_From_PCORI_EALOI_SCS = "\* \(1\) Enter the amount of funding requested from PCORI for this application Do not exceed \$125,000 total costs; total costs are all direct costs and indirect costs combined "
Public const Project_Start_Date_EALOI_SCS = "\* \(4\) Enter the project's requested Start Date Projects for the Engagement Award: Stakeholder Convening Support in the April 2024 Cycle may begin between December 1, 2024 and May 1, 2025\. PCORI prefers all contract periods to begin on the first day of the month\. " 
Public const Project_End_Date_EALOI_SCS = "\* \(5\) Enter the project's requested End Date Projects for the Engagement Award: Stakeholder Convening Support may be up to 12 months in duration\. The length of the project period must be justified by the level of activity that will occur during the project\. "
Public const Dates_stakeholder_convenings_EALOI_SCS = "\* \(6\) Identify the anticipated date\(s\) for the stakeholder convening\(s\) during your project period\. You must provide the estimated month and year of each planned convening\. 0 of 3050 Characters "
Public const Vulnerable_Underserved_Pop_Focus_EALOI_SCS = "\* \(7\) Does your project focus on any of the following vulnerable or underserved populations\? The PCORI Engagement Award Program is open to all applicants; however, we would like to know if your project focuses on any of the following vulnerable or underserved populations\. \(Please select all that apply\) Children 0-12 Children 13-18 Children 18-21 Adults >65 Disabled persons African-Americans Hispanic/Latino American American Indian/Alaska Native Pacific Islander Asian Residents of rural areas Resident of urban areas Veterans Women LGBT Low income groups Patients w/ low health literacy/numeracy N/A Other \(please specify\) Multiple chronic conditions Rare diseases Genetic make-up affects medical outcomes "
Public const Vulnerable_Underserved_Other_EALOI_SCS = "\(7a\) If other vulnerable/underserved population, please specify "
Public const Proposal_Focus_Stakeholder_Community_EALOI_SCS = "\* \(8\) Which stakeholder community\(-ies\) does your project focus on\? Select all that apply\. Patient/Consumer Caregiver/Family member of patient Patient/Caregiver Advocacy Organization Clinician Clinic/Hospital/Health System Payer Industry Research Policy Maker Training Institution "
Public const Experience_engaging_population_EALOI_SCS = "\* \(9\) What is your organization's experience working with the population\(s\) you seek to engage\? Why is it important to fund engagement efforts with the population\(s\)\? Provide evidence of established relationships with these stakeholders\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "
Public const Collaborator_Partner_Orgs_Role_EALOI_SCS = "\* \(10\) Provide the name of organizational collaborators or partners and describe the role the organization\(s\) will play in meeting the goals and objectives of the project\. \(Applicants may not exceed 2,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 5,000\.\) 0 of 5000 Characters "
Public const Existing_Project_Funded_by_Others_EALOI_SCS = "\* \(11\) Is this a previously existing project that has been funded by others\? --None-- Yes No " 
Public const Describe_funders_explain_EALOI_SCS = "\(11a\) If yes, identify the previous funders and explain "
Public const PCORnet_involvement_EALOI_SCS = "\* \(12\) Does your project include collaborations with existing PCORnet Network Partners \(i\.e\., Clinical Research Networks, Coordinating Center for PCORnet\)\? --None-- Yes No "

'''''''''''''''''''Additional Project Information Tab for EA CB'''''''''''''''''''''
Public const Amount_Requested_From_PCORI_EALOI_CB ="\* \(1\) Enter the amount of funding requested from PCORI for this application Do not exceed \$300,000 total costs; total costs are all direct costs and indirect costs combined "
Public const Project_Start_Date_EALOI_CB ="\* \(4\) Enter the project's requested Start Date Projects for the Engagement Award: Capacity Building in the Fall 2024 Cycle may begin between June 1, 2025 and November 1, 2025\. PCORI prefers all contract periods to begin on the first day of the month\. "
Public const Project_End_Date_EALOI_CB = "\* \(5\) Enter the project's requested End Date Projects for the Engagement Award: Capacity Building may be up to 24 months in duration\. The length of the project period must be justified by the level of activity that will occur during the project\. "
Public const Vulnerable_Underserved_Pop_Focus_EALOI_CB = "\* \(6\) Does your project focus on any of the following vulnerable or underserved populations\? The PCORI Engagement Award Program is open to all applicants; however, we would like to know if your LOI focuses on any of the following vulnerable or underserved populations\. \(Please select all that apply\) Children 0-12 Children 13-18 Children 18-21 Adults >65 Disabled persons African-Americans Hispanic/Latino American American Indian/Alaska Native Pacific Islander Asian Residents of rural areas Resident of urban areas Veterans Women LGBT Low income groups Patients w/ low health literacy/numeracy N/A Other \(please specify\) Multiple chronic conditions Rare diseases Genetic make-up affects medical outcomes "
Public const Vulnerable_Underserved_Other_EALOI_CB = "\(6a\) If other vulnerable/underserved population, please specify "
Public const Proposal_Focus_Stakeholder_Community_EALOI_CB ="\* \(7\) Which stakeholder community\(-ies\) does your project focus on\? Select all that apply\. Patient/Consumer Caregiver/Family member of patient Patient/Caregiver Advocacy Organization Clinician Clinic/Hospital/Health System Payer Industry Research Policy Maker Training Institution "
Public const Experience_engaging_population_EALOI_CB ="\* \(8\) What is your organization's experience working with the population\(s\) you seek to engage\? Why is it important to fund engagement efforts with the population\(s\)\? Provide evidence of established relationships with these stakeholders\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 0 of 1000 Characters "
Public const Collaborator_Partner_Orgs_Role_EALOI_CB ="\* \(9\) Provide the name of organizational collaborators or partners and describe the role the organization\(s\) will play in meeting the goals and objectives of the project\. \(Applicants may not exceed 2,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 5,000\.\) 0 of 5000 Characters "
Public const Existing_Project_Funded_by_Others_EALOI_CB = "\* \(10\) Is this a previously existing project that has been funded by others\? --None-- Yes No "
Public const Describe_funders_explain_EALOI_CB = "\(10a\) If yes, identify the previous funders and explain "
Public const PCORnet_involvement_EALOI_CB ="\* \(11\) Does your project include collaborations with existing PCORnet Network Partners \(i\.e\., Clinical Research Networks, Coordinating Center for PCORnet\)\? --None-- Yes No "



'''''''''''''''''''Additional Project Information Tab for EA App DI''''''''''''''''''''
Public const Budget_Description_EAApp_DI = "\* \(3\) Provide a brief budget narrative describing how the amount requested from PCORI will be utilized\. Your response must include an estimated dollar amount or percentage of the budget for each category and identify the percentage of indirect costs being required\. The budget categories include: Personnel Costs, Consultant Costs, Supply Costs, Travel Costs, Other Expenses, Subcontractor Costs, and Indirect Costs\. Where possible, identify specific expenses within the budget categories\. Describe if patient/stakeholder partners will receive compensation for their role on the project\. If patient/stakeholder partners will be compensated, include the number of patient/stakeholder partners to be compensated and include an estimated dollar amount or percentage of the total budget allocated to patient/stakeholder compensation\. If compensation of patient/stakeholder partners will occur but details are not yet available, explain how compensation rates will be determined and who will be involved in those decisions\. If patient/stakeholder partners will not be compensated, explain why\. Project budgets should reflect the time and contributions of all partners, including patients and stakeholders\. Fair financial compensation demonstrates that patients, caregivers, and patient/caregiver organizations’ contributions to the project, including related commitments of time and effort, are valuable and valued\. \(See our Financial Compensation of Partners Framework, Budgeting for Engagement Activities, and Cost Principles for more information\. Though these documents discuss compensation in research, the concepts are relevant for Engagement Award projects\.\) Budgets should account for all costs associated with proposed activities and note any in-kind support or external funding\. Applicants should keep personnel costs \(applicant organization/institution staff\) below 50% of the total project budget\. However, higher personnel costs may be considered with strong justification included in application documents\. \(Applicants may not exceed 2,000 characters, including spaces, for this response\. Please disregard the automated counter’s listed maximum of 3,050\.\) 33 of 3050 Characters "
Public const Vulnerable_Underserved_Pop_Focus_EAApp_DI = "\* \(6\) Does your project focus on any of the following vulnerable or underserved populations\? The PCORI Engagement Award Program is open to all applicants; however, we would like to know if your proposal focuses on any of the following vulnerable or underserved populations\. \(Please select all that apply\) Children 0-12 Children 13-18 Adults >65 Disabled persons African-Americans Hispanic/Latino American American Indian/Alaska Native Pacific Islander Asian Residents of rural areas Resident of urban areas Veterans Women LGBT Low income groups Patients w/ low health literacy/numeracy N/A Other \(please specify\) Multiple chronic conditions Rare diseases Genetic make-up affects medical outcomes Children 18-21 "
Public const Proposal_Focus_Stakeholder_Community_EAApp_DI = "\* \(7\) Which stakeholder community\(-ies\) does your project focus on\? Select all that apply\. Patient/Consumer Caregiver/Family member of patient Patient/Caregiver Advocacy Organization Clinic/Hospital/Health System Payer Industry Research Policy Maker Training Institution Clinician "
Public const Experience_engaging_population_EAApp_DI = "\* \(8\) What is your organization's experience working with the population\(s\) you seek to engage\? Why is it important to fund dissemination efforts with the population\(s\)\? For what target audience\(s\) and in what ways is your organization a trusted source of information\? Please describe the type of information your organization generally communicates to your target audience\. What approaches does your organization use to communicate information to your target audience\? Please describe the number of people you reach using these approaches\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 52 of 1000 Characters "
Public const Collaborator_Partner_Orgs_Role_EAApp_DI = "\* \(9\) Provide the name of organizational collaborators or partners and describe the role the organization\(s\) will play in meeting the goals and objectives of the project\. \(Applicants may not exceed 2,000 characters including spaces for this response\. Disregard the automated counter’s listed maximum of 5,050\.\) 51 of 5050 Characters "

'''''''''''''''''''Additional Project Information Tab for EA CB Application '''''''''''''''''''''
Public const Total_Project_Budget_EAApp_CB = "\* \(2\) Enter the project's total budget Include the amount requested from PCORI plus any additional funding from sources other than PCORI that will support the proposed project\. Funding from non-PCORI sources to support this project should be described in the appropriate section of the Budget Justification applicant template\. "
Public const Project_Start_Date_EAApp_CB = "\* \(4\) Enter the project's requested Start Date Projects for the Engagement Award: Capacity Building in the Fall 2024 Cycle may begin between June 1, 2025 and November 1, 2025\. PCORI prefers all contract periods to begin on the first day of the month\. "
Public const Project_End_Date_EAApp_CB = "\* \(5\) Enter the project's requested End Date Projects for the Engagement Award: Capacity Building may be up to 24 months in duration\. The length of the project period must be justified by the level of activity that will occur during the project\. "
Public const Experience_engaging_population_EAApp_CB = "\* \(8\) What is your organization's experience working with the population\(s\) you seek to engage\? Why is it important to fund engagement efforts with the population\(s\)\? Provide evidence of established relationships with these stakeholders\. \(Applicants may not exceed 1,000 characters including spaces for this response\.\) 52 of 1000 Characters "

'''''''''''''''''''''Using PCORI-funded Evidence & Tools EA DI '''''Tab
Public const Disseminating_PCORI_funded_evidence_EALOI_DI = "\* \(1\) The intent of an Dissemination Initiative project is to disseminate PCORI-funded research findings\. Applicants must identify each piece of eligible PCORI-funded research findings proposed for dissemination\. All eligible PCORI-funded research findings must have been published in a peer-reviewed journal \(primary CER results\) or posted on PCORI’s website \(systematic reviews and evidence updates\) by the LOI submission deadline\. Full evidence eligibility criteria is available in the Engagement Award: Dissemination Initiative - Fall 2024 Cycle funding announcement\. Will your proposed Engagement Award Project actively disseminate eligible PCORI-funded research findings\? --None-- Yes No "
Public const Link1_UPFET_Tab_EALOI_DI = "Engagement Award: Dissemination Initiative - Fall 2024 Cycle funding announcement"
Public const Link1_UPFET_Tab_URL_EALOI_DI = "https://www\.pcori\.org/funding-opportunities/announcement/engagement-award-dissemination-initiative-fall-2024-cycle"
Public const Evidence_Dissemination_EALOI_DI = "\* \(1a\) If yes, identify the eligible PCORI-funded research findings that this project will be actively disseminating\. Each piece of eligible PCORI-funded research findings should be listed on its own line in the box below \(hit 'Enter' after completing each entry to move to the next line\)\. If proposing to disseminate PCORI-funded CER results, provide the following information for each: \(a\) Title of the eligible PCORI-funded CER results publication, \(b\) Hyperlink to the publication on PCORI's website, and \(c\) Name of the principal investigator \(PI\) on the related PCORI-funded project \(this information can be found at the bottom of the publication page on PCORI’s website and the PI may not necessarily be the lead author on the publication\)\. If proposing to disseminate PCORI-funded Systematic Reviews, Systematic Review Updates, or PCORI Evidence Updates, provide the following information for each: \(a\) Title of the Systematic Review, Systematic Review Update, or Evidence Update, and \(b\) Hyperlink to the Systematic Review, Systematic Review Update, or Evidence Update on PCORI’s website\. Please contact the Engagement Award program at ea@pcori\.org for questions regarding evidence eligibility\. Please note: A letter of support will be required at the time of full proposal submission for eligible PCORI-funded research findings identified here\. Please see the ""Evidence Eligible for Dissemination Initiative Projects"" section of the full funding announcement for more information\. 0 of 3050 Characters "
Public const Link2_UPFET_Tab_EALOI_DI = "PCORI-funded CER results"
Public const Link2_UPFET_Tab_URL_EALOI_DI = "https://www\.pcori\.org/research-results/pcori-literature\?keyword=&f\[0\]=article_type:2900"
Public const Link3_UPFET_Tab_EALOI_DI = "PCORI-funded Systematic Reviews, Systematic Review Updates"
Public const Link3_UPFET_Tab_URL_EALOI_DI = "https://www\.pcori\.org/impact/evidence-synthesis-reports-and-interactive-visualizations/pcori-ahrq-systematic-reviews"
Public const Link4_UPFET_Tab_EALOI_DI = "PCORI Evidence Updates"
Public const Link4_UPFET_Tab_URL_EALOI_DI = "https://www\.pcori\.org/impact/evidence-updates"
Public const Link5_UPFET_Tab_EALOI_DI = "ea@pcori\.org"
Public const Link5_UPFET_Tab_URL_EALOI_DI = "mailto:ea@pcori\.org"
Public const Active_Dissemination_Key_Messages_EALOI_DI = "\* \(1b\) If yes, what are the key messages from the identified eligible PCORI-funded research findings that you plan to actively disseminate\? Key messages should be provided for each PCORI-funded research finding identified in the prior question\. 0 of 3050 Characters "
Public const Building_capacity_dissemination_EALOI_DI = "\* \(2\) The intent of a Building Capacity for Dissemination project is to strengthen the infrastructure and relationships necessary to actively disseminate or implement research results, products, or programs tested within PCORI-funded research studies\. Applicants must identify an area of PCORI's existing or emerging PCOR/CER findings highly relevant to their target population and in alignment with their organization/institution's goals\. Will your proposed Engagement Award Project build capacity for dissemination of PCORI-funded research findings\? --None-- Yes No "
Public const Evidence_Area_Build_Capacity_EALOI_DI = "\* \(2a\) If yes, identify an area of PCORI's existing or emerging PCOR/CER findings highly relevant to the target population and in alignment with the organization/institution's goals\. For reference, you may explore conditions and topic areas and populations of interest to PCORI, along with PCORI’s portfolio of funded PCOR/CER studies and existing or emerging PCOR/CER findings\. If you are not building capacity for dissemination as part of your project, enter N/A in the box below\. 0 of 3050 Characters "
Public const Link6_UPFET_Tab_EALOI_DI = "conditions and topic areas"
Public const Link6_UPFET_Tab_URL_EALOI_DI = "https://www\.pcori\.org/topics"
Public const Link7_UPFET_Tab_EALOI_DI = "populations of interest"
Public const Link7_UPFET_Tab_URL_EALOI_DI = "https://www\.pcori\.org/funding-opportunities/what-who-we-fund"
Public const Link8_UPFET_Tab_EALOI_DI = "portfolio of funded PCOR/CER studies"
Public const Link8_UPFET_Tab_URL_EALOI_DI = "https://www\.pcori\.org/explore-our-portfolio\?keyword=&keyword=&f\[0\]=award_types:1719"
Public const Use_PCORI_funded_tool_resource_EALOI_DI = "\* \(2\) Will your proposed Engagement Award Project use a PCORI-funded research or engagement tool or resource\? --None-- Yes No "
Public const Link9_UPFET_Tab_EALOI_DI = "project page"
Public const Link9_UPFET_Tab_URL_EALOI_DI = "https://www\.pcori\.org/research-results"
Public const PCORI_funded_tool_resource_EALOI_DI = "\* \(2a\) If yes, identify the PCORI-funded tool\(s\)/resource\(s\) that this project will be using by listing the following information: \(a\) Tool Name, \(b\) Project Title, \(c\) Project Lead/PI Name, and \(d\) URL to the tool or resource's project page on the PCORI website\. Each tool/resource should be listed on its own line in the box below \(hit 'Enter' after completing the entry for each individual tool to move to the next line in the box\)\. If you are not using a PCORI-funded research or engagement tool or resource, enter N/A in the box below\. Please note: A letter of support will be required at the time of full proposal submission for each PCORI-funded tool/resource identified here\. Please see the ""Use of a PCORI-Funded Tool or Resource"" section of the full funding announcement for more information\. 0 of 3050 Characters "
Public const Description_tool_resource_purpose_EALOI_DI = "\* \(2b\) If yes, describe the tool\(s\) or resource\(s\) selected and justification for its continued and expanded use\. Provide information about prior use/implementation of the tool\(s\) or resource\(s\), including any data on its effectiveness\. If you are not using a PCORI-funded research or engagement tool or resource, enter N/A in the box below\. 0 of 3050 Characters "

'''''''''''''''''''''Using PCORI-funded Evidence & Tools EA SCS '''''Tab
Public const Disseminating_PCORI_funded_evidence_EALOI_SCS = "\* \(1\) The intent of convening around dissemination of research findings is to disseminate the results of PCORI-funded studies\. Applicants proposing to convene around dissemination of research findings must identify each piece of eligible PCORI-funded research findings proposed for dissemination\. All eligible PCORI-funded research findings must have been published in a peer-reviewed journal \(primary CER results\) or posted on PCORI’s website \(systematic reviews and evidence updates\) by the LOI submission deadline\. Full evidence eligibility criteria is available in the Engagement Award: Stakeholder Convening Support - April 2024 Cycle funding announcement\. Will your proposed Engagement Award Project convene around the dissemination of eligible PCORI-funded research findings\? --None-- Yes No "
Public const Link1_UPFET_Tab_EALOI_SCS = "Engagement Award: Stakeholder Convening Support - April 2024 Cycle funding announcement"
Public const Link1_UPFET_Tab_URL_EALOI_SCS = "https://www\.pcori\.org/funding-opportunities/announcement/engagement-award-stakeholder-convening-support-april-2024-cycle"
Public const Evidence_Dissemination_EALOI_SCS = "\* \(1a\) If yes, identify the eligible PCORI-funded research findings that this project will be convening around dissemination of\. Each piece of eligible PCORI-funded research findings should be listed on its own line in the box below \(hit 'Enter' after completing each entry to move to the next line\)\. If proposing to convene around the dissemination of PCORI-funded CER results, provide the following information for each: \(a\) Title of the eligible PCORI-funded CER results publication, \(b\) Hyperlink to the publication on PCORI's website, and \(c\) Name of the principal investigator \(PI\) on the related PCORI-funded project \(this information can be found at the bottom of the publication page on PCORI’s website and the PI may not necessarily be the lead author on the publication\)\. If proposing to convene around the dissemination of PCORI-funded Systematic Reviews, Systematic Review Updates, or PCORI Evidence Updates, provide the following information for each: \(a\) Title of the Systematic Review, Systematic Review Update, or Evidence Update, and \(b\) Hyperlink to the Systematic Review, Systematic Review Update, or Evidence Update on PCORI’s website\. Please contact the Engagement Award program at ea@pcori\.org for questions regarding evidence eligibility\. If you are not convening around the dissemination of eligible PCORI-funded research findings as part of your project, enter N/A in the box below\. 0 of 3050 Characters "
Public const Active_Dissemination_Key_Messages_EALOI_SCS = "\* \(1b\) If yes, what are the key messages from the identified eligible PCORI-funded research findings that you plan to convene around the dissemination of\? Key messages must be provided for each PCORI-funded research finding identified in the prior question\. If you are not convening around the dissemination of eligible PCORI-funded research findings as part of your project, enter N/A in the box below\. 0 of 3050 Characters "

'''''''''''''''''''''Using PCORI-funded Evidence & Tools EA CB '''''Tab
Public const Use_PCORI_funded_tool_resource_EALOI_CB ="\* \(1\) Will your proposed Engagement Award Project use a PCORI-funded engagement tool or resource\? --None-- Yes No "
Public const PCORI_funded_tool_resource_EALOI_CB ="\* \(1a\) If yes, identify the PCORI-funded tool\(s\)/resource\(s\) that this project will be using by listing the following information: \(a\) Tool Name, \(b\) Project Title, \(c\) Project Lead/PI Name, and \(d\) URL to the tool or resource's project page on the PCORI website\. Each tool/resource should be listed on its own line in the box below \(hit 'Enter' after completing the entry for each individual tool to move to the next line in the box\)\. If you are not using a PCORI-funded engagement tool or resource, enter N/A in the box below\. Please note: A letter of support will be required at the time of full proposal submission for each PCORI-funded tool/resource identified here\. Please see the ""Use of a PCORI-Funded Tool or Resource"" section of the full funding announcement for more information\. 0 of 3050 Characters "
Public const Link1_UPFET_Tab_EALOI_CB = "project page"
Public const Link1_UPFET_Tab_URL_EALOI_CB = "https://www\.pcori\.org/research-results"
Public const Description_tool_resource_purpose_EALOI_CB ="\* \(1b\) If yes, describe the tool\(s\) or resource\(s\) selected and justification for its continued and expanded use\. Provide information about prior use/implementation of the tool\(s\) or resource\(s\), including any data on its effectiveness\. If you are not using a PCORI-funded engagement tool or resource, enter N/A in the box below\. 0 of 3050 Characters "


'''''''''''''''''''''Using PCORI-funded Evidence & Tools EA DI Application  '''''Tab
Public const Evidence_Dissemination_EAApp_DI = "\* \(1a\) If yes, identify the eligible PCORI-funded research findings that this project will be actively disseminating\. Each piece of eligible PCORI-funded research findings should be listed on its own line in the box below \(hit 'Enter' after completing each entry to move to the next line\)\. If proposing to disseminate PCORI-funded CER results, provide the following information for each: \(a\) Title of the eligible PCORI-funded CER results publication, \(b\) Hyperlink to the publication on PCORI's website, and \(c\) Name of the principal investigator \(PI\) on the related PCORI-funded project \(this information can be found at the bottom of the publication page on PCORI’s website and the PI may not necessarily be the lead author on the publication\)\. If proposing to disseminate PCORI-funded Systematic Reviews, Systematic Review Updates, or PCORI Evidence Updates, provide the following information for each: \(a\) Title of the Systematic Review, Systematic Review Update, or Evidence Update, and \(b\) Hyperlink to the Systematic Review, Systematic Review Update, or Evidence Update on PCORI’s website\. Please contact the Engagement Award program at ea@pcori\.org for questions regarding evidence eligibility\. 76 of 3050 Characters "
Public const Active_Dissemination_Key_Messages_EAApp_DI = "\* \(1b\) If yes, what are the key messages from the identified eligible PCORI-funded research findings that you plan to actively disseminate\? Key messages should be provided for each PCORI-funded research finding identified in the prior question\. 110 of 3050 Characters "
Public const PCORI_funded_tool_resource_EAApp_DI = "\* \(2a\) If yes, identify the PCORI-funded tool\(s\)/resource\(s\) that this project will be using by listing the following information: \(a\) Tool Name, \(b\) Project Title, \(c\) Project Lead/PI Name, and \(d\) URL to the tool or resource's project page on the PCORI website\. Each tool/resource should be listed on its own line in the box below \(hit 'Enter' after completing the entry for each individual tool to move to the next line in the box\)\. If you are not using a PCORI-funded research or engagement tool or resource, enter N/A in the box below\. 56 of 3050 Characters "
Public const Description_tool_resource_purpose_EAApp_DI = "\* \(2b\) If yes, describe the tool\(s\) or resource\(s\) selected and justification for its continued and expanded use\. Provide information about prior use/implementation of the tool\(s\) or resource\(s\), including any data on its effectiveness\. If you are not using a PCORI-funded research or engagement tool or resource, enter N/A in the box below\. 90 of 3050 Characters "

'''''''''''''''''''''Using PCORI-funded Evidence & Tools EA CB Application  '''''Tab

Public const Using_PCORI_funded_evidence_EAApp_CB = "\* \(1\) Will your proposed Engagement Award Project use a PCORI-funded engagement tool or resource\? --None-- Yes No "
Public const PCORI_funded_tool_resource_EAApp_CB = "\* \(1a\) If yes, identify the PCORI-funded tool\(s\)/resource\(s\) that this project will be using by listing the following information: \(a\) Tool Name, \(b\) Project Title, \(c\) Project Lead/PI Name, and \(d\) URL to the tool or resource's project page on the PCORI website\. Each tool/resource should be listed on its own line in the box below \(hit 'Enter' after completing the entry for each individual tool to move to the next line in the box\)\. If you are not using a PCORI-funded engagement tool or resource, enter N/A in the box below\. Please note: A letter of support is required for each PCORI-funded tool/resource identified here\. Please see the ""Use of a PCORI-Funded Tool or Resource"" section of the full funding announcement for more information\. 76 of 3050 Characters "
Public const Description_tool_resource_purpose_EAApp_CB = "\* \(1b\) If yes, describe the engagement tool\(s\) or resource\(s\) selected and justification for its continued and expanded use\. Provide information about prior use/implementation of the tool\(s\) or resource\(s\), including any data on its effectiveness to build capacity for patient and stakeholder engagement in patient-centered comparative clinical effectiveness research \(CER\)\. If you are not using a PCORI-funded engagement tool or resource, enter N/A in the box below\. 110 of 3050 Characters "


'''''''''''''''''''Key Personnel Tab EA application DI''''''''''''''''''
Public const Key_Personnel_InsText_EAApp_DI = "Please utilize the following resources when completing the proposal: PCORI Funding Announcement: Engagement Award: Dissemination Initiative - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources   Limit your Key Personnel entries to five \(5\), not including the Project Lead"


'''''''''''''''''''Key Personnel Tab EA application CB''''''''''''''''''
Public const Key_Personnel_InsText_EAApp_CB = "Please utilize the following resources when completing the proposal: PCORI Funding Announcement: Engagement Award: Capacity Building - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Limit your Key Personnel entries to five \(5\), not including the Project Lead"


''''''''''''''''Attachments Tab EA DI LOI''''''''''''''''
Public const Attachment_tab_InsText1_EALOI_DI = "The Engagement Award: Dissemination Initiative requires the completion of a supplementary document that includes a few additional questions about your project\. This document is required to be uploaded to PCORI Online as part of your LOI submission\. Download and complete the Engagement Award: Dissemination Initiative LOI Supplemental Template for the Fall 2024 Cycle\. Upload the file as a PDF document\."
Public const Templates_EALOI_DI = "\* Engagement Award: Dissemination Initiative LOI Supplemental Template Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const EALOI_DI_Template_upload_link = "LOI Supplemental Template"
Public const EALOI_DI_Template_upload_URL = "https://www\.pcori\.org/sites/default/files/PCORI-engagement-award-DI-LOI-supplemental-template-fall-2024-cycle\.docx"
Public const EALOI_DI_Template_Title_for_InternalVerification = "PCORI-engagement-award-DI-LOI-supplemental-template-fall-2024-cycle\.docx"

''''''''Template upload file Path'''''''''''''
Public const attachment_LOI_EADI = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST\PCORI-engagement-award-DI-LOI-supplemental-template-april-2024-cycle.docx"

''''''''''''Attachment Tab EA DI Application'''''''''''
Public const attachment_Instext_EAApp_DI = "All required templates are listed in the Engagement Award: Dissemination Initiative - April 2024 Cycle funding announcement under ""Applicant Resources\."" Applicants must follow all instructions listed in the individual required templates\. Applicant templates are unique to each funding cycle\. Do not re-use an applicant template from a previous PCORI application\. Please ensure you are using the applicant template attached to the PFA for the funding cycle you are applying to\. Please also ensure the file you uploaded is in the correct file format as specified in the prompts below\."
Public const EAApp_DI_Template_resource_Name = "Engagement Award: Dissemination Initiative - April 2024 Cycle funding announcement"
Public const EAApp_DI_Template_resource_Link = "https://www\.pcori\.org/funding-opportunities/announcement/engagement-award-dissemination-initiative-april-2024-cycle"
Public const Biosketch_file_Attachment_Section = "\* Please submit a Biosketch file Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Budget_file_Attachment_Section = "\* Please submit a Budget file \(as a Microsoft Excel file\) Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Budget_Justification_Attachment_Section = "\* Please submit a Budget Justification file Note: If requesting greater than 10 percent indirect costs, provide a copy of the organization’s federally negotiated or independently audited indirect cost rate letter\. This document must be attached to the Budget Justification and submitted as one single PDF file\. Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Milestone_Attachment_Section = "\* Please submit a Milestone/Deliverable Table file \(as a Microsoft Excel file\) Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Workplan_Attachment_Section = "\* Please submit a Workplan file \(as a Microsoft Word file\) Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Letters_Of_Support_Attachment_Section = "\* Please submit a Letters of Support file Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Board_Of_Directors_Attachment_Section = "Please submit a Board of Directors file \(optional\) Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "
Public const Recent_Evaluations_Attachment_Section = "Please submit a Recent Articles/Evaluations file \(optional\) Sr\. Number File Name - Created Date Action No Attachments Choose file No file chosen "

'''''''''''''''''' Attachments upload File Path For EA DI Application'''''''''''''''
Public const attachment_Biosketch_file = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST EA\PCORI-engagement-award-biosketch-template-october-2023-cycle.docx"
Public const attachment_Budget_file = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST EA\PCORI-engagement-award-budget-two-year-template-october-2023-cycle.xlsx"
Public const attachment_Budget_Justification = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST EA\PCORI-engagement-award-budget-justification-two-year-template-october-2023-cycle.docx"
Public const attachment_Milestone = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST EA\PCORI-engagement-award-milestone-deliverable-table-template-DI-BC4D-october-2023-cycle.xlsx"
Public const attachment_Workplan = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST EA\PCORI-engagement-award-workplan-template-DI-building-capacity-dissemination-october-2023-cycle.docx"
Public const attachment_Letters_Of_Support = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST EA\PCORI-engagement-award-letters-of-support-table-october-2023-cycle.docx"
Public const attachment_Board_Of_Directors = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST EA\PCORI-engagement-award-board-of-directors-october-2023-cycle.docx"
Public const attachment_Recent_Evaluation = "C:\Users\sahad.PCORI\OneDrive - PCORI\AUTOMATION docs\1. LOI and Application\Attachments for TEST EA\PCORI-engagement-award-evaluation-reporting-tool.docx"

''''''''''''Attachment Tab EA CB Application'''''''''''
Public const attachment_Instext_EAApp_CB = "All required templates are listed in the Engagement Award: Capacity Building - Fall 2024 Cycle funding announcement under ""Applicant Resources\."" Applicants must follow all instructions listed in the individual required templates\. Applicant templates are unique to each funding cycle\. Do not re-use an applicant template from a previous PCORI application\. Please ensure you are using the applicant template attached to the PFA for the funding cycle you are applying to\. Please also ensure the file you uploaded is in the correct file format as specified in the prompts below\."
Public const EAApp_CB_Template_resource_Name = "Engagement Award: Capacity Building - Fall 2024 Cycle funding announcement"
Public const EAApp_CB_Template_resource_Link = "https://www\.pcori\.org/funding-opportunities/announcement/engagement-award-capacity-building-fall-2024-cycle"

''''''''''''''''''Budget Tab EA DI Application'''''''''''
Public const Budgettab_Instext_EAAPP_DI = "Please utilize the following resources when completing the proposal: PCORI Funding Announcement: Engagement Award: Dissemination Initiative - April 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - April 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources   Upload the required Budget file to the Attachments tab\. Then, please insert budget subtotals for each category in the fields below\."

''''''''''''''''''Budget Tab EA CB Application'''''''''''
Public const Budgettab_Instext_EAAPP_CB = "Please utilize the following resources when completing the proposal:   PCORI Funding Announcement: Engagement Award: Capacity Building - Fall 2024 Cycle Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Upload the required Budget file to the Attachments tab\. Then, please insert budget subtotals for each category in the fields below\."

'''''''''''''''''''''''''Authorizations Tab  EA DI'''''''''''''''
Public const Authorization_DILOI_InsText1 = "Please consult the following resources to complete the authorizations section and submit the LOI: Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Authorizations Note: There are no formal steps for the Administrative Official to take to submit the LOI to PCORI\. The Project Lead should share the LOI with the Administrative Official for review prior to submitting the LOI\. By agreeing to the statement below, the Project Lead acknowledges they have authorization to submit the LOI to PCORI\. "
Public const Authorized_to_Submit_LOI_EADI = "\* I hereby certify that, to the best of my knowledge, the information in this PCORI LOI is true and accurate, and I am authorized by my organization to submit this LOI to PCORI\. Yes No "
Public const Authorization_DILOI_InsText2 = "Select the ‘Save’ button below\. Then, scroll to the top of the page to select the ‘Review/Submit’ button\. "
Public const Authorization_DILOI_InsText2_OuterHtml = "<b>Select the ‘Save’ button below\. Then, scroll to the top of the page to select the ‘Review/Submit’ button\.</b>"

'''''''''''''''''''''''''Authorizations Tab  EA SCS'''''''''''''''
Public const Authorization_SCS_LOI_InsText1 = "Please utilize the following resources when completing the LOI: Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Note: There are no formal steps for the Administrative Official to take to submit the LOI to PCORI\. The Project Lead should share the LOI with the Administrative Official for review prior to submitting the LOI\. By agreeing to the statement below, the Project Lead acknowledges they have authorization to submit the LOI to PCORI\. "

''''''''''''''''''''''''Authorizations Tab  EA DI Application'''''''''''''''
Public const Authorization_DIApp_InsText1 = "Please consult the following resources to complete the authorizations section and submit the proposal: Submission Instructions: Engagement Award Submission Instructions - April 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Note: After the Project Lead selects to submit, the proposal will be transferred to the Administrative Official for final review\. The Administrative Official must log in and complete the authorization form in order to formally submit the proposal to PCORI\. This step must be completed by the 5:00 p\.m\. \(ET\) proposal submission deadline\. More information on this required step is included in the Engagement Award Submission Instructions linked to the funding announcement you are applying to\. "
Public const Authorized_to_Submit_App_EADI = "\* I hereby certify that, to the best of my knowledge, the information in this PCORI proposal is true and accurate\. Yes No "

''''''''''''''''''''''''Authorizations Tab  EA CB Application'''''''''''''''
Public const Authorization_CBApp_InsText1 = "Please consult the following resources to complete the authorizations section and submit the proposal: Submission Instructions: Engagement Award Submission Instructions - Fall 2024 Cycle PCORI Online Resources for Applicants: Engagement Award Program Resources Note: After the Project Lead selects to submit, the proposal will be transferred to the Administrative Official for final review\. The Administrative Official must log in and complete the authorization form in order to formally submit the proposal to PCORI\. This step must be completed by the 5:00 p\.m\. \(ET\) proposal submission deadline\. More information on this required step is included in the Engagement Award Submission Instructions linked to the funding announcement you are applying to\. "




''''''''''''************************************************************************
''''''''''''''''''''''''EA Review Forms''''''''''''''''''''''
''''''''''''''EACB LOI Review Form''''''''''
Public const EACB_LOI_Review_Required_Field_Error = "Review the following fields Project Team Experience Adequate\? Sufficient Fit\? Proposed Collaborators Appropriate\? Additional Reviewer Needed\? Innovative Idea\? Engagement in Patient-Centered CER\? Engagement Plan Adequate\? Applicant Feedback, if Rejected Is this nonresponsive\? Applicant Feedback, if Invited Reviewer Disposition Reasonable Goals, Objectives, Outcomes\? Enough Financial Information Provided\? Cooperation Concerns\? I Have a Conflict with Application Project Lead Qualification Appropriate\? Are Cost Designations Reasonable\? Level of Enthusiasm LOI Review Comments Activities & Deliverables Support Goal\?"

'''''COI Section'''
Public const COI_Section_Text_LOI_Review = "CONFLICT OF INTEREST After reading the application, please indicate whether or not you have a conflict of interest by selecting one of the checkboxes below\. If you do have a conflict of interest, please save the record immediately and do not continue with the review\. Otherwise, please continue to fill out the form\."
Public const Conflict_with_Application = "I Have a Conflict with Application I Have a Conflict with Application Help Info Please check Either I Do Have a Conflict with Application or I Do Not Have a Conflict with Application"

''''''Review Criteria Section''''
Public const Review_Criteria_Text_LOI_Review = "REVIEW CRITERIA Please read the reviewer guidance for each criterion and respond to the question prompts in each section\. Then, follow the ""Reviewer Disposition and Feedback Instructions"" to provide your disposition and feedback on the LOI\. Please complete all prompts\."

'''''''''''''''''Criteria 1'''''''''''CB
Public const Criteria1_Program_Fit_CBLOI_Review = "Criteria 1: Program Fit and Balance How well does the application demonstrate program fit or how well does the application proffer an innovative idea that supports the organization’s mission\? • Does not include any non-responsive activities from 'Categories of Nonresponsiveness' list • Describes the problem and proposed solution  • Explains how the project is clearly focused on and supports engagement in patient-centered comparative clinical effectiveness research \(CER\) • Demonstrates an understanding of patient-centered CER \(This includes understanding the characteristics specific to patient-centeredness as well as an understanding that CER is not simply engaged research of any kind\.\)  • Rationale that there is a need for additional patient and/or stakeholder capacity to participate in patient-centered CER\. • Information and learnings generated by the project must be transferrable and of interest or use not just to the applicant organization but to others doing related work  • Clearly describes an innovative idea that supports PCORI’s mission  • \[REVIEWER TWO ONLY\] In an effort to avoid redundancy, do the goals or objectives of the project resemble those of a current or past EA project, to the best of your knowledge\? \(If uncertain, say that you are unsure in your response\)"
Public const Criteria1_Program_Fit_CBLOI_Review_Italic_Text = "<em>How well does the application demonstrate program fit or how well does the application proffer an innovative idea that supports the organization’s mission\?</em>" 
Public const Program_Fit_Criteria_LOI_Review = "How well does the LOI demonstrate program fit with the Engagement Award Program and the PFA to which they are applying\?"
Public const Sufficient_Criteria_LOI_Review = "Sufficient Fit\? --None-- Please pick a value in the Field Prior To Submission"
Public const Sufficient_Criteria_LOI_Review_acc_Name = "Sufficient Fit\?"

''''''''''This below line is captured innertext'''''
Public const Engagement_CER_Crit_LOI_Review = "Engagement in Patient-Centered CER Crit\.Is the project clearly focused on supporting engagement in patient-centered CER\?"

Public const Engagement_Patient_Centered_CER_LOI_Review = "Engagement in Patient-Centered CER\? --None-- You must fill out the ""Engagement in PCOR/CER\?"" field prior to submission"
Public const Engagement_Patient_Centered_CER_LOI_Review_acc_name = "Engagement in Patient-Centered CER\?"


''''''''''This below line is captured innertext'''''
Public const Innovative_Idea_Criteria_LOI_Review = "Innovative Idea CriteriaDoes the LOI offer an innovative idea that supports integrating patients and other stakeholders in the patient-centered CER process\?"

Public const Innovative_Idea_LOI_Review = "Innovative Idea\? --None-- Please specify an appropriate value whether or not for ""Innovative Idea"" prior to submission"
Public const Innovative_Idea_LOI_Review_acc_Name = "Innovative Idea\?"

''''''''''''''Criteria 1 CB/DI/SCS acc_Name of dropsown field value for Prod''''''''
Public const Sufficient_Criteria_LOI_Review_acc_Name_Prod = "Sufficient Fit\?"
Public const Engagement_Patient_Centered_CER_LOI_Review_acc_name_Prod = "Engagement in Patient-Centered CER\?"
Public const Innovative_Idea_LOI_Review_acc_Name_Prod = "Innovative Idea\?"


''''''''''''''''Criteria 2 CB''''''
Public const Criteria2_Project_Plan_CBLOI_Review = "Criteria 2: Project Plan How reasonable is the applicant's project plan\? Are there any concerns that the applicant has not provided a reasonable set of activities to achieve the objectives and outcomes in the project plan\? • Describes the aims and goals of the project, including the objectives • Includes objectives that support the goal of building capacity to engage in patient-centered CER  • Describes the project methods, activities, and strategies that will be employed to support the overall goal • Specifies the projected short- \(during the PCORI-funded project\), medium- \(0-2 years post-funding\), and long-term \(3\+ years post-funding\) outcomes; States the significance of the outcomes  • Clearly describes projected outputs \(including sustainability plans, evaluation analysis, and resources\) • Includes a brief outline of an evaluation plan focused on knowledge sharing and transfer  • If the applicant proposes creating a research agenda, does the LOI indicate that the agenda will be focused on or lend itself specifically to patient-centered CER\? • Demonstration of a clear and direct link to future opportunities for participation in patient-centered CER\. • For LOIs focused on IDD: Does the project address IDD in general or a subtopic of IDD \(e\.g\., autism\)\? If the project focuses on IDD in general, are the project activities and outputs relevant to all conditions\? Does the applicant discuss how the project will address the differences among populations and remain relevant to all\? • Describes any tools/trainings/programs that will be used as part of the project \(if applicable\)\. Are those resources appropriate for the proposed project\? • \[REVIEWER TWO ONLY\] If this is a resubmission of either an LOI or a full proposal, did the applicant adequately address the questions posed in the feedback letter\(s\) they previously received\?"
Public const Criteria2_Project_Plan_CBLOI_Review_Italic_Text = "<em>How reasonable is the applicant's project plan\? Are there any concerns that the applicant has not provided a reasonable set of activities to achieve the objectives and outcomes in the project plan\?</em>" 

''''''''''This below line is captured innertext'''''
Public const Goals_Objective_Outcomes_Criteria_LOI_Review = "Goals, Objective, Outcomes CriteriaDoes the applicant offer a reasonable set of goals, objectives, and expected outcomes for the project\?"

Public const Reasonable_Goals_Objective_Outcomes_Criteria_LOI_Review = "Reasonable Goals, Objectives, Outcomes\? --None-- You must choose a specific value for the ""Reasonable Goals, Objectives, Outcomes"" field prior to submission\."
Public const Reasonable_Goals_Objective_Outcomes_Criteria_LOI_Review_acc_name = "Reasonable Goals, Objectives, Outcomes\?"
''''''''''This below line is captured innertext'''''
Public const Activities_Deliverables_Criteria_LOI_Review = "Activities and Deliverables CriteriaDoes the applicant provide a reasonable set of activities and deliverables to achieve the proposed goals, objectives, and outcomes of the project\?"

Public const Activities_Deliverables_SupportGoal_LOI_Review = "Activities & Deliverables Support Goal\? --None-- Please choose a suitable value for the ""Activities and Deliverables Support Goal"" field"
Public const Activities_Deliverables_SupportGoal_LOI_Review_acc_name = "Activities & Deliverables Support Goal\?"

''''''''''''''Criteria 2 CB/DI/SCS acc_Name of drop down field value for Prod''''''''
Public const Reasonable_Goals_Objective_Outcomes_Criteria_LOI_Review_acc_name_Prod = "Reasonable Goals, Objectives, Outcomes\?"
Public const Activities_Deliverables_SupportGoal_LOI_Review_acc_name_Prod = "Activities & Deliverables Support Goal\?"


'''''''''''''''''Criteria 3 CB'''''''''''''''
Public const Criteria3_Lead_Experience_Organizational_Capabilities_CBLOI_Review = "Criteria 3: Project Lead Previous Experience and Organizational Capabilities Do the qualifications of the project lead align with the scope of the project\? Does the applicant demonstrate sufficient background/organizational capabilities for related projects that are aligned with an emphasis on patient-centered CER\?   • Describes the project lead’s previous experience related to patient-centered CER • Demonstrates sufficient organizational capabilities for projects with an emphasis on patient-centered CER • Demonstrates involvement of adequate and qualified personnel in conducting the project  • Explains involvement and achievements in other patient-centered CER or PCORI-funded projects"
Public const Criteria3_Lead_Experience_Organizational_Capabilities_CBLOI_Review_Italic_text = "<em>Do the qualifications of the project lead align with the scope of the project\? Does the applicant demonstrate sufficient background/organizational capabilities for related projects that are aligned with an emphasis on patient-centered CER\?&nbsp;</em>"

''''''''''This below line is captured innertext'''''
Public const Project_Lead_Qualification_Exp_Crit_CBLOI_Review = "Project Lead Qualification & Exp\. Crit\.Do the proposed project lead's qualifications and previous work/educational experience align with the scope of the project\?"

Public const Project_Lead_Qualification_Appropriate_CBLOI_Review = "Project Lead Qualification Appropriate\? --None-- You must fill out the ""Lead Qualification Appropriate"" field prior to submission"
Public const Project_Lead_Qualification_Appropriate_CBLOI_Review_acc_name = "Project Lead Qualification Appropriate\?"

''''''''''This below line is captured innertext'''''
Public const Project_Team_CER_Exp_CBLOI_Review = "Project Team Patient-Centered CER Exp\.Does the applicant organization/proposed project team demonstrate sufficient background/organizational capabilities for related projects that are aligned with an emphasis on patient-centered CER\?"

Public const Project_Team_Experiance_Adequate_CBLOI_Review = "Project Team Experience Adequate\? --None-- You must fill out the ""Project Team Experience Adequate\?"" field prior to submission"
Public const Project_Team_Experiance_Adequate_CBLOI_Review_acc_name = "Project Team Experience Adequate\?"

''''''''''''''Criteria 3 CB/DI/SCS acc_Name of drop down field value for Prod''''''''
Public const Project_Lead_Qualification_Appropriate_CBLOI_Review_acc_name_Prod = "Project Lead Qualification Appropriate\?"
Public const Project_Team_Experiance_Adequate_CBLOI_Review_acc_name_Prod = "Project Team Experience Adequate\?"


''''''''''''''''''Criteria 4 CB LOi Review''''''
Public const Criteria4_Patient_Stakeholder_Engagement_Plan_CBLOI_Review = "Criteria 4: Patient and Stakeholder Engagement Plan Is there an adequate plan for engaging patients and other stakeholders in the conduct of the proposed project\? Are the proposed collaborators meaningful and appropriate based on aligning the interest, expertise and scope of work of patients and other stakeholders involved in the project\? Are there any concerns about the ability of the collaborators to work together\? • Explains who the patient and stakeholder partners are and any preexisting engagement that has taken place before • Describes how stakeholders \(i\.e\., patients, caregivers, and clinicians\) and the organizations that represent them were involved in developing the proposed project • Proposed collaborators are meaningful and appropriate based on aligning the interest, expertise, and scope of work of patients and other stakeholders involved in the project • Means of collaboration \(i\.e\., meetings, etc\.\) and frequency of interactions have been addressed • Explains how the capacity that is developed through the award will be applied to existing or planned patient-centered CER partnership opportunities"
Public const Criteria4_Patient_Stakeholder_Engagement_Plan_CBLOI_Review_Italic_Text = "<em>Is there an adequate plan for engaging patients and other stakeholders in the conduct of the proposed project\? Are the proposed collaborators meaningful and appropriate based on aligning the interest, expertise and scope of work of patients and other stakeholders involved in the project\? Are there any concerns about the ability of the collaborators to work together\?</em>"

''''''''''This below line is captured innertext'''''
Public const Engagement_Plan_Criteria_CBLOI_Review = "Engagement Plan CriteriaIs there an adequate plan described for meaningful engagement of patients and/or other stakeholders, as appropriate, in conducting the proposed project\?"

Public const Engagement_Plan_Adequate_CBLOI_Review = "Engagement Plan Adequate\? --None-- You must specify whether or not the ""Engagement Plan Adequate"" field prior to submission"
Public const Engagement_Plan_Adequate_CBLOI_Review_acc_name = "Engagement Plan Adequate\?"

''''''''''This below line is captured innertext'''''
Public const Proposed_Collaborators_Criteria_CBLOI_Review = "Proposed Collaborators CriteriaAre the proposed collaborators meaningful and appropriate based on aligning the interest, expertise, and scope of work of patients and other stakeholders involved in the project\?"

Public const Proposed_Collaborators_Appropriate_CBLOI_Review = "Proposed Collaborators Appropriate\? --None-- Please pick a value in the Field Prior To Submission"
Public const Proposed_Collaborators_Appropriate_CBLOI_Review_acc_name = "Proposed Collaborators Appropriate\?"

''''''''''This below line is captured innertext'''''
Public const Cooperation_Criteria_CBLOI_Review = "Cooperation CriteriaAre there any concerns about the ability of patients/stakeholders to be involved in the project, or for collaborators to successfully work together\?"
Public const Cooperation_Concerns_CBLOI_Review = "Cooperation Concerns\? --None-- Please choose a desired value for ""Any cooperation Concerns"" field Prior to submission"
Public const Cooperation_Concerns_CBLOI_Review_acc_name = "Cooperation Concerns\?"

''''''''''''''Criteria 4 CB/DI/SCS acc_Name of drop down field value for Prod''''''''
Public const Engagement_Plan_Adequate_CBLOI_Review_acc_name_Prod = "Engagement Plan Adequate\?"
Public const Proposed_Collaborators_Appropriate_CBLOI_Review_acc_name_Prod = "Proposed Collaborators Appropriate\?"
Public const Cooperation_Concerns_CBLOI_Review_acc_name_Prod = "Cooperation Concerns\?"

''''''''''''Criteria 5 CB LOI Review section''''''''''''
Public const Criteria5_Budget_Cost_Proposal_CBLOI_Review = "Criteria 5: Budget/Cost Proposal How well does the applicant’s budget proposal demonstrate reasonable cost designations\? • Provides clear amounts & justification for each budget item • Reflects the time and contributions of all partners, including patients and stakeholders\. Fair financial compensation demonstrates that patients, caregivers, and patient/caregiver organizations’ contributions to the project, including related commitments of time and effort, are valuable and valued\."
Public const Criteria5_Budget_Cost_Proposal_CBLOI_Review_Italic_Text = "<em>How well does the applicant’s budget proposal demonstrate reasonable cost designations\?</em>"

''''''''''This below line is captured innertext'''''
Public const Financial_Support_Criteria_CBLOI_Review = "Financial Support CriteriaHas the applicant provided enough information to make an assessment of the financial support being requested through the Engagement Award Program\?"

Public const Enough_Financial_Information_CBLOI_Review  = "Enough Financial Information Provided\? --None-- You must specify whether or not the reviewer has provided the Enough Financial information\."
Public const Enough_Financial_Information_CBLOI_Review_acc_name = "Enough Financial Information Provided\?"

''''''''''This below line is captured innertext'''''
Public const Reasonable_Budget_Criteria_CBLOI_Review = "Reasonable Budget CriteriaDoes the applicant’s budget proposal demonstrate reasonable cost designations given the project scope, activities and expected duration\?"

Public const Are_Cost_Designations_Reasonable_CBLOI_Review = "Are Cost Designations Reasonable\? --None-- You must indicate whether or not or if you are unsure if the ""Costs Designations Are Reasonable"" prior to submission"
Public const Are_Cost_Designations_Reasonable_CBLOI_Review_acc_name = "Are Cost Designations Reasonable\?" 

''''''''''''''Criteria 5 CB/DI/SCS acc_Name of drop down field value for Prod''''''''
Public const Enough_Financial_Information_CBLOI_Review_acc_name_Prod = "Enough Financial Information Provided\?"
Public const Are_Cost_Designations_Reasonable_CBLOI_Review_acc_name_Prod = "Are Cost Designations Reasonable\?" 


'''''''''''''''''Categories of nonresponsiveness''''''''''''''''''''
Public const Categories_of_nonresponsiveness_CBLOI_Review = "CATEGORIES OF NONRESPONSIVENESS The Engagement Award Program does not fund projects that: • Do not have a clear focus on patient-centered CER\. • Include a cost-effectiveness analysis of alternative approaches to providing care\. • Solely intend to increase patient engagement in health care or healthcare systems rather than healthcare research, particularly patient-centered CER\. • Design or test healthcare interventions\. • Involve the use of a drug or medical device\. • Create clinical practice guidelines, care protocols, or decision support tools\. • Create coverage, payment, or policy recommendations or guidelines\. • Address only quality measures, quality improvement, or engagement around quality measures\. • Recruit and enroll patients for clinical trials\. • Involve patients only as subjects\. • Are research studies, such as randomized controlled trials, observational studies, pragmatic clinical studies, and systematic reviews\. • Create or maintain a registry or recruit people for a registry\. • Are designed solely to validate tools or instruments not created through a PCORI-funded project\. • Develop funding proposals, grant applications, or similar requests for project or research funding\. • Involve grantmaking, such as giving grants using PCORI's project funds\. • Focus solely on social determinants of health, with no focus on patient-centered CER\. • Plan to disseminate research results without including PCORI-funded research or related products\. • Implement PCORI-funded findings in a clinical practice setting\. PCORI funds this type of project in a different program\. • Aim to influence, directly or indirectly, any federal, state, or local laws, regulations, judicial decisions, or the like, including preparation or planning activities, research, and other background work related to or in contemplation of lobbying activities\. • Aim to create an independent corporation \(nonprofit or for-profit\), limited liability company, partnership, or other legal entity\. If any listed category applies to this LOI, it may be considered nonresponsive for an Engagement Award\. If you are unsure, please explain your thoughts in the ""LOI Review Feedback"" field and mark your Preliminary Disposition as ""Unsure\."""
Public const Is_this_nonresponsive_CBLOI_Review = "Is this nonresponsive\? --None-- Please specify whether the LOI is non responsive or not before submitting the review\."
Public const Is_this_nonresponsive_CBLOI_Review_acc_name = "Is this nonresponsive\?"
Public const all_activities_that_apply_CBLOI_Review_picklist = "Do not have a clear focus on patient-centered CER\. Include a cost-effectiveness analysis of alternative approaches to providing care\. Solely intend to increase patient engagement in health care or healthcare systems rather than healthcare research, particularly patient-centered CER\. Design or test healthcare interventions\. Involve the use of a drug or medical device\. Create clinical practice guidelines, care protocols, or decision support tools\. Create coverage, payment, or policy recommendations or guidelines\. Address only quality measures, quality improvement, or engagement around quality measures\. Recruit and enroll patients for clinical trials\. Involve patients only as subjects\. Are research studies, such as randomized controlled trials, observational studies, pragmatic clinical studies, and systematic reviews\. Create or maintain a registry or recruit people for a registry\. Are designed solely to validate tools or instruments not created through a PCORI-funded project\. Develop funding proposals, grant applications, or similar requests for project or research funding\. Involve grantmaking, such as giving grants using PCORI's project funds\. Focus solely on social determinants of health, with no focus on patient-centered CER\. Plan to disseminate research results without including PCORI-funded research or related products\. Implement PCORI-funded findings in a clinical practice setting\. PCORI funds this type of project in a different program\. Aim to influence, directly or indirectly, any fed\., state, or local laws, regulations, judicial decisions, or the like, including preparation or planning activities, research, & other background work related to or in contemplation of lobbying activities\. Aim to create an independent corporation \(nonprofit or for-profit\), limited liability company, partnership, or other legal entity\."
Public const all_activities_that_apply_CBLOI_Review_field = "If yes, select all activities that apply"

''''''''''''''Categories of nonresponsiveness' CB/DI/SCS acc_Name of drop down field value for Prod''''''''
Public const Is_this_nonresponsive_CBLOI_Review_acc_name_Prod = "Is this nonresponsive\?"

''''''''''''''''''Reviewer Disposition Section''''''
Public const Reviewer_Disposition_Feedback_CBLOI_Review = "REVIEWER DISPOSITION AND FEEDBACK INSTRUCTIONS • Reviewer Disposition: Select ""Invite to Submit a Full Proposal,"" ""Unsure/Further Discussion Needed,"" ""Reject,"" or ""Nonresponsive"" • Level of Enthusiasm: Select ""High,"" ""Medium,"" or ""Low"" • LOI Review Comments: Provide your comments and feedback on the LOI based on the review criteria listed above\. What is the main reason or rationale for your recommendation\? Explain your level of enthusiasm\. • Applicant Feedback, if Invited: What questions or feedback do you have for the applicant to address in their full proposal\? \(Enter ""n/a"" if recommendation is to reject\) • Applicant Feedback, if Rejected: What feedback do you want to provide the applicant to explain your review decision and/or assist them with a future submission\? \(Enter ""n/a"" if recommendation is to invite\) • Additional Reviewer Needed\? Select ""Yes"" or ""No"" • Additional Reviewer Request: If ""Yes,"" provide the name of the individual who you recommend review the LOI, or identify the subject matter expertise that is needed\. Briefly explain why\."
Public const Reviewer_Disposition_CBLOI_Review = "Reviewer Disposition --None-- You must choose a specific value for ""Reviewer Disposition"" field prior to submission"
Public const Reviewer_Disposition_CBLOI_Review_acc_name = "Reviewer Disposition"
Public const Level_of_Enthusiasm_CBLOI_Review = "Level of Enthusiasm --None-- Please pick a value in the Field Prior To Submission"
Public const Level_of_Enthusiasm_CBLOI_Review_acc_name = "Level of Enthusiasm"
Public const LOI_Review_Comments_EALOI_Review = "LOI Review Comments Please Specify The LOI Review Feedback Prior To Submission"
Public const Applicant_Feedback_Invited_EALOI_Review = "Applicant Feedback, if Invited Please provide feedback prior to submission\. Enter “n/a” if this field is not applicable to your disposition\."
Public const Applicant_Feedback_Rejected_EALOI_Review = "Applicant Feedback, if Rejected Please provide feedback prior to submission\. Enter “n/a” if this field is not applicable to your disposition\."
Public const Additional_Reviewer_Needed_EALOI_Review = "Additional Reviewer Needed\? --None-- If any Additional Reviewer is Needed Please Specify the Name By choosing a value from ""Additional Reviewer Needed"" field"
Public const Additional_Reviewer_Needed_EALOI_Review_acc_name = "Additional Reviewer Needed\?"
Public const Additional_Reviewer_Request_EALOI_Review = "Additional Reviewer Request"

'''''''''''''Reviewer Disposition Section' CB/DI/SCS acc_Name of drop down field value for Prod''''''''
Public const Reviewer_Disposition_CBLOI_Review_acc_name_Prod = "Reviewer Disposition"
Public const Level_of_Enthusiasm_CBLOI_Review_acc_name_Prod = "Level of Enthusiasm"
Public const Additional_Reviewer_Needed_EALOI_Review_acc_name_Prod = "Additional Reviewer Needed\?"


''''''''''''EA LOI Review Status dropdown Field acc_name for DI/CB/SCS in PQT
Public const EA_LOI_Review_Status_acc_name = "Status"

''''''''''''EA LOI Review Status dropdown Field acc_name for DI/CB/SCS in PRod
Public const EA_LOI_Review_Status_acc_name_Prod = "Status"


''''''''''''''''''''''''''''''''**************************'''''''''''''''''''''''''''''''''''''''''
''''''''''''''EADI LOI Review Form''''''''''

Public const EADI_LOI_Review_Required_Field_Error = "Review the following fields Project Team Experience Adequate\? Sufficient Fit\? Proposed Collaborators Appropriate\? Additional Reviewer Needed\? Innovative Idea\? Engagement Plan Adequate\? Is this nonresponsive\? Reviewer Disposition Reasonable Goals, Objectives, Outcomes\? Enough Financial Information Provided\? Cooperation Concerns\? I Have a Conflict with Application Project Lead Qualification Appropriate\? Are Cost Designations Reasonable\? Level of Enthusiasm LOI Review Comments Activities & Deliverables Support Goal\?"

'''''''''''''''''Criteria 1'''''''''''DI
Public const Criteria1_Program_Fit_DILOI_Review = "Criteria 1: Program Fit How well does the application demonstrate program fit or how well does the application proffer an innovative idea that supports the organization’s mission\? • Does not include any non-responsive activities from 'Categories of Nonresponsiveness' list • Describes the problem and proposed solution  • Explains how the project is clearly focused on and connected to patient-centered comparative clinical effectiveness research \(CER\)\. • Information and/or tools generated by the project must be transferrable and of interest or use not just to the applicant organization but to others doing related work • Clearly describes an effective strategy that actively disseminates PCORI’s evidence • Demonstrates an opportunity to effectively leverage PCORI’s research findings • Clearly identifies which eligible PCORI-funded study\(ies\) will be disseminated and why the findings are relevant to the specific targeted end users for dissemination • \[REVIEWER TWO ONLY\] In an effort to avoid redundancy, do the goals or objectives of the project resemble those of a current or past EA project, to the best of your knowledge\? \(If uncertain, say that you are unsure in your response\)"

'''''''''''''''''Criteria 2 DI'''''''''''''''
Public const Criteria2_Project_Plan_DILOI_Review = "Criteria 2: Project Plan How reasonable is the applicant's project plan\? Are there any concerns that the applicant has not provided a reasonable set of activities to achieve the objectives and outcomes in the project plan\? • Describes the aims and goals of the project, including the objectives  • Describes the project methods, activities, and strategies that will be employed to support the overall goal \(includes alternative plans for convening should an in-person meeting not be feasible\) • Specifies the projected short- \(during the PCORI-funded project\), medium- \(0-2 years post-funding\), and long-term \(3\+ years post-funding\) outcomes; States the significance of the outcomes • Clearly states the primary goal of the dissemination effort • Describes active dissemination strategies to targeted end users and justifies the choice of these strategies • Clearly identifies capacity building versus active dissemination activities and commits to primarily focusing on dissemination by emphasizing strategies for disseminating eligible PCORI-funded research results • Provides an estimate of the number of people who will be reached \(short-, medium-, and long-term\) during the dissemination project • Describes a plan for evaluating the impact of the dissemination effort, including metrics for measuring the success of the dissemination strategy • Clearly describes projected outputs \(including sustainability plans, evaluation analysis, and resources\) • Includes a brief outline of an evaluation plan focused on knowledge sharing and transfer  • Describes plan to use a tool/resource within the population \(if applicable\)  • Describes any tools/trainings/programs that will be used as part of the project, and the evidence base for the resources that will be used  • Provides a reasonable timeline to complete activities  • \[REVIEWER TWO ONLY\] If this is a resubmission of either an LOI or a full proposal, did the applicant adequately address the questions posed in the feedback letter\(s\) they previously received\?"

'''''''''''''''''Criteria 3 DI'''''''''''''''
Public const Criteria3_Lead_Experience_Organizational_Capabilities_DILOI_Review = "Criteria 3: Project Lead Previous Experience and Organizational Capabilities Do the qualifications of the project lead align with the scope of the project\? Does the applicant demonstrate sufficient background/organizational capabilities for related projects that are aligned with an emphasis on patient-centered CER\? • Describes the project lead’s previous experience related to patient-centered CER • Demonstrates sufficient organizational capabilities for projects with an emphasis on patient-centered CER • Demonstrates sufficient organizational capabilities for projects with an emphasis on disseminating evidence of any kind • Explains the project lead, team, and organization’s relationship to the targeted end users and experience in successfully disseminating research findings and bringing evidence to them • Demonstrates involvement of adequate and qualified personnel in conducting the project • The involvement of the project team demonstrates a reasonable level of effort • Explains involvement and achievements in other patient-centered CER or PCORI-funded projects • Provides clear examples of successful prior active dissemination efforts"

'''''''''''''''''Criteria 4 DI'''''''''''''''
Public const Criteria4_Patient_Stakeholder_Engagement_Plan_DILOI_Review = "Criteria 4: Patient and Stakeholder Engagement Plan Is there an adequate plan for engaging patients and other stakeholders in the conduct of the proposed project\? Are the proposed collaborators meaningful and appropriate based on aligning the interest, expertise and scope of work of patients and other stakeholders involved in the project\? Are there any concerns about the ability of the collaborators to work together\? • Explains who the patient and stakeholder partners are and any preexisting engagement that has taken place before • Describes how stakeholders \(i\.e\., patients, caregivers, and clinicians\) and the organizations that represent them were involved in developing the proposed project • Proposed collaborators are meaningful and appropriate based on aligning the interest, expertise, and scope of work of patients and other stakeholders involved in the project • Means of collaboration \(i\.e\., meetings, etc\.\) and frequency of interactions have been addressed • Describes the specific responsibilities that key stakeholder partners will have in executing the proposed dissemination strategy • Describes organization’s experience in reaching stakeholders proposed to actively disseminate PCORI-funded research findings to"

'''''''''''''''''Criteria 5 DI'''''''''''''''
Public const Criteria5_Budget_Cost_Proposal_DILOI_Review = "Criteria 5: Budget/Cost Proposal How well does the applicant’s budget proposal demonstrate reasonable cost designations\? • Provides clear amounts & justification for each budget item • A reasonable percentage of the budget is dedicated to dissemination activities • Staff compensation is not an excessive proportion of the budget request; Applicants should keep personnel costs \(applicant organization/institution staff\) below 50% of the total project budget\. However, higher personnel costs may be considered with strong justification\. • Reflects the time and contributions of all partners, including patients and stakeholders\. Fair financial compensation demonstrates that patients, caregivers, and patient/caregiver organizations’ contributions to the project, including related commitments of time and effort, are valuable and valued\."

''''''''Reviewer Disposation and feedback instructions LOI Review DI ''''''
Public const Reviewer_Disposition_Feedback_DILOI_Review = "REVIEWER DISPOSITION AND FEEDBACK INSTRUCTIONS Note: Before you determine a disposition and complete the fields below, please ensure you have reviewed the LOI and the ""Dissemination Initiative LOI Supplemental Template"" that is attached to the LOI in the ""Files"" section of Salesforce\.  • Reviewer Disposition: Select ""Invite to Submit a Full Proposal,"" ""Unsure/Further Discussion Needed,"" ""Reject,"" or ""Nonresponsive"" • Level of Enthusiasm: Select ""High,"" ""Medium,"" or ""Low"" • LOI Review Comments: Provide your comments and feedback on the LOI based on the review criteria listed above\. Use this space to also provide feedback on the ""Dissemination Initiative LOI Supplemental Template"" that is attached to the LOI in the ""Files"" section of Salesforce\. What is the main reason or rationale for your recommendation\?  • Applicant Feedback, if Invited: What questions or feedback do you have for the applicant to address in their full proposal\? • Applicant Feedback, if Rejected: What feedback do you want to provide the applicant to explain your review decision and/or assist them with a future submission\? • Additional Reviewer Needed\? Select ""Yes"" or ""No"" • Additional Reviewer Request: If ""Yes,"" provide the name of the individual who you recommend review the LOI, or identify the subject matter expertise that is needed\. Briefly explain why\."

''''''''''''''SUBMIT YOUR REVIEW''' section'''''
Public const Submit_Your_Review_EALOIReview = "SUBMIT YOUR REVIEW When you are finished entering information into this form, check the box next to ""Review Complete\?"" and click ""Save"" to formally submit your review to the Engagement Award Program\. Need to Edit After Submitting\? Uncheck the box next to ""Review Complete\?"" and click ""Save"" to restore edit capabilities\. Then, edit the form and follow the steps above to resubmit the review form\."
Public const Submit_Your_Review_EALOIReview_ItalicText = "<em>Need to Edit After Submitting\?</em>"
Public const Review_Complete_EALOIReview = "Review Complete\? Review Complete\? Help Info The ""Review Complete"" checkbox must be unselected to make additional edits\."

''''''''''''''''''''''''''''''''**************************'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''EA LOI Review form SCS''''''''''''''''''

'''''''''''''''''Criteria 1'''''''''''SCS
Public const Criteria1_Program_Fit_SCSLOI_Review = "Criteria 1: Program Fit How well does the application demonstrate program fit or how well does the application proffer an innovative idea that supports the organization’s mission\? • Does not include any non-responsive activities from 'Categories of Nonresponsiveness' list • Describes the problem and proposed solution  • Explains how the project is clearly focused on and supports engagement in patient-centered comparative clinical effectiveness research \(CER\) • Demonstrates an understanding of patient-centered CER \(This includes understanding the characteristics specific to patient-centeredness as well as an understanding that CER is not simply engaged research of any kind\.\) • Information and/or tools generated by the project must be transferrable and of interest or use not just to the applicant organization but to others doing related work  • Describes the opportunity to promote the use of a PCORI-funded tool\(s\)/resource\(s\) to expand participation in patient-centered CER to address the problem \(if applicable\)  • Describes the opportunity or need to convene the proposed stakeholders around the proposed topic area • All meetings under this PFA, no matter the health topic, should include one or more of the following focus areas: \(a\) Facilitating Discussion Around Development of Topics, Questions, and Engagement Plans for Patient-Centered CER, \(b\) Sharing Engagement Methods and Fostering Partnerships, \(c\) Convening Around the Life Cycle of Patient-Centered CER, Including Strategies for Dissemination, or \(d\) Dissemination of Research Findings • Applicants who plan to convene around dissemination of PCORI-funded research findings must identify the eligible, completed PCORI-funded evidence in the LOI and full proposal and provide the key messages they plan to disseminate from the evidence • \[REVIEWER TWO ONLY\] In an effort to avoid redundancy, do the goals or objectives of the project resemble those of a current or past EA project, to the best of your knowledge\? \(If uncertain, say that you are unsure in your response\)"

''''''''''''''''Criteria 2 SCS''''''
Public const Criteria2_Project_Plan_SCSLOI_Review = "Criteria 2: Project Plan How reasonable is the applicant's project plan\? Are there any concerns that the applicant has not provided a reasonable set of activities to achieve the objectives and outcomes in the project plan\? • Describes the aims and goals of the project, including the objectives  • Describes the project methods, activities, and strategies that will be employed to support the overall goal \(includes alternative plans for convening should an in-person meeting not be feasible\)  • Specifies the projected short- \(during the PCORI-funded project\), medium- \(0-2 years post-funding\), and long-term \(3\+ years post-funding\) outcomes; States the significance of the outcomes  • Clearly describes projected outputs \(including sustainability plans, evaluation analysis, and resources\) • Includes a brief outline of an evaluation plan focused on knowledge sharing and transfer  • If the applicant proposes creating a research agenda, does the LOI indicate that the agenda will be focused on or lend itself specifically to patient-centered CER\? • Describes the central focus or shared priority for the convening that unifies the stakeholders \(i\.e\. geography, health condition, population\) to explore issues related to patient-centered CER, or communicate PCORI-funded research findings to targeted end-user audiences\. • For LOIs focused on IDD: Does the project address IDD in general or a subtopic of IDD \(e\.g\., autism\)\? If the project focuses on IDD in general, are the project activities and outputs relevant to all conditions\? Does the applicant discuss how the project will address the differences among populations and remain relevant to all\? • Describes plan to use a PCORI-funded tool/resource within the population \(if applicable\)  • Describes any tools/trainings/programs that will be used as part of the project, and the evidence base for the resources that will be used  • Provides a reasonable timeline to complete activities  • Convening date\(s\) should be planned so that there is adequate time to allow for meaningful stakeholder engagement prior to the convening, PCORI involvement in the planning, and to provide ample time for evaluation and resource development post-convening\(s\)\. Convenings should not occur at the start of the project period\. • \[REVIEWER TWO ONLY\] If this is a resubmission of either an LOI or a full proposal, did the applicant adequately address the questions posed in the feedback letter\(s\) they previously received\?"

''''''''''''''''Criteria 3 SCS''''''
Public const Criteria3_Lead_Experience_Organizational_Capabilities_SCSLOI_Review = "Criteria 3: Project Lead Previous Experience and Organizational Capabilities Do the qualifications of the project lead align with the scope of the project\? Does the applicant demonstrate sufficient background/organizational capabilities for related projects that are aligned with an emphasis on patient-centered CER\? • Describes the project lead’s previous experience related to patient-centered CER • Demonstrates sufficient organizational capabilities for projects with an emphasis on patient-centered CER • Demonstrates involvement of adequate and qualified personnel in conducting the project • The involvement of the project team demonstrates a reasonable level of effort • If planning to convene around dissemination of research findings: \(a\) Explains the project lead, team, and organization’s relationship to the targeted end users and experience in successfully disseminating research findings and bringing evidence to them, and \(b\) If planning to convene around dissemination of research findings, applicant provides examples of successful prior dissemination efforts • Explains involvement and achievements in other patient-centered CER or PCORI-funded projects • Provides clear examples of successful prior efforts in planning and facilitating convenings"
Public const Criteria3_Lead_Experience_Organizational_Capabilities_SCSLOI_Review_Italic_text = "<em>Do the qualifications of the project lead align with the scope of the project\? Does the applicant demonstrate sufficient background/organizational capabilities for related projects that are aligned with an emphasis on patient-centered CER\?</em>"


''''''''''''''''Criteria 4 SCS''''''
Public const Criteria4_Patient_Stakeholder_Engagement_Plan_SCSLOI_Review = "Criteria 4: Patient and Stakeholder Engagement Plan Is there an adequate plan for engaging patients and other stakeholders in the conduct of the proposed project\? Are the proposed collaborators meaningful and appropriate based on aligning the interest, expertise and scope of work of patients and other stakeholders involved in the project\? Are there any concerns about the ability of the collaborators to work together\? • Explains who the patient and stakeholder partners are and any preexisting engagement that has taken place before • Describes how stakeholders \(i\.e\., patients, caregivers, and clinicians\) and the organizations that represent them were involved in developing the proposed project • Proposed collaborators are meaningful and appropriate based on aligning the interest, expertise, and scope of work of patients and other stakeholders involved in the project • Means of collaboration \(i\.e\., meetings, etc\.\) and frequency of interactions have been addressed • If planning to convene around dissemination of PCORI-funded research findings: \(a\) Describes the specific responsibilities that key stakeholder partners will have in executing the proposed dissemination strategy, \(b\) Describes organization’s experience in reaching stakeholders proposed to disseminate PCORI research findings to, and \(c\) Explains how the proposed stakeholders can benefit from the evidence proposed for dissemination﻿"

Public const Criteria4_Patient_Stakeholder_Engagement_Plan_CBSCS_Review_Italic_Text = "<em>Is there an adequate plan for engaging patients and other stakeholders in the conduct of the proposed project\? Are the proposed collaborators meaningful and appropriate based on aligning the interest, expertise and scope of work of patients and other stakeholders involved in the project\? Are there any concerns about the ability of the collaborators to work together\?</em>"













'''''''''''''-------------------------------------------------------- FUNCTIONS START --------------------------------------------------------------------------------------------------------------------------------

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

Public Function Open_ReviewersPortal()
Err.Clear
On Error Resume Next
Wait 3

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)			 
	'SystemUtil.CloseProcessByName IEBrowser
	'SystemUtil.Run IEBrowser , Reviewerurl,,,3
	SystemUtil.CloseProcessByName ChromeBrowser
	SystemUtil.Run ChromeBrowser, Reviewerurl,,,3 
	wait 3

				If err <> 0 Then    	
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"Opening the Portal" & "Failed","Opening the Portal" & "-" & "Failed to open"
				Else 
					LogReport 0,"Opening the Portal" & "Success","Opening the Portal" & "-" & "Was successful"
				End If   

End Function

''------------------------------------------   IMPERSONATION FUNCTIONS START    --------------------------------------------------------
Function impersonate_as(userName)
Err.Clear
On error resume next
Wait 3

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

				If err <> 0 Then    	
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"impersonate_as - " & userName,"impersonate as" & "-" &userName& "Failed to impersonate"
				Else 
					LogReport 0,"impersonate_as - " & userName,"impersonate as" & "-" &userName& "successfully impersonated"
				End If
				
End Function

'' Function to Impersonate as an external user
Function ImpersonateExternalUserLogin (UserName)
Err.Clear
On error resume next
Wait 3

	Dim Brwser
	Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
	
	navigateToAccoutsPageThrughTopSearchBox UserName
	                
	Brwser.Webtable("html tag:=TABLE", "column names:=Action;Name;Awardee Institution/Organization;Primary Reviewer Role;Reviewer Status;Do Not Invite Qualifier;Email;PCORI Identified Programs").Link("innertext:="&UserName).Click
	
	wait 5
	If Browser("micclass:=browser").Page("micclass:=page").Link("html tag:=A", "innertext:="&UserName).Exist(2) Then
	          Browser("micclass:=browser").Page("micclass:=page").Link("html tag:=A", "innertext:="&UserName).Click
	End If
	
	wait 2
	click_webElement_any "workWithPortalLabel","Manage External User"
	clk_link_Object2 "Log in to Community as User"

				If err <> 0 Then    	
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"ImpersonateExternalUserLogin - " & UserName,"Impersonate as External User" & "-" &UserName& " - Failed to impersonate"
				Else 
					LogReport 0,"ImpersonateExternalUserLogin - " & UserName,"Impersonate as External User" & "-" &UserName& " - successfully impersonated"
				End If


End Function

Function logOut_asUser(userName)
Err.Clear
On error resume next
Wait 3

	HtmlID = "globalHeaderNameMink"
	clk_onlink_forimpersonateinternal HtmlID,userName
	wait 2
	clk_onlogoutwhendone_impersonating "Logout"
	wait 2

				If err <> 0 Then    	
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"logOut_asUser - " & userName,"Log OUT as User" & "-" &userName& " - Failed to log out"
				Else 
					LogReport 0,"logOut_asUser - " & userName,"Log OUT as User" & "-" &userName& " - successfully log out"
				End If
				
End Function

Function clk_link_Objectforimpersonate2()
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
	l("micclass").value = "Link"
	l("html id").value = "moderatorMutton"
	l("html tag").value = "A"
	l("title").value = "User Action Menu"
  
  
 Set lo =  getParentObject().ChildObjects(l)
 print lo.count
  
			    If lo.count = 0 Then
			  		LogReport 1,"clk_link_Objectforimpersonate - " & "to impersonate","Failed to find link" & "-" & "to impersonate"
			  	else
			  		LogReport 0,"clk_link_Objectforimpersonate - " & "to impersonate","The Link" & "-" & "to impersonate" & "-" & "is found succesfully"
			    End If

 lo(0).click
 
End Function
     
Function clk_onlink_forimpersonateinternal(HtmlID,strName)
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
	l("micclass").value = "Link"
	l("innertext").value = strName
	l("html id").value = HtmlID
	l("html tag").value = "A"
		  
		  
Set lo =  getParentObject().ChildObjects(l)
print lo.count
  
			If lo.count = 0 Then
				LogReport 1,"clk_onlink_forimpersonateinternal - " & strName,"Link with innertext" & "-" & strName &"NOT found - NOT expected"
			else
				LogReport 0,"clk_onlink_forimpersonateinternal - " & strName,"Link with innertext" & "-" & strName & "-" & "FOUND succesfully - Expected"
			End If

lo(0).click

End Function
     
''------------------------------------------   IMPERSONATION FUNCTIONS  END   --------------------------------------------------------



'Public Function Open_SalesForce_Application_PartnerUsers()
'On Error Resume Next
'Set btncalc = Description.Create()  
'btncalc("micclass").value = "Browser"
'
'Set btn =DeskTop.ChildObjects(btncalc)			 
'SystemUtil.CloseProcessByName IEBrowser
'SystemUtil.CloseProcessByName ChromeBrowser
'SystemUtil.Run ChromeBrowser ,url,,,3 
'SystemUtil.Run IEBrowser , "https://sit-pcori.cs91.force.com/engagement",,,3
'wait 3
'
'If err <> 0 Then    	
'LogReport 4,"Error", err.number & "-" & err.description
'LogReport 1,"Opening the Sales Force Application" , "-" & "Failed to open"
'Else 
'LogReport 0,"Opening the Sales Force Application" , "-" & "Was successfull"
'End If   
'End Function

Function populateKeyProgramOfficerResearchAwards()
Err.Clear
On error resume next
Wait 3

		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("column names").value="Program Officer.*"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 		  	       
				For x  = 1 to a     
						For j  = 1 to b
							c = l_link(0).GetCellData(x,j)
							print c
							
							Select Case c
							
						    	Case "Program Officer"
                                 	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit1.set "Geeta Bhat"
                                       
                                Case "Program Associate"
                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit1.set "Evelyn Whitlock"
                                       
                                Case "Engagement Officer"
                                  	   Set oEdit4 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit4.set "Chinenye Anyanwu"
                                       
                                Case "Contract Administrator"
                                  	   Set oEdit3 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit3.set "CMA Admin ( Test User )"
                                       
                                Case "Contract Coordinator"
                                  	   Set oEdit3 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit3.set "CMA Coordinator (Test User)"
							                                                                    
							End Select     
                         						
						
						Next
						
				Next                               
End Function 

Function PopulateProjectReaserchAwardSummary()
Err.Clear
On error resume next
Wait 3

		intnumber =  TEST & RandomString(3)   '''Int((1000000-1)*Rnd+1)
		
		d = Date + 365
		'''''strNameProject = "Reasearch Award_" & RandomString(3)
		
				
		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("column names").value="Project Record Type.*"
		
		Set l_link=  getParentObject().ChildObjects(odesc)
		print l_link.count
		a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 		  	       
				For x  = 1 to a     
						For j  =1 to b
							c = l_link(0).GetCellData(x,j)
							print c
							Select Case c
						    	
                                    Case "Contract Number"
                                  	   Set oEdit4 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit4.set intnumber
                                       
                                    Case "*Status"
                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
                                       oEdit1.select "Executed"
                                       
'''''                                 	Case "*Short Project Title"
'''''                                 	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
'''''                                       oEdit1.set strNameProject 
'''''                                       
'''''                                     Case "*Full Project Title5000 remaining"
'''''                                 	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
'''''                                       oEdit1.set strNameProject
                                       
                                     Case "Awardee Institution/Organization"
                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit1.set InstitOrg
                                      
                                     Case "Contract End Date*"
                                  	   Set oEdit3 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit3.set d
                                     
                                      Case "*Award Date"
                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit1.set "1/7/2017" 
                                       
                                      Case "Contract Start Date*"
                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit1.set "1/7/2017" 
                                       
                                      'Case "Program"
                                  	   'Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
                                       'oEdit1.select "Clinical Effectiveness Research"                                       
                                       
                                      Case "Application Number"
                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit1.set intnumber 
                                       
                                     Case "Kickoff"
                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit1.set Date +30
                                       
                                    Case "Award Date"
                                 	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                     oEdit1.set "5/7/2017"
                                     
                                     Case "Research Period End Date"
                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit1.set Date +300 
                                       
'''''                                     Case "PFA"
'''''                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
'''''                                       oEdit1.select testPFA     
'''''                                       
'''''                                       
'''''                                     Case "Primary Campaign Source"
'''''                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
'''''                                       oEdit1.set testCampaign 
                                       
                                     Case "CER Category"
                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 0)
                                       oEdit1.select "Studies of Interventions for Caregivers" 
                                       
                                    Case "*Invoice Type"
'''                                  	   Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebList", 1)
'''                                       oEdit1.select InvoiceType
                                       HtmlID = "00N39000004WxDx"
                                       strData = "Research Awards – Cost Reimbursable"
                                       selectWeblist_PI_Information strData,HtmlID
							                                                                    
							End Select     
                         						
						
						Next
						
				Next                             
        wait 5
PopulateProjectReaserchAwardSummary = strNameProject 
     	
End Function  
  
''Function to click on ANY button based on specific button NAME as parameter
Public Function clk_Button_usingName(strName)
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
	print pHwnd
	
	Set editO = Description.Create
	
	editO("micclass").value = "WebButton"
	editO("visible").value = true  
	editO("name").value = strName
	
Set editObject =  getParentObject().ChildObjects(editO) 
print editObject.count
editObject(0).highlight

			If editObject.count = 0 Then
				LogReport 1,"clk_Button_usingName - " & strName,"Button with name" & "-" & strName &"NOT found - NOT expected"
			else
				LogReport 0,"clk_Button_usingName - " & strName,"Button with name" & "-" & strName & "-" & "is FOUND succesfully"
			End If

editObject(0).click

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

''Function to click link based on Name and INDEX (if there are many links with the same name on the page)
Function clk_link_byName_Index(strName,i)
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
lo(i).Highlight

		    If lo.count = 0 Then
		  		LogReport 1,"clk_link_byName_Index - " & strName,"Link with name" & "-" & strName &"NOT found - NOT expected"
		  	else
		  		LogReport 0,"clk_link_byName_Index - " & strName,"Link with Name" & "-" & strName & "-" & "is FOUND succesfully"
		  	End If

lo(i).click

End Function

'' Function to search for anything using GLOBAL Search internally
Public Sub navigateToAccoutsPageThrughTopSearchBox(strData)
Err.Clear
On error resume next
Wait 3

	setWebEditBox(strData)
	wait 2
	'clk_Button_usingName "Search"
	wait 2
	
		    If err <> 0 Then
		  		LogReport 4,"Error", err.number & "-" & err.description
		  		LogReport 1,"navigateToAccoutsPageThrughTopSearchBox - " & strData,"Failed to navigate through Global search to " & "-" & strData
		  	else
		  		LogReport 0,"navigateToAccoutsPageThrughTopSearchBox - " & strData,"Navigated through Global search to " & "-" & strData & "-" & "succesfully"
		  	End If
		  	
End Sub 

''Function to set Web Edit box to certain value or Values - if Multiple it will fill out in a sequence one by one (use coma to separate input value parameters)
Function setWebEditBox(strData)
Err.Clear
On error resume next
Wait 3

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
	
	Set edObj =  getParentObject().ChildObjects(editO) 

for i = 0 To uBound(strArr)
items = edObj(i).GetRoProperty("all items")
print items
print strArr(i)
If Not(strArr(i) = "")Then

edObj(i).set strArr(i)  
oShell.SendKeys strSearchText				            
End If

Next
		
		
		    If err <> 0 Then
		  		LogReport 4,"Error", err.number & "-" & err.description
		  		LogReport 1,"setWebEditBox - " & strData,"Failed to set web edit box to " & "-" & strData
		  	else
		  		LogReport 0,"setWebEditBox - " & strData,"Set web edit box to " & "-" & strData & "-" & "was done succesfully"
		  	End If

End Function

''Function to select certain value from WebList - if Multiple web Lists on the page it will fill out in a sequence one by one (use coma to separate input value parameters)
Function selectWeblist(strData)
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

''Function to verify if specific Web Element with specific Class parameter does NOT exist - used mostly on Portal for Paper/Pencil Icon
Public Function verify_ifWebElement_NOT_exist_byClass (strName)
Err.Clear
On error resume next
Wait 3

	Set l = Description.Create
	l("micclass").value = "WebElement"
	l("class").value = strName 

Set lo =  getParentObject().ChildObjects(l)

				If lo.count = 0 Then      
					LogReport 0,"verify_ifWebElement_NOT_exist_byClass - " & strName, "WebElement with class--" & strName & "- NOT exists - Expected"
				else
					LogReport 1,"verify_ifWebElement_NOT_exist_byClass - " & strName, "Webelement with class--" & strName & "- Exists - Not Expected"			
				End If

End Function

''Function to verify that ANY web button with specified Name as parameter Exists
Public Function verifyIfButtonDoesExist(strName)
Err.Clear
On error resume next
Wait 3

	Set btncalc = Description.Create() 
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	
	Set pa = Description.Create
	pa("micclass").value = "Page"
	
	Set l = Description.Create
	l("micclass").value = "WebButton"
	l("name").value = strName

Set lo =  getParentObject().ChildObjects(l)
lo(0).highlight 

				If lo.count = 0 Then      
					LogReport 1," verifyIfButtonDoesExist - "&strName, "WebButton --" & strName & "- does NOT exist - Not Expected"
				else
					LogReport 0," verifyIfButtonDoesExist - "&strName, "WebButton --" & strName & "- Exists - Expected"
				End If

End Function

''Function to verify that ANY web button with specified Name and Class as parameters Exists
Public Function verify_IfButtonExist_class_name(nclass,strName)
Err.Clear
On error resume next
Wait 3

	Set btncalc = Description.Create() 
	btncalc("micclass").value = "Browser"
	
	Set btn = DeskTop.ChildObjects(btncalc)
	
	Set pa = Description.Create
	pa("micclass").value = "Page"
	
	Set l = Description.Create
	l("micclass").value = "WebButton"
	l("class").value = nclass
	l("name").value = strName

Set lo =  getParentObject().ChildObjects(l)
lo(0).highlight 

				If lo.count = 0 Then      
					LogReport 1," verify_IfButtonExist_class_name - "&strName, "WebButton --" & strName & "- does NOT exist - Not Expected"
				else
					LogReport 0," verify_IfButtonExist_class_name - "&strName, "WebButton --" & strName & "- Exists - Expected"
				End If

End Function

''Function to verify that Web Button with specified Name parameter doe NOT exist
Public Function verifyIfButtonDoesnotExist(strName)
Err.Clear
On error resume next
Wait 3

	Set btncalc = Description.Create() 
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	
	Set pa = Description.Create
	pa("micclass").value = "Page"
	
	Set l = Description.Create
	l("micclass").value = "WebButton"
	l("name").value = strName

Set lo =  getParentObject().ChildObjects(l)
print lo.count

				If lo.count > 0 Then      
					LogReport 1,"verifyIfButtonDoesnotExist- "&strName, "WebButton --" & strName & "- exists - Not Expcted"
				else
					LogReport 0,"verifyIfButtonDoesnotExist- "&strName, "WebButton --" & strName & "- does not Exist - Expected"
				
				End If

End Function

''Function to verify that Web Element with specified Innertext parameter Exists
Public Function verifyIfWebElementDoesExist2(innertext)
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
	l("innertext").value = innertext
	l("visible").value = True 
	
Set lo = getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

				If lo.count = 0 Then      
					LogReport 1,"verifyIfWebElementDoesExist2 - " & innertext, "WebElement --" & innertext & "- Doesn't Exist - Not Expected"
				else
					LogReport 0,"verifyIfWebElementDoesExist2 - " & innertext, "WebElement --" & innertext & "- Exists - Expected"
				End If

End Function

''Function to verify if Web Element with specified Html ID as parameter exists
Public Function verifyIfWebElementDoesExist_HtmlId (HtmlId)
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
	l("visible").value = True 
	l("html id").value = HtmlId
	
	strcount = l.count
	print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

				If lo.count = 0 Then      
				LogReport 1,"verifyIfWebElementDoesExist_HtmlId - " & HtmlId, "The WebElement --" & HtmlId & "- Doesn't Exist - Not Expected"
				else
				LogReport 0,"verifyIfWebElementDoesExist_HtmlId - " & HtmlId, "The WebElement --" & HtmlId & "- Exists - Expected"
				End If

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

''Function to verify if specified App as parameter selected from Blue Round at the top right Internally, if NOT - select the right App
Function BlueRoundValue_Exist (Innertext)
Err.Clear
On error resume next
Wait 3

	Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
	
	'''
	'''Set l = Description.Create
	'''l("micclass").value = "WebElement"
	'''l("html tag").value = "SPAN"
	'''l("class").value = "menuButtonLabel"
	'''l("html id").value = "tsidLabel"
	'''l("innertext").value = Innertext
	'''	
	'''	Set lo =  getParentObject().ChildObjects(l)
	
	click_webElement_BlueRound ()
	Wait 1
	            If Brwser.WebElement("class:=menuButtonLabel","html id:=tsidLabel","html tag:=SPAN", "innertext:="&Innertext).Exist(3) Then
	            		click_webElement_BlueRound ()
                		''lo(0).Highlight
                		Wait 1                      
                Else
                    	clk_link_Object2 Innertext
						Wait 2 
                End If
                
End Function

''Function to verify if Web Element with specified parameters Exists
Public Function verifyIfWebElementDoesExist1(HtmlTag, Innertext, HtmlID)
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
	l("html tag").value = HtmlTag
	l("innertext").value = Innertext
	l("html id").value = HtmlID
	 
	strcount = l.count
	print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

				If lo.count = 0 Then      
				LogReport 1,"verifyIfWebElementDoesExist1 - "& Innertext, "WebElement --" & Innertext & "- Doesn't Exist - Not Expected"
				else
				LogReport 0,"verifyIfWebElementDoesExist1 - " & Innertext, "WebElement --" & Innertext & "- Exists - Expected"
				End If

End Function

''Function to verify that Web Element with specified Class and Innertext as parameters Exists
Public Function verifyIfWebElementDoesExist_class_Innertext (classL, Innertext)
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
	l("class").value = classL
	l("innertext").value = Innertext
 
	strcount = l.count
	print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

				If lo.count = 0 Then      
				LogReport 1,"verifyIfWebElementDoesExist_class_Innertext - " & Innertext, "WebElement --" & Innertext & "- Doesn't Exist - Not Expected"
				else
				LogReport 0,"verifyIfWebElementDoesExist_class_Innertext - " & Innertext, "WebElement --" & Innertext & "- Exists - Expected"
				End If

End Function

''Function to Wait until specified Button appears on the page (page completed Loading)
Function waitForButton(strName)
Err.Clear
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

Print "Waiting for button " & "-" & strName& "-" &"Time elapsed is " & TimeDiff
'Logger Information: 
'                                                     ' "INFO","WaitForNextSpanObject","Waiting for " & "-" & strWebName & "-" &"Time elapsed is " &TimeDiff
'exit the loop if the object is not found for more than a minute
If  TimeDiff> 20 Then
'Log the results
Print "Button" & "-" & strName & "was NOT found"

Exit Do
End If    

If status >  0 Then
'Log the results   
wait 1                
print "Button was found"                    
status = 1

End If
Loop  
waitForButton =  TimeDiff       
End function


Public Function login_intoSalesForce_Application(userid,password)
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
    
    editO("micclass").value = "SFLEdit"
    editO("visible").value = true
    editO("html tag").value = "INPUT"
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
Err.Clear
On error resume next
Wait 3

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
					LogReport 1,"Get Parent Object","The parent object was not found"
				else
					LogReport 0,"Get Parent Object","The parent object was found succesfully"
				End If  

End Function

''Function to open up Salesforce and Log in: parameters are User Name and Password
Public Function navigateAndLoginToSalesForce(strUserName,strPassword)
Err.Clear
On error resume next
Wait 3

refresh_Chrome_browser()
Open_SalesForce_Application()
wait 2

'Enter valid ID and password.
login_intoSalesForce_Application strUserName, strPassword 

'Click “Log in” button.
clk_Button_usingName logInButton

				If err <> 0 Then
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"navigateAndLoginToSalesForce","Was not able to log in"
				else
					LogReport 0,"navigateAndLoginToSalesForce","Was able to log in"
				End If

End Function 
	
''Accenture function
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
Err.Clear
On error resume next
Wait 3

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

				If err <> 0 Then
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"fillForm_eventDetails","Issue filling out the event details form"
				else
					LogReport 0,"fillForm_eventDetails","NO Issue filling out the event details form"
				End If
End Function

''Function to click on a Single checkbox on a Page - very generic and works ONLY if there is ONE checkbox on a page
Function clk_singelCheckbox()
Err.Clear
On error resume next
Wait 3

	Set ck = Description.Create
	
	ck("micclass").value = "WebCheckBox"
	ck("visible").value = true  
	
	
	Set co =   getParentObject().ChildObjects(ck)
	print co.count
	
				If co.count = 0 Then
					LogReport 1,"clk_singelCheckbox - "&"only checkbox", "Single Check Box not found on a page"
				else
					LogReport 0,"clk_singelCheckbox - "&"only checkbox", "Single Check Box Found on a page"					
				End If

co(0).click
	
End Function

Function clk_Checkbox_byIndex(i)
Err.Clear
On error resume next
Wait 3

	Set ck = Description.Create
	
	ck("micclass").value = "WebCheckBox"
	ck("visible").value = true  
	
	
	Set co =   getParentObject().ChildObjects(ck)
	print co.count
	
				If co.count = 0 Then
					LogReport 1,"clk_Checkbox_byIndex - "&"Index - "&i, "Check Box not found on a page"
				else
					LogReport 0,"clk_Checkbox_byIndex - "&"Index - "&i, "Check Box Found on a page and clicked"					
				End If

co(i).click
	
End Function

''Function to click on a single checkbox in a New Window
Function clk_singelCheckbox_newWindow()
Err.Clear
On error resume next
Wait 3

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	
	Set ck = Description.Create
	ck("micclass").value = "WebCheckBox"
	ck("visible").value = true  

Set co =  btn(1).Page("micclass:=Page").ChildObjects(ck)
print co.count

				If co.count = 0 Then
					LogReport 1,"clk_singelCheckbox_newWindow ", "Single Check Box NOT found in a new window"
				else
					LogReport 0,"clk_singelCheckbox_newWindow ", "Single Check Box found in a new window"
				End If

co(0).click

End Function

''Function to click on a Button with specified Name as parameter in a New window
Public Function clk_Button_usingName_NewWindow(strName)
Err.Clear
On error resume next
Wait 3

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	
	Set editO = Description.Create
	editO("micclass").value = "WebButton"
	editO("visible").value = true  
	editO("name").value = strName  

Set editObject = btn(1).Page("micclass:=Page").ChildObjects(editO)
Print editObject.count

				If editObject.count = 0 Then
					LogReport 1,"clk_Button_usingName_NewWindow" & strName,"Button with name " & "-" & strName & "NOT found in a new window, NOT Expected"
				else
					LogReport 0,"clk_Button_usingName_NewWindow" & strName,"Button with name" & "-" & strName & "-" & "is found in a new window"
				End If

editObject(0).click

End Function

Function RandomString( ByVal strLen ) 
Dim str
Const LETTERS = "abcdefghijklmnopqrstuvwxyz0123456789" 
For i = 1 to strLen
str = str & Mid( LETTERS, RandomNumber( 1, Len( LETTERS ) ), 1 )
Next
RandomString = str

End Function

''Function to set Web Edit box or multiple boxes in a new window
Function setWebEditBox_NewWindow(strData)
Err.Clear
On error resume next
Wait 3

	strArr = split(strData,",")
	
	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	
	Set editO = Description.Create
	editO("micclass").value = "WebEdit"
	editO("visible").value = true  
	
	Set edObj =  btn(1).Page("micclass:=Page").ChildObjects(editO) 
	
	for i = 0 To uBound(strArr)
	
	If Not(strArr(i) = "")Then
	
	edObj(i).set strArr(i)        
	End If
	Next
	
				If err <> 0 Then
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"setWebEditBox_NewWindow" & strData,"Failed to set Web Edit box as" & "-" & strData & "in a new window"
				else
					LogReport 0,"setWebEditBox_NewWindow" & strData,"The Web Edit box is set as" & "-" & strData & "in a new window successfully"
				End If	

End Function

'''Function to Clone Project
Function cloneProject()
Err.Clear
On error resume next
Wait 3

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
Err.Clear
On error resume next
Wait 3

Set odesc=Description.Create
odesc("micclass").value="WebTable"

'Set l_link=  getParentObject().ChildObjects(odesc)
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
'print c2
'                                     
'If c = stragree  And c2 = strstat Then
Set oEdit1 = l_link(i).ChildItem(2,1, "Link", 0)
oEdit1.click
'End If

Next
Next                            


Exit for
End if

End If
End If

next
End function

'Modify Hardcoded values each time you create Campaign - NAME change for Past Due
Function create_Campaign(t)
Err.Clear
On error resume next
Wait 3

			strNameEvent = "Campaign_AD_Test" & t & "_" & RandomString(3) 
			strArrN = split(straddinfo,",")
			
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
			If trim(strArr(0)) = "*Campaign Name" Then
			l_link(i).GetROProperty("rows") 
			
			a = l_link(i).GetROProperty("rows")  
			b = l_link(i).GetROProperty("cols") 		  	       
			For x  = 1 to a     
			For j  = 1 to b
			c = l_link(i).GetCellData(x,j)
			print c
			Select Case c
				Case "*Campaign Name"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set strNameEvent
				Case "Cycle"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set TestCycleR1
				Case "Program"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set AD_program
				Case "LOI Form"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set LOI_form
				Case "Application Form"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set APP_form
				Case "PFA Contact Email"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "pcori.reg@gmail.com"
                                  
			End Select                               
			
			Next
			Next                             
				
			Exit for
			
			End If
			End If
			Next
			
create_Campaign = strNameEvent

End Function

Function create_Campaign_Methods()
Err.Clear
On error resume next
Wait 3

			strNameEvent = "Campaign_Methods_Test" & "_" & RandomString(3) 
			strArrN = split(straddinfo,",")
			
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
			If trim(strArr(0)) = "*Campaign Name" Then
			l_link(i).GetROProperty("rows") 
			
			a = l_link(i).GetROProperty("rows")  
			b = l_link(i).GetROProperty("cols") 		  	       
			For x  = 1 to a     
			For j  = 1 to b
			c = l_link(i).GetCellData(x,j)
			print c
			Select Case c
				Case "*Campaign Name"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set strNameEvent
				Case "Cycle"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set TestCycleR1
				Case "Program"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set Program_Methods                       
				Case "LOI Form"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "18C3 - LOI Form - Methods"                                  
				Case "Application Form"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "18C3-App-Methods"                                       
					Case "Application Form"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "18C3-App-Methods"                                          	
					Case "PFA Contact Email"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set "pcori.reg@gmail.com"
       
			End Select                               
			
			Next
			Next                             
				
			Exit for
			
			End If
			End If
			Next
			
create_Campaign_Methods = strNameEvent

End Function


''''Function create_Campaign_CDR ()
''''				On error resume next
''''
''''			strNameEvent = "Campaign_CDR_Test_Automation" & "_" & RandomString(3) 
''''			strArrN = split(straddinfo,",")
''''			
''''			Set odesc=Description.Create
''''			odesc("micclass").value="WebTable"
''''			
''''			Set l_link=  getParentObject().ChildObjects(odesc)
''''			print l_link.count
''''			For i = 0 to l_link.count - 1
''''			strName = l_link(i).GetROProperty("column names")
''''			print strName
''''			strArr = split(strName,";")
''''			If Not(strName = "") Then
''''			print strArr(0)
''''			If trim(strArr(0)) = "*Campaign Name" Then
''''			l_link(i).GetROProperty("rows") 
''''			
''''			a = l_link(i).GetROProperty("rows")  
''''			b = l_link(i).GetROProperty("cols") 		  	       
''''			For x  = 1 to a     
''''			For j  = 1 to b
''''			c = l_link(i).GetCellData(x,j)
''''			print c
''''			Select Case c
''''				Case "*Campaign Name"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set strNameEvent
''''				Case "Cycle"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "Cycle 3 2017"
''''				Case "Program"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "Communication and Dissemination Research"                      
''''				Case "LOI Form"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "17C3 - LOI Form - Broads"                                  
''''				Case "Application Form"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "17C3 - App - Broads"                                       
''''					Case "Application Form"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "17C3 - App - Broads"                                          	
''''			Case "PFA Contact Email"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "pcori.reg@gmail.com"
''''                                  
''''			End Select                               
''''			
''''			Next
''''			Next                             
''''				
''''			Exit for
''''			
''''			End If
''''			End If
''''			Next
''''			
''''			create_Campaign_CDR = strNameEvent
''''End Function

''''Function create_Campaign_CER ()
''''				On error resume next
''''
''''			strNameEvent = "Campaign_CER_Test_Automation" & "_" & RandomString(3) 
''''			strArrN = split(straddinfo,",")
''''			
''''			Set odesc=Description.Create
''''			odesc("micclass").value="WebTable"
''''			
''''			Set l_link=  getParentObject().ChildObjects(odesc)
''''			print l_link.count
''''			For i = 0 to l_link.count - 1
''''			strName = l_link(i).GetROProperty("column names")
''''			print strName
''''			strArr = split(strName,";")
''''			If Not(strName = "") Then
''''			print strArr(0)
''''			If trim(strArr(0)) = "*Campaign Name" Then
''''			l_link(i).GetROProperty("rows") 
''''			
''''			a = l_link(i).GetROProperty("rows")  
''''			b = l_link(i).GetROProperty("cols") 		  	       
''''			For x  = 1 to a     
''''			For j  = 1 to b
''''			c = l_link(i).GetCellData(x,j)
''''			print c
''''			Select Case c
''''				Case "*Campaign Name"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set strNameEvent
''''				Case "Cycle"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "Cycle 3 2017"
''''				Case "Program"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "Clinical Effectiveness and Decision Science"                      
''''				Case "LOI Form"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "17C3 - LOI Form - Broads"                                  
''''				Case "Application Form"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "17C3 - App - Broads"                                       
''''					Case "Application Form"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "17C3 - App - Broads"                                          	
''''			Case "PFA Contact Email"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "pcori.reg@gmail.com"
''''                                  
''''			End Select                               
''''			
''''			Next
''''			Next                             
''''				
''''			Exit for
''''			
''''			End If
''''			End If
''''			Next
''''			
''''			create_Campaign_CER = strNameEvent
''''End Function

''''Function create_Campaign_IHS ()
''''				On error resume next
''''
''''			strNameEvent = "Campaign_IHS_Test_Automation" & "_" & RandomString(3) 
''''			strArrN = split(straddinfo,",")
''''			
''''			Set odesc=Description.Create
''''			odesc("micclass").value="WebTable"
''''			
''''			Set l_link=  getParentObject().ChildObjects(odesc)
''''			print l_link.count
''''			For i = 0 to l_link.count - 1
''''			strName = l_link(i).GetROProperty("column names")
''''			print strName
''''			strArr = split(strName,";")
''''			If Not(strName = "") Then
''''			print strArr(0)
''''			If trim(strArr(0)) = "*Campaign Name" Then
''''			l_link(i).GetROProperty("rows") 
''''			
''''			a = l_link(i).GetROProperty("rows")  
''''			b = l_link(i).GetROProperty("cols") 		  	       
''''			For x  = 1 to a     
''''			For j  = 1 to b
''''			c = l_link(i).GetCellData(x,j)
''''			print c
''''			Select Case c
''''				Case "*Campaign Name"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set strNameEvent
''''				Case "Cycle"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "Cycle 3 2017"
''''				Case "Program"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "Improving Healthcare Systems"                      
''''				Case "LOI Form"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "17C3 - LOI Form - Broads"                                  
''''				Case "Application Form"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "17C3 - App - Broads"                                       
''''					Case "Application Form"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "17C3 - App - Broads"                                          	
''''			Case "PFA Contact Email"
''''			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
''''			oEdit1.set "pcori.reg@gmail.com"
''''                                  
''''			End Select                               
''''			
''''			Next
''''			Next                             
''''				
''''			Exit for
''''			
''''			End If
''''			End If
''''			Next
''''			
''''			create_Campaign_IHS = strNameEvent
''''End Function

''Function to login to Portal as RAPI and navigate to LOIs and Applications Tab
Function RAPI_loginPortal_goToLOIsandApps()
Err.Clear
On error resume next
Wait 3

		Open_ReviewersPortal()
		Wait 3
					''-------------- Read User Name = User Email from Text File saved previously
					If runOn="Methods" Then
						RAPI_email = read_FromFile_anyFilePath (TextfilePathForRAPI_email_M)
					else
						RAPI_email = read_FromFile_anyFilePath (TextfilePathForRAPI_email)
					End If
		

		login_intoSalesForce_Application RAPI_email, passWord1
		Wait 2
		clk_link_Object2 portalLoginbtn
		Wait 4

		''-------------- Verify if User is on the right Page
		verifyIfWebElementDoesExist2 verifyUseronPortal
		Wait 3

''-------------- Navigate to LOI Saved in Draft Status under Home Dashboard 
clk_link_Object2 ResearchAwardsbtn
Wait 1
clk_link_Object2 linkMyLoisandApps
Wait 2

				If err <> 0 Then
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"RAPI_loginPortal_goToLOIsandApps" & RAPI_email,"Failed to login as RAPI" & "-" & RAPI_email
				else
					LogReport 0,"RAPI_loginPortal_goToLOIsandApps" & RAPI_email,"Logged in as RAPI" & "-" & RAPI_email &" and navigated successfully"
				End If	
				
End Function

''Function to login to Portal as RAAO and navigate to LOIs and Applications Tab
Function RAAO_loginPortal_goToLOIsandApps()
Err.Clear
On error resume next
Wait 3

		Open_ReviewersPortal()
		Wait 3
					''-------------- Read User Name = User Email from Text File saved previously
					If runOn="Methods" Then
						RAAO_email = read_FromFile_anyFilePath (TextfilePathForRAAO_email_M)
					else
						RAAO_email = read_FromFile_anyFilePath (TextfilePathForRAAO_email)
					End If
		
		login_intoSalesForce_Application RAAO_email, passWord1
		Wait 2
		clk_link_Object2 portalLoginbtn
		Wait 4

		''-------------- Verify if User is on the right Page
		verifyIfWebElementDoesExist2 verifyUseronPortal
		Wait 3

''-------------- Navigate to LOI Saved in Draft Status under Home Dashboard 
clk_link_Object2 ResearchAwardsbtn
Wait 1
clk_link_Object2 linkMyLoisandApps
Wait 2

				If err <> 0 Then
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"RAAO_loginPortal_goToLOIsandApps" & RAAO_email,"Failed to login as RAAO" & "-" & RAAO_email
				else
					LogReport 0,"RAAO_loginPortal_goToLOIsandApps" & RAAO_email,"Logged in as RAAO" & "-" & RAAO_email &" and navigated successfully"
				End If	
				
End Function

''Function to click on any Link with specified Html ID and Innertext parameters
Function clk_link_Object_HtmlID(strName, HtmlID)
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
	l("micclass").value = "Link"
	l("innertext").value = strName
	l("html id").value = HtmlID

Set lo =  getParentObject().ChildObjects(l)
print lo.count

				If lo.count = 0 Then
					LogReport 1,"clk_link_Object_HtmlID" & strName,"Link was NOT found with innertext" & "-" & strName
				else
					LogReport 0,"clk_link_Object_HtmlID" & strName,"The Link with innertext" & "-" & strName & "-" & "was found succesfully"
				End If

lo(0).click

End Function

''Function to select Single radio button on a page
Function selectRadioButton(strData)
Err.Clear
On error resume next
Wait 3

Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True

Set O =  getParentObject().ChildObjects(odesc)
print O.count

				If O.count = 0 Then
					LogReport 1,"selectRadioButton" & strData,"Radiobutton is NOT found with Value" & "-" & strData
				else
					LogReport 0,"selectRadioButton" & strData,"Radiobutton with value" & "-" & strData &" found successfully"
				End If	
				
O(0).select strData

End Function

''Function to verify if Link with specified Innertext exists
Public Function verifyIfLinkDoesExist1 (strName)
Err.Clear
On error resume next
Wait 3

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
lo(0).Highlight

				If lo.count = 0 Then   	
				LogReport 1,"verifyIfLinkDoesExist1 - "&strName, "The Link --" & strName & "- does not exist _ Not Expected"
				else
				LogReport 0,"verifyIfLinkDoesExist1 - "&strName, "The Link --" & strName & "-Exists - Expected"
				End If

End Function

'' Function to verify that LINK with specified Innertext does NOT exist
Public Function verifyIfLink_NOT_Exist( strName)
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
	l("micclass").value = "Link"
	l("html tag").value = "A"
	l("innertext").value = strName
	
	strcount = l.count
	print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count

				If lo.count > 0 Then      
				LogReport 1,"verifyIfLink_NOT_Exist - "&strName, "The Link --" & strName & "- Exists - NOT Expected"
				else
				LogReport 0,"verifyIfLink_NOT_Exist - "&strName, "The Link --" & strName & "- Doesn't Exist - Expected"
				End If

End Function


Function click_applicationInListView()
Err.Clear
On error resume next
Wait 3

Set odesc=Description.Create
odesc("micclass").value="WebTable"

'Set l_link=  getParentObject().ChildObjects(odesc)
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
Err.Clear
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

Function clickApplicationsInAList(strStatus)
Err.Clear
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

End Function 

'Function clk_MagnifyingClass_StationLookUP()
'
'On error resume next 
'Set btncalc = Description.Create()  
'btncalc("micclass").value = "Browser"
'
'Set btn =DeskTop.ChildObjects(btncalc)
'
'print strHwnd
'
'Set pa = Description.Create
'pa("micclass").value = "Page"
'
'Set l = Description.Create
'l("micclass").value = "Image"
'l("image type").value = "Image Link"
'l("visible").value = true
'
'Set lo =  getParentObject().ChildObjects(l)
'print lo.count
'For i = 0 to lo.count - 1
'name = lo(i).GetRoProperty("alt")
'If name = "Panel Lookup (New Window)" Then
'lo(i).click
'Exit for
'End If
'Next
'End Function

''Function to click on a Link with specified Innertext as parameter in a new window
Function clk_Link_ObjectInApage_NewWindow (strData)
Err.Clear
On error resume next
Wait 3

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	wait 5
	
	Set l = Description.Create
	l("micclass").value = "Link"
	l("visible").value = True
	l("innertext").value = strData
	
	Set lo = btn(1).Page("micclass:=Page").ChildObjects(l)
	print lo.count

lo(0).Highlight

				If lo.count = 0 Then
					LogReport 1,"clk_Link_ObjectInApage_NewWindow "&strData, "The Link with innertext - "&strData & "NOT Found - NOT Expected"
				else
					LogReport 0,"clk_Link_ObjectInApage_NewWindow "&strData, "The Link with innertext - "&strData & "Found - Expected"				
				End If

lo(0).click

End Function

Function LogReport(micPass,strTestStepName,strResultDesc)

Reporter.ReportEvent micPass, strTestStepName, strResultDesc
End Function


Sub click_magnifyinglass(strStatus)
Err.Clear
On error resume next
Wait 3

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
	od("class").value = "fa fa-search"
	od("abs_y").value= y

Set Ob =  getParentObject().ChildObjects(od)
print Ob.count

				If Ob.count = 0 Then
					LogReport 1,"click_magnifyinglass "&"magnifying glass", "magnifying glass" & "NOT Found - NOT Expected"
				else
					LogReport 0,"click_magnifyinglass "&"magnifying glass","magnifying glass" & "Found - Expected"				
				End If

Ob(0).click

End Sub

''Before using this Function - make sure it works as expected
Public Function verifyIfWebElementDoesexit_NewWindow(strName)
Err.Clear
On error resume next
Wait 3

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
					LogReport 0,"verifyIfWebElementDoesexit_NewWindow - "&strName, "Webelement --" & strName & "-  exists - Expected"
				else
					LogReport 1,"verifyIfWebElementDoesexit_NewWindow - "&strName, "Webelement --" & strName & "- doesn't Exist - Not Expected"			
				End If

End Function

''LOI
Function clickReviewRecordBasedOnStatus(strStatus)
On error resume next
	Set odesc=Description.Create
    odesc("micclass").value="WebTable"
    odesc("cols").value= 12  
                     
    Set l_link=  getParentObject().ChildObjects(odesc)
    print l_link.count
                 a = l_link(0).GetROProperty("rows") 
                 b = l_link(0).GetROProperty("cols") 
                           
                        For x  = 1 to a    
                             For j  = 1 to b
                                      c =  l_link(0).GetCellData(x,j)
                                      print c
                                      
                                      If Trim(c) = strStatus Then
                                      	 Set o = l_link(i).ChildItem(x,2, "Link", 0)
                                      	 o.click
                                      	 intText =  o.getRoProperty("innertext")
                                      	 Exit function
                                      End If          
                            next
                       next    
clickReviewRecordBasedOnStatus = intText

End Function 

Function  getProjectNamefromList(strTocompare)
On error resume next
Set editO = Description.Create
editO("micclass").value = "WebList"
editO("visible").value = true  
 
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


''' ---------------------------------------------------------------START: Functions to Create a New User - Modify Name and email if needed

Function fillUserNameInformation (FirstName)  '(strFirstName, strLastName,strEmailID,strpassword)
On error resume next
			   Randomize
         intnumber =  Int((1000000-1)*Rnd+1)
         
          strFirstName = FirstName & "_" & RandomString(4)
          strLastName = "User" & "_" & RandomString(4)
          strEmailID = "test" & "_" & RandomString(2)&"@yopmail.com"
          'strFirstName = "RAPI" & "_" & RandomString(3)
				Set editO = Description.Create
				editO("micclass").value = "WebEdit"
				editO("visible").value = true               
               				
				Set O =  getParentObject().ChildObjects(editO) 
				print O.count
				
				For i = 0 to O.count - 1
				
						strName = O(i).getRoProperty("name")
						
						If trim(strName) = "communitiesSelfRegPage:theForm:firstName" Then
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
     
Function fillAddress()
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
				
				 'First Name
				 O(i).set "1345678986"
				 'Last Name
				 O(i + 1).set "123 main st"
				 'Email id
				 O(i + 2).set "Alexander City"
				 'Email Confirmation
				
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
         
Function searchForTimeBasedWorkflow()
On error resume next
    	' go to the set up
		clk_link_Object2 "Setup"
	
		'click the time based workflow
		clk_link_Object2 "Time-Based Workflow"
	
		'click the search button
		clk_Button_usingName "Search"
End Function
         
''Function to verify if Link with specified Innertext and Html ID as parameters exist - reusable
Public Function verify_IfLinkDoesExist_HtmlID(strName, HtmlID)
Err.Clear
On error resume next
Wait 3

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	
	print strHwnd
	
	Set pa = Description.Create
	pa("micclass").value = "Page"
	
	Set l = Description.Create
	l("micclass").value = "Link"
	l("html tag").value = "A"
	l("html id").value = HtmlID
	l("innertext").value = strName

Set lo =  getParentObject().ChildObjects(l)
lo(0).Highlight

				If lo.count > 0 Then   	
					LogReport 0,"verify_IfLinkDoesExist_HtmlID - "&strName, "The Link --" & strName & "- exist _ Expected"
				else
					LogReport 1,"verify_IfLinkDoesExist_HtmlID - "&strName, "The Link --" & strName & "- NOT Exists - NOT Expected"
				End If

End Function


Public Function searchForReports(strData)    

wait 5
setWebEditBox(strData)

clk_Button_usingName "Search"

End Function   

    
Function NavigateToReviewerPortalnewuser()
Err.Clear
On error resume next
		Open_ReviewersPortal()
		
''		verifyIfWebElementDoesExist2 webElement_onPortal_LogInPage

'''c = webElement_onPortal_LogInPage
'''
'''If (InStr(c,"PCORI Online is now open for:") > 0) Then
'''		wait 5
'''		clickNewuserlink()
'''	else
'''	MsgBox "Text: PCORI Online is now open for: - is NOT Found"
'''	
'''End If

		wait 5
		clickNewuserlink()

End Function
    
Function Clickingbackonbrowser()
On error resume next

   Set oWshShell = CreateObject("WScript.Shell")
   oWshShell.SendKeys "{BS}"  
                     
End Function

Function refresh_Chrome_browser()
On error resume next

   Set oWshShell = CreateObject("WScript.Shell")
   wait 3
   oWshShell.SendKeys "{F5}" 
                     
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


''Function to click on a first checkbox on a Page
Function selectewebCheckbox()
Err.Clear
On error resume next
Wait 2

	Set odesc=Description.Create
	odesc("micclass").value="WebCheckBox"		        
	
	Set l_link=  getParentObject().ChildObjects(odesc)
	print l_link.count
				
	Set oEdit1 = l_link(i).ChildItem("WebCheckBox", 0)
	
				If oEdit1.count > 0 Then   	
				LogReport 0,"selectewebCheckbox - "&"webcheckbox", "Checkbox on a page exists -- Expected"
				else
				LogReport 1,"selectewebCheckbox - "&"webcheckbox", "Checkbox on a page NOT exists -- NOT Expected"
				End If
	
oEdit1.click
											
End Function 

''Function to click on a single specific checkbox when create new Portal user - HARDCODED value inside
Function ClickWebCheckbox()
Err.Clear
On error resume next
Wait 3

Set MyBrowser = Browser("micclass:=Browser").Page("micclass:=Page")
    
'Set oWebckbox = Description.Create
'oWebckbox("micclass").Value = "WebCheckBox"
'oWebckbox("visible").Value = True
'oWebckbox ("name").Value= "communitiesSelfRegPage:theForm:j_id48"
'oWebckbox("type").Value= "checkbox"
'oWebckbox("html tag").Value= "INPUT"
'
'MyBrowser.oWebckbox.click

				If MyBrowser.WebCheckBox("html tag:=INPUT","name:=communitiesSelfRegPage:theForm:j_id48","type:=checkbox", "title:=Acknowledgement required").Exist(3) Then
					LogReport 0,"ClickWebCheckbox - "&"webcheckbox", "Checkbox on a new user page exists -- Expected"
				else
					LogReport 1,"ClickWebCheckbox - "&"webcheckbox", "Checkbox on a new user page NOT exists -- NOT Expected"
				End If

MyBrowser.WebCheckBox("html tag:=INPUT","name:=communitiesSelfRegPage:theForm:j_id48","type:=checkbox", "title:=Acknowledgement required").Click

End Function



''*************************************************************************************************************************************************************************************************************
''*************************************************************************************************************************************************************************************************************
''*************************************************************************************************************************************************************************************************************
''--------------  *****      From this point down are FUNCTIONS Added by Alena starting 1/23/2018     *****   -----------------------------


''------------------------------------------------  ***********   FUNCTIONS to Write to Text Files and READ from Text Files (can be replaced by Excel document in the future    *********** --------------------


''''--------------------------------  *** GENERIC FUNCTIONS to use for any file path - will auto create Text file and write to / read from ***   ----------------------

Function write_ToAfile_anyFilePath_fileName(anyFilePath,strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile(anyFilePath,2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function

Public Function read_FromFile_anyFilePath (anyFilePath)
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile(anyFilePath,1)
          
          sContent = objTxtFile.Readline
         read_FromFile_anyFilePath = sContent
End Function


''-------------------------------------------------------------------------  ******   FUNCTIONS to Write and Read from Text Files are DONE    ******   ------------------------------------------

Function fillOut_CMA_LOIprescreenSection()

	''---CMA Owner Field
dblClick_webElement_RAapp "CF00N39000003LYaJ_ileinner"
Wait 2
setWebEditBox_HtmlID "CMA Admin (Test User)", "CF00N39000003LYaJ"
Wait 2
clk_Image_Link "CMA Owner Lookup \(New Window\)"
Wait 2
clk_Link_ObjectInApage_NewWindow "CMA Admin \(Test User\)"
Wait 3

''---CMA comments
dblClick_webElement_RAapp "00N39000003LYaI_ileinner"
Wait 2
setWebEditBox_HtmlID "CMA Comments - TEST", "00N39000003LYaI"
Wait 2
clk_Button_usingName "OK"
Wait 2

''---Compliance
dblClick_webElement_RAapp "00N39000003LYaH_ileinner"
Wait 2
selectWeblist_PI_Information "Compliant", "00N39000003LYaH"
Wait 2

''clk_Button_usingName editBtn
''setWebEditBox_HtmlID "CMA Admin (Test User)", "CF00N39000003LYaJ" 
''setWebEditBox_HtmlID "CMA COMMENTS - TEST Automation run", "00N39000003LYaI"
''selectWeblist_PI_Information "Compliant", "00N39000003LYaH"

clk_Button_usingName saveBtn
waitForButton editBtn

End Function


Function fillMailingAddress()  '(strPhone, strStreet,strCity,strZip)
     	
     	On error resume next
			   Randomize
         intnumber =  Int((9999-1)*Rnd+1)
         
          strPhone = "525-236-6958"
          strStreet = "Street" & RandomString(5)
          strCity = "City" & RandomString(5)
        '  strState = "Virginia"
          strZip = "20138"
          
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
 
 'Function to select 8 Radio Buttons when create a New User
Function selectRadioButton1()
On error resume next

Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True


Set O =  getParentObject().ChildObjects(odesc)
print O.count

O(0).select "Research Awards"
Wait 2
O(1).select "Mr."
Wait 2
O(2).select "Woman"
Wait 2
O(3).select "Yes"
Wait 2
O(4).select "Asian"
Wait 2
O(5).select "Patient/Consumer"
Wait 2
O(6).select "No"
Wait 2
O(7).select "Opt out"
Wait 2

End Function


'Function to Capture Email Address from User Profile - New User
Public Function captureLinkText()
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

Set lo =  getParentObject().ChildObjects(l)
Print lo.count

emailNewUser = lo(8).getRoProperty("innertext")
Print emailNewUser
wait 2
'Write captured Email Address to Function itself in order to call it later in a script
captureLinkText = emailNewUser

End Function

''Function to capture Link Text based on HtmlID - internally
Public Function capture_LinkText_HtmlID (HtmlID)
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
l("html id").value = HtmlID

Set lo =  getParentObject().ChildObjects(l)
Print lo.count

OwnerName = lo(0).getRoProperty("innertext")
Print OwnerName

capture_LinkText_HtmlID = OwnerName

End Function


'function to set Web Edit when Update User Info
Function setWebEditBoxNewUser(strData)
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
editO("html tag").value = "TEXTAREA" 
editO("kind").value = "multiline" 
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

'Function to Capture RAPI Name from User Profile - New User
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

		Dim arrCorrectName
		arrCorrectName = split(nameNewUser, ". ")
		'Print arrCorrectName(1)
'USE this name later to search for RAPI internally - when verify as Admin - covers TC1-TC2
nameRAPI = arrCorrectName(1)
'Print nameRAPI
'Write captured Email Address to Function itself in order to call it later in a script
captureWebElementText_RAPI_Name = nameRAPI


End Function

'Function to select/change Gender for New User through the Portal
Function selectWeblistGender(strData)
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
editO("html id").value = "00N70000003BwzM" 
'editO("type").value = "text" 

Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
for i = 0 To uBound(strArr)
items = edObj(i).GetRoProperty("all items")
print items
print strArr(i)


edObj(i).select strData       

Next

End Function

'''------------------------------------------------------------- Functions to Create Campaign Timeframe current and past: START

''Function to Select Time Frame - TABLE2: Modify dates for Test Needs manually here
Function create_CampaignTimeframe_Current()
On error resume next
	
			'DATEs in sequence - CURRENT
			k = Date - 10
			n = Date + 30
			s = Date + 20
			t = Date + 22
			v = Date - 1
			
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
			If trim(strArr(0)) = "Start Date" Then
			l_link(i).GetROProperty("rows") 
			
			a = l_link(i).GetROProperty("rows")  
			b = l_link(i).GetROProperty("cols") 		  	       
			For x  = 1 to a     
			For j  = 1 to b
			c = l_link(i).GetCellData(x,j)
			print c
			Select Case c
			Case "Start Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set k
			Case "End Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set n
			Case "LOI Submission Deadline Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set s
			Case "Board Meeting Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set t
			Case "Application Start Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set v			
			                                 
			End Select                               
			
			Next
			Next                             
				
			Exit for
			
			End If
			End If
			Next

End Function

Function create_CampaignTimeframe_PAST()
On error resume next
	
			'DATEs in sequence - PAST DUE
			k = Date - 45
			n = Date -25 
			s = Date -27 
			t = Date -12
			v = Date - 16
			
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
			If trim(strArr(0)) = "Start Date" Then
			l_link(i).GetROProperty("rows") 
			
			a = l_link(i).GetROProperty("rows")  
			b = l_link(i).GetROProperty("cols") 		  	       
			For x  = 1 to a     
			For j  = 1 to b
			c = l_link(i).GetCellData(x,j)
			print c
			Select Case c
			Case "Start Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set k
			Case "End Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set n
			Case "LOI Submission Deadline Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set s
			Case "Board Meeting Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set t
			Case "Application Start Date"
			Set oEdit1 = l_link(i).ChildItem(x,j + 1, "WebEdit", 0)
			oEdit1.set v			
			                                 
			End Select                               
			
			Next
			Next                             
				
			Exit for
			
			End If
			End If
			Next

End Function
'''------------------------------------------------------------- Functions to Create Campaign Timeframe current and past: END


''Fill out Contact Information tab on Quiz - Portal
Function setWebEditBox_ContactInfo_Portal(RAPI, RAAO, PID1, PID2, FCofficer, Org, Distr, Dept)
On error resume next

wait 3
strArr = split(RAPI,",")

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
'editO("html id").value = "j_id0:mainForm:j_id360:5:inputFieldId" 
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 

for i = 0 To uBound(strArr)
'id =   edObj(i).GetRoProperty("html id")
items = edObj(i).GetRoProperty("all items")
print items

print RAPI(i)
If Not(RAPI(i) = "")Then

'edObj(i).set strArr(i) 
edObj(i).set RAPI
edObj(i+1).set "Test Dual PI Name"
edObj(i+2).set "dualpitest@nomail.com"
edObj(i+3).set RAAO
edObj(i+4).set PID1
edObj(i+5).set PID2
edObj(i+6).set FCofficer
edObj(i+7).set Org
edObj(i+8).set Distr
edObj(i+9).set Dept

oShell.SendKeys strSearchText				            
End If

Next
If err <> 0 Then
err.clear   

End If
End Function


'Find if LOI link appears on All list views
Public Function FindAndClickLinkInList_LOI (LinkInnertext)
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")

Err.Clear
On Error Resume Next   
                'Capture the page no value and convert it into Integer
PageStr = Brwser.WebElement("class:=right", "html tag:=SPAN").GetRoProperty("innertext")
PageStr = Split(PageStr, "Pageof")
PageNo = Cint (PageStr(1))

'Create the description Link object
                
Link2 = Split(LinkInnertext, " ")
Print Link2()
Link3 = Link2(1) & ", " & Link2(0)
Print Link3

		Set objDesc = Description.Create
objDesc("micclass").value = "WebElement"
objDesc("visible").value = True
objDesc("html tag").value = "SPAN"
objDesc("innertext").value = Link3

'Search the link each until the link appears
                For i = 1 To PageNo
                
                If Brwser.WebElement(objDesc).Exist(5) Then
                                         
                                click_webElement_InListLOI Link3
                                Exit Function
                Else
                                clk_Image_Link "Next"
                End If
                Next
                
                If err <> 0 Then
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"Finding the WebElement object" & Link3,"Failed to click the WebElement" & "-" & Link3
				else
					LogReport 0,"cliking the WebElement object" & Link3,"The WebElement" & "-" & Link3 & "-" & "is clicked succesfully"
				End If

Set objDesc = Nothing

End Function


''Function to Find if Project link appears on All list view in Projects Tab and click on it to Open
Public Function FindAndClickLinkInList_Project (LinkInnertext)
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")

Err.Clear
On Error Resume Next   
                'Capture the page no value and convert it into Integer
PageStr = Brwser.WebElement("class:=right", "html tag:=SPAN").GetRoProperty("innertext")
PageStr = Split(PageStr, "Pageof")
PageNo = Cint (PageStr(1))

''Create the description Link object
 
	Set objDesc = Description.Create
objDesc("micclass").value = "WebElement"
objDesc("visible").value = True
objDesc("html tag").value = "SPAN"
objDesc("innertext").value = LinkInnertext

''Search the link each until the link appears
                For i = 1 To PageNo
                
                If Brwser.WebElement(objDesc).Exist(5) Then
                
                                Brwser.WebElement(objDesc).highlight 
                                
                                click_webElement_InListLOI LinkInnertext
                                Exit Function
                                
                Else
                                Print "WebElement = Project Name - NOT Exist"
                                clk_Image_Link "Next"
                End If
                Next
                
                If err <> 0 Then
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"Finding the WebElement object" & LinkInnertext,"Failed to click the WebElement" & "-" & LinkInnertext
				else
					LogReport 0,"cliking the WebElement object" & LinkInnertext,"The WebElement" & "-" & LinkInnertext & "-" & "is clicked succesfully"
				End If

Set objDesc = Nothing

End Function

Public Function click_webElement_InListLOI (strData)
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
l("html tag").value = "SPAN"
l("innertext").value = strData

Set edObj =  getParentObject().ChildObjects(l)
edObj(i).click

End Function

''Enter text in Web Edit box of newly opened window
Public Function setWebEditBox_NewOpenWindow (StrData)
Err.Clear
On Error Resume Next
Wait 3

Dim Brwser
    Set Brwser = Browser("micclass:=browser","creationtime:=1","Title:=.*").Page("micclass:=page","creationtime:=1","Title:=.*")
    
    Set oDesc = Description.Create
    oDesc("micclass").value = "WebEdit"
    oDesc("html tag").value = "INPUT" 
    oDesc("name").value = "lksrch"
    
    Brwser.WebEdit(oDesc).Highlight
    Brwser.WebEdit(oDesc).Set StrData
    
				If err > 0 Then
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"setWebEditBox_NewOpenWindow " & StrData,"Web Edit was NOT found in a new window" & "-" & StrData
				else
					LogReport 0,"setWebEditBox_NewOpenWindow " & StrData,"Web Edit " & "-" & StrData & "-" & "Found successfully in a new window"
				End If
	
End Function

''Click on web button with specified Name as parameter in a new open window
Public Function clk_Button_usingName_NewOpenWindow (StrName)
Err.Clear
On Error Resume Next
Wait 3

Dim Brwser
    Set Brwser = Browser("micclass:=browser","creationtime:=1","Title:=.*").Page("micclass:=page","creationtime:=1","Title:=.*")
    
    Set oDesc = Description.Create
    
    oDesc("micclass").value = "WebButton"
    oDesc("html tag").value = "INPUT" 
    oDesc("name").value = StrName
    
    Brwser.WebButton(oDesc).Highlight
    Brwser.WebButton(oDesc).Click
    
				If err > 0 Then
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"clk_Button_usingName_NewOpenWindow " & StrName,"Failed to Click on the Button " & "-" & StrName
				else
					LogReport 0,"clk_Button_usingName_NewOpenWindow " & StrName,"Button with name" & "-" & StrName & "-" & "Clicked successfully"
				End If

End Function

''Function to click on WebElement = Paper/Pencil Icon on the Portal
Public Function click_webElementPortal_Paper_Pencil_Icon ()
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
l("class").value = htmlIDpaperPencilIconPortal
l("outerhtml").value = "<td class=""slds-truncate sorting_1"" style=""color: black; white-space: initial;""><a onclick=""selectedAction\('EDIT'\)""><img src=""https://www\.foundationconnect\.org/staticresources/SLDS/icons/utility/edit_form_60\.png""></a></td>"
'l("visible").value = False
l("html tag").value = "TD"

Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click

End Function


'Function to click on WebElement = Paper/Pencil Icon on the Portal
Public Function click_webElementPortal_Paper_Pencil_Icon1 ()
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
l("class").value = htmlIDpaperPencilIconPortal
l("outerhtml").value = "<i class=""fa fa-pencil-square-o""></i>"
l("visible").value = False
l("html tag").value = "I"



Set edObj =  getParentObject().ChildObjects(l)
edObj(i+1).Highlight
edObj(i+1).click

End Function

'Function to click on WebElement = Paper/Pencil Icon on the Portal
Public Function click_webElementPortal_RAAO_Icon (i)
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
			l("class").value = "fa fa-file-text-o"
			l("outerhtml").value = "<i class=""fa fa-file-text-o""></i>"
			''l("visible").value = False
			l("html tag").value = "I"
			
			Set edObj =  getParentObject().ChildObjects(l)
			edObj(i).Highlight
			edObj(i).click

End Function

'''Function to click SORT button on Portal LOI - Open Items to sort LOIs from the Highest # to Lowest #
Public Function click_webElement_sort_LOI_Portal ()
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
l("class").value = "DataTables_sort_icon css_right ui-icon ui-icon-triangle-1-s"
l("visible").value = True
l("html tag").value = "SPAN"

Set edObj =  getParentObject().ChildObjects(l)
edObj(0).Highlight
edObj(0).click

End Function

''Function to Capture Property Value of WebElement = created LOI Number on Portal in LOIs Tab
Public Function capture_webElement_onPortalLOIsTab_byColumnNumber(j)
Err.Clear
On error resume next 
Wait 3

	Set odesc=Description.Create
    odesc("micclass").value="WebTable"
    odesc("cols").value= 10
    odesc("column names").value = "Edit;View;-;LOI Number;PFA;LOI Amount Requested from PCORI*;PI/Project Lead Name;Awardee Institution/Organization;Project Name*;External Status"
                     
    Set l_link=  getParentObject().ChildObjects(odesc)
    print l_link.count
	
i=0
         a = l_link(i).GetROProperty("rows") 
         Print "rows on LOI tab - "&a
         b = l_link(i).GetROProperty("cols") 
         Print "columns on LOI tab - "&b
                 
					For x  = 2 to a  	
				
                       c =  l_link(i).GetCellData(x,j)
                       print "captured value on LOIs tab is - "&c

					Next

capture_webElement_onPortalLOIsTab_byColumnNumber = c

End Function

''Generic function to Capture WebElement Value
Public Function capture_webElement_value (HtmlID, HtmlTag)
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
			l("html id").value = HtmlID
			l("html tag").value = HtmlTag
			
			Set edObj =  getParentObject().ChildObjects(l)
			Captured_V = edObj(0).GetRoProperty("innertext")
			Print Captured_V
			capture_webElement_value = Captured_V

End Function

''Function to Capture WebElement value by CLASS
Public Function capture_webElement_value_byTag(HtmlTag)
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
			'l("class").value = wclass
			l("visible").value = True
			l("html tag").value = HtmlTag
			
			Set edObj =  getParentObject().ChildObjects(l)
			Captured_V = edObj(0).GetRoProperty("innertext")
			Print Captured_V
			capture_webElement_value_byClass = Captured_V

End Function

Public Function click_webElement_byClass (WClass, HtmlTag)
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
			l("class").value = WClass
			l("visible").value = True
			l("html tag").value = HtmlTag
			
			Set edObj =  getParentObject().ChildObjects(l)
			edObj(0).Highight
			edObj(0).click
			Wait 2
End Function

''Internally select LOI web list with certain Status
Function selectWeblist_LOI(strData)

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
                      editO("select type").value = "ComboBox Select" 
			
			Set edObj =  getParentObject().ChildObjects(editO) 
			print edObj.count
			
			for i = 0 To uBound(strArr)
			items = edObj(i).GetRoProperty("all items")
			print items
			print strArr(i)
			     
			If Not(strArr(i) = "")Then    
			
			edObj(0).select trim(strArr(i))        
			End If
			
			Next

End Function

'' Function to select certain web list value for Applications 
Function selectWeblist_APP(strData)
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
editO("html id").value = "selecttab" 
editO("html tag").value = "SELECT" 

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

End Function

''Internally - function to click on Blue round internally at the top right
Public Function click_webElement_Settinggear()
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
'''l("class").value = "menuButtonButton"
l("class").value = "headerTrigger  tooltip-trigger uiTooltip"
l("outertext").value = "Setup"
'''l("html id").value = "tsidButton"
l("innertext").value = "Setup"
l("visible").value = True
l("html tag").value = "DIV"

Set edObj =  getParentObject().ChildObjects(l)
edObj(0).click

			If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"click_webElement_Settinggear - " & "Setup","Failed to click the webElement" & "-" & "Setup"
			else
			  	LogReport 0,"click_webElement_Settinggear - " & "Setup","The webElement" & "-" & "Setup" & "-" & "is clicked successfully"
			End If

End Function


''Function to Click WebElement
Public Function click_webElement_top_email_personnel (Innertext, HtmlTag)
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
 
 Public Function click_webElement_byHtmlID_innertext (Innertext, HtmlID)
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
			l("html id").value = HtmlID
			l("innertext").value = Innertext
			
			Set edObj =  getParentObject().ChildObjects(l)
			edObj(0).Highlight
			edObj(0).click

			If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"click_webElement_byHtmlID_innertext - " & Innertext,"Failed to click the webElement" & "-" & Innertext
			else
			  	LogReport 0,"click_webElement_byHtmlID_innertext - " & Innertext,"The webElement" & "-" & Innertext & "-" & "is clicked successfully"
			End If
			
End Function
 
Public Function clk_Image_Link (AltImage)
Err.Clear
On error resume next

			wait 2
			Set btncalc = Description.Create()  
			btncalc("micclass").value = "Browser"
			
			Set btn =DeskTop.ChildObjects(btncalc)
			strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
			print strHwnd
			
			Set pa = Description.Create
			pa("micclass").value = "Page"
			
			Set l = Description.Create
			l("micclass").value = "Image"
			l("alt").value = AltImage
			l("image type").value = "Image Link"
			
			Set lo =  getParentObject().ChildObjects(l)
			
			lo(0).highlight
			lo(0).click
			
			If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"clk_Image_Link - " & AltImage,"Failed to click the link" & "-" & AltImage
			else
			  	LogReport 0,"clk_Image_Link - " & AltImage,"The Link" & "-" & AltImage & "-" & "is clicked successfully"
			End If
End Function 
 
 
''Function to verify that Checkbox is Checked
Public Function verify_checkbox_ISchecked (AltImage, HtmlID)
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
l("alt").value = AltImage
l("image type").value = "Plain Image"
l("html tag").value = "IMG"
l("html id").value = HtmlID

Set lo =  getParentObject().ChildObjects(l)

lo(0).highlight

If lo.count = 0 Then
	LogReport 1,"verify_checkbox_ISchecked" & AltImage,"Checkbox " & AltImage & " - is NOT checked - NOT EXPECTED"
else
  	LogReport 0,"verify_checkbox_ISchecked" & AltImage,"Checkbox " & AltImage & " - is checked - EXPECTED"
End If

End Function


''Mohammed - Find the Project link appears on All list views
Public Function FindAndClickLinkInList (LinkInnertext)
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")

Err.Clear
On Error Resume Next   
                'Capture the page number value and convert it into Integer
PageStr = Brwser.WebElement("class:=right", "html tag:=SPAN").GetRoProperty("innertext")
PageStr=Split(PageStr, "Pageof")
PageNo = Cint (PageStr(1))

'Create the description Link object
                Set oDesc = Description.Create
oDesc("micclass").value = "Link"
oDesc("html tag").value = "A" 
oDesc("innertext").value = LinkInnertext

'Search the link on each page until the link appears
                For i = 1 To PageNo
                
                If Brwser.Link(oDesc).Exist(5) Then
                                clk_link_Object2 LinkInnertext
                Exit Function
                Else
                                clk_Image_Link "Next"
                End If
                Next
                
                If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"FindAndClickLinkInList - " & LinkInnertext,"Failed to click the link" & "-" & LinkInnertext
				else
				LogReport 0,"FindAndClickLinkInList - " & LinkInnertext,"The link" & "-" & LinkInnertext & "-" & "is clicked succesfully"
				End If

Set oDesc = Nothing

End Function


'Function to verify Status when Admin Click LOI and Opens it
Function Verify_LOI_Status_Internally (HtmlID, Status)
Err.Clear
On Error Resume Next  

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
'print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


'Create description Link object
                Set stDesc = Description.Create
				stDesc("micclass").value = "WebElement"
				stDesc("html tag").value = "DIV" 
				stDesc("html id").value = HtmlID
				'stDesc("innertext").value = Status
				stDesc("visible").value = True

Set edObj =  getParentObject().ChildObjects(stDesc)
LOI_Status = edObj(i).GetRoProperty("innertext")
Print LOI_Status

		If LOI_Status = Status Then
                 Print "PASS - Status is correct - "&LOI_Status
                 LogReport 0,"Verify_LOI_Status_Internally - " & LOI_Status, "LOI Status is correct" & "-" &LOI_Status& " - Expected"
        Else
                 Print = "FAIL - Status is not correct - "&LOI_Status
                 LogReport 1,"Verify_LOI_Status_Internally - " & LOI_Status,"LOI Status is NOT correct" & "-" &LOI_Status& " - NOT Expected"
        End If
		
			
        If err <> 0 Then
		LogReport 4,"Error", err.number & "-" & err.description
	    	   	    
		End If
          
	
	Set stDesc = Nothing
	Set edObj = Nothing
	
End Function

''Mohammed - Internal Error msg validation Function
Function VerifyErrorMessage(StrErrMsg)
Err.Clear
On Error Resume Next
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
                Set oDesc = Description.Create
oDesc("micclass").value = "WebElement"
oDesc("html tag").value = "DIV" 
oDesc("html id").value = "errorDiv_ep"
oDesc("Visible").value = "True"

                Brwser.WebElement(oDesc).highlight
                ErrMsg = Brwser.WebElement(oDesc).GetRoProperty("innertext")
                
                CapErrmsg = Split(ErrMsg,".")
                
                
                
                for i = 0 To uBound(CapErrmsg)
                'Print "Capture ---" & CapErrmsg(i)

                If StrErrMsg = Trim (CapErrmsg(i)) Then
                                Print "Capture and validation of Error Message ----- " & StrErrMsg & "-------- done successfully"
                                Exit Function
                'Else
                                'Print "Capture and validation of Error Message------- " & StrErrMsg & "-------- Failed"
                End If
Next
                
                If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"Capture and validate Error Message " & StrErrMsg,"Validation of error message " & "-" & StrErrMsg & "-" & " Failed"
                Else
                LogReport 0,"Capture and validate Error Message " & StrErrMsg,"Validation of error message " & "-" & StrErrMsg & "-" & " done successfully"
Set oDesc = Nothing
End If
End Function

'Function to Select from WebLists on LOI Form Portal - Resubmission Tab
Function selectWeblist_Resubmission(strData)
Err.Clear
On error resume next
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

edObj(i+1).select trim(strArr(i))   
edObj(i+2).select trim(strArr(i+1)) 
edObj(i+3).select trim(strArr(i+2)) 

End If

Next

End Function


''GOOD ONE - Very specific - Function verifies on LOI Form Resubmission Tab if User Select NO for the first question - the following questions are not visible
Public Function verifyIfWebListDoes_NOT_Exist(strName)
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
l("micclass").value = "WebList"
l("html tag").value = "SELECT"
l("innertext").value = "--None--YesNo"
l("visible").value = True 
l("selected item index").value = strName

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count

If lo.count = 0 Then      
LogReport 0,"verifyIfWebListDoes_NOT_Exist - "& strName, "The WebElement --" & strName & "- Doesn't Exist - Expected"
else
LogReport 1,"verifyIfWebListDoes_NOT_Exist - "& strName, "The WebElement --" & strName & "- Exists - NOT Expected"

End If

End Function

''Function to verify that WebList does not exist on LOI Pre Screen section
Public Function verify_WebListDoes_NOT_Exist (HtmlID)
Err.Clear
On error resume next

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn = DeskTop.ChildObjects(btncalc)
strHwnd = btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebList"
l("html tag").value = "SELECT"
l("visible").value = True 
l("html id").value = HtmlID

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count

If lo.count = 0 Then
LogReport 0,"verify_WebListDoes_NOT_Exist - " & HtmlID, "The List --" & HtmlID & "- does NOT exist - Expected"
else
LogReport 1,"verify_WebListDoes_NOT_Exist - " & HtmlID, "The List --" & HtmlID & "-Exists - NOT Expected"
End If
End Function


''Function to verify that WebEdit does not exist on LOI Pre Screen section
Public Function verify_WebEditDoes_NOT_Exist (HtmlID)
Err.Clear
On error resume next

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn = DeskTop.ChildObjects(btncalc)
strHwnd = btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebEdit"
l("visible").value = True 
l("html id").value = HtmlID

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count

If lo.count = 0 Then
LogReport 0,"verify_WebEditDoes_NOT_Exist - " & HtmlID, "The Edit --" & HtmlID & "- does NOT exist - Expected"
else
LogReport 1,"verify_WebEditDoes_NOT_Exist - " & HtmlID, "The Edit --" & HtmlID & "-Exists - NOT Expected"
End If

End Function


'GOOD ONE - Function verifies on LOI Form Resubmission Tab if User Select NO for the first question - the following questions are not visible
Public Function verifyIfWebList_Resub_Exist1(strData)
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
l("micclass").value = "WebList"
l("innertext").value = "--None--"
l("all items").value = "--None--"
l("visible").value = True 
l("html id").value = strData

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count

If lo.count = 0 Then      
LogReport 1,"verifyIfWebList_Resub_Exist1 - "&strData, "WebList --" & strData & "- Doesn't Exist - NOT Expected"
else
LogReport 0,"verifyIfWebList_Resub_Exist1 - "&strData, "WebList --" & strData & "- Exists - Expected"

End If

End Function

'Function to Select from WebList on LOI Form - PI Information Tab
Function selectWeblist_PI_Information(strData, HtmlID)
Err.Clear
On error resume next

strArr = split(strData,";")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")

Set pa = Description.Create
pa("micclass").value = "Page"

Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
'editO("visible").value = true  
editO("html id").value = HtmlID
 

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
	LogReport 1,"selectWeblist_PI_Information - " & HtmlID,"Failed to select weblist value " & strData & " with html ID" & "-" & HtmlID
else
  	LogReport 0,"selectWeblist_PI_Information - " & HtmlID,"The weblist value "&strData&" with html ID " & "-" & HtmlID & "-" & "selected successfully"
End If

End Function

''function to select from weblist based on all items
Function selectWeblist_PI_Information_allItems(strData, allItems)
Err.Clear
On error resume next
strArr = split(strData,";")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")

Set pa = Description.Create
pa("micclass").value = "Page"

Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("all items").value = allItems

Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count

'''for i = 0 To uBound(strArr)
'''items = edObj(i).GetRoProperty("all items")
'''print items
'''print strArr(i)
'''If Not(strArr(i) = "")Then	

edObj(0).select strData'''trim(strArr(i))   
 
'''End If
'''
'''Next

If err <> 0 Then
	LogReport 4,"Error", err.number & "-" & err.description
	LogReport 1,"selectWeblist_PI_Information_allItems - " & allItems,"Failed to select weblist value " & strData & " from all items" & "-" & allItems
else
  	LogReport 0,"selectWeblist_PI_Information_allItems - " & allItems,"The weblist value "&strData&" from all items " & "-" & allItems & "-" & "selected successfully"
End If

End Function



''function to select from weblist based on index
Function selectWeblist_PI_Information_index(strData, index)
Err.Clear
On error resume next

strArr = split(strData,";")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")

Set pa = Description.Create
pa("micclass").value = "Page"

Set editObject =  getParentObject()

Set editO = Description.Create

editO("micclass").value = "WebList"
editO("html tag").value = "SELECT" 


Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count
	
edObj(index).select strData  

If err <> 0 Then
	LogReport 4,"Error", err.number & "-" & err.description
	LogReport 1,"selectWeblist_PI_Information_index - " & index,"Failed to select weblist value " & strData & " with index" & "-" & index
else
  	LogReport 0,"selectWeblist_PI_Information_index - " & index,"The weblist value "&strData&" with index " & "-" & index & "-" & "selected successfully"
End If

End Function

'Function to CLick Arrow Link Image to move Pick List values from Left to right - PI-Information Tab - Portal
Public Function clk_Image_Link_Portal (HtmlID)
Err.Clear
On error resume next

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

lo(0).highlight
lo(0).click

If err <> 0 Then
	LogReport 4,"Error", err.number & "-" & err.description
	LogReport 1,"clk_Image_Link_Portal - " & HtmlID,"Failed to click the link" & "-" & HtmlID
else
  	LogReport 0,"clk_Image_Link_Portal - " & HtmlID,"The Link" & "-" & HtmlID & "-" & "is clicked successfully"
End If

End Function 

Public Function clk_Image_Link_Portal_BYindex (index)
Err.Clear
On error resume next

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
l("html tag").value = "IMG"
l("class").value = "picklistArrowRight"

Set lo =  getParentObject().ChildObjects(l)

lo(index).highlight
'lo(index).click
lo(index).fireEvent("onClick")
If err <> 0 Then
	LogReport 4,"Error", err.number & "-" & err.description
	LogReport 1,"clk_Image_Link_Portal_BYindex - " & index,"Failed to click the link with index" & "-" & index
else
  	LogReport 0,"clk_Image_Link_Portal_BYindex - " & index,"The Link with index" & "-" & index & "-" & "is clicked successfully"
End If

End Function


''function to click Yes on Key Personnel - Portal - Can make it more reusable by combining with LOI one
Public Function click_webElement_KeyPersonnel ()
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
l("html tag").value = "LABEL"
l("innertext").value = "Yes"

Set edObj =  getParentObject().ChildObjects(l)
edObj(i).click

End Function

'Function to Attach file from Computer on Portal
Function attachFileFromFileSystem_Portal(strFilePath)
Err.Clear
On error resume next
                  Set oShell = CreateObject("WScript.Shell")
       
                  
					Set btncalc = Description.Create()  
                    btncalc("micclass").value = "Browser"
                  
                    Set btn =DeskTop.ChildObjects(btncalc)
                    strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
                    print strHwnd

                     strSearchText = strFilePath
                     oShell.SendKeys strSearchText 
                     wait 2
                     ''''''''''''''''Browser("hwnd:="&strHwnd).Dialog("text:=Choose File to Upload").WinButton("text:=&Open").click
                     Browser("hwnd:="&strHwnd).Dialog("text:=Open").WinButton("text:=&Open").click
                     
If err <> 0 Then
	LogReport 4,"Error", err.number & "-" & err.description
	LogReport 1,"Failed to attach file"
else
  	LogReport 0,"attached file successfully"
End If

End Function

Function attach_FileFromFileSystem_works(strFilePath)
Err.Clear
On error resume next
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

        Set edObj =  getParentObject().ChildObjects(editO) 
         
        edObj(0).set strFilePath


End Function

'function to attach File on Notes and Attachments - Internally
Public Function Attach_File_internally(filePath)
Err.Clear
On error resume next
'Scroll down to 'Notes & Attachments' related list and select 'Attach File' button

clk_Button_usingName "Attach File"
waitForButton  "Attach File"

'Click 'Browse...' button
					'''''Browser("micclass:=Browser").Page("micclass:=Page").WebTable("column names:=1\.;Select the File", "html tag:=TABLE", "name:=file").WebFile("html tag:=INPUT", "html id:=file", "type:=file").Click
					'''''
					'''''attachFileFromFileSystem_Portal filePath
					'''''wait 2

attach_FileFromFileSystem_works filePath

'Click "Attach File" button to attach.
clk_Button_usingName "Attach File" 
waitForButton  "Attach File"

'Click "Done" button.
clk_Button_usingName  " Done "

  End Function

''Function to verify all the fiedls and sections on LOI Detail Page = PFA Methods
Function verify_LOI_Methods_PageLayout()

		verifyIfButtonDoesExist editBtn
		verifyIfButtonDoesExist deleteBtn

		verifyIfButtonDoesExist sharing_btn 
		verifyIfButtonDoesExist generatePDF_RAloiMethods
		
verifyIfWebElementDoesExist2 "LOI Detail"

		verifyIfWebElementDoesExist2 "Pre Screen Questionnaire:"
verify_IfWebElementDoesExist_byOutertext "Decision Aid"
verify_IfWebElementDoesExist_byOutertext "New Intervention"
verifyIfWebElementDoesExist2 "Does it involve a foreign organization"
verify_IfWebElementDoesExist_byOutertext "Practice Guidelines"
verify_IfWebElementDoesExist_byOutertext "Cost Effective Analysis"
verifyIfWebElementDoesExist_class_Innertext "last labelCol", "Project Personnel"


		verifyIfWebElementDoesExist2 "Contacts:"
verifyIfWebElementDoesExist2 "Principal Investigator"
verifyIfWebElementDoesExist2 "Administrative Official"
verifyIfWebElementDoesExist2 "Dual PI Name"
verifyIfWebElementDoesExist2 "Dual PI Email"
verifyIfWebElementDoesExist2 "PI Designee 1"
verifyIfWebElementDoesExist2 "PI Designee 2"
verifyIfWebElementDoesExist2 "Financial Officer"

		verifyIfWebElementDoesExist2 "PI Information:"
verify_IfWebElementDoesExist_byOutertext "Primary Identification"
verifyIfWebElementDoesExist2 "Primary Group Identification\*"
verify_IfWebElementDoesExist_byOutertext "Previous involvement with PCORI\*"
verifyIfWebElementDoesExist2 "Previous involvement with PCORI - Other"
verifyIfWebElementDoesExist2 "PI Work Telephone"
verify_IfWebElementDoesExist_byOutertext "Position Title"
verifyIfWebElementDoesExist2 "Project Lead Degree\*"
verifyIfWebElementDoesExist2 "Project Lead Degree - Other"
verify_IfWebElementDoesExist_byOutertext "Relevant Exp after Terminal Degree"
verifyIfWebElementDoesExist2 "How many years of relevant experience\?"

verify_IfWebElementDoesExist_byOutertext "Field Relevant Exp"
verify_IfWebElementDoesExist_byOutertext "Grants Funded as PI"
verify_IfWebElementDoesExist_byOutertext "Largest Previous Grant/Contract Fund"
verify_IfWebElementDoesExist_byOutertext "Previous Grants/Contracts"
verify_IfWebElementDoesExist_byOutertext "Other \(please specify\)"

		verifyIfWebElementDoesExist2 "Organization Information:"
verifyIfWebElementDoesExist2 "Awardee Institution/Organization"
verifyIfWebElementDoesExist2 "D-U-N-S Number"
verifyIfWebElementDoesExist2 "Location/satellite"
verifyIfWebElementDoesExist2 "Department"
verifyIfWebElementDoesExist2 "Congressional District"
verifyIfWebElementDoesExist2 "Congressional District\*"
verifyIfWebElementDoesExist2 "Street Address"
verifyIfWebElementDoesExist2 "City"
verifyIfWebElementDoesExist2 "State/Province"
verifyIfWebElementDoesExist2 "Zip Code"
verifyIfWebElementDoesExist2 "Country"

		verifyIfWebElementDoesExist2 "Project Information:"
'''SPS-4416		
''verifyIfWebElementDoesExist2 "PFA"
verifyIfWebElementDoesExist2 "Primary Campaign Source"
verifyIfWebElementDoesExist2 "Cycle"
verifyIfWebElementDoesExist2 "Program"
verifyIfWebElementDoesExist2 "PFA Type"
verify_IfWebElementDoesExist_byOutertext "IHS Award Size"
verifyIfWebElementDoesExist2 "Project Name\*"
verify_IfWebElementDoesExist_byOutertext "LOI Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Invited"
verify_IfWebElementDoesExist_byOutertext "App Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Bypass Review"
verify_IfWebElementDoesExist_byOutertext "Previous Application"

verifyIfWebElementDoesExist2 "Total Direct Costs"
verifyIfWebElementDoesExist2 "Total Indirect Costs"
verifyIfWebElementDoesExist2 "LOI Amount Requested from PCORI\*"
verify_IfWebElementDoesExist_byOutertext "Project Duration"
verifyIfWebElementDoesExist2 "Key Word One"
verifyIfWebElementDoesExist2 "Key Word Two"
verifyIfWebElementDoesExist2 "Key Word Three"
verifyIfWebElementDoesExist2 "Key Word Four"
verifyIfWebElementDoesExist2 "Key Word Five"
verifyIfWebElementDoesExist2 "Key Word Six"

		verifyIfWebElementDoesExist2 "Project Focus:"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Rare Disease Focus"
verify_IfWebElementDoesExist_byOutertext "Population Focus"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Focus"
verify_IfWebElementDoesExist_byOutertext "Study Design"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Population Focus Other"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Other"
verify_IfWebElementDoesExist_byOutertext "Other Study Design"
	
verifyIfWebElementDoesExist2 "Primary Patient Partner\(s\)"
verifyIfWebElementDoesExist2 "Primary Stakeholder Partner\(s\)"

		verifyIfWebElementDoesExist2 "Methods Information:"

verify_IfWebElementDoesExist_byOutertext "Methods Research Areas"
verify_IfWebElementDoesExist_byOutertext "Methods Used One"
verify_IfWebElementDoesExist_byOutertext "Methods Used Two"
verify_IfWebElementDoesExist_byOutertext "Methods Used Three"
verify_IfWebElementDoesExist_byOutertext "Methods Used Four"
verify_IfWebElementDoesExist_byOutertext "Methods Used Five"
verify_IfWebElementDoesExist_byOutertext "Methods Used Six"
verify_IfWebElementDoesExist_byOutertext "Methods Research Areas Other"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced One"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced Two"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced Three"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced Four"


		verifyIfWebElementDoesExist2 "LOI Pre Screen \(For Internal Use\):"
verifyIfWebElementDoesExist2 "CMA Owner"
verifyIfWebElementDoesExist2 "CMA Comments"
verifyIfWebElementDoesExist2 "Administrative Compliance Flag"
verifyIfWebElementDoesExist2 "Program Owner"
verifyIfWebElementDoesExist2 "Program Response"
verifyIfWebElementDoesExist2 "Program Non Responsiveness Decision"
verifyIfWebElementDoesExist2 "MRO Owner"
verifyIfWebElementDoesExist2 "MRO Response"
verifyIfWebElementDoesExist2 "MRO Non Responsiveness Flag"

		verifyIfWebElementDoesExist2 "Final LOI Decision \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Recommended for Alternate PFA"
verifyIfWebElementDoesExist2 "Proposed PFA"
verify_IfWebElementDoesExist_byOutertext "Alternate PFA Rationale"
verifyIfWebElementDoesExist2 "Lead Reviewer"
verifyIfWebElementDoesExist2 "Reviews Status"
verifyIfWebElementDoesExist2 "Draft Comments for Public Reply"
verifyIfWebElementDoesExist2 "Consolidated Internal Comments"
verify_IfWebElementDoesExist_byOutertext "Final Decision Status"
verifyIfWebElementDoesExist2 "Final Comments Screened"
verifyIfWebElementDoesExist2 "Final Comments for Public Reply"
verifyIfWebElementDoesExist2 "Email Proof"
verifyIfWebElementDoesExist2 "Preview Email Communications"


		verifyIfWebElementDoesExist2 "Status \(For Internal Use\):"
verifyIfWebElementDoesExist2 "LOI Number"
verifyIfWebElementDoesExist2 "LOI Status"
verifyIfWebElementDoesExist2 "External Status"
verifyIfWebElementDoesExist2 "Name"
verifyIfWebElementDoesExist2 "LOI Owner"
verifyIfWebElementDoesExist2 "LOI Review1"
verifyIfWebElementDoesExist2 "Review Roll-up 1"
verifyIfWebElementDoesExist2 "LOI Review2"
verifyIfWebElementDoesExist2 "Review Roll-up 2"
verifyIfWebElementDoesExist2 "LOI Review3"
verifyIfWebElementDoesExist2 "Review Roll-up 3"
verifyIfWebElementDoesExist2 "LOI Review4"
verifyIfWebElementDoesExist2 "Review Roll-up 4"
verifyIfWebElementDoesExist2 "Project Lead Email Address\*"
verifyIfWebElementDoesExist2 "Legal name of organization\*"
verifyIfWebElementDoesExist2 "LOI Submission Deadline Date"
verifyIfWebElementDoesExist2 "Custom Links"
verifyIfWebElementDoesExist2 "Created By"
verify_IfWebElementDoesExist_byOutertext "Submitted By"
verify_IfWebElementDoesExist_byOutertext "Submission Date"
verifyIfWebElementDoesExist2 "Last Modified By"
verifyIfWebElementDoesExist2 "LOI Record Type"
verifyIfWebElementDoesExist2 "Application Start Date"


	
End Function

''Function to verify bottom part of LOI Deail Page - Methods - as CMA USer
Function verify_LOI_Methods_PageLayout__bottomPart_asCMA()

	verifyIfButtonDoesExist Convertbtn
	
		verifyIfWebElementDoesExist_HtmlId "00Q0x000001HTuP_00N39000003LYa3_title"
verifyIfButtonDoesExist "New Project Personnel"
		verifyIfWebElementDoesExist_HtmlId "00Q0x000001HTuP_00N39000003LYZU_title"
		verify_IfWebElementDoesExist_byOutertext "Question Attachments"
verifyIfButtonDoesExist "New Question Attachment"
		verify_IfWebElementDoesExist_byOutertext "Notes & Attachments"
verifyIfButtonDoesExist "New Note"
verifyIfButtonDoesExist "Attach File"
verifyIfButtonDoesExist "View All"
		verify_IfWebElementDoesExist_byOutertext "Activity History"
verifyIfButtonDoesExist "Log a Call"
verifyIfButtonDoesExist "Mail Merge"
verifyIfButtonDoesExist "Send an Email"
		verify_IfWebElementDoesExist_byOutertext "Campaign History"
verifyIfButtonDoesExist "Add to Campaign"
		verify_IfWebElementDoesExist_byOutertext "Lead History"
	
		
End Function


''Function to verify bottom part of LOI Deail Page - Methods - as Science PO USer
Function verify_LOI_Methods_PageLayout__bottomPart_asSciencePO()

	
		verifyIfWebElementDoesExist_HtmlId "00Q0x000001HTuP_00N39000003LYa3_title"
verifyIfButtonDoesExist "New Project Personnel"
		verifyIfWebElementDoesExist_HtmlId "00Q0x000001HTuP_00N39000003LYZU_title"
verifyIfButtonDoesExist "New LOI Review"
		verify_IfWebElementDoesExist_byOutertext "Question Attachments"
		verify_IfWebElementDoesExist_byOutertext "Notes & Attachments"
verifyIfButtonDoesExist "New Note"
verifyIfButtonDoesExist "Attach File"
verifyIfButtonDoesExist "View All"
		verify_IfWebElementDoesExist_byOutertext "Activity History"
verifyIfButtonDoesExist "Log a Call"
verifyIfButtonDoesExist "Mail Merge"
verifyIfButtonDoesExist "Send an Email"
		verify_IfWebElementDoesExist_byOutertext "Campaign History"
verifyIfButtonDoesExist "Add to Campaign"
		verify_IfWebElementDoesExist_byOutertext "Lead History"
	
		
End Function

Function REM_Verify_Project_Status_Internally (Status)
Err.Clear
On Error Resume Next  

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
'print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


'Create description Link object
                Set stDesc = Description.Create
				stDesc("micclass").value = "WebElement"
				stDesc("html tag").value = "DIV" 
				stDesc("html id").value = "opp11_ileinner"
				'stDesc("innertext").value = Status
				stDesc("visible").value = True

Set edObj =  getParentObject().ChildObjects(stDesc)
Project_Status = edObj(i).GetRoProperty("innertext")
Print Project_Status

		If Project_Status = Status Then
                 Print "PASS - Project Status is correct = "&Status
        Else
                 Print = "FAIL - Project Status is not correct = "&Project_Status
        End If
		
			
        If err <> 0 Then
		LogReport 4,"Error", err.number & "-" & err.description
	    LogReport 1,"REM_Verify_Project_Status_Internally - " & Project_Status,"is NOT correct" & "-" & "NOT Expected"
	    Else 
	    LogReport 0,"REM_Verify_Project_Status_Internally - " & Project_Status, "is correct" & "-" & "Expected"
		End If
          
	
	Set stDesc = Nothing
	Set edObj = Nothing
	
End Function



''Function to verify LOI Deail Page Layout - Methods - as MRO USer
Function verify_LOI_Methods_PageLayout_asMRO()

		verifyIfButtonDoesExist editBtn
		verifyIfButtonDoesExist deleteBtn
		verifyIfButtonDoesExist Convertbtn
		verifyIfButtonDoesExist sharing_btn 
		verifyIfButtonDoesExist generatePDF_RAloiMethods
		
verifyIfWebElementDoesExist2 "LOI Detail"

		verifyIfWebElementDoesExist2 "Pre Screen Questionnaire:"
verify_IfWebElementDoesExist_byOutertext "Decision Aid"
verify_IfWebElementDoesExist_byOutertext "New Intervention"
verify_IfWebElementDoesExist_byOutertext "PCORnet involvement"
verifyIfWebElementDoesExist2 "Does it involve a foreign organization"
verify_IfWebElementDoesExist_byOutertext "Practice Guidelines"
verify_IfWebElementDoesExist_byOutertext "Cost Effective Analysis"
verifyIfWebElementDoesExist_class_Innertext "labelCol", "Project Personnel"


		verifyIfWebElementDoesExist2 "Contacts:"
verifyIfWebElementDoesExist2 "Principal Investigator"
verifyIfWebElementDoesExist2 "Administrative Official"
verifyIfWebElementDoesExist2 "Dual PI Name"
verifyIfWebElementDoesExist2 "Dual PI Email"
verifyIfWebElementDoesExist2 "PI Designee 1"
verifyIfWebElementDoesExist2 "PI Designee 2"
verifyIfWebElementDoesExist2 "Financial Officer"

		verifyIfWebElementDoesExist2 "PI Information:"
verify_IfWebElementDoesExist_byOutertext "Primary Identification"
verifyIfWebElementDoesExist2 "Primary Group Identification\*"
verify_IfWebElementDoesExist_byOutertext "Previous involvement with PCORI\*"
verifyIfWebElementDoesExist2 "Previous involvement with PCORI - Other"
verifyIfWebElementDoesExist2 "PI Work Telephone"
verify_IfWebElementDoesExist_byOutertext "Position Title"
verifyIfWebElementDoesExist2 "Project Lead Degree\*"
verifyIfWebElementDoesExist2 "Project Lead Degree - Other"
verify_IfWebElementDoesExist_byOutertext "Relevant Exp after Terminal Degree"
verifyIfWebElementDoesExist2 "How many years of relevant experience\?"

verify_IfWebElementDoesExist_byOutertext "Field Relevant Exp"
verify_IfWebElementDoesExist_byOutertext "Grants Funded as PI"
verify_IfWebElementDoesExist_byOutertext "Largest Previous Grant/Contract Fund"
verify_IfWebElementDoesExist_byOutertext "Previous Grants/Contracts"
verify_IfWebElementDoesExist_byOutertext "Other \(please specify\)"

		verifyIfWebElementDoesExist2 "Organization Information:"
verifyIfWebElementDoesExist2 "Awardee Institution/Organization"
verifyIfWebElementDoesExist2 "D-U-N-S Number"
verifyIfWebElementDoesExist2 "Location/satellite"
verifyIfWebElementDoesExist2 "Department"
verifyIfWebElementDoesExist2 "Congressional District"
verifyIfWebElementDoesExist2 "Street Address"
verifyIfWebElementDoesExist2 "City"
verifyIfWebElementDoesExist2 "State/Province"
verifyIfWebElementDoesExist2 "Zip Code"
verifyIfWebElementDoesExist2 "Country"

		verifyIfWebElementDoesExist2 "Project Information:"
'''SPS-4416		
''verifyIfWebElementDoesExist2 "PFA"
verifyIfWebElementDoesExist2 "Primary Campaign Source"
verifyIfWebElementDoesExist2 "Cycle"
verifyIfWebElementDoesExist2 "Program"
verifyIfWebElementDoesExist2 "PFA Type"
verify_IfWebElementDoesExist_byOutertext "IHS Award Size"
verifyIfWebElementDoesExist2 "Project Name\*"
verify_IfWebElementDoesExist_byOutertext "LOI Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Invited"
verify_IfWebElementDoesExist_byOutertext "App Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Bypass Review"
verify_IfWebElementDoesExist_byOutertext "Previous Application"

verifyIfWebElementDoesExist2 "Total Direct Costs"
verifyIfWebElementDoesExist2 "Total Indirect Costs"
verifyIfWebElementDoesExist2 "LOI Amount Requested from PCORI\*"
verify_IfWebElementDoesExist_byOutertext "Project Duration"
verifyIfWebElementDoesExist2 "Key Word One"
verifyIfWebElementDoesExist2 "Key Word Two"
verifyIfWebElementDoesExist2 "Key Word Three"
verifyIfWebElementDoesExist2 "Key Word Four"
verifyIfWebElementDoesExist2 "Key Word Five"
verifyIfWebElementDoesExist2 "Key Word Six"

		verifyIfWebElementDoesExist2 "Project Focus:"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Rare Disease Focus"
verify_IfWebElementDoesExist_byOutertext "Population Focus"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Focus"
verify_IfWebElementDoesExist_byOutertext "Study Design"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Population Focus Other"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Other"
verify_IfWebElementDoesExist_byOutertext "Other Study Design"
	
verifyIfWebElementDoesExist2 "Primary Patient Partner\(s\)"
verifyIfWebElementDoesExist2 "Primary Stakeholder Partner\(s\)"

		verifyIfWebElementDoesExist2 "Methods Information:"

verify_IfWebElementDoesExist_byOutertext "Methods Research Areas"
verify_IfWebElementDoesExist_byOutertext "Methods Used One"
verify_IfWebElementDoesExist_byOutertext "Methods Used Two"
verify_IfWebElementDoesExist_byOutertext "Methods Used Three"
verify_IfWebElementDoesExist_byOutertext "Methods Used Four"
verify_IfWebElementDoesExist_byOutertext "Methods Used Five"
verify_IfWebElementDoesExist_byOutertext "Methods Used Six"
verify_IfWebElementDoesExist_byOutertext "Methods Research Areas Other"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced One"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced Two"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced Three"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced Four"


		verifyIfWebElementDoesExist2 "LOI Pre Screen \(For Internal Use\):"
verifyIfWebElementDoesExist2 "CMA Owner"
verifyIfWebElementDoesExist2 "CMA Comments"
verifyIfWebElementDoesExist2 "Administrative Compliance Flag"
verifyIfWebElementDoesExist2 "Program Owner"
verifyIfWebElementDoesExist2 "Program Response"
verifyIfWebElementDoesExist2 "Program Non Responsiveness Decision"
verifyIfWebElementDoesExist2 "MRO Owner"
verifyIfWebElementDoesExist2 "MRO Response"
verifyIfWebElementDoesExist2 "MRO Non Responsiveness Flag"

		verifyIfWebElementDoesExist2 "Final LOI Decision \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Recommended for Alternate PFA"
verifyIfWebElementDoesExist2 "Proposed PFA"
verify_IfWebElementDoesExist_byOutertext "Alternate PFA Rationale"
verifyIfWebElementDoesExist2 "Lead Reviewer"
verifyIfWebElementDoesExist2 "Reviews Status"
verifyIfWebElementDoesExist2 "Draft Comments for Public Reply"
verifyIfWebElementDoesExist2 "Consolidated Internal Comments"
verify_IfWebElementDoesExist_byOutertext "Final Decision Status"
verifyIfWebElementDoesExist2 "Final Comments Screened"
verifyIfWebElementDoesExist2 "Final Comments for Public Reply"
verifyIfWebElementDoesExist2 "Email Proof"
verifyIfWebElementDoesExist2 "Preview Email Communications"


		verifyIfWebElementDoesExist2 "Status \(For Internal Use\):"
verifyIfWebElementDoesExist2 "LOI Number"
verifyIfWebElementDoesExist2 "LOI Status"
verifyIfWebElementDoesExist2 "External Status"
verifyIfWebElementDoesExist2 "Name"
verifyIfWebElementDoesExist2 "LOI Owner"
verifyIfWebElementDoesExist2 "LOI Review1"
verifyIfWebElementDoesExist2 "Review Roll-up 1"
verifyIfWebElementDoesExist2 "LOI Review2"
verifyIfWebElementDoesExist2 "Review Roll-up 2"
verifyIfWebElementDoesExist2 "LOI Review3"
verifyIfWebElementDoesExist2 "Review Roll-up 3"
verifyIfWebElementDoesExist2 "LOI Review4"
verifyIfWebElementDoesExist2 "Review Roll-up 4"
verifyIfWebElementDoesExist2 "Project Lead Email Address\*"
verifyIfWebElementDoesExist2 "Legal name of organization\*"
verifyIfWebElementDoesExist2 "LOI Submission Deadline Date"
verifyIfWebElementDoesExist_class_Innertext "last labelCol", "Custom Links"
verifyIfWebElementDoesExist2 "Created By"
verify_IfWebElementDoesExist_byOutertext "Submitted By"
verify_IfWebElementDoesExist_byOutertext "Submission Date"
verifyIfWebElementDoesExist2 "Last Modified By"
verifyIfWebElementDoesExist2 "LOI Record Type"
verifyIfWebElementDoesExist2 "Application Start Date"

	
		verifyIfWebElementDoesExist_HtmlId "00Q0x000001HTuP_00N39000003LYa3_title"
verifyIfButtonDoesExist "New Project Personnel"
		verifyIfWebElementDoesExist_HtmlId "00Q0x000001HTuP_00N39000003LYZU_title"
verifyIfButtonDoesExist "New LOI Review"
		verify_IfWebElementDoesExist_byOutertext "Question Attachments"
verifyIfButtonDoesExist "New Question Attachment"
		verify_IfWebElementDoesExist_byOutertext "Notes & Attachments"
verifyIfButtonDoesExist "New Note"
verifyIfButtonDoesExist "Attach File"
verifyIfButtonDoesExist "View All"
		verify_IfWebElementDoesExist_byOutertext "Activity History"
verifyIfButtonDoesExist "Log a Call"
verifyIfButtonDoesExist "Mail Merge"
verifyIfButtonDoesExist "Send an Email"
		verify_IfWebElementDoesExist_byOutertext "Campaign History"
verifyIfButtonDoesExist "Add to Campaign"
		verify_IfWebElementDoesExist_byOutertext "Lead History"
	
		
End Function


''Function to verify Application = Methods Page Layout by CMA
Function verify_App_Methods_PageLayout()

verifyIfButtonDoesExist editBtn
verifyIfButtonDoesExist "Print Application"
verifyIfButtonDoesExist "Generate Online Critique"
verifyIfButtonDoesExist "Generate Summary Statement"

		verifyIfWebElementDoesExist2 project_Detail
		verifyIfWebElementDoesExist2 "LOI Pre Screen Questionnaire:"

verify_IfWebElementDoesExist_byOutertext "Decision Aid"
verify_IfWebElementDoesExist_byOutertext "New Intervention"
verifyIfWebElementDoesExist2 "Partial funding Agency Name"
verifyIfWebElementDoesExist2 "Network Type"
verifyIfWebElementDoesExist2 "Total Cumulative Actual Enrolled"
verifyIfWebElementDoesExist2 "Does it involve a foreign organization"
verify_IfWebElementDoesExist_byOutertext "Practice Guidelines"
verify_IfWebElementDoesExist_byOutertext "Cost Effective Analysis"


		verifyIfWebElementDoesExist2 "Authorized Users:"
verify_IfWebElementDoesExist_byOutertext "PI/Project Lead 1 Name"
verifyIfWebElementDoesExist2 "Administrative Official Name"
verifyIfWebElementDoesExist2 "PI/Project Lead Designee 1 Name"
verifyIfWebElementDoesExist2 "PI/Project Lead Designee 2 Name"
verifyIfWebElementDoesExist2 "Financial Contact"


		verifyIfWebElementDoesExist2 "PI Information:"
verify_IfWebElementDoesExist_byOutertext "Primary Identification"
verify_IfWebElementDoesExist_byOutertext "Previous involvement with PCORI\*"
verifyIfWebElementDoesExist2 "Previous involvement with PCORI - Other"
verifyIfWebElementDoesExist2 "PI Work Telephone"
verify_IfWebElementDoesExist_byOutertext "Position Title"
verifyIfWebElementDoesExist2 "Project Lead Degree\*"
verifyIfWebElementDoesExist2 "Project Lead Degree - Other"
verify_IfWebElementDoesExist_byOutertext "Relevant Exp after Terminal Degree"
verify_IfWebElementDoesExist_byOutertext "Field Relevant Exp"
verifyIfWebElementDoesExist2 "How many years of relevant experience\?"
verify_IfWebElementDoesExist_byOutertext "Grants Funded as PI"
verify_IfWebElementDoesExist_byOutertext "Largest Previous Grant/Contract Fund"
verify_IfWebElementDoesExist_byOutertext "Previous Grants/Contracts"
verify_IfWebElementDoesExist_byOutertext "Other \(please specify\)"

		verifyIfWebElementDoesExist2 "Organization Information:"
verifyIfWebElementDoesExist2 "Awardee Institution/Organization"
verifyIfWebElementDoesExist2 "D-U-N-S Number"
verifyIfWebElementDoesExist2 "Location/satellite"
verifyIfWebElementDoesExist2 "Department"
verifyIfWebElementDoesExist2 "Congressional District"
verifyIfWebElementDoesExist2 "Street Address"
verifyIfWebElementDoesExist2 "City"
verifyIfWebElementDoesExist2 "State/Province"
verifyIfWebElementDoesExist2 "Zip Code"
verifyIfWebElementDoesExist2 "Country"

		verifyIfWebElementDoesExist2 "Project Information:"
verifyIfWebElementDoesExist2 "Primary Campaign Source"
verifyIfWebElementDoesExist2 "Cycle"

verifyIfWebElementDoesExist2 "PFA Type"
verifyIfWebElementDoesExist2 "Priority Area"
verify_IfWebElementDoesExist_byOutertext "IHS Award Size"
verifyIfWebElementDoesExist2 "Short Project Title"
verifyIfWebElementDoesExist2 "Project Name\*-off layout"
verifyIfWebElementDoesExist2 "Full Project Title"
verify_IfWebElementDoesExist_byOutertext "LOI Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Invited"
verify_IfWebElementDoesExist_byOutertext "App Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Bypass Review"
verify_IfWebElementDoesExist_byOutertext "Previous Application"
verifyIfWebElementDoesExist2 "Total Direct Costs"
verifyIfWebElementDoesExist2 "Total Indirect Costs"

verifyIfWebElementDoesExist2 "LOI Amount Requested from PCORI\*"
verify_IfWebElementDoesExist_byOutertext "Project Duration"
verify_IfWebElementDoesExist_byOutertext "PCORnet involvement"
verifyIfWebElementDoesExist2 "Key Word One"
verifyIfWebElementDoesExist2 "Key Word Two"
verifyIfWebElementDoesExist2 "Key Word Three"
verifyIfWebElementDoesExist2 "Key Word Four"
verifyIfWebElementDoesExist2 "Key Word Five"
verifyIfWebElementDoesExist2 "Key Word Six"

		verifyIfWebElementDoesExist2 "Project Focus:"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Rare Disease Focus"
verify_IfWebElementDoesExist_byOutertext "Population Focus"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Focus"
verify_IfWebElementDoesExist_byOutertext "Study Design SC"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Population Focus Other"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Other"
verify_IfWebElementDoesExist_byOutertext "Other Study Design"
	

		verifyIfWebElementDoesExist2 "Methods Information:"
verify_IfWebElementDoesExist_byOutertext "Methods Research Areas"
verify_IfWebElementDoesExist_byOutertext "Methods Used One"
verify_IfWebElementDoesExist_byOutertext "Methods Used Two"
verify_IfWebElementDoesExist_byOutertext "Methods Used Three"
verify_IfWebElementDoesExist_byOutertext "Methods Used Four"
verify_IfWebElementDoesExist_byOutertext "Methods Used Five"
verify_IfWebElementDoesExist_byOutertext "Methods Used Six"
verify_IfWebElementDoesExist_byOutertext "Methods Research Areas Other"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced One"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced Two"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced Three"
verify_IfWebElementDoesExist_byOutertext "Methods Advanced Four"

		verifyIfWebElementDoesExist2 "Project Narratives:"
verifyIfWebElementDoesExist2 "Contract Start Date\*"
verifyIfWebElementDoesExist2 "Close Date"
verifyIfWebElementDoesExist2 "Technical Abstract"
verifyIfWebElementDoesExist2 "Public Abstract"
verifyIfWebElementDoesExist2 "Project Narrative"
verify_IfWebElementDoesExist_byOutertext "Study Comparators"

verify_IfWebElementDoesExist_byOutertext "Engagement Plan"
verify_IfWebElementDoesExist_byOutertext "Trial Arms"
verify_IfWebElementDoesExist_byOutertext "Trial Length"
verify_IfWebElementDoesExist_byOutertext "Primary Outcome"
verify_IfWebElementDoesExist_byOutertext "Secondary Outcome"
verify_IfWebElementDoesExist_byOutertext "Recruitment/Retention Plan"
verify_IfWebElementDoesExist_byOutertext "Sample Size"
verify_IfWebElementDoesExist_byOutertext "Health Systems Factors"
verify_IfWebElementDoesExist_byOutertext "Health Systems Factors Other"
verify_IfWebElementDoesExist_byOutertext "Data Sources"
verify_IfWebElementDoesExist_byOutertext "Data Sources Other"

		verifyIfWebElementDoesExist2 "PCORI Staff \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Program Officer"
verifyIfWebElementDoesExist2 "Program Associate"
verifyIfWebElementDoesExist2 "Contract Administrator"
verifyIfWebElementDoesExist2 "Contract Coordinator"

		verifyIfWebElementDoesExist2 "Application Responsiveness CMA \(For Internal Use\):"
verifyIfWebElementDoesExist2 "CMA Administratively Compliant"
verify_IfWebElementDoesExist_byOutertext "Scrub Project Personnel Contact"
verifyIfWebElementDoesExist2 "CMA Feedback"
verifyIfWebElementDoesExist2 "CMA Budget Feedback"

		verifyIfWebElementDoesExist2 "Application Responsiveness Program \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Program Decision"
verifyIfWebElementDoesExist2 "Ready for Merit Review"
verifyIfWebElementDoesExist2 "Program Response to CMA"
verifyIfWebElementDoesExist2 "Program Response to MRO"
verifyIfWebElementDoesExist2 "Science Program Justification"

		verifyIfWebElementDoesExist2 "Merit Review Responsiveness Screen \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Cost Effectiveness Analysis"
verifyIfWebElementDoesExist2 "Changed Aims/Comparators"
verifyIfWebElementDoesExist2 "Guidelines Proposed"
verifyIfWebElementDoesExist2 "Responsiveness to LOI Feedback"
verifyIfWebElementDoesExist2 "Move Forward to Online Review"
verifyIfWebElementDoesExist2 "Cost Effectiveness Comments"
verifyIfWebElementDoesExist2 "Aims/Comparators Comments"
verifyIfWebElementDoesExist2 "Guidelines Comments"
verifyIfWebElementDoesExist2 "Responsiveness Comments"
verifyIfWebElementDoesExist2 "Move Forward Comments"

		verifyIfWebElementDoesExist2 "PIR \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Flag for PIR"

		verifyIfWebElementDoesExist2 "Merit Review - Online Review \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Reviewer 1"
verifyIfWebElementDoesExist2 "Reviewer 2"
verifyIfWebElementDoesExist2 "Reviewer 3"
verifyIfWebElementDoesExist2 "Reviewer 4"
verifyIfWebElementDoesExist2 "Reviewer 5"
verifyIfWebElementDoesExist2 "Panel"
verifyIfWebElementDoesExist2 "Panel Finalized\?"
verifyIfWebElementDoesExist2 "Online Review Record Type"
verifyIfWebElementDoesExist2 "Online Review Deadline"
verifyIfWebElementDoesExist2 "Create Online Reviews"

		verifyIfWebElementDoesExist2 "Merit Review Summary \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Add to Discussion Line"
verifyIfWebElementDoesExist2 "Discussion Line Comments"
verify_IfWebElementDoesExist_byOutertext "Discussion Order Ranking"
verifyIfWebElementDoesExist2 "Application Download - MR"
verifyIfWebElementDoesExist2 "Combined Online Critique Download"
verifyIfWebElementDoesExist2 "Summary Statement Download"
verifyIfWebElementDoesExist2 "In-Person Discussion Notes"
verifyIfWebElementDoesExist2 "All Online Reviews Completed\?"
verifyIfWebElementDoesExist2 "Average Online Review Score"
verifyIfWebElementDoesExist2 "Average In-Person Score"
verifyIfWebElementDoesExist2 "Quartile Included in Summary Statement"
verifyIfWebElementDoesExist2 "Quartile"
verify_IfWebElementDoesExist_byOutertext "Re-Review"

		verifyIfWebElementDoesExist2 "Funding Slate:"
verifyIfWebElementDoesExist2 "Scores are Ready"
verifyIfWebElementDoesExist2 "Funding Slate"
verifyIfWebElementDoesExist2 "Amount Proposed to Board"
verifyIfWebElementDoesExist2 "Funding Slate Communications"
verifyIfWebElementDoesExist2 "Funding Slate Stage"
verifyIfWebElementDoesExist2 "Funding Slate Change Rationale"
verifyIfWebElementDoesExist2 "Exception"
verifyIfWebElementDoesExist2 "Exception Criteria"
verifyIfWebElementDoesExist2 "Exception Rationale"
verifyIfWebElementDoesExist2 "Exception Action"
verifyIfWebElementDoesExist2 "Selection Committee Notes"

		verifyIfWebElementDoesExist2 "System Information \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Status"
verifyIfWebElementDoesExist2 "External Status"
verifyIfWebElementDoesExist2 "Application Number"
verifyIfWebElementDoesExist2 "Project Record Type"
verify_IfWebElementDoesExist_byOutertext "LOI Record Type"
verifyIfWebElementDoesExist2 "LOI"
verifyIfWebElementDoesExist2 "PI Approval"
verifyIfWebElementDoesExist2 "Project Owner"
verifyIfWebElementDoesExist2 "Created By"
verifyIfWebElementDoesExist2 "Last Modified By"
verifyIfWebElementDoesExist2 "Application Submission Deadline Date"
verifyIfWebElementDoesExist2 "AO Approval"

		verifyIfWebElementDoesExist2 "Custom Links:"
		verifyIfWebElementDoesExist2 "Applicant Attachments:"


	
End Function

''Function to verify RA Application Methods Page Layout by CMA USer
Function verify_RAapp_Methods_pageLayout_asCMA()
	
	verify_App_Methods_PageLayout()
verifyIfWebElementDoesExist2 "Owning Program"
verifyIfWebElementDoesExist2 "Application Amount"
verifyIfWebElementDoesExist2 "Goals"

verify_IfWebElementDoesExist_byOutertext "Budgets"
verify_IfWebElementDoesExist_byOutertext "COI & Expertise"
verify_IfWebElementDoesExist_byOutertext "Milestones - Deliverables"
	verifyIfButtonDoesExist "New Milestone - Deliverable"
verifyIfWebElementDoesExist_HtmlId "0060x0000056UP8_00N70000003CCov_title"
	verifyIfButtonDoesExist "New Project Personnel"
verify_IfWebElementDoesExist_byOutertext "Open Activities"
	verifyIfButtonDoesExist "New Task"
	verifyIfButtonDoesExist "New Event"
verify_IfWebElementDoesExist_byOutertext "External Reviews"
verify_IfWebElementDoesExist_byOutertext "PIRs"
	verifyIfButtonDoesExist "New PIR"
verify_IfWebElementDoesExist_byOutertext "Question Attachments"
	verifyIfButtonDoesExist "New Question Attachment"
verify_IfWebElementDoesExist_byOutertext "Files"
	verifyIfButtonDoesExist "Upload Files"
verify_IfWebElementDoesExist_byOutertext "Notes & Attachments"
	verifyIfButtonDoesExist "New Note"
	verifyIfButtonDoesExist "Attach File"
	verifyIfButtonDoesExist "View All"
verify_IfWebElementDoesExist_byOutertext "Activity History"
verify_IfWebElementDoesExist_byOutertext "Project Team"
verify_IfWebElementDoesExist_byOutertext "Project Field History"
verify_IfWebElementDoesExist_byOutertext "Audits"
	verifyIfButtonDoesExist "New Audit"
verify_IfWebElementDoesExist_byOutertext "MR Appeals Inquiries"
verify_IfWebElementDoesExist_byOutertext "CMA Application Reviews"
verifyIfButtonDoesExist "New CMA Application Review"



	
End Function


''Function to verofy RA App Methods Page Layout by Science PO user
Function verify_RAapp_Methods_pageLayout_asSciencePO()
	
	verify_App_Methods_PageLayout()
verifyIfWebElementDoesExist2 "Owning Program"
verifyIfWebElementDoesExist2 "Application Amount"
verifyIfWebElementDoesExist2 "Goals"
	
	
verify_IfWebElementDoesExist_byOutertext "Budgets"

verify_IfWebElementDoesExist_byOutertext "Milestones - Deliverables"
	verifyIfButtonDoesExist "New Milestone - Deliverable"
	
verifyIfWebElementDoesExist_HtmlId "0060x0000056UP8_00N70000003CCov_title"
	verifyIfButtonDoesExist "New Project Personnel"
	
verify_IfWebElementDoesExist_byOutertext "Open Activities"
	verifyIfButtonDoesExist "New Task"
	verifyIfButtonDoesExist "New Event"
	
verify_IfWebElementDoesExist_byOutertext "External Reviews"
	verifyIfButtonDoesExist "New Review"

verify_IfWebElementDoesExist_byOutertext "PIRs"
	verifyIfButtonDoesExist "New PIR"

verify_IfWebElementDoesExist_byOutertext "Question Attachments"
	
verify_IfWebElementDoesExist_byOutertext "Files"
	verifyIfButtonDoesExist "Upload Files"
	
verify_IfWebElementDoesExist_byOutertext "Notes & Attachments"
	verifyIfButtonDoesExist "New Note"
	verifyIfButtonDoesExist "Attach File"
	verifyIfButtonDoesExist "View All"
	
verify_IfWebElementDoesExist_byOutertext "Activity History"

verify_IfWebElementDoesExist_byOutertext "Project Team"

verify_IfWebElementDoesExist_byOutertext "Project Field History"

verify_IfWebElementDoesExist_byOutertext "Audits"
	verifyIfButtonDoesExist "New Audit"
	
verify_IfWebElementDoesExist_byOutertext "MR Appeals Inquiries"
verify_IfWebElementDoesExist_byOutertext "CMA Application Reviews"


	
End Function


''Function to verify Application Methods Page Layout by MRO User
Function verify_RAapp_Methods_pageLayout_asMRO()

verify_App_Methods_PageLayout()
verifyIfWebElementDoesExist2 "Previous Fluxx Application"

verify_IfWebElementDoesExist_byOutertext "Budgets"
		verifyIfButtonDoesExist "New Budget"
		
verify_IfWebElementDoesExist_byOutertext "COI & Expertise"
verifyIfButtonDoesExist "New COI & Expertise"

verify_IfWebElementDoesExist_byOutertext "Milestones - Deliverables"
	verifyIfButtonDoesExist "New Milestone - Deliverable"
	
verifyIfWebElementDoesExist_HtmlId "0060x0000056UP8_00N70000003CCov_title"
	verifyIfButtonDoesExist "New Project Personnel"
	
verify_IfWebElementDoesExist_byOutertext "Open Activities"
	verifyIfButtonDoesExist "New Task"
	verifyIfButtonDoesExist "New Event"
	
verify_IfWebElementDoesExist_byOutertext "External Reviews"
verifyIfButtonDoesExist "New Review"

verify_IfWebElementDoesExist_byOutertext "PIRs"
		verifyIfButtonDoesExist "New PIR"
	
verify_IfWebElementDoesExist_byOutertext "Question Attachments"
		verifyIfButtonDoesExist "New Question Attachment"
	
verify_IfWebElementDoesExist_byOutertext "Files"
		verifyIfButtonDoesExist "Upload Files"
	
verify_IfWebElementDoesExist_byOutertext "Notes & Attachments"
		verifyIfButtonDoesExist "New Note"
		verifyIfButtonDoesExist "Attach File"
		verifyIfButtonDoesExist "View All"
	
verify_IfWebElementDoesExist_byOutertext "Activity History"

verify_IfWebElementDoesExist_byOutertext "Amendments"
		verifyIfButtonDoesExist "New Amendment"

verify_IfWebElementDoesExist_byOutertext "Project Team"
		verifyIfButtonDoesExist "Add"
		verifyIfButtonDoesExist "Add Default Team"
		verifyIfButtonDoesExist "Display Access"
		verifyIfButtonDoesExist "Delete All"

verify_IfWebElementDoesExist_byOutertext "Project Field History"

verify_IfWebElementDoesExist_byOutertext "Audits"
		verifyIfButtonDoesExist "New Audit"
	
verify_IfWebElementDoesExist_byOutertext "MR Appeals Inquiries"
		verifyIfButtonDoesExist "New MR Appeals Inquiry"
		
verify_IfWebElementDoesExist_byOutertext "CMA Application Reviews"
		verifyIfButtonDoesExist "New CMA Application Review"
		

End Function


Public Function PIR_AttachFile_Portal_Chrome (FilePath)
                
Err.Clear
On Error Resume Next
Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
                Set oShell = CreateObject("WScript.Shell")
                
                Brwser.WebFile("html id:=file", "html tag:=INPUT", "disabled:=0", "visible:=True").Click
wait 2
                oShell.SendKeys "File Path"

wait 3
                Brwser.Dialog("text:=Choose File to Upload", "visible:=True").WinEdit("attached text:=File &name:", "nativeclass:=Edit", "enabled:=True").Set FilePath

Wait 2
				Brwser.Dialog("text:=Choose File to Upload", "visible:=True", "enabled:=True").WinObject("nativeclass:=Button","text:=&Open", "enabled:=True").Click

If err <> 0 Then
    LogReport 4,"Error", err.number & "-" & err.description
    LogReport 1,"Failed to attach file"
else
      LogReport 0,"attached file successfully"
End If

End Function



'Function to Attach file on Portal
Public Function Attach_File_Portal(strFilePath)
On error resume next

clk_Button_usingName "Choose file"
Wait 1
attachFileFromFileSystem_Portal strFilePath
Wait 3
clk_link_Object2 "Upload"
Wait 2

  End Function
  
  'Function to attach files When Fill Out RA Appliation record on Portal
  Function Attach_many_Files_RAapp_Portal (strFilePath, i)
  	
  	clk_Button_usingName_Index "Choose file", i 
  	Wait 2
  	attachFileFromFileSystem_Portal strFilePath
	Wait 5
	clk_Button_usingName_Index "Upload",i
	'''''clk_link_byName_Index "Upload",i
	Wait 3

  End Function
  
  
Function Attach_many_Files_RAapp_Portal_Chrome (strFilePath, i)
  	
  	clk_Button_usingName_Index "Choose file", i 
  	Wait 2
  	PIR_AttachFile_Portal_Chrome strFilePath
	Wait 15
	clk_link_byName_Index "Upload",i
	Wait 3

End Function
  
  'Function to click on WebElement = Magnifying Glass Icon on the Portal
Public Function click_webElementPortal_Magnif_Glass_Icon (i)
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
l("class").value = "fa fa-search"
l("outerhtml").value = "<i class=""fa fa-search""></i>"
'l("visible").value = False
l("html tag").value = "I"

Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click

End Function

'Function to verify that 2 values for Resubmittion = Previous LOI # are matching
Function compare_resub_values()
	On Error Resume Next
''Read Value of "Previous Application" that RAPI entered in LOI Form
Resub_N = readFromFile_Resubmission()

''Capture value of "Previous Application" that is presented Internally on LOI Detail Page
Captured_Element = capture_webElement_value ("00N39000003LYbb_ileinner", "DIV")
Print Captured_Element

'''Compare 2 values to verify they match

If 	CInt(Resub_N) = CInt(Captured_Element) Then
	compare_resub_values() = "Values are matching"
	Print "Values are matching"
	else
	compare_resub_values() = "Values are NOT matching"
	Print "Values are NOT matching"
End If

End Function

Function compare_date_values(compWith, htmlID)
	
	Captured_Element = capture_webElement_value (htmlID, tag)
Print Captured_Element
	
	
If 	CDate(compWith) = CDate(Captured_Element) Then
	
	Print "Values are matching"
	else
	
	Print "Values are NOT matching"	
	
End If
	
	
	
End Function

Function setWebEditBox_HtmlID(strData, HtmlID)
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
editO("html id").value = HtmlID

Set edObj =  getParentObject().ChildObjects(editO) 

for i = 0 To uBound(strArr)

items = edObj(i).GetRoProperty("all items")
print items
print strArr(i)
If Not(strArr(i) = "")Then

edObj(i).set strArr(i)  
oShell.SendKeys strSearchText				            
End If

Next
If err <> 0 Then
err.clear   

End If
End Function


Function setWebEditBox_index(strData, index)
On error resume next

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
''editO("html tag").value = "TEXTAREA"

Set edObj = getParentObject().ChildObjects(editO) 

edObj(index).set strData 

If err <> 0 Then
err.clear   

End If

End Function



'Function to Click WebEdit Box
Function Click_WebEditBox_BYindex (index)
On error resume next

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



Set edObj =  getParentObject().ChildObjects(editO) 

edObj(index).click 
				            
If err <> 0 Then
err.clear   

End If
End Function

Function Click_WebEditBox_byHtmlID (htmlID)
On error resume next

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
editO("html id").value = htmlID



Set edObj =  getParentObject().ChildObjects(editO) 

edObj(0).click 
				            
If err <> 0 Then
err.clear   

End If
End Function


'Function to Capture innertext of the Link

Public Function captureLinkText_QuestionAttachment_LOI(m)
On error resume next 

Set odesc=Description.Create
odesc("micclass").value="WebTable"
''''odesc("column names").value = "Action;Description;Question Attachment Name"
'''''odesc("column names").value =";Choose a RowSelect All;DescriptionShow actions;Sort by:Question Attachment NameSorted: NoneShow actions;"
odesc("column names").value =";Choose a RowSelect All;DescriptionShow Description column actions;Sort by:Question Attachment NameSorted: NoneShow Question Attachment Name column actions;"

odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count


		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
			print strArr(3)
          '''''If trim(strArr(3)) = "Sort by:Question Attachment NameSorted: NoneShow actions" Then
		If trim(strArr(3)) = "Sort by:Question Attachment NameSorted: NoneShow Question Attachment Name column actions" Then
			l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows") 
print a		
		b = l_link(i).GetROProperty("cols") 		
print b		
		For x  = 1 to a     
				For j  = 1 to b
				c = l_link(i).GetCellData(x,j-m)
				print c
				Next
		Next
			End if
		End If
captureLinkText_QuestionAttachment_LOI = c

End Function


'Function to Capture innertext of the Link based on Table Column and Index
''k -  starts with 2 Always, it is LOI Record line
''m - is Column in LOI Review Table based on PI Last Name FC = 0,  "Minus" m = Column # we want
Public Function captureLinkText_LOIReview_Number(k, m)
On error resume next 
Set odesc=Description.Create
odesc("micclass").value="WebTable"

''''''''''''After Each SF releases need to update this value  - This is all the column names for LOi Review Webtable''''''
odesc("column names").value = ";Choose a RowSelect All;Sort by:LOI Accepted or Denied\?Sorted: NoneShow LOI Accepted or Denied\? column actions;Sort by:SF LOI Review NumberSorted: NoneShow SF LOI Review Number column actions;Sort by:Reviewer NameSorted: NoneShow Reviewer Name column actions;Sort by:Reviewer LabelSorted: NoneShow Reviewer Label column actions;Sort by:COISorted: NoneShow COI column actions;Sort by:ProgramSorted: NoneShow Program column actions;Sort by:Institution Name FCSorted: NoneShow Institution Name FC column actions;Sort by:PI First Name FCSorted: NoneShow PI First Name FC column actions;Sort by:PI Last Name FCSorted: NoneShow PI Last Name FC column actions;Sort by:Record TypeSorted: NoneShow Record Type column actions;"

''odesc("column names").value = ";Choose a RowSelect All;Sort by:LOI Accepted or Denied\?Sorted: NoneShow actions;Sort by:SF LOI Review NumberSorted: NoneShow actions;Sort by:Reviewer NameSorted: NoneShow actions;Sort by:Reviewer LabelSorted: NoneShow actions;Sort by:COISorted: NoneShow actions;Sort by:ProgramSorted: NoneShow actions;Sort by:Institution Name FCSorted: NoneShow actions;Sort by:PI First Name FCSorted: NoneShow actions;Sort by:PI Last Name FCSorted: NoneShow actions;Sort by:Record TypeSorted: NoneShow actions;"
odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count


		strName = l_link(i).GetROProperty("column names")
		''print "Table Column Names: "& strName
		
		strArr = split(strName,";")
		If Not(strName = "") Then
			'print strArr(2)
		
		''''''''''''After Each SF releases need to update this value  - This is the column name for LOI review Number''''''
		If trim(strArr(3)) = "Sort by:SF LOI Review NumberSorted: NoneShow SF LOI Review Number column actions" Then
		'If trim(strArr(3)) = "Sort by:SF LOI Review NumberSorted: NoneShow actions" Then
				
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 
		
		For x  = 1 to k     
				For j  = 1 to b
				c = l_link(i).GetCellData(x,b-m)
				''print c
				Next
		Next
			End if
		End If
captureLinkText_LOIReview_Number = c

End Function

'Function to Capture innertext of the Link based on Table Column and Index
''k -  starts with 2 Always, it is LOI Record line
''m - is Column in LOI Review Table based on PI Last Name FC = 0,  "Minus" m = Column # we want
''Public Function captureLinkText_ProjectPersonnelNumber_onRAapp(k, m)
''On error resume next 
''Set odesc=Description.Create
''odesc("micclass").value="WebTable"
''odesc("column names").value = "Action;Project Personnel Number;First Name;Last Name;Status;Role;Key Personnel Flag;Email;Telephone;Institution/Org"
''odesc("html tag").value = "TABLE"
''
''Set l_link=  getParentObject().ChildObjects(odesc)
''''print l_link.count
''
''
''		strName = l_link(i).GetROProperty("column names")
''		''print "Table Column Names: "& strName
''		
''		strArr = split(strName,";")
''		If Not(strName = "") Then
''			'print strArr(2)
''		
''		
''		If trim(strArr(2)) = "Project Personnel Number" Then
''				
''		a = l_link(i).GetROProperty("rows")  
''		b = l_link(i).GetROProperty("cols") 
''		
''		For x  = 1 to k     
''				For j  = 1 to b
''				c = l_link(i).GetCellData(x,b-m)
''				''print c
''				Next
''		Next
''			End if
''		End If
''captureLinkText_ProjectPersonnelNumber_onRAapp = c
''
''End Function




''Function to Capture innertext of the Link based on Table Column and Index in Activity History Section of LOI Detail page or Project Detail Page
''k -  starts with 2 Always, it is Row #, starts always with 2.
''m = 4 = is Column # we want to use in Activity History section; count from the last column => "Last Modified Date/Time" = Arr(0)

Public Function captureLinkText_inActivityHistory_byRow_onLOI_andApp(k, m, PageLayout)
On error resume next 

If PageLayout="LOI" Then

	Set odesc=Description.Create
	odesc("micclass").value="WebTable" 
	odesc("column names").value = "Action;Name;Task;Due Date;Assigned To;Last Modified Date/Time" 
	odesc("html tag").value = "TABLE"
else

	Set odesc=Description.Create
	odesc("micclass").value="WebTable" 
	odesc("column names").value = "Item Number;NameColumn Actions;Assigned ToColumn Actions;Event TypeColumn Actions;StartColumn Actions;Due DateColumn Actions;CompleteColumn Actions;DescriptionColumn Actions;SummaryColumn Actions;Action"
	odesc("html tag").value = "TABLE"

End If

Set l_link=  getParentObject().ChildObjects(odesc)
''print l_link.count


		strName = l_link(i).GetROProperty("column names")
		''print "Table Column Names: "& strName
		
		strArr = split(strName,";")
		If Not(strName = "") Then
			'print strArr(2)
		
		
		If trim(strArr(1)) = "Name" Then
				
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 
		''Print "b - "&b
		
		For x  = 1 to k     
				For j  = 1 to b
				c = l_link(i).GetCellData(x,b-m)
				''print c
				Next
		Next
			End if
		End If
		
captureLinkText_inActivityHistory_byRow_onLOI_andApp = c

End Function

'''''''Function to Capture innertext of the Link based on Table Column and Index in Activity History Section of Project Detail page
'''''''k -  starts with 2 Always, it is Row #, starts always with 2.
'''''''m = 4 = is Column # we want to use in Activity History section; count from the last column => "Last Modified Date/Time" = Arr(0)
'''''
'''''Public Function captureLinkText_inActivityHistory_byRow_onApp(k, m)
'''''On error resume next 
'''''Set odesc=Description.Create
'''''odesc("micclass").value="WebTable" 
'''''odesc("column names").value = "Action;Name;Assigned To;Type;Start;Due Date;Complete;Description;Summary"
'''''odesc("html tag").value = "TABLE"
'''''
'''''Set l_link=  getParentObject().ChildObjects(odesc)
'''''''print l_link.count
'''''
'''''
'''''		strName = l_link(i).GetROProperty("column names")
'''''		''print "Table Column Names: "& strName
'''''		
'''''		strArr = split(strName,";")
'''''		If Not(strName = "") Then
'''''			'print strArr(2)
'''''		
'''''		
'''''		If trim(strArr(1)) = "Name" Then
'''''				
'''''		a = l_link(i).GetROProperty("rows")  
'''''		b = l_link(i).GetROProperty("cols") 
'''''		''Print "b - "&b
'''''		
'''''		For x  = 1 to k     
'''''				For j  = 1 to b
'''''				c = l_link(i).GetCellData(x,b-m)
'''''				''print c
'''''				Next
'''''		Next
'''''			End if
'''''		End If
'''''		
'''''captureLinkText_inActivityHistory_byRow_onApp = c
'''''
'''''End Function


'Function to Capture innertext of the Link=Question Attachment Name based on Question Text in Question Attachments Section of Application Detail page

Public Function captureLinkText_inQuestionAttachments_byRow(Qtext)
On error resume next 
Set odesc=Description.Create
odesc("micclass").value="WebTable"
odesc("column names").value = ";Choose a RowSelect All;Sort by:Question Attachment NameSorted AscendingShow Question Attachment Name column actions;Sort by:Question TextSorted: NoneShow Question Text column actions;"
odesc("html tag").value = "TABLE"
odesc("acc_name").value = "Question Attachments"
Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count


		strName = l_link(i).GetROProperty("column names")
		print "Table Column Names: "& strName
		
		strArr = split(strName,";")
		If Not(strName = "") Then
			print strArr(3)
		
		
		'If trim(strArr(2)) = "Question Text" Then
		'If trim(strArr(3)) = "Sort by:Question TextSorted: NoneShow actions" Then
		''''If trim(strArr(3)) = "	Sort by:Question TextSorted: NoneShow Question Text column actions" Then	
		a = l_link(i).GetROProperty("rows")  
		Print "rows in table - a - "&a
		'''writeToAfile_attachmentsonAPP a
		
		b = l_link(i).GetROProperty("cols") 
		Print "column b - "&b
		
		For x  = 1 to a     
				For j  = 1 to 1
				c = l_link(i).GetCellData(x,b-1)
				print "Captured Question Text in row - "&x&" - "&c
				
					If CStr(c)=CStr(Qtext) Then
					
						m = l_link(i).GetCellData(x,b-2)
						Arr1 = split(m, "Open ")
						Print Arr1(0)
						Print Arr1(1)
New_QsName = Arr1(0)
Print New_QsName
						Print "Captured Question Attachment Name found in row - "&x&" - "&New_QsName                        '''New_QsName
												
						else
						
						Print "Attachment named - "&Qtext&" - NOT found in row - "&x
						
					End If
				
				
				Next
		Next
			End if
		''''End If		
	

captureLinkText_inQuestionAttachments_byRow = New_QsName

End Function



Public Function captureWebElement_inProjectsRelatedList(appName)
On error resume next 
Set odesc=Description.Create
odesc("micclass").value="WebTable"
odesc("column names").value = "Action;Short Project Title;Status;Amount Requested From PCORI;Close Date"
odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
'print l_link.count


		strName = l_link(i).GetROProperty("column names")
		'print "Table Column Names: "& strName
		
		strArr = split(strName,";")
		If Not(strName = "") Then
			'print strArr(2)
		
		
		If trim(strArr(1)) = "Short Project Title" Then
				
		a = l_link(i).GetROProperty("rows")  
		Print "rows in table - a - "&a
		'''writeToAfile_attachmentsonAPP a
		
		b = l_link(i).GetROProperty("cols") 
		Print "b - "&b
		
		For x  = 2 to a     
				For j  = 1 to 1
				c = l_link(i).GetCellData(x,b-3)
				print "Captured Short Project Title - "&x&" - "&c
				
					If CStr(c)=CStr(appName) Then
					
						m = l_link(i).GetCellData(x,b-2)
						Print "Captured App Status in row - "&x&" - "&m
						captureWebElement_inProjectsRelatedList = CStr(m)
						Exit Function
						else
						
						Print "Application with name - "&appName&" - NOT found in row - "&x
						
					End If
				
				
				Next
		Next
			End if
		End If
		


End Function


''Write # of Question Attachments on Application record to test file -----------------------------------------------------*****
Function writeToAfile_attachmentsonAPP (strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile(TextfilePathFor_attachmentsAPP,2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function
'Read from text file - # of Attachments on Application record
Public Function readFromFile_attachmentsonAPP()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile(TextfilePathFor_attachmentsAPP,1)
          
          sContent = objTxtFile.Readline
          readFromFile_attachmentsonAPP = sContent
          
End Function


''Write LOI Template # on Application record to test file -----------------------------------------------------*****
Function writeToAfile_attachmentsonAPP_LOItemplate (strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile(TextfilePathFor_attachmentsAPP_LOItemplate,2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function
'Read from text file - LOI Template #  on Application record
Public Function readFromFile_attachmentsonAPP_LOItemplate()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile(TextfilePathFor_attachmentsAPP_LOItemplate,1)
          
          sContent = objTxtFile.Readline
          readFromFile_attachmentsonAPP_LOItemplate = sContent
          
End Function


''Write Lettersofsupport # on Application record to test file -----------------------------------------------------*****
Function writeToAfile_attachmentsonAPP_Lettersofsupport (strContent)
	On error resume next
	Set Fo = createobject("Scripting.FilesystemObject")
            Set f = Fo.openTextFile(TextfilePathFor_attachmentsAPP_Lettersofsupport,2,true)     'open in write mode
            f.Write (strContent)
            f.Close
            Set f = nothing
End Function
'Read from text file - LOI Template #  on Application record
Public Function readFromFile_attachmentsonAPP_Lettersofsupport()
	On error resume next
	 Set objFSO = CreateObject("Scripting.FileSystemObject")
          Set objTxtFile = objFSO.OpenTextFile(TextfilePathFor_attachmentsAPP_Lettersofsupport,1)
          
          sContent = objTxtFile.Readline
          readFromFile_attachmentsonAPP_Lettersofsupport = sContent
          
End Function

Function capture_numberOfRows_inQuestionAttachments_table()

	On error resume next 
Set odesc=Description.Create
odesc("micclass").value="WebTable"
odesc("column names").value = "Action;Question Attachment Name;Question Text"
odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
'print l_link.count


		strName = l_link(i).GetROProperty("column names")
		'print "Table Column Names: "& strName
		
		strArr = split(strName,";")
		If Not(strName = "") Then
			'print strArr(2)
		
		
		If trim(strArr(2)) = "Question Text" Then
				
		a = l_link(i).GetROProperty("rows")  
		Print "rows in table - a - "&a
	
	capture_numberOfRows_inQuestionAttachments_table = a
	End If
	
End  If

End Function


''Function to verify email notification copy populated in Activity History section of LOI Detail page - always starts with 2 (for both Do Not Invite and Convert) SO-1365
''Make sure that Function to Capture Link text is working properly - Table columns change from time to time
Function verify_EmailNotification_populatedInActivityHistorySection_byRow (i,pageLayout, Actual_Email_name)

	'' 4 - is Column # where Name located if to count from the last column on the right -> to left
If pageLayout = "LOI" Then
	task_name = captureLinkText_inActivityHistory_byRow_onLOI_andApp (i, 4, "LOI")
	Print "Captured Task Name: "& task_name
else
	task_name = captureLinkText_inActivityHistory_byRow_onLOI_andApp (i, 7, "Application")
	Print "Captured Task Name: "& task_name
End If


If task_name = Actual_Email_name Then
	Print "Email Notification populated correct in Activity History section = "&task_name
	LogReport 0,"verify_EmailNotification_populatedInActivityHistorySection_byRow" & i,"Email Notification populated correct in Activity History section = "&task_name
	
	else
	Print "Email Notification did NOT populate in Activity History section"
	LogReport 1,"verify_EmailNotification_populatedInActivityHistorySection_byRow" & i,"Email Notification did NOT populate in Activity History section"
End If

End Function

'Function to click Checkbox with specified Html ID as parameter
Function click_Checkbox_HtmlID (HtmlID)
On error resume next
wait 3
Set ck = Description.Create

ck("micclass").value = "WebCheckBox"
ck("visible").value = true  
ck("html id").value = HtmlID

Set co =   getParentObject().ChildObjects(ck)
print co.count

				If co.count > 0 Then
					LogReport 0,"click_Checkbox_HtmlID "&HtmlID, "The Check Box exists"&HtmlID&"- Expected"
				else
					LogReport 1,"click_Checkbox_HtmlID "&HtmlID, "The Check Box does NOT exist "&HtmlID&"- NOT expected"
				End If

co(0).click

End Function

'''************************************************************************************************************************************************************************************************************
'''''''''''''''''''''''''''''''''''' --------------------- QUIZ functions: START


'' ***  !!!  ***  Functions which commented out are NOT used - those page layout verifications are done during Cycle testing of new Quizes - no need to test it again. Test only functionality.

''Function to Verify that on LOI Record all Resubmission fields under Project Information section are populated with values from LOI Form
Public Function verify_resubmission_fields ()
On error resume next
Value1 = capture_webElement_value ("00N39000003LYay_ileinner", "DIV")
Value2 = capture_webElement_value ("00N39000003LYbd_ileinner", "DIV")
Value3 = capture_webElement_value ("00N39000003LYa9_ileinner", "DIV")
Value4 = capture_webElement_value ("00N39000003LYba_ileinner", "DIV")

If Value1 = "Yes" Then
	Print "LOI Resubmission field is populated with CORRECT value"
Else
	Print "LOI Resubmission field is populated with WRONG value"
End If

If Value2 = "No" Then
	Print "Previous LOI Invited field is populated with CORRECT value"
Else
	Print "Previous LOI Invited field is populated with WRONG value"
End If
If Value3 = "No" Then
	Print "App Resubmission field is populated with CORRECT value"
Else
	Print "App Resubmission field is populated with WRONG value"
End If
If Value4 = "No" Then
	Print "Previous LOI Bypass Review field is populated with CORRECT value"
Else
	Print "Previous LOI Bypass Review field is populated with WRONG value"
End If

If err <> 0 Then
	LogReport 4,"Error", err.number & "-" & err.description
	LogReport 1, "verify_resubmission_fields" & "Failed to verify" & "NOT Expected"
else
  	LogReport 0,"verify_resubmission_fields" & "Able to verify" & "Expected"
End If

End Function 


''Function to Verify all Tabs on LOI Form - Broad 
Function Verify_LOIForm_Tabs()

verifyIfLinkDoesExist1 "Contact Information"
verifyIfLinkDoesExist1 "Pre Screen Questionnaire"
verifyIfLinkDoesExist1 "Resubmission"
verifyIfLinkDoesExist1 "PI Information"
verifyIfLinkDoesExist1 "Project Information"
verifyIfLinkDoesExist1 "Project Personnel"
verifyIfLinkDoesExist1 "Templates and Uploads"

End Function

'Function to verify all Tabs on RA App
Function Verify_RAappForm_Tabs()
	'' Cycle 3
	
verifyIfLinkDoesExist1 "Contact Information"
verifyIfLinkDoesExist1 "Pre Screen Questionnaire"
verifyIfLinkDoesExist1 "Resubmission"
verifyIfLinkDoesExist1 "PI Information"
verifyIfLinkDoesExist1 "Project Information"
verifyIfLinkDoesExist1 "Project Personnel"
verifyIfLinkDoesExist1 "Budget"
''''' SPS-3712 - starting from Cycle 3 - Milestones Tab is removed from Quiz
'''verifyIfLinkDoesExist1 "Milestones"
verifyIfLinkDoesExist1 "Templates & Uploads"
verifyIfLinkDoesExist1 "Certification"

End Function


'''''Function to verify Error messages that pop-up when user tries to save Blank LOI Form
'''Function Verify_ErrorMsg_ContactInfoPage_LOIForm ()
'''verifyIfWebElementDoesExist2 "Error: Please make sure to include a Principal Investigator and Organization for this LOI"
'''Wait 2
'''verifyIfWebElementDoesExist2 "This is a required field\. Principal Investigator \(Contact\)"
'''verifyIfWebElementDoesExist2 "This is a required field\. Administrative Official"
'''verifyIfWebElementDoesExist2 "This is a required field\. Organization"
'''verifyIfWebElementDoesExist2 "This is a required field\. Congressional District"
'''verifyIfWebElementDoesExist2 "This is a required field\. Department"
'''
'''End Function
'''
'''''Function to verify that all fields Exist on Contact Information Page - LOI Form - Broad [modify fields based on the form changes]
'''Function Verify_allFields_ContactInfoPage_LOIForm ()
'''
'''verifyIfWebElementDoesExist2 "Click 'Save & Next' to continue to the next tab. Otherwise you could receive an error message."
'''verifyIfWebElementDoesExist2 "- A Principal Investigator \(PI\) and an Administrative Official \(AO\) with a valid email address are required to submit your LOI to PCORI."
'''verifyIfWebElementDoesExist2 "- To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking “New User\.”"
'''verifyIfWebElementDoesExist2 "- User login information from the previous PCORI Online were not migrated to the new PCORI Online."
'''verifyIfWebElementDoesExist2 "- The AO and the PI cannot be the same individual."
'''verifyIfWebElementDoesExist2 "- Individuals assigned at the “Contact Information” tab will have access to the LOI and Application\."
'''verifyIfWebElementDoesExist2 "- To find your Congressional District please click here."
'''verifyIfWebElementDoesExist2 "- Fields marked with \(\*\) are required"
'''verifyIfWebElementDoesExist2 "Principal Investigator \(Contact\)"
'''verifyIfWebElementDoesExist2 "Dual Principal Investigator Name"
'''verifyIfWebElementDoesExist2 "Dual Principal Investigator Email"
'''verifyIfWebElementDoesExist2 "Administrative Official"
'''verifyIfWebElementDoesExist2 "PI Designee 1"
'''verifyIfWebElementDoesExist2 "PI Designee 2"
'''verifyIfWebElementDoesExist2 "Financial Officer"
'''verifyIfWebElementDoesExist2 "Organization"
'''verifyIfWebElementDoesExist2 "Congressional District"
'''verifyIfWebElementDoesExist2 "Department"
'''Wait 2
'''
'''End Function

'''''METHODS
'''Function verify_fields_onContractInfoPage_LOImethods()
'''
'''verifyIfWebElementDoesExist2 "- A Principal Investigator \(PI\) and an Administrative Official \(AO\) with a valid email address are required to submit your LOI to PCORI. "
'''verifyIfWebElementDoesExist2 "- To assign a user, click the lookup icon and start to type their name\. If the user does not exist in our system, they must register in PCORI Online by clicking ""New User\."" "
'''verifyIfWebElementDoesExist2 "- User login information from the previous PCORI Online were not migrated to the new PCORI Online. "
'''verifyIfWebElementDoesExist2 "- The AO and the PI cannot be the same individual. "
'''verifyIfWebElementDoesExist2 "- Individuals assigned at the ""Contact Information"" tab will have access to the LOI and Application. "
'''verifyIfWebElementDoesExist2 "- To find your Congressional District please click here. "
'''verifyIfWebElementDoesExist2 "- Fields marked with \(\*\) are require"
'''
'''verifyIfWebElementDoesExist2 "Principal Investigator \(Contact\)"
'''verifyIfWebElementDoesExist2 "Dual Principal Investigator Name"
'''verifyIfWebElementDoesExist2 "Dual Principal Investigator Email"
'''verifyIfWebElementDoesExist2 "Administrative Official"
'''verifyIfWebElementDoesExist2 "PI Designee 1"
'''verifyIfWebElementDoesExist2 "PI Designee 2"
'''verifyIfWebElementDoesExist2 "Financial Officer"
'''verifyIfWebElementDoesExist2 "Organization"
'''verifyIfWebElementDoesExist2 "Congressional District"
'''verifyIfWebElementDoesExist2 "Department"
'''Wait 2
'''End Function


''''Function to verify ALL Fields are present on Pre Screen Questionnaire Tab
Function Verify_allFields_PreScreenQPage_LOIForm ()
verifyIfWebElementDoesExist2 Instruction_Text1
verifyIfWebElementDoesExist2 Instruction_Text2
verifyIfWebElementDoesExist2 PreScQIns1
verifyWebElement_By_Outertext Decision_Aid
verifyWebElement_By_Outertext  New_Intervention
verifyWebElement_By_Outertext Practice_Guidelines
verifyWebElement_By_Outertext Cost_Effective_Analysis
verifyIfWebElementDoesExist2 PreScQIns2
verifyWebElement_By_Outertext foreign_organization
verify_IfLinkDoesExist_ByURL "here", Link1URL
verify_IfLinkDoesExist_ByURL "here", Link2URL
verify_IfLinkDoesExist_ByURL "PCORI's award eligibility requirements", PreScQLinkURL1
verify_IfLinkDoesExist_ByURL "FAQ\.", PreScQLinkURL2
Wait 2
End Function

'''''Function to verify Error messages that come up when User tries to save blank Pre Screen Q Tab
Function Verify_errorMsg_PreScreenQPage_LOIFrom ()
verifyWebElement_By_Outertext PScQ_Tab_err1
verifyWebElement_By_Outertext PScQ_Tab_err2
verifyWebElement_By_Outertext PScQ_Tab_err3
verifyWebElement_By_Outertext PScQ_Tab_err4
verifyWebElement_By_Outertext PScQ_Tab_err5
End Function

'''''Function to verify all the fields on REsubmission Tab of LOI Form - Broad
Function Verify_allFields_ResubmissionPage_LOIForm ()
verifyIfWebElementDoesExist2 Instruction_Text1
verifyIfWebElementDoesExist2 Instruction_Text2
verifyIfWebElementDoesExist2 ResubIns1
verifyIfWebElementDoesExist2 ResubIns2
verifyWebElement_By_Outertext LOI_Resubmission
verifyWebElement_By_Outertext Previous_LOI_Invited
verifyWebElement_By_Outertext App_Resubmission
verifyWebElement_By_Outertext Previous_LOI_Bypass_Review
verifyWebElement_By_Outertext Previous_Application
End Function


'''''METHODS - Resubmission Page
'''Function Verify_allFields_ResubmissionPage_LOIForm_Methods ()
'''
'''verifyIfWebElementDoesExist2 "Listed below are questions regarding the submission of an LOI and/or Application to a previous PCORI funding announcement. If you require additional assistance locating a previous application title and/or ID number, please contact pfa@pcori.org"
'''verifyIfWebElementDoesExist2 "Have you submitted this project to PCORI before as an LOI\?"
'''verifyIfWebElementDoesExist2 "If you answered 'Yes' to this question, please answer the following three questions regarding your LOI Resubmission\. If you do not see values in the drop-down lists for the fields below, click ""Save"" at the bottom of the page to respond to the additional questions about your resubmission\."
'''verifyIfWebElementDoesExist2 "Was that LOI invited for a full application\?"
'''verifyIfWebElementDoesExist2 "Have you submitted this project to PCORI before as a full application\?"
'''verifyIfWebElementDoesExist2 "After previous Merit Review cycle, did you receive an invitation from PCORI inviting you to bypass the LOI process for this project\?"
'''verifyIfWebElementDoesExist2 "If answered ""yes"", please attach the invitation letter in the ""Templates \& Uploads"" section"
'''verifyIfWebElementDoesExist2 "Please enter the ID of your prior application\(s\)"
'''Wait 3
'''End Function

'''''Function to verify on Resubmission Tab of LOI Form - if User select "No" to the first Question -> the rest of the questions is not accessible
'''Function Verify_otherFields_ResubmissionPage_LOIForm ()
'''
'''verifyIfWebList_Resub_Exist1 "j_id0:mainForm:j_id360:2:inputFieldId"
'''verifyIfWebList_Resub_Exist1 "j_id0:mainForm:j_id360:3:inputFieldId"
'''verifyIfWebList_Resub_Exist1 "j_id0:mainForm:j_id360:4:inputFieldId"
'''Wait 2
'''
'''verifyIfWebListDoes_NOT_Exist "2"
'''verifyIfWebListDoes_NOT_Exist "3"
'''verifyIfWebListDoes_NOT_Exist "4"
'''Wait 2
'''
'''End Function

'''''Function to verify Error messages that come up on PI Information Page when click to Save blanc form - LOI Form
Function Verify_ErrorMsg_PIinfoPage_LOIForm ()
verifyWebElement_By_Outertext PIinfo_Tab_err1
verifyWebElement_By_Outertext PIinfo_Tab_err2
verifyWebElement_By_Outertext PIinfo_Tab_err3
verifyWebElement_By_Outertext PIinfo_Tab_err4
verifyWebElement_By_Outertext PIinfo_Tab_err5
verifyWebElement_By_Outertext PIinfo_Tab_err6
verifyWebElement_By_Outertext PIinfo_Tab_err7
verifyWebElement_By_Outertext PIinfo_Tab_err8
verifyWebElement_By_Outertext PIinfo_Tab_err9
verifyWebElement_By_Outertext PIinfo_Tab_err10

End Function


'''''Function to verify Error messages that come up on PI Information Page when click to Save blanc form - LOI Form - METHODS
'''Function Verify_ErrorMsg_PIinfoPage_LOIForm_Methods ()
'''	verifyIfWebElementDoesExist2 "Required field : PI Work Telephone Number"
'''
'''verifyIfWebElementDoesExist2 "Required field : For the purpose of this project, with which group does the PI or project lead identify"
'''
'''verifyIfWebElementDoesExist2 "Required field : Have you interacted with PCORI in the past in the following ways\? \(Select all that apply\) "
'''
'''verifyIfWebElementDoesExist2 "Required field : Position Title"
'''
'''verifyIfWebElementDoesExist2 "Required field : Degree"
'''
'''verifyIfWebElementDoesExist2 "Required field : How many years of research experience do you have after attaining your terminal degree\?"
'''
'''verifyIfWebElementDoesExist2 "Required field : How many years of research experience do you have related to this field\?"
'''
'''verifyIfWebElementDoesExist2 "Required field : As the PI or project lead, approximately how many grants/contracts have you had funded"
'''
'''verifyIfWebElementDoesExist2 "Required field : Total dollar amount \(direct cost\) for largest grants/contract for which you were the PI "
'''
'''verifyIfWebElementDoesExist2 "Required field : Have you received grants or contracts from: \(Choose all that apply\)"
'''
'''End Function


'''''Function to Verify all the fields on PI Information PAge - LOI Form
'''Function Verify_allFields_PIinfoPage_LOIForm_Methods ()
'''	verifyIfWebElementDoesExist2 "PI Work Telephone Number"
'''verifyIfWebElementDoesExist2 "For the purpose of this project, with which group does the PI or project lead identify"
'''verifyIfWebElementDoesExist2 "Have you interacted with PCORI in the past in the following ways\? \(Select all that apply\)"
'''verifyIfWebElementDoesExist2 "Please describe ""Other"" previous interactions with PCORI"
'''verifyIfWebElementDoesExist2 "Position Title"
'''verifyIfWebElementDoesExist2 "Degree"
'''verifyIfWebElementDoesExist2 "Please describe ""Other"" degree"
'''verifyIfWebElementDoesExist2 "How many years of research experience do you have after attaining your terminal degree\?"
'''verifyIfWebElementDoesExist2 "How many years of research experience do you have related to this field\?"
'''verifyIfWebElementDoesExist2 "As the PI or project lead, approximately how many grants/contracts have you had funded"
'''verifyIfWebElementDoesExist2 "Total dollar amount \(direct cost\) for largest grants/contract for which you were the PI"
'''verifyIfWebElementDoesExist2 "Have you received grants or contracts from: \(Choose all that apply\)"
'''verifyIfWebElementDoesExist2 "Please describe ""Other"" organizations from which you have received grants/contracts"
'''Wait 2
'''End Function

Function Verify_allFields_PIinfoPage_LOIForm_AD ()
verifyIfWebElementDoesExist2 Instruction_Text1
verifyIfWebElementDoesExist2 Instruction_Text2
wait 2
verify_IfLinkDoesExist_ByURL "here", Link1URL
verify_IfLinkDoesExist_ByURL "here", Link2URL
verifyWebElement_By_Outertext PI_Work_Telephone
verifyWebElement_By_Outertext Primary_Group_Identification
verifyWebElement_By_Outertext Previous_involvement_PCORI
verifyIfWebElementDoesExist2 "Please describe ""Other"" previous interactions with PCORI"
verifyWebElement_By_Outertext Position_Title
verifyWebElement_By_Outertext Project_Lead_Degree
verifyIfWebElementDoesExist2 Project_Lead_Degree_Other
verifyWebElement_By_Outertext Relevant_Exp_Terminal_Degree
verifyWebElement_By_Outertext Relevant_experience
Wait 2
verifyWebElement_By_Outertext Grants_Funded_as_PI
verifyWebElement_By_Outertext Contract_Fund 
verifyWebElement_By_Outertext Previous_Contracts
Wait 2
verifyIfWebElementDoesExist2 "Please describe ""Other"" organizations from which you have received grants/contracts"
Wait 2
End Function


''Function to fill out all the fields on PI Information Page - LOI Form
Function fill_Out_allFields_PIinfoPage_LOIForm ()
selectWeblist "Clinician"

selectWeblist_PI_Information_index "0-4 years", 7
'''selectWeblist_PI_Information_allItems "0-4 years","--None--;0-4 years;5-9 years;10\+ years"

selectWeblist_PI_Information_allItems "3–5 years","--None--;0–2 years;3–5 years;6–9 years;10–15 years;16 years \+"

selectWeblist_PI_Information_allItems "6-10","--None--;0;1-5;6-10;11-15;16-20;21-25;26 or greater"

selectWeblist_PI_Information_allItems "$500,000 - 1 million","--None--;N/A;Less than \$500,000;\$500,000 - 1 million;\$1\.1 - 5 million;\$5\.1 - 10 million;Greater than \$10 million"

setWebEditBox "571-436-9999,Test - Other previous interactions with PCORI,Test - Position Title,Test - Other Degree,Test - Other Organizations"
Wait 2
selectWeblist_PI_Information_allItems "Visited PCORI’s website", "Joined a PCORI email list;Visited PCORI’s website;Participated in applicant training;Watched a PCORI webinar;Attended PCORI sponsored event in-person;Attended event where PCORI was featured;Met with PCORI staff;Met with a PCORI Ambassador;Applied to review PCORI funding app;Applied for PCORI funding;Received PCORI funding;Served as a PCORI Merit reviewer;Participated in a PCORI Advisory Panel;Other \(please specify\);None of the above"

clk_Image_Link_Portal_BYindex 0

selectWeblist_PI_Information_allItems "BCH", "AAS;AB;APRN;BA;BC;BCH;BCHIR;BM;BMBC;BMEDS;BPHAR;BS;BSC;BSN;CHB;CHM;DA;DBA;DC;DCH;DDS;DES;DM;DMD;DMH;DMS;DNSC;DO;DPH;DPHIL;DMP;DSC;DVM;EDD;HS;JD;LPN;MA;MB;MBA;MBBC;MBCHB;MCHIR;MD;MED;MIS;MLIS;MN;MPA;MPH;MPHIL;MS;MSN;MSURG;MSW;ND;NN;NP;PA;PHD;PHRMD;PTA;RN;SB;SCD;Other \(please specify\)"
clk_Image_Link_Portal_BYindex 1

selectWeblist_PI_Information_allItems "AHRQ", "PCORI;AHRQ;CDC;NIH;RWJF;Other \(please specify\);None of the above"

clk_Image_Link_Portal_BYindex 2
Wait 2

End Function

'''''Function to verify error messages that come up when user tries to save blank Project Information Page - LOI Form
Function Verify_errorMsg_ProjectInfoPage_LOIForm ()
verifyWebElement_By_Outertext Proj_Tab_error1
verifyWebElement_By_Outertext Proj_Tab_error2
verifyWebElement_By_Outertext Proj_Tab_error3
verifyWebElement_By_Outertext Proj_Tab_error4
verifyWebElement_By_Outertext Proj_Tab_error5
verifyWebElement_By_Outertext Proj_Tab_error6
verifyWebElement_By_Outertext Proj_Tab_error7
verifyWebElement_By_Outertext Proj_Tab_error7
verifyWebElement_By_Outertext Proj_Tab_error9
verifyWebElement_By_Outertext Proj_Tab_error10
verifyWebElement_By_Outertext Proj_Tab_error11
verifyWebElement_By_Outertext Proj_Tab_error12
verifyWebElement_By_Outertext Proj_Tab_error13
verifyWebElement_By_Outertext Proj_Tab_error14
verifyWebElement_By_Outertext Proj_Tab_error15
verifyWebElement_By_Outertext Proj_Tab_error16
verifyWebElement_By_Outertext Proj_Tab_error17
verifyWebElement_By_Outertext Proj_Tab_error18
verifyWebElement_By_Outertext Proj_Tab_error19
verifyWebElement_By_Outertext Proj_Tab_error20
verifyWebElement_By_Outertext Proj_Tab_error21
verifyWebElement_By_Outertext Proj_Tab_error22
verifyWebElement_By_Outertext Proj_Tab_error23
verifyWebElement_By_Outertext Proj_Tab_error24
End Function

'''''Methods
'''Function verify_errorMsg_ProjectInfoPage_LOImethods()
'''verifyIfWebElementDoesExist2 "Required field : Project Title"
'''verifyIfWebElementDoesExist2 "Required field : Is the primary focus of your study on a rare disease\?"
'''verifyIfWebElementDoesExist2 "Required field : Total direct cost"
'''verifyIfWebElementDoesExist2 "Required field : Total indirect cost "
'''verifyIfWebElementDoesExist2 "Required field : Total amount requested "
'''verifyIfWebElementDoesExist2 "Required field : Please select your estimated project length \(in months\) "
'''verifyIfWebElementDoesExist2 "Required field : Primary disease or condition"
'''verifyIfWebElementDoesExist2 "Required field : Primary disease or condition focus"
'''verifyIfWebElementDoesExist2 "Required field : Secondary disease or condition"
'''verifyIfWebElementDoesExist2 "Required field : Secondary disease or condition focus"
'''verifyIfWebElementDoesExist2 "Required field : Does your proposal focus on any of the following populations\? \(Select all that apply\) "
'''
'''verifyIfWebElementDoesExist2 "Required field : Racial or ethnic minorities"
'''verifyIfWebElementDoesExist2 "Required field : Which of the following healthcare topics is a primary focus of your proposal\? \(Select the best choice\) "
'''verifyIfWebElementDoesExist2 "Required field : Which of the following healthcare topics is a secondary focus of your proposal\? \(Select the best choice\) "
'''verifyIfWebElementDoesExist2 "Required field : Please select the research area\(s\) of interest listed in the Methods Program Cycle 1 2017 PFA to which your application is most responsive \(Select all that apply\) "
'''verifyIfWebElementDoesExist2 "Required field : Methods used- One "
'''verifyIfWebElementDoesExist2 "Required field : Methods used- Two "
'''verifyIfWebElementDoesExist2 "Required field : Methods used- Three "
'''verifyIfWebElementDoesExist2 "Required field : Methods advanced - One"
'''verifyIfWebElementDoesExist2 "Required field : Does any portion of your study proposal include collaborations with existing PCORnet entities \(including CDRNs, PPRNs, or Collaborative Research Groups\)\? "
'''verifyIfWebElementDoesExist2 "Required field : Does your proposed project involve a foreign organization as the prime applicant or foreign organization\(s\) as sub-contractor\(s\) or project performance site\(s\)\? "
'''
'''End Function

'''''Function to verify all the fields on Project Information Page
Function verify_allTheFields_onProjectInfo_page()
verifyIfWebElementDoesExist2 Instruction_Text1
verifyIfWebElementDoesExist2 Instruction_Text2
verifyWebElement_By_Outertext Project_Name
verifyWebElement_By_Outertext Rare_Disease_Focus
verifyWebElement_By_Outertext BPS_National_Priority_Instext
verify_IfLinkDoesExist_ByURL "goal of each National Priority here",BPS_National_Priority_URL
verifyWebElement_By_Outertext National_Priorities_primary_BPS
verifyWebElement_By_Outertext National_Priorities_secondary_BPS
verifyWebElement_By_Outertext National_Priorities_tertiary_BPS
verifyIfWebElementDoesExist2Topic_Themes_InsText
verify_IfLinkDoesExist_ByURL "text of each topic here",Topic_Themes_URL
verifyWebElement_By_Outertext Topic_Themes_Primary
verifyIfWebElementDoesExist2 Topic_Themes_Secondary
verifyIfWebElementDoesExist2 Topic_Themes_Tertiary
verifyWebElement_By_Outertext BPS_Categories
verifyWebElement_By_Outertext Total_Direct_Costs
verifyWebElement_By_Outertext Total_Indirect_Costs
verifyWebElement_By_Outertext LOI_Amount_Requested_from_PCORI
verifyWebElement_By_Outertext Patient_Care_Costs
verifyWebElement_By_Outertext Total_Patient_Care_Costs
verifyWebElement_By_Outertext Project_Duration
verifyIfWebElementDoesExist2 ProjIns1
verifyWebElement_By_Outertext Primary_Disease_Condition
verifyWebElement_By_Outertext Primary_Disease_Condition_Focus
verifyWebElement_By_Outertext Primary_Disease_Condition_Other
verifyIfWebElementDoesExist2 ProjIns2
verifyWebElement_By_Outertext Secondary_Disease_Condition
verifyWebElement_By_Outertext Secondary_Disease_Focus
verifyWebElement_By_Outertext Secondary_Disease_Other
verifyWebElement_By_Outertext Population_Focus
verifyIfWebElementDoesExist2 "Please describe ""Other"" population focus"
verifyWebElement_By_Outertext Racial_Minority_Focus
verifyIfWebElementDoesExist2 "Please describe ""Other"" racial or ethnic minority focus"
verifyWebElement_By_Outertext Healthcare_Primary_Focus
verifyIfWebElementDoesExist2 Healthcare_Primary_Other
verifyWebElement_By_Outertext Healthcare_Secondary_Focus	
verifyWebElement_By_Outertext Healthcare_Secondary_Other
verifyWebElement_By_Outertext Sample_Size
verifyWebElement_By_Outertext PCORnet_involvement
verifyWebElement_By_Outertext PCORnet_Front_Door 
verifyWebElement_By_Outertext PCORnet_Front_Door_Number 
verifyWebElement_By_Outertext PCORnet_ID_Network
verifyWebElement_By_Outertext Address_SAE
verifyWebElement_By_Outertext Which_SAE
verifyWebElement_By_Outertext PCORI_Research_Area_Topic
	End Function

'''''METHODS:
'''Function verify_allTheFields_onProjectInfoPage_LOImethods()
'''
'''		verifyIfWebElementDoesExist2 "Project Title"
'''	verifyIfWebElementDoesExist2 "Is the primary focus of your study on a rare disease\?"
'''	verifyIfWebElementDoesExist2 "Total direct cost"
'''	verifyIfWebElementDoesExist2 "Total indirect cost"
'''	verifyIfWebElementDoesExist2 "Total amount requested"
'''	verifyIfWebElementDoesExist2 "Please select your estimated project length \(in months\)"
'''	verifyIfWebElementDoesExist2 "Which disease or condition is the primary focus of your proposal\? \(Select the best option\)"
'''	verifyIfWebElementDoesExist2 "Primary disease or condition"
'''	
'''	verifyIfWebElementDoesExist2 "If you do not see values in the drop-down list for the field below, click ""Save"" at the bottom of the page to respond to the additional questions\."
'''	verifyIfWebElementDoesExist2 "Primary disease or condition focus"
'''	verifyIfWebElementDoesExist2 "Please describe ""Other"" primary disease or condition"
'''	verifyIfWebElementDoesExist2 "Which disease or condition is the secondary focus of your proposal\? \(Select the best option\)"
'''	verifyIfWebElementDoesExist2 "Secondary disease or condition"
'''	
'''	verifyIfWebElementDoesExist2 "Secondary disease or condition focus"
'''	
'''	verifyIfWebElementDoesExist2 "Please describe ""Other"" secondary disease or condition"
'''	verifyIfWebElementDoesExist2 "Does your proposal focus on any of the following populations\? \(Select all that apply\)"
'''	verifyIfWebElementDoesExist2 "Please describe ""Other"" population focus"
'''	verifyIfWebElementDoesExist2 "Racial or ethnic minorities"
'''	verifyIfWebElementDoesExist2 "Please describe ""Other"" racial or ethnic minority focus"
'''	verifyIfWebElementDoesExist2 "Which of the following healthcare topics is a primary focus of your proposal\? \(Select the best choice\)"
'''	verifyIfWebElementDoesExist2 "Please describe ""Other"" primary focus healthcare topic"
'''	verifyIfWebElementDoesExist2 "Which of the following healthcare topics is a secondary focus of your proposal\? \(Select the best choice\)"
'''	verifyIfWebElementDoesExist2 "Please describe ""Other"" secondary focus healthcare topic"
'''	verifyIfWebElementDoesExist2 "Please select the research area\(s\) of interest listed in the Methods Program Cycle 1 2017 PFA to which your application is most responsive \(Select all that apply\)"
'''	verifyIfWebElementDoesExist2 "Please describe ""Other"" research area\(s\)"
'''	
'''	verifyIfWebElementDoesExist2 "Methods used- One"
'''	verifyIfWebElementDoesExist2 "Methods used- Two"
'''	verifyIfWebElementDoesExist2 "Methods used- Three"
'''	verifyIfWebElementDoesExist2 "Methods used- Four"
'''	verifyIfWebElementDoesExist2 "Methods used- Five"
'''	verifyIfWebElementDoesExist2 "Methods used- Six"
'''	verifyIfWebElementDoesExist2 "What method\(s\) does the application aim to advance\? \(Please identify at least one method\)"
'''	verifyIfWebElementDoesExist2 "Methods advanced - One"
'''	verifyIfWebElementDoesExist2 "Methods advanced - Two"
'''	verifyIfWebElementDoesExist2 "Methods advanced - Three"
'''	verifyIfWebElementDoesExist2 "Methods advanced - Four"
'''	
'''	verifyIfWebElementDoesExist2 "Does any portion of your study proposal include collaborations with existing PCORnet entities \(including CDRNs, PPRNs, or Collaborative Research Groups\)\?"
'''	verifyIfWebElementDoesExist2 "Does your proposed project involve a foreign organization as the prime applicant or foreign organization\(s\) as sub-contractor\(s\) or project performance site\(s\)\?"
'''	verifyIfWebElementDoesExist2 "If yes, please make sure that you review PCORI's award eligibility requirements and the guidance detailed in this FAQ\."
'''	
'''End Function

''Function to fill out all fields on Project Information Tab
Function fill_Out_allFields_ProjectInfoPage_LOIForm ()
	''Cycle 3
	
setWebEditBox "SA Broad LOI Test RA Project Auto_"& RandomString(3)
		
setWebEditBox ",23000,45000,70000,100000,Test Other primary disease or condition,Test Other secondary secondary disease or condition,Test Other population population focus,Test Other racial -  racial or ethnic minority focus,Test Other primary focus healthcare topic,Test Other secondary focus healthcare topic,2500,Test PCORnet Front Door Number new text field"
Wait 3
selectWeblist "No,Achieve Health Equity,Achieve Health Equity,Achieve Health Equity,Promoting health for older adults,Promoting healthy children and youth,Improving cardiovascular health,Category 3 (PCORnet® Study),No,30,Kidney Disease"
Wait 5

selectWeblist_PI_Information_index "Other",11
selectWeblist_PI_Information_index "Cancer", 12
selectWeblist_PI_Information_index "Bladder Cancer", 13

selectWeblistany "j_id0:j_id2:j_id3:mainForm:j_id713:28:inputFieldId", "Disparities: Other"
selectWeblistany "j_id0:j_id2:j_id3:mainForm:j_id713:30:inputFieldId", "Disparities: Demographic"
selectWeblistany "j_id0:j_id2:j_id3:mainForm:j_id713:33:inputFieldId", "Yes"
selectWeblistany "j_id0:j_id2:j_id3:mainForm:j_id713:34:inputFieldId", "No"

''''''''''This is for Does your study proposal address a PCORI Special Area of Emphasis (SAE)? picklist field''''
selectWeblistany "j_id0:j_id2:j_id3:mainForm:j_id713:37:inputFieldId", "Yes"

''''''''''''This is for If Yes, which SAE does your study address?  field'''''
selectWeblistany "j_id0:j_id2:j_id3:mainForm:j_id713:38:inputFieldId", "C3 BPS: Long COVID"

'''''''''''This is for Does your proposal focus on any of the following populations? (Select all that apply) question'''''
selectWeblistany "j_id0:j_id2:j_id3:mainForm:j_id713:24:inputFieldId_unselected", "Children 13-18"
clk_Image_Link_Portal_BYindex 0
wait 2
''''''''''''''''''This is for Racial or ethnic minorities question
selectWeblistany "j_id0:j_id2:j_id3:mainForm:j_id713:26:inputFieldId_unselected", "Asian"
clk_Image_Link_Portal_BYindex 1

'''''''''''''This is for Category 3: Please identify at least one PCORnet® Clinical Research Network that will submit a Letter of Support on your behalf if you are invited to submit a full application question
wait 2 
selectWeblistany "j_id0:j_id2:j_id3:mainForm:j_id713:36:inputFieldId_unselected", "GPC"
clk_Image_Link_Portal_BYindex 2

'''''''''''''''''''''''This is for Does your study proposal address a PCORI Research Priority Area (IDD/MMM/COVID-19/Rare Disease)? question

wait 2 
selectWeblistany "j_id0:j_id2:j_id3:mainForm:j_id713:39:inputFieldId_unselected", "4) Rare Disease"
clk_Image_Link_Portal_BYindex 3
End Function

''Function to fill out all fields on Project Information Tab - LOI Form Methods
Function fill_Out_allFields_ProjectInfoPage_LOIForm_Methods ()
'' UPDATED on 2/19/19
'' Cycle 3:

setWebEditBox "RA App Methods Auto_"& RandomString(3)

setWebEditBox ",22000,41000,77000,Test Other primary disease,Test Other secondary disease,Test Other population,Test Other racial,Test Other primary healthcare,Test Other secondary healthcare,TEST - Methods used- One,TEST - Methods used- Two,TEST - Methods used- Three,TEST - Methods used- Four,TEST - Methods used- Five,TEST - Methods used- Six,TEST - Methods advanced - One,TEST - Methods advanced - Two,TEST - Methods advanced - Three,TEST - Methods advanced - Four"
Wait 2

selectWeblist "No,20,Kidney Disease"
Wait 5
selectWeblist_PI_Information_index "Other",3

selectWeblist_PI_Information_index "Blood Disorders", 4

selectWeblist_PI_Information_index "Anemia", 5

selectWeblist_PI_Information_allItems "Children 13-18", "N/A - this proposal does not focus on a population;Children 0-12;Children 13-18;Children 18-21;Adults 21-64;Adults >65;Disabled persons;Racial or ethnic minorities:;Residents of rural areas;Residents of urban areas;Veterans;Women;LGBTQ;Low income groups;Patients with low health literacy/numeracy and/or limited English proficiency;Individuals with multiple chronic conditions;Individuals with rare or genetic disease;Other \(please specify\)"
selectWeblist_PI_Information_allItems "Asian", "American Indian or Alaska Native;Asian;Black or African American;Hispanic/Latino;Native Hawaiian or Pacific Islander;White;Two or more races;Other \(please specify\)"
selectWeblist_PI_Information_allItems "Methods to Improve Study Design", "Methods Related to Ethics & Human Subjects Protections;Methods to Improve Study Design;Methods to Support Data Research Networks;Methods to Improve the Use of NLP;Methods - Other"
clk_Image_Link_Portal_BYindex 0
clk_Image_Link_Portal_BYindex 1
clk_Image_Link_Portal_BYindex 2

selectWeblist_PI_Information_index "Disparities: Other",12
selectWeblist_PI_Information_index "Disparities: Demographic", 13
Wait 2

selectWeblist_PI_Information_index "No", 17

End Function

'''''''''''''''Function to verify all the text on Project Personnel Tab LOI Form
Function Verify_text_ProjectPersonnelPage_LOIForm ()
verifyWebElement_By_Outertext Pro_Personnel_InText
verifyWebElement_By_OuterHTML "STRONG",Pro_Personnel_BoldedText 

End Function

Function Verify_fields_onNewProjectPersonnel_LOImethods()
''**************Clicking on the link within instructions text and verifying**********'''''''''
wait 2
verify_IfLinkDoesExist_ByURL "here", Link1URL
verify_IfLinkDoesExist_ByURL "here", Link2URL
wait 2
clk_link_byName_Index "here", 0
wait 3
CloseLatestOpenedBrowser()

wait 2
clk_link_byName_Index "here", 1
wait 3
CloseLatestOpenedBrowser()
wait 2
verifyIfWebElementDoesExist2 "For instructions on using our system, click here to access our PCORI Portal User Guide\.To view additional information about the PCORI application process, click here to access our Applicant FAQs page\."
verifyWebElement_By_OuterHTML "P", Pro_Personnel_childrecord_Text
verifyWebElement_By_OuterHTML  "P","<p>At least one key personnel entry is required</p>"
verifyWebElement_By_OuterHTML "P","<p>If stakeholder is selected, you may enter ""N/A"" for institution\.</p>"
verifyIfWebElementDoesExist2 "PCORI is committed to recognizing the contributions of all of the members of the research team, including patient and stakeholder partners, and makes information about the research projects it funds publicly available through various means including but not limited to, press releases and other communications, posting on PCORI’s website, and responding to requests for information about research projects and partners\.  By providing the names of the members of the research team \(including individuals and partnering organizations\), Applicant understands and acknowledges that PCORI may use such information as described above in the event that Applicant is awarded a contract\.  Please contact PCORI staff at pfa@pcori\.org for guidance on how to appropriately recognize partners who want to remain anonymous\."
verifyWebElement_By_Outertext "\* First Name "
verifyWebElement_By_Outertext "\* Last Name "
verifyWebElement_By_Outertext "Institution or Org "
verifyWebElement_By_Outertext "\* Primary perspective on the research team --None-- Patient Stakeholder Scientific "
verifyWebElement_By_Outertext "\* Project Role --None-- PI/Project Lead 1 Name Dual PI Co-PI Co-Investigator Stakeholder Partner Patient Partner Project Manager Other "
verifyWebElement_By_Outertext "Please describe OTHER role "
verifyWebElement_By_Outertext "\* For the purposes of this project, which of the following patient or stakeholder communities reflects this person's primary affiliation\? --None-- Caregiver/Family member of patient Patient/Caregiver Advocacy Organization Clinician Clinic/Hospital/Health System Purchaser Payer Industry Research Policy Maker Training Institution N/A "
verifyWebElement_By_Outertext "\* Degrees AAS AB APRN BA BC BCH BCHIR BM BMBC BMEDS BPHAR BS BSC BSN CHB CHM DA DBA DC DCH DDS DES DM DMD DMH DMP DMS DNSC DO DPH DPHIL DSC DVM EDD HS JD LPN MA MB MBA MBBC MBCHB MCHIR MD MED MIS MLIS MN MPA MPH MPHIL MS MSN MSURG MSW ND NN NP Other \(please specify\) PA PHARMD PHD PTA RN SB SCD "
verifyWebElement_By_Outertext "Please describe ""Other"" degree 0 of 255 Characters "
verifyWebElement_By_Outertext "\* Telephone "
verifyWebElement_By_Outertext "\* Email "
verifyWebElement_By_Outertext "\* Key Personnel Yes No "

End Function

''Function works for both RA LOI and LOI Methods
Function create_newProjectPersonnel_LOImethods()

	'''-------------- Create 1 record for Project Personnel
clk_Button_usingName newBtn

''''-------------- Verify Fields on New Project Personnel Page
Verify_fields_onNewProjectPersonnel_LOImethods()


setWebEditBox projectPersonnel_firstRecord
click_webElement_KeyPersonnel ()
selectWeblist projectPersonnel_fillOutwebLists

'' Click on arrow to the right to ADD Degree
clk_Image_Link_Portal_BYindex 0
Wait 1
clk_Button_usingName saveBtn
Wait 2

End Function

'Function to Submit LOI Application that is in Draft Status
Function submit_LOI_inDraft_Status ()

		clk_link_Object2 tabLOIs
		Wait 2
		clk_link_Object2 tabOpenItems
		Wait 2
		click_webElementPortal_Paper_Pencil_Icon ()
		Wait 3
		clk_Button_usingName reviewSubmitBtn
		Wait 3
		clk_Button_usingName "Submit"
		Wait 5
		
		Set oShell = CreateObject("WScript.Shell") 
		oShell.SendKeys "{ENTER}"
		Wait 6
		
End Function


'''************************************************************************************************************************************************* QUIZ functions: END


'''' INTERNALLY: STARTS

Function fillOut_MROloiPreScreenSection()
	clk_Button_usingName editBtn
setWebEditBox_HtmlID MROname, htmlID_foredit_LOI_preScreen_MRO_webEdit 
setWebEditBox_HtmlID "MRO COMMENTS - TEST Automation run", htmlID_foredit_LOI_preScreen_MRO_webEditcomments
selectWeblist_PI_Information "Responsive", htmlID_foredit_LOI_preScreen_MRO_webList
clk_Button_usingName saveBtn
waitForButton editBtn

''-------------- Verify all the fields are saved with entered values:
verifyIfWebElementDoesExist1 "DIV", MROname, htmlID_forView_LOIPreScreen_MRO_webEdit
verifyIfWebElementDoesExist1 "DIV", "MRO COMMENTS - TEST Automation run", htmlID_forView_LOIPreScreen_MRO_webEditComments
verifyIfWebElementDoesExist1 "DIV", "Responsive", htmlID_forView_LOIPreScreen_MRO_webList

End Function

Function verify_RAEAofficer_cantEditLOIpreScreenFields()
	verify_WebListDoes_NOT_Exist "CF00N39000003LYbp_mlktp"
Wait 2
verify_WebListDoes_NOT_Exist "00N39000003LYbn"
Wait 2
verify_WebEditDoes_NOT_Exist "00N39000003LYbo"
Wait 2
'''-------------- Verify that CMA fields of LOI Pre Screen section are NOT editable by current USer
verify_WebListDoes_NOT_Exist "CF00N39000003LYaJ_mlktp"
Wait 2
verify_WebListDoes_NOT_Exist "00N39000003LYaH"
Wait 2
verify_WebEditDoes_NOT_Exist "00N39000003LYaI"
Wait 2
'''Deprecated SPS-3871
''''''''-------------- Verify that MRO  fields of LOI Pre Screen section are NOT editable by current USer
'''''verify_WebListDoes_NOT_Exist "CF00N39000003LYb5_mlktp"
'''''Wait 2
'''''verify_WebListDoes_NOT_Exist "00N39000003LYb3"
'''''Wait 2
'''''verify_WebEditDoes_NOT_Exist "00N39000003LYb4"
'''''Wait 2

End Function

Function fillOut_SciencePO_loiPreScreenSection()

	clk_Button_usingName editBtn
setWebEditBox_HtmlID "Geeta Bhat", "CF00N39000003LYbp" 
setWebEditBox_HtmlID "Science PO COMMENTS - TEST Automation run", "00N39000003LYbo"
selectWeblist_PI_Information "Responsive", "00N39000003LYbn"
clk_Button_usingName saveBtn
waitForButton editBtn

''-------------- Verify that all entered values are saved:
verifyIfWebElementDoesExist1 "DIV", "Geeta Bhat", "CF00N39000003LYbp_ileinner"
verifyIfWebElementDoesExist1 "DIV", "Science PO COMMENTS - TEST Automation run", "00N39000003LYbo_ileinner"
verifyIfWebElementDoesExist1 "DIV", "Responsive", "00N39000003LYbn_ileinner"

End Function

'Function to Double Click
Function dblClick_webElement_RAapp (HtmlID)

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
	odesc("html id").value= HtmlID
	
	Set O =  getParentObject().ChildObjects(odesc)
	print O.count
	
	O(0).DoubleClick
	Wait 2
	
End Function

''Function to Doubleclick on WebElement based on HtmlID and index 
Function dblClick_webElement_HtmlID_index (HtmlID, index)

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
	odesc("html id").value= HtmlID
	
	Set O =  getParentObject().ChildObjects(odesc)
	print O.count
Wait 3
	
	O(index).DoubleClick
	Wait 5
	
End Function

Function dblClick_image_RAapp (HtmlID)

	On error resume next
	
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

	Set odesc=Description.Create
	odesc("micclass").value="Image"
	odesc("visible").value= True
	odesc("html tag").value= "IMG"
	odesc("html id").value= HtmlID
	
	Set O =  getParentObject().ChildObjects(odesc)
	print O.count

	
	O(0).DoubleClick
	Wait 5
	
End Function

'Function to write text in Frame/WebElement in Edit mode - Internally - NOT Working
Function edit_WebElement_EditMode_Internally (StrComments)
On error resume next
wait 3
Set Brwsr = Browser("micclass:=Browser").Page("micclass:=Page").Frame("micclass:=Frame")

Set ck = Description.Create

ck("micclass").value = "WebElement"
ck("html id").value = "00N39000003LYaeEAG_rta_body"
ck("html tag").value = "BODY"

Set co =   getParentObject().ChildObjects(ck)
co(0).setTOProperty StrComments

If co.count > 0 Then
LogReport 0,"Able to find WebElement and edit it"
else
LogReport 1,"NOT Able to find WebElement and edit it"
End If

End Function

Function selectWeblist_LO_Final_Decision_Status(HtmlID, strData)
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
editO("all items").value = "--None--;Accept;Do Not Invite"
'editO("html tag").value = "INPUT" 
'editO("type").value = "text" 

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

End Function

'Click on CheckBox according to LOI name 
Public Function Click_Checkbox_LOI (LinkInnertext)
                
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
                                
Err.Clear
On Error Resume Next   
                'Capture the page no value and convert it into Integer
PageStr = Brwser.WebElement("class:=right", "html tag:=SPAN").GetRoProperty("innertext")
PageStr=Split(PageStr, "Pageof")
PageNo = Cint (PageStr(1))
'print "Page No is " &PageNo

For i = 1 To PageNo
'Create the description WebTable object
                Set oWTable = Description.Create
                    oWTable("micclass").value = "WebTable"
                    oWTable("html tag").value = "TABLE" 
                    oWTable("class").value = "x-grid3-row-table"

                Set Tables = Brwser.ChildObjects(oWTable)
'Print "No of Table " & Tables.Count

'Create the description Link object - RAPI name needs to be reversed
                
Link2 = Split(LinkInnertext, " ")
Print Link2()
Link3 = Link2(1) & ", " & Link2(0)
Print Link3
                Set oDesc = Description.Create
                oDesc("micclass").value = "Link"
                oDesc("html tag").value = "A" 
                oDesc("visible").value = "True"
                oDesc("innertext").value = Link3

                For j = 0 To Tables.Count-1
                
                                If Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link(oDesc).Exist(3) Then
                                                
                                                WebTableIndex = j
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link(oDesc).Highlight
                                                Brwser.WebCheckbox("class:=checkbox","html tag:=INPUT","name:=ids", "index:="&j).Click
                                
                                        Print " Click on the checkbox according to LOI link name done successfully"   
                                                                
                                        LogReport 0,"Click on the checkbox according to LOI link name " & LinkInnertext," Project " & "-" & LinkInnertext & "-" & " Selected Successfully"                                                          
                                End If
                
                
                Next

Next 
                                                
Set oDesc = Nothing
Set oWTable = Nothing
                
End Function

'Function to verify that LOI with the new status is no longer in specified list view 
Public Function Verify_LOI_is_NOT_inListView (LinkInnertext)
                
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
                                
Err.Clear
On Error Resume Next   
                'Capture the page no value and convert it into Integer
PageStr = Brwser.WebElement("class:=right", "html tag:=SPAN").GetRoProperty("innertext")
PageStr=Split(PageStr, "Pageof")
PageNo = Cint (PageStr(1))
'print "Page No is " &PageNo

For i = 1 To PageNo
'Create the description WebTable object
                Set oWTable = Description.Create
                    oWTable("micclass").value = "WebTable"
                    oWTable("html tag").value = "TABLE" 
                    oWTable("class").value = "x-grid3-row-table"

                Set Tables = Brwser.ChildObjects(oWTable)
'Print "No of Table " & Tables.Count

'Create the description Link object - RAPI name needs to be reversed
                
Link2 = Split(LinkInnertext, " ")
Print Link2()
Link3 = Link2(1) & ", " & Link2(0)
Print Link3 
                Set oDesc = Description.Create
                oDesc("micclass").value = "Link"
                oDesc("html tag").value = "A" 
                oDesc("visible").value = "True"
                oDesc("innertext").value = Link3

                For j = 0 To Tables.Count-1
                
                                If Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link(oDesc).Exist(3) Then
                                                
                                                WebTableIndex = j
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link(oDesc).Highlight
                                                                                               
                                        LogReport 1,"Verifying if " & LinkInnertext," LOI Exists " & "-" & LinkInnertext & "-" & " Exists - NOT Expected"   
								Else 
										LogReport 0,"Verifying if " & LinkInnertext," LOI Exists " & "-" & LinkInnertext & "-" & " Doesn't Exist - Expected"  										
                                End If
                
                
                Next

Next 
                                                
Set oDesc = Nothing
Set oWTable = Nothing
                
End Function

'Function to verify that LOI with the new status is no longer in specified list view 
Public Function Verify_LOI_is_inListView (LinkInnertext)
                
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
                                
Err.Clear
On Error Resume Next   
'                'Capture the page no value and convert it into Integer
PageStr = Brwser.WebElement("class:=right", "html tag:=SPAN").GetRoProperty("innertext")
PageStr=Split(PageStr, "Pageof")
PageNo = Cint (PageStr(1))
'print "Page No is " &PageNo

For i = 1 To PageNo
'Create the description WebTable object
                Set oWTable = Description.Create
                    oWTable("micclass").value = "WebTable"
                    oWTable("html tag").value = "TABLE" 
                    oWTable("class").value = "x-grid3-row-table"

                Set Tables = Brwser.ChildObjects(oWTable)
'Print "No of Table " & Tables.Count

'Create the description Link object - RAPI name needs to be reversed
Link2 = Split(LinkInnertext, " ")
Print Link2()
Link3 = Link2(1) & ", " & Link2(0)
Print Link3            

                Set oDesc = Description.Create
                oDesc("micclass").value = "Link"
                oDesc("html tag").value = "A" 
                oDesc("visible").value = "True"
                oDesc("innertext").value = Link3

                For j = 0 To Tables.Count-1
                
                                If Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).WebElement(oDesc).Exist(3) Then
                                                
                                                WebTableIndex = j
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link(oDesc).Highlight
                                                                                               
                                        LogReport 0,"Verifying if " & LinkInnertext," LOI Exists " & "-" & LinkInnertext & "-" & " Exists - Expected"   
										Exit Function
                                End If
										'LogReport 1,"Verifying if " & LinkInnertext," LOI Exists " & "-" & LinkInnertext & "-" & " Doesn't Exist - NOT Expected"  
                
                Next

Next 
                                                
Set oDesc = Nothing
Set oWTable = Nothing
                
End Function

'Function to click Radio Button
Function click_RadioButton (HtmlID)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True
odesc("html id").value = HtmlID


Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(0).click

End Function

Function click_RadioButton_convert_LOI (i)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True
odesc("name").value = "dupconid"
odesc("html tag").value = "INPUT"

Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(i).click

End Function


Function click_RadioButton_byHtmlID_index (HtmlID, i)
On error resume next
Set odesc=Description.Create
odesc("micclass").value="WebRadioGroup"
odesc("visible").value= True
odesc("html id").value = HtmlID


Set O =  getParentObject().ChildObjects(odesc)
print O.count
O(i).click

End Function


''''Click on Paper/Pencil Icon to Open Draft Status Project in Portal
'''Public Function Click_PaperPencil_Portal_raAPP (StrInnertext)
'''                
'''                Dim Brwser
'''                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
'''                                
'''Err.Clear
'''On Error Resume Next   
'''
'''For i = 1 To 2
''''Create the description WebTable object
'''                Set oWTable = Description.Create
'''                                oWTable("micclass").value = "WebTable"
'''                                oWTable("html tag").value = "TABLE" 
'''                                oWTable("class").value = "datatableRequestOpen dataTable"
'''
'''                Set Tables = Brwser.ChildObjects(oWTable)
'''Print "No of Table " & Tables.Count
'''
'''                Set oDesc = Description.Create
'''                oDesc("micclass").value = "WebElement"
'''                oDesc("html tag").value = "SPAN" 
'''                'oDesc("visible").value = False
'''                oDesc("innertext").value = StrInnertext
'''
'''                For j = 0 To Tables.Count-1
'''                
'''                                If Brwser.WebTable("class:=datatableRequestOpen dataTable","html tag:=TABLE", "index:="&j).WebElement(oDesc).Exist(3) Then
'''                                                
'''                                                WebTableIndex = j
'''                                                Brwser.WebTable("class:=datatableRequestOpen dataTable","html tag:=TABLE", "index:="&j).WebElement(oDesc).Highlight
'''                                                Brwser.WebElement("class:=fa fa-pencil-square-o","html tag:=I", "index:="&j).Click
'''                                
'''                                                                Print " Click on Paper/Pencil Icon according to Project name done successfully"   
'''                                                                
'''                                                                                LogReport 0," Click on Paper/Pencil Icon according to Project name " & StrInnertext," Project " & "-" & StrInnertext & "-" & " Clicked Successfully"                                                          
'''                                End If
'''                
'''                
'''                Next
'''
'''Next 
'''
'''                                                
'''                If err <> 0 Then
'''LogReport 4,"Error ", err.number & "-" & err.description
'''LogReport 1,"Click on the checkbox according to Project link name " & StrInnertext," Project " & "-" & StrInnertext & "-" & " Not Selected"         
'''                
'''                End If
'''
'''Set oDesc = Nothing
'''Set oWTable = Nothing
'''                
'''End Function

'''
''''Click on Paper/Pencil Icon to Open Draft Status LOI Record
'''Public Function Click_PaperPencil_Portal_RA_LOI (StrInnertext)
'''                
'''                Dim Brwser
'''                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
'''                                
'''Err.Clear
'''On Error Resume Next   
'''
'''
''''Create the description WebTable object
'''                Set oWTable = Description.Create
'''                                oWTable("micclass").value = "WebTable"
'''                                oWTable("html tag").value = "TABLE" 
'''                                oWTable("class").value = "datatableInquiryOpen dataTable"
'''
'''                Set Tables = Brwser.ChildObjects(oWTable)
'''Print "No of Tables " & Tables.Count
'''
'''                Set oDesc = Description.Create
'''                oDesc("micclass").value = "WebElement"
'''                oDesc("html tag").value = "SPAN" 
'''                oDesc("visible").value = False
'''                oDesc("innertext").value = StrInnertext
'''
'''                For j = 0 To Tables.Count-1
'''                
'''                                If Brwser.WebTable("class:=datatableInquiryOpen dataTable","html tag:=TABLE").WebElement(oDesc).Exist(3) Then
'''                                                
'''                                                WebTableIndex = j
'''                                                Brwser.WebTable("class:=datatableInquiryOpen dataTable","html tag:=TABLE").WebElement(oDesc).Highlight
'''                                                Brwser.WebElement("class:=fa fa-pencil-square-o","html tag:=I").Click
'''                                
'''                                                      Print " Click on Paper/Pencil Icon according to Project name done successfully"   
'''                                                                
'''                                                     '' LogReport 0," Click on Paper/Pencil Icon according to Project name " & StrInnertext," Project " & "-" & StrInnertext & "-" & " Clicked Successfully"                                                          
'''                                End If
'''                
'''                
'''                Next
'''
'''
'''                                                
'''                If err <> 0 Then
'''LogReport 4,"Error ", err.number & "-" & err.description
'''LogReport 1,"Click on the checkbox according to Project link name " & StrInnertext," Project " & "-" & StrInnertext & "-" & " Not Selected"         
'''                
'''                End If
'''
'''Set oDesc = Nothing
'''Set oWTable = Nothing
'''                
'''End Function

''Function to Fill Out Project Information Tab on RA APP - Portal
Function fill_Out_ProjectInfoTab_RAapp_Portal_Methods ()
On Error Resume Next

'''Projected Start Date
Click_WebEditBox_BYindex 1
Browser("micclass:=Browser").Page("micclass:=Page").WebElement("class:=day", "html tag:=TD","innertext:=23").click
Wait 2

'''Projected End Date
Click_WebEditBox_BYindex 2
Browser("micclass:=Browser").Page("micclass:=Page").WebElement("class:=day", "html tag:=TD","innertext:=28").click
Wait 2



'''Name the study comparators
setWebEditBox_index "TEST Name the study comparators", 3
Wait 2
'''Provide the primary outcome for your proposed study. (Identify only one outcome)
setWebEditBox_index "TEST primary outcome", 4
Wait 2
'''Provide the secondary outcome for your proposed study (if applicable)
setWebEditBox_index "TEST secondary outcome", 5
Wait 2
'''If your project is a health systems project, which factors that drive health system change will your proposal test? (Select all that apply)
selectWeblist_PI_Information_index "N/A- this proposal is not for a health systems project", 0
Wait 2
clk_Image_Link_Portal_BYindex 0
Wait 2
'''Please describe "Other" health systems factors
setWebEditBox_index "TEST Other health systems factors", 6
Wait 2
'''Does your proposed research use any of the following data sources? (Select all that apply)
selectWeblist_PI_Information_index "Medicaid data", 3
Wait 2
clk_Image_Link_Portal_BYindex 1
Wait 2
'''Please describe "Other" data sources
setWebEditBox_index "TEST Other data sources", 7
Wait 2
'''Primary disease or condition focus
selectWeblist_PI_Information_index "Other", 10
Wait 2
'''Total amount requested
setWebEditBox_index "55000", 10
Wait 1

''not in Cycle 3
''''''Secondary disease or condition focus
'''selectWeblist_PI_Information_index "Bladder Cancer", 12
'''Wait 2

''''''Does your proposed project involve a foreign organization as the prime applicant or foreign organization(s) as sub-contractor(s) or project performance site(s)?
'''selectWeblist_PI_Information_index "Yes", 21
'''Wait 2

End Function

Function fillOut_App_abstractFields_IE()

''-------------- Technical Abstract Field
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, j_id0:mainForm:j_id.*", "index:=0").WebElement("html tag:=P", "visible:=True").Object.innertext = "TEST TEST - Technical Abstract"
Wait 3
'''-------------- Public Abstract Field
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, j_id0:mainForm:j_id.*","index:=1").WebElement("html tag:=P", "visible:=True").Object.innertext = "TEST TEST - Public Abstract"
Wait 3

End Function

Function fillOut_App_abstractFields_Chrome()

''-------------- Technical Abstract Field
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, j_id0:mainForm:j_id360:4:inputRichTextAreaId").WebElement("html tag:=P", "visible:=True").Object.innertext = "TEST TEST - Technical Abstract"
Wait 3
'''-------------- Public Abstract Field
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, j_id0:mainForm:j_id360:5:inputRichTextAreaId").WebElement("html tag:=P", "visible:=True").Object.innertext = "TEST TEST - Public Abstract"
Wait 3

End Function


'Function to find WebElement in List and click the link to open it
Public Function FindAndClickLinkInList (LinkInnertext)
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")

Err.Clear
On Error Resume Next   
PageNo = 2
'Create the description Link object
                Set oDesc = Description.Create
oDesc("micclass").value = "Link"
oDesc("html tag").value = "A" 
oDesc("innertext").value = LinkInnertext

'Search the link each until the link appears
                For i = 1 To PageNo
                
                If Brwser.Link(oDesc).Exist(5) Then
                                clk_link_Object2 LinkInnertext
                Exit Function
                Else
                                clk_Image_Link "Next"
                End If
                Next
                
                If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"Finding the link object" & LinkInnertext,"Failed to click the link" & "-" & LinkInnertext
				else
				LogReport 0,"cliking the link object" & LinkInnertext,"The link" & "-" & LinkInnertext & "-" & "is clicked succesfully"
				End If

Set oDesc = Nothing

End Function


'Function to click webelement based on the Link kin the same row
Public Function Click_onProjectPersonnel_ByEmailName (StrInnertext)
                
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
                                
Err.Clear
On Error Resume Next   

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
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link("html tag:=A","visible:=True", "location:="&k+1).Highlight
                                                Wait 3
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link("html tag:=A","visible:=True", "location:="&k+1).Click
                                                Wait 3
                                
                                        Print " Click on Webelement according to Email link name done successfully"   
                                                                
                                        LogReport 0,"Click on the link according to email link name " & StrInnertext," email " & "-" & StrInnertext & "-" & " Selected Successfully"  
				Exit Function    '' to make sure script stop searching for record after record is found (otherwise it will go to every table and keep searching)
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


'Function to click webelement based on the Link in the same row - Contract Administrator for RA App
Public Function Click_onContactAdmin_byProjectName (StrInnertext,p)
                
                Dim Brwser
                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
                                
Err.Clear
On Error Resume Next   

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
                                                
                                                ''CLick on Checkbox for the specified Project
                                                WebTableIndex = j
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link(oDesc).Highlight
                                                Brwser.WebCheckbox("class:=checkbox","html tag:=INPUT","name:=ids", "index:="&j).Click
                                                
                                                
                                                k=0
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link(oDesc).Highlight
                                                Wait 3
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).WebElement("html tag:=DIV","visible:=True", "location:="&k+p).Highlight
                                                Wait 3
                                                Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).WebElement("html tag:=DIV","visible:=True", "location:="&k+p).DoubleClick
                                                Wait 3
                                
                                        Print " Click on Webelement according to Email link name done successfully"   
                                                                
                                        LogReport 0,"Click on the WebElement according to  link name " & StrInnertext," project " & "-" & StrInnertext & "-" & " Selected Successfully"  
						Exit Function
				     End If
                  
                Next
Exit for
Next 
                                     
                                                
                If err <> 0 Then
LogReport 4,"Error ", err.number & "-" & err.description
LogReport 1,"Click on the checkbox according to Project link name " & StrInnertext," Project " & "-" & StrInnertext & "-" & " Not Selected"         
                
                End If

Set oDesc = Nothing
Set oWTable = Nothing
                
End Function


Function click_WinButton (dialog, text)

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd
            
        Browser("hwnd:="&strHwnd).Dialog("text:="&dialog).WinButton("text:="&text).click    
        Wait 3                                                           
                                                                  
End Function

'function to verify if WebElement NOT exist
Public Function verify_webElement_NOT_exist (Innertext)
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
	l("innertext").value = Innertext
	l("visible").value = True 

Set lo =  getParentObject().ChildObjects(l)

				If lo.count = 0 Then
				LogReport 0,"verify_webElement_NOT_exist - " & Innertext, "The Webelement --" & Innertext & "- does not exist - Expected"
				else
				LogReport 1,"verify_webElement_NOT_exist - " & Innertext, "The Webelement --" & Innertext & "- Exists - NOT Expected"	
				End If

End Function

'function to verify if Link NOT exist on the page
Public Function verify_Link_NOT_exist (Innertext)
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
	l("micclass").value = "Link"
	l("innertext").value = Innertext
	l("visible").value = True 

Set lo = getParentObject().ChildObjects(l)
Print lo.count&" - that many links found"

				If lo.count = 0 Then
				LogReport 0,"verify_Link_NOT_exist - " & Innertext, "The Link --" & Innertext & "- NOT exist - Expected"
				else
				LogReport 1,"verify_Link_NOT_exist - " & Innertext, "The Link --" & Innertext & "- Exists - NOT Expected"	
				End If

End Function


'Function to select Weblist using name as parameter
Function selectWeblist_PI_Information_Name(strData, Name)
Err.Clear
On error resume next 
Wait 3

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
editO("name").value = Name

Set edObj =  getParentObject().ChildObjects(editO) 
print edObj.count

for i = 0 To uBound(strArr)
items = edObj(i).GetRoProperty("all items")
print items
print strArr(i)
If Not(strArr(i) = "strData")Then	

edObj(i).select trim(strArr(i))   
 

End If

Next

End Function


''Function to select weblist on a new Window using specified Html ID as parameter
Function select_webList_NEWwindow (strData, HtmlID)

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
wait 5

Set l = Description.Create
l("micclass").value = "WebList"
l("visible").value = True
l("html id").value = HtmlID

Set lo = btn(1).Page("micclass:=Page").ChildObjects(l)
print lo.count

				If lo.count > 0 Then
					LogReport 0,"select_webList_NEWwindow "&strData, "WebList" &strData& " exists and can select"
				else
					LogReport 1,"select_webList_NEWwindow "&strData, "WebList " &strData& " NOT exist and can not select"
				End If

lo(0).select strData
		
End Function


'Function to click Web Button on Budget Tab RA Application
Public Function clk_webButton_usingName3(strName, HtmlID)
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
	print pHwnd
	
	Set editO = Description.Create
	
	editO("micclass").value = "WebButton"
	editO("html tag").value = "INPUT"
	editO("name").value = strName 
	editO("html id").value = HtmlID
	
	Set editObject =  getParentObject().ChildObjects(editO) 
	print editObject.count

editObject(0).highlight
editObject(0).click

				If editObject.count = 0 Then
				LogReport 1,"clk_webButton_usingName3" & strName,"Failed to click the button, button not Found" & "-" & strName
				else
				LogReport 0,"clk_webButton_usingName3" & strName,"The button" & "-" & strName & "-" & "is Found & clicked succesfully"
				End If

End Function


Function fill_out_Budget_Tab_RAapp (strData)
On error resume next
strArr = split(strData,",")

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"

print pHwnd

For k = 0 To 3

clk_webButton_usingName3 "Add Row", "j_id0:budgetForm:j_id64:"&k&":j_id65:j_id87"

waitForButton "Review/Submit"

click_webCheckbox_Name "j_id0:budgetForm:j_id64:"&k&":j_id65:j_id73:0:j_id75"
Wait 2
o = date + 10
m = date + 63

setWebEditBox_HtmlID o, "j_id0:budgetForm:j_id64:"&k&":j_id65:j_id67"
'Wait 2
setWebEditBox_HtmlID m, "j_id0:budgetForm:j_id64:"&k&":j_id65:j_id70"
'Wait 2

						For n = 0 To 2    '# of WebEdits : Hourly Unit Rate & Total
						
								HtmlID = "j_id0:budgetForm:j_id64:"&k&":j_id65:j_id73:0:j_id76:"&n&":j_id78"
						
								Set editO = Description.Create
								editO("micclass").value = "WebEdit"
								editO("visible").value = true
								editO("html id").value = HtmlID
								editO("html tag").value = "INPUT"
						
						Set edObj =  getParentObject().ChildObjects(editO) 
						
												for i = 0 To uBound(strArr)
												items = edObj(i).GetRoProperty("all items")
												print items
												print strArr(i)
												
															If Not(strArr(i) = "")Then
															
															edObj(i).set strArr(n) 
															Wait 2												
																			            
															End If
												
												Next
						
        
						Next

			''clk_webButton_usingName3 "Save Rows", "j_id0:budgetForm:j_id64:"&k&":j_id65:j_id85"
			clk_Button_usingName "Save All Rows"
			waitForButton "OK"
			clk_Button_usingName "OK"
			
Next

	
End Function


'Function to click web checkbox based on Name as parameter
Function click_webCheckbox_Name (Name)
Err.Clear
On error resume next
Wait 3

	Set ck = Description.Create
	
	ck("micclass").value = "WebCheckBox"
	ck("visible").value = true  
	ck("html tag").value = "INPUT"
	ck("name").value = Name

Set co =   getParentObject().ChildObjects(ck)
print co.count

				If co.count > 0 Then
					LogReport 0,"click_webCheckbox_Name "& Name, "The Check Box with the name - "&Name&" - found - Expected"
				else
					LogReport 1,"click_webCheckbox_Name "& Name, "The Check Box with the name - "&Name&" - NOT found - not expected"
				End If

co(0).click

End Function


'Function to click on "+" internally to find more Tabs
Function click_plusIcon_internally ()
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
lo(0).click

				If lo.count = 0 Then
				LogReport 1,"click_plusIcon_internally" & "+ icon","Failed to click the link" & "-" & "+ icon"
				else
				LogReport 0,"click_plusIcon_internally" & "+ icon","The Link" & "-" & "+ icon" & "-" & "is clicked succesfully"
				End If

End Function

''''Function to create Milestones in RA Project - Deprecated
'''Function create_2Milestones_RAapp ()
'''	On Error Resume Next
''''clk_Button_usingName newBtn
''''Wait 4
''''setWebEditBox "Milestone1 - TEST,Test Description - Milestone1"
''''Wait 2
''''Click_WebEditBox_BYindex 2
''''Browser("micclass:=Browser").Page("micclass:=Page").WebElement("class:=day", "html tag:=TD","innertext:=26").click
''''Wait 3
''''clk_Button_usingName saveBtn
''''Wait 5
'''	
'''clk_Button_usingName newBtn
'''Wait 2
'''setWebEditBox "X. 1 Milestone2 - TEST,Test Description - Milestone2"
'''Wait 2
'''Click_WebEditBox_BYindex 2
'''Browser("micclass:=Browser").Page("micclass:=Page").WebElement("class:=day", "html tag:=TD","innertext:=24").click
'''Wait 3
'''clk_Button_usingName saveBtn
'''Wait 5
'''
'''	
'''End Function

'Enter value on Web Edit box by index if many web edit boxes present
Function setWebEditBox_ByIndex(strData, i)
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
	editO("visible").value = true  
	editO("html tag").value = "TEXTAREA" 

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

'''''Function to search for Milestone record in list of many many records - Internally (function will Navigate to the right location on Milestone) - Milestone Tab in Application Deprecated
'''Public Function find_record_in_list ()
'''                
'''                Dim Brwser
'''                Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
'''                                
'''Err.Clear
'''On Error Resume Next   
'''
'''StrInnertext = "2017 Conference for Community Engagement and Healthcare Improvement"
'''
'''For i = 1 To 1
''''Create the description WebTable object
'''                Set oWTable = Description.Create
'''                                oWTable("micclass").value = "WebTable"
'''                                oWTable("html tag").value = "TABLE" 
'''                                oWTable("class").value = "x-grid3-row-table"
'''
'''                Set Tables = Brwser.ChildObjects(oWTable)
'''Print "# of Tables = " & Tables.Count
'''
'''
'''
'''                Set oDesc = Description.Create
'''                oDesc("micclass").value = "Link"
'''                oDesc("html tag").value = "A" 
'''                oDesc("visible").value = "True"
'''                oDesc("innertext").value = StrInnertext
'''
'''                For j = 0 To 0
'''                
'''                                If Brwser.WebTable("class:=x-grid3-row-table","html tag:=TABLE", "index:="&j).Link(oDesc).Exist(5) Then
'''                                            '''''click_webElement_top_email_personnel "R", "SPAN"
'''                                            
'''                                            click_webElement_top_email_personnel "X", "SPAN"
'''											Wait 2
'''								Else
'''								
'''								click_webElement_byClass "x-grid3-hd-inner x-grid3-hd-00N39000003APrd", "DIV"
'''
'''								clk_link_Object2 "X"
'''								
'''''											click_webElement_top_email_personnel "Project Title", "DIV"
'''''											Wait 2
'''''											click_webElement_top_email_personnel "R", "SPAN"
'''''											Wait 2
'''				
'''                                End If
'''                                                
'''                Next
'''
'''Next 
'''                                     
'''                                                
'''                If err <> 0 Then
'''LogReport 4,"Error ", err.number & "-" & err.description
'''LogReport 1,"Click on the checkbox according to Project link name " & StrInnertext," Project " & "-" & StrInnertext & "-" & " Not Selected"         
'''                
'''                End If
'''
'''Set oDesc = Nothing
'''Set oWTable = Nothing
'''                
'''End Function

'Function to CLick WebButton based on Index if Name is the same and many on the Page
Public Function clk_Button_usingName_Index(strName,i)
On error resume next
wait 3

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

editObject(i).highlight
editObject(i).click

				If editObject.count = 0 Then
				LogReport 1,"clk_Button_usingName_Index" & strName,"Failed to click the button" & "-" & strName
				else
				LogReport 0,"clk_Button_usingName_Index" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
				End If

End Function

'Function to create LOI record Fast for general purpose to use it - Status Submitted
Function create_new_LOI_fast ()
''-------------- Login to Portal and create LOI

									Open_ReviewersPortal()
									Wait 3
									''-------------- Read User Name = User Email from Text File saved previously
									readFromFile_RAPI_Email()
									RAPI_email = readFromFile_RAPI_Email
						
									login_intoSalesForce_Application RAPI_email, passWord1
									Wait 2
									clk_link_Object2 portalLoginbtn
									Wait 4
						
''-------------- Navigate to Funding Opportunities
clk_link_Object2 ResearchAwardsbtn
Wait 1
clk_link_Object2 linkFundingOpports
Wait 1
''-------------- Read saved Campaigns names from Text files
Camp1 = readFromFile_Campaign_Name_Methods()

'''-------------- Click on Current campaign and verify that User is able to Apply
clk_link_Object2 Camp1
Wait 2
clk_Button_usingName "Apply"
Wait 2
''-------------- Fill out required fields on Contact Information section and click Save 
readFromFile_RAPI_Name()
UserName = readFromFile_RAPI_Name()

setWebEditBox ",Dual PI Name,dualPI34@yopmail.com,,Alena Aranda,,Financial Contact Smoke Test"
Wait 3
setWebEditBox_ContactInfo_Portal UserName, "RAAO Portal UAT", InstitOrg, "Test DIstrict", "Test Dept"
Wait 3
clk_link_Object2 saveAndNextBtn
Wait 2

''-------------- Select "No" for all the questions and click Save
selectWeblist "No,No,No,No"
Wait 2
clk_link_Object2 saveAndNextBtn
Wait 2

''-------------- Select "Yes" for the first Question => verify User is able to answer the following questions
selectWeblist "Yes"
Wait 2
selectWeblist_Resubmission "No,No,No"
setWebEditBox resubmissionAppfield 
clk_link_Object2 saveAndNextBtn
Wait 2

''-------------- Fill Out all the fields on PI Information Tab and Click Save & Next at the bottom of the Page
fill_Out_allFields_PIinfoPage_LOIForm ()
Wait 5
clk_link_Object2 saveAndNextBtn
Wait 2

''------------- Fill Out all the fields on Project Information Tab and Click Save & Next at the bottom of the Page
fill_Out_allFields_ProjectInfoPage_LOIForm_Methods ()
Wait 5
clk_link_Object2 saveAndNextBtn
Wait 2

''-------------- Create 1 record for Project Personnel
clk_link_Object2 newBtn
setWebEditBox projectPersonnel_firstRecord
click_webElement_KeyPersonnel ()
selectWeblist projectPersonnel_fillOutwebLists
clk_Image_Link_Portal htmlIDrightArrowOnPPNewRecord
Wait 2
clk_Button_usingName saveBtn
Wait 2
clk_link_Object2 "Next"
Wait 2

''-------------- Attach 3 files for future test steps
Attach_File_Portal attachmentResubmission
Attach_File_Portal attachmentAutomation
Attach_File_Portal attachmentPDF
clk_link_Object2 saveBtn
Wait 3

''-------------- SUBMIT LOI Record
clk_Button_usingName reviewSubmitBtn
Wait 3
clk_Button_usingName "Submit"
Wait 3
clk_Button_usingName "OK"
Wait 2

clk_link_Object2 tabLOIs
Wait 2
clk_link_Object2 tabOpenItems
Wait 2

LOI_Num = capture_webElement_value ("j_id0:mainForm:j_id282:iInquirySection:j_id283:2:j_id301:0:j_id312", "SPAN")
writeToAfileLOI_New LOI_Num
Wait 2

create_new_LOI_fast = LOI_Num
End Function

''Function to create LOI Form Template
Function create_LOI_form_Template ()

click_plusIcon_internally ()
Wait 2
clk_link_Object2 "Community Manager"
Wait 2
clk_link_Object2 "Design an Application/View existing Application"
Wait 2
setWebEditBox "New Test Template TC42"
Wait 2
clk_Button_usingName "Create New"
Wait 2
click_webElement_byClass "fa fa-caret-down", "I"
Wait 2
clk_link_Object2 "Portal Tab"
Wait 3
setWebEditBox "TAB1 Alena,6"
Wait 2
selectWeblist_PI_Information_allItems "Inquiry", "Request;Inquiry"
Wait 2
clk_Button_usingName saveBtn
Wait 2
'''-------------- Add New Question
click_RadioButton "newQstn"
Wait 2
setWebEditBox "Question1 Test Alena,Pre-Text1 Alena,Post-Text1 Alena"
Wait 2
selectWeblist_PI_Information_allItems "Instruction", "--None--;Salesforce Data Type;Instruction;Attachment"
Wait 2
clk_Button_usingName saveBtn
Wait 2
clk_Checkbox_byIndex 0
Wait 2
clk_Button_usingName "Save Questions"
Wait 2
''-------------- Add new Question
click_RadioButton "newQstn"
Wait 2
setWebEditBox "Question2 Test Alena,Pre-Text2 Alena,Post-Text2 Alena"
Wait 2
selectWeblist_PI_Information_allItems "Instruction", "--None--;Salesforce Data Type;Instruction;Attachment"
Wait 2
clk_Button_usingName saveBtn
Wait 2
clk_Button_usingName "Save Questions"
Wait 2
''-------------- Add new Tab
click_webElement_byClass "menuButtonLabel dropdown-toggle btn btn-primary", "SPAN"
Wait 3
clk_link_Object2 "Organization Tab"
Wait 2

setWebEditBox_ByIndex "TAB2 Alena", 2
setWebEditBox_ByIndex "7", 3
Wait 2
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "html id:=j_id.*").WebElement("html tag:=P", "visible:=True").Object.innertext = "TEST TEST - Instruction Text Alena"
Wait 3
clk_Button_usingName saveBtn
Wait 2
clk_Button_usingName "Back To Configuration"
Wait 2

End Function

'Function to Edit new LOI Form Template
Function edit_LOI_Form_Template ()
click_RadioButton "newQstn"
Wait 2
setWebEditBox "Question3 Test Alena,Pre-Text3 Alena,Post-Text3 Alena"
Wait 2
selectWeblist_PI_Information_allItems "Instruction", "--None--;Salesforce Data Type;Instruction;Attachment"
Wait 2
clk_Button_usingName saveBtn
Wait 2
clk_Button_usingName "Save Questions"
Wait 2
End Function

'Function to add LOI Form to Campaign record
Function add_LOI_Form_Template_to_Campaign ()
clk_Image_Link "LOI Form Lookup \(New Window\)"
Wait 2
setWebEditBox_NewOpenWindow "New Test Template TC42"
Wait 2
clk_Button_usingName_NewOpenWindow " Go! "
Wait 2
clk_Link_ObjectInApage_NewWindow "New Test Template TC42"
Wait 2

'''				clk_Image_Link "Application Form Lookup \(New Window\)"
'''				Wait 2
'''				setWebEditBox_NewOpenWindow "New Test Template TC38"
'''				Wait 2
'''				clk_Button_usingName_NewOpenWindow " Go! "
'''				Wait 2
'''				clk_Link_ObjectInApage_NewWindow "New Test Template TC38"
'''				Wait 2
'''								
clk_Button_usingName saveBtn
Wait 2

End Function

''''------------------------------------------------- FUNCTIONS to IMPERSONATE: START
Function clk_link_Objectforimpersonate()
Err.Clear
On error resume next
wait 3

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
  
				If lo.count = 0 Then
				LogReport 1,"cliking the link" & "User Action Menu","Failed to click the link" & "-" & "User Action Menu"
				else
				LogReport 0,"cliking the link" & "User Action Menu","The Link" & "-" & "User Action Menu" & "-" & "is clicked succesfully"
				End If

     End Function
     
Function clk_onlink_forimpersonateinternal(HtmlID,strName)
Err.Clear
On error resume next
wait 3

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
  
				If lo.count = 0 Then
				LogReport 1,"clk_onlink_forimpersonateinternal" & strName,"Failed to click the link" & "-" & strName
				else
				LogReport 0,"clk_onlink_forimpersonateinternal" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
				End If

End Function
     
''''''------------------------------------------''''''''This Part: logout as the impersonated user'''''''''''''''''''

Function clk_onlogoutwhendone_impersonating(strName)
Err.Clear
On error resume next
wait 3

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
      l("html tag").value = "A"
                   
Set lo =  getParentObject().ChildObjects(l)
print lo.count
lo(0).click
  
				If lo.count = 0 Then
				LogReport 1,"clk_onlogoutwhendone_impersonating" & strName,"Failed to click the link" & "-" & strName
				else
				LogReport 0,"clk_onlogoutwhendone_impersonating" & strName,"The Link" & "-" & strName & "-" & "is clicked succesfully"
				End If
  
End Function

'''------------------------------------------------- Functions to IMPERSONATE: DONE

Function capture_LOIreview_numbers()

captureLinkText_LOIReview_Number 2,9
ReviewNum1 = captureLinkText_LOIReview_Number(2,9)
Print "Here is the review number 1: " & ReviewNum1
write_ToAfile_anyFilePath_fileName TextfilePathForReviewer1_Number, ReviewNum1
Wait 3

captureLinkText_LOIReview_Number 3,9
ReviewNum2 = captureLinkText_LOIReview_Number(3,9)
Print "Here is the review number 2: " & ReviewNum2
write_ToAfile_anyFilePath_fileName TextfilePathForReviewer2_Number, ReviewNum2
Wait 3

End Function


Function capture_LOIreview_numbers_methods()

captureLinkText_LOIReview_Number 2,9
ReviewNum1 = captureLinkText_LOIReview_Number(2,9)
Print "Here is the review number 1: " & ReviewNum1
write_ToAfile_anyFilePath_fileName TextfilePathForReviewer1_Number_M, ReviewNum1
Wait 3

captureLinkText_LOIReview_Number 3,9
ReviewNum2 = captureLinkText_LOIReview_Number(3,9)
Print "Here is the review number 2: " & ReviewNum2
write_ToAfile_anyFilePath_fileName TextfilePathForReviewer2_Number_M, ReviewNum2
Wait 3

End Function


Function Page_Layout_RA_LOI_review_Methods()

''-------------- LOI Information section Page Layout Verification
verifyIfWebElementDoesExist2 "LOI Review Detail"

verifyIfWebElementDoesExist2 "LOI Information:"

verifyIfWebElementDoesExist2 "LOI"

verifyIfWebElementDoesExist2 "PFA FC"

verifyIfWebElementDoesExist2 "Cycle FC"

verifyIfWebElementDoesExist2 "Program"

verifyIfWebElementDoesExist2 "PFA Type FC"

verifyIfWebElementDoesExist2 "LOI ResubmissionsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZT', 'Have you submitted this project to PCORI before as an LOI\?'\);"

verifyIfWebElementDoesExist2 "Previous LOI InvitedsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZe', 'Was that LOI invited for a full application\?'\);"

verifyIfWebElementDoesExist2 "App ResubmissionsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYYq', 'Have you submitted this project to PCORI before as a full application\?'\);"

verifyIfWebElementDoesExist2 "Previous LOI Bypass ReviewsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZd', 'After previous Merit Review cycle, did you receive an invitation from PCORI inviting you to bypass the LOI process for this project\?'\);"

verifyIfWebElementDoesExist2 "Previous Application Id\(s\) FC"

verifyIfWebElementDoesExist2 "LOI Number"

verifyIfWebElementDoesExist2 "Project Title FC"

verifyIfWebElementDoesExist2 "Institution Name FC"

verifyIfWebElementDoesExist2 "PI First Name FC"

verifyIfWebElementDoesExist2 "PI Last Name FC"

verifyIfWebElementDoesExist2 "Reviewer Name"

verifyIfWebElementDoesExist2 "Reviewer Label"

verifyIfWebElementDoesExist2 "COIsfdcPage\.setHelp\('01I7000000084cu\.00N70000003DUue', 'Conflict of Interest'\);"


''-------------- Review section Page Layout Verification
verifyIfWebElementDoesExist2 "Review:"

verifyIfWebElementDoesExist2 "Responsive Research Areas of Interest"

verifyIfWebElementDoesExist2 "Responsive RAI\(s\) CommentssfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZn', 'If responsive to RAI\(s\), please specify'\);"

verifyIfWebElementDoesExist2 "Includes Clinical Practice Guidelines"

verifyIfWebElementDoesExist2 "Includes CPG CommentssfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZM', 'If LOI includes CPG, please explain why responsive'\);"

verifyIfWebElementDoesExist2 "Does LOI include CER or CEA\?"

verifyIfWebElementDoesExist2 "Includes CER or CEA commentssfdcPage\.setHelp\('01I7000000084cu\.00N39000003Lp1z', 'Includes Cost Effectiveness \(CER\) or Cost Effectiveness Analysis \(CEA\)\?'\);"

verifyIfWebElementDoesExist2 "Includes PCORnetsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZP', 'Includes PCORnet\?'\);"

verifyIfWebElementDoesExist2 "Includes PCORnet CommentssfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZO', 'If LOI includes PCORnet, please provide explanation'\);"

verifyIfWebElementDoesExist2 "Foreign InstitutionsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZG', 'Is the institution applying foreign\?'\);"

verifyIfWebElementDoesExist2 "Greater Than Time Approval"

verifyIfWebElementDoesExist2 "Greater Than Budget Approval"


''-------------- Review Notes section Page Layout Verification
verifyIfWebElementDoesExist2 "Review Notes:"

verifyIfWebElementDoesExist2 "Internal comments for other reviewer"

verifyIfWebElementDoesExist2 "Feedback for public replysfdcPage\.setHelp\('01I7000000084cu\.00N70000003DUv8', 'Use specific and appropriate language\.'\);"

verifyIfWebElementDoesExist2 "LOI Accepted or Denied\?"


''-------------- "Programmatic Responsiveness / Alternate PFA" Notes section Page Layout Verification
verifyIfWebElementDoesExist2 "Programmatic Responsiveness / Alternate PFA:"

verifyIfWebElementDoesExist2 "Programmatically Responsive\?"

verifyIfWebElementDoesExist2 "If ""No"" can it move to alternate PFA\?"

verifyIfWebElementDoesExist2 "If ""Yes"" then which PFA\?sfdcPage\.setHelp\('01I7000000084cu\.00N70000003DUup', 'If &quot;Yes&quot; which PFA should it move to\?'\);"

''--------------- "Check and Click Save to Submit Review:" section Page Latout Verification
verifyIfWebElementDoesExist2 "Check and Click Save to Submit Review:"

verifyIfWebElementDoesExist2 "Review CompletedsfdcPage\.setHelp\('01I7000000084cu\.00N70000003DUv3', 'When review is completed check this box and save\.'\);"

verifyIfWebElementDoesExist2 "System Information:"

verifyIfWebElementDoesExist2 "Owner"

verifyIfWebElementDoesExist2 "Record Type"

verifyIfWebElementDoesExist2 "Last Modified By"

verifyIfWebElementDoesExist2 "Created By"

verifyIfWebElementDoesExist2 "LOI Review History"


End Function


'Function to verify PAGE LAyout for RA LOI Review AD
Function Page_Layout_RA_LOI_review_AD()

	''-------------- LOI Information section Page Layout Verification
verifyIfWebElementDoesExist2 "LOI Review Detail"

verifyIfWebElementDoesExist2 "LOI Information:"

verifyIfWebElementDoesExist2 "LOI"

verifyIfWebElementDoesExist2 "PFA FC"

verifyIfWebElementDoesExist2 "Cycle FC"

verifyIfWebElementDoesExist2 "Program"

verifyIfWebElementDoesExist2 "PFA Type FC"

verifyIfWebElementDoesExist2 "LOI ResubmissionsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZT', 'Have you submitted this project to PCORI before as an LOI\?'\);"

verifyIfWebElementDoesExist2 "Previous LOI InvitedsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZe', 'Was that LOI invited for a full application\?'\);"

verifyIfWebElementDoesExist2 "App ResubmissionsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYYq', 'Have you submitted this project to PCORI before as a full application\?'\);"

verifyIfWebElementDoesExist2 "Previous LOI Bypass ReviewsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZd', 'After previous Merit Review cycle, did you receive an invitation from PCORI inviting you to bypass the LOI process for this project\?'\);"

verifyIfWebElementDoesExist2 "Previous Application Id\(s\) FC"

verifyIfWebElementDoesExist2 "LOI Number"

verifyIfWebElementDoesExist2 "Project Title FC"

verifyIfWebElementDoesExist2 "Institution Name FC"

verifyIfWebElementDoesExist2 "PI First Name FC"

verifyIfWebElementDoesExist2 "PI Last Name FC"

verifyIfWebElementDoesExist2 "Reviewer Name"

verifyIfWebElementDoesExist2 "Reviewer Label"

verifyIfWebElementDoesExist2 "COIsfdcPage\.setHelp\('01I7000000084cu\.00N70000003DUue', 'Conflict of Interest'\);"

''-------------- Review section Page Layout Verification
verifyIfWebElementDoesExist2 "Review:"

verifyIfWebElementDoesExist2 "AD Specific AimssfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZu', 'Should be CER, no CEA'\);"

verifyIfWebElementDoesExist2 "Significance"

verifyIfWebElementDoesExist2 "Background"

verifyIfWebElementDoesExist2 "Study Design"

verifyIfWebElementDoesExist2 "Participants/sites appropriate\?"

verifyIfWebElementDoesExist2 "Outcomes"

verifyIfWebElementDoesExist2 "Reviewer AssessmentsfdcPage\.setHelp\('01I7000000084cu\.00N39000003LYZq', 'The Reviewer\\&#39;s Assessment of the LOI'\);"

'verifyIfWebElementDoesExist2 "Power calculations/Sample Size"

verifyIfWebElementDoesExist2 "Analytic Plan"

verifyIfWebElementDoesExist2 "Comparators"

verifyIfWebElementDoesExist2 "Engagement"

verifyIfWebElementDoesExist2 "Greater Than Time Approval"

verifyIfWebElementDoesExist2 "Greater Than Budget Approval"

''-------------- Review Notes section Page Layout Verification
verifyIfWebElementDoesExist2 "Review Notes:"

verifyIfWebElementDoesExist2 "Internal comments for other reviewer"

verifyIfWebElementDoesExist2 "Feedback for public replysfdcPage\.setHelp\('01I7000000084cu\.00N70000003DUv8', 'Use specific and appropriate language\.'\);"

verifyIfWebElementDoesExist2 "LOI Accepted or Denied\?"

''-------------- "Programmatic Responsiveness / Alternate PFA" Notes section Page Layout Verification
verifyIfWebElementDoesExist2 "Programmatic Responsiveness / Alternate PFA:"

verifyIfWebElementDoesExist2 "Programmatically Responsive\?"

verifyIfWebElementDoesExist2 "If ""No"" can it move to alternate PFA\?"

verifyIfWebElementDoesExist2 "If ""Yes"" then which PFA\?sfdcPage\.setHelp\('01I7000000084cu\.00N70000003DUup', 'If &quot;Yes&quot; which PFA should it move to\?'\);"

''--------------- "Check and Click Save to Submit Review:" section Page Latout Verification
verifyIfWebElementDoesExist2 "Check and Click Save to Submit Review:"

verifyIfWebElementDoesExist2 "Review CompletedsfdcPage\.setHelp\('01I7000000084cu\.00N70000003DUv3', 'When review is completed check this box and save\.'\);"

verifyIfWebElementDoesExist2 "System Information:"

verifyIfWebElementDoesExist2 "Owner"

verifyIfWebElementDoesExist2 "Record Type"

verifyIfWebElementDoesExist2 "Last Modified By"

verifyIfWebElementDoesExist2 "Created By"

verifyIfWebElementDoesExist2 "LOI Review History"

	
	
End Function

''Function to verify that LOI Review Decision has value ACCEPT based on which Row it is - always starts with 2
Function verify_LOIreview_desicion_isACCEPT_byRow (i)
	
	LOI_review_decision = captureLinkText_LOIReview_Number (i, 8)

If LOI_review_decision = "Accept" Then
	Print "LOI Review Decision field saved with correct value"
	LogReport 0,"verify_LOIreview_desicion_isACCEPT_byRow" & i,"Decision:" & "LOI Review Decision field saved with correct value"
	
	else
	Print "LOI Review Decision field saved with NOT correct value"
	LogReport 1,"verify_LOIreview_desicion_isACCEPT_byRow" & i,"Decision:" & "LOI Review Decision field saved with NOT correct value"
End If

End Function

''Function to verify that LOI Review Decision has value ACCEPT based on which Row it is - always starts with 2
Function verify_LOIreview_desicion_isDENY_byRow (i)
	
	LOI_review_decision = captureLinkText_LOIReview_Number (i, 8)

If LOI_review_decision = "Deny" Then
	Print "LOI Review Decision field saved with correct value"
	LogReport 0,"verify_LOIreview_desicion_isACCEPT_byRow" & i,"Decision:" & "LOI Review Decision field saved with correct value"
	
	else
	Print "LOI Review Decision field saved with NOT correct value"
	LogReport 1,"verify_LOIreview_desicion_isACCEPT_byRow" & i,"Decision:" & "LOI Review Decision field saved with NOT correct value"
End If

End Function

''Function to verify LOI Review Owner Name
Function verify_LOIreview_OwnerName (name, HtmlID)

OwnerName = capture_LinkText_HtmlID (HtmlID)
Wait 2

If OwnerName = name Then
	Print "LOI Review Owner populated with current User"
	LogReport 0,"verify_LOIreview_OwnerName - " & name,"Owner name " & name & " = LOI Review Owner is CORRECT"
	
	else
	Print "LOI Review Owner populated with WRONG User"
	LogReport 1,"verify_LOIreview_OwnerName - " & name,"Owner name " & name & " = LOI Review Owner is NOT correct"
End If

End Function

''Function to verify that all Reviewer Labels are in Sequence
Function verify_LOI_Reviewer_Labels (Reviewer_Label, i)
																					'captureLinkText_LOIReview_Number 2, 5
																					'Wait 3
																					'Reviewer_Label1 = captureLinkText_LOIReview_Number (2, 5)
																					'Print "Reviewer1 = " & Reviewer_Label1
																					'
																					'captureLinkText_LOIReview_Number 3, 5
																					'Reviewer_Label2 = captureLinkText_LOIReview_Number (3, 5)
																					'Print "Reviewer2 = " & Reviewer_Label2
																					'
																					'captureLinkText_LOIReview_Number 4, 5
																					'Reviewer_Label3 = captureLinkText_LOIReview_Number (4, 5)
																					'Print "Reviewer3 = " & Reviewer_Label3
																					'
																					'captureLinkText_LOIReview_Number 5, 5
																					'Reviewer_Label4 = captureLinkText_LOIReview_Number (5, 5)
																					'Print "Reviewer4 = " & Reviewer_Label4
																					'
																					'verify_LOI_Reviewer_Labels Reviewer_Label1, 1
																					'Wait 2
																					'verify_LOI_Reviewer_Labels Reviewer_Label2, 2
																					'Wait 2
																					'verify_LOI_Reviewer_Labels Reviewer_Label3, 3
																					'Wait 2
																					'verify_LOI_Reviewer_Labels Reviewer_Label4, 4
																					'Wait 2
'For i = 1 To 4
'
captureLinkText_LOIReview_Number (i+1), 5
Wait 3
Reviewer = captureLinkText_LOIReview_Number ((i+1), 5)

''"Reviewer "&i
	If Reviewer = Reviewer_Label Then
	Print "LOI Reviewer Label" & Reviewer_Label & " - is in sequence"
	LogReport 0,"verify_LOI_Reviewer_Labels" & Reviewer_Label & " - Reviewer Label","Result - " & "LOI Reviewer Label "& Reviewer_Label & " is in Sequence"
	
	else
	Print "LOI Reviewer Label " & Reviewer_Label & "- is NOT in sequence"
	LogReport 1,"verify_LOI_Reviewer_Labels" & Reviewer_Label & "Reviewer Label","Result - " & "LOI Reviewer Label" & Reviewer_Label & " is NOT in Sequence"
	
	End If		
'Next
	
End Function


'Create cycle function
Public Function CreateNewCycle (TestCycle)
'Click all tab '+' sign to see all tab
clk_Image_Link AllTab

'Click the Funding Slates tab
clk_link_Object2 CyclTab

'Click on New Button
clk_Button_usingName newBtn

'Set Cycle Name 
setWebEditBox_ByIndex TestCycle, 1

'Set COI Due Date 
setWebEditBox_ByIndex Date+20, 2

'Click the Save Button
clk_Button_usingName saveBtn
Wait 2

End Function

''Function to verify RAAO approval page on RA APplication form - fields
Function verify_RAAO_approvalPageFields()

'''''verifyIfWebElementDoesExist2 "Only the Administrative Official can access these fields\. Please click the Cancel button if you are not the AO\. Note: Once you select Approved and then SAVE, the application cannot be edited to withdrawn without contacting PCORI\. Please see your PFA for contact information\."
'''''VerifyIfWebElementDoesExist2 "\* I hereby certify that, to the best of my knowledge, the information in this PCORI application is true and accurate\. \* I understand that the discovery of false or fictitious information included in the application may result in rejection from the review process or termination of an award\. \* I certify that the funds applied for will be used as outlined in the proposed budget, and in accordance with the contract terms and conditions\. \* To the best of my knowledge, none of the key personnel and /or collaborating institutions are currently banned from receiving federal funds due to debarment or engagement in research misconduct\. \* By submitting this application, I attest that I am recognized by my institution as an official authorized to enter into contractual agreements and commit institution resources\. "
'''''
''Completely changed after refresh of Sandbox - June 2018:
'''''VerifyIfWebElementDoesExist2 "I hereby certify that, to the best of my knowledge, the information in this PCORI application is true and accurate\."
'''''Wait 2
'''''VerifyIfWebElementDoesExist2 "I understand that the discovery of false or fictitious information included in the application may result in rejection from the review process or termination of an award\."
'''''Wait 2
'''''VerifyIfWebElementDoesExist2 "I certify that the funds applied for will be used as outlined in the proposed budget, and in accordance with the contract terms and conditions\."
'''''Wait 2
'''''VerifyIfWebElementDoesExist2 "To the best of my knowledge, none of the key personnel and /or collaborating institutions are currently banned from receiving federal funds due to debarment or engagement in research misconduct\."
'''''Wait 2
'''''VerifyIfWebElementDoesExist2 "By submitting this application, I attest that I am recognized by my institution as an official authorized to enter into contractual agreements and commit institution resources\."
'''''Wait 2
'''''VerifyIfWebElementDoesExist2 "Once you select Approved/Reject and then SAVE, the application cannot be edited to withdrawn without contacting PCORI\. Please see your PFA for contact information"
'''''Wait 2
'''''VerifyIfWebElementDoesExist2 "\*I Agree "
'''''Wait 2
'''''VerifyIfWebElementDoesExist2 "AO Approval "
'''''Wait 2
'''''VerifyIfWebElementDoesExist2 "Withdraw Application "
'''''Wait 2
'''''VerifyIfWebElementDoesExist2 "Withdraw Reasons "
'''''Wait 2
	VerifyIfWebElementDoesExist2 "Please scroll to the top of the page to select the “Review/Submit” button "
	
End Function

Function verify_RAAO_page_layout_IE(browser)
		
	If browser="IE" Then
		
''VerifyIfWebElementDoesExist2 "Only the Administrative Official can access these fields\. Please click the Cancel button if you are not the AO\.  Note: Once you select Approved and then SAVE, the application cannot be edited to withdrawn without contacting PCORI\. Please see your PFA for contact information\."
''VerifyIfWebElementDoesExist2 "\* I hereby certify that, to the best of my knowledge, the information in this PCORI application is true and accurate\.  \* I understand that the discovery of false or fictitious information included in the application may result in rejection from the review process or termination of an award\.  \* I certify that the funds applied for will be used as outlined in the proposed budget, and in accordance with the contract terms and conditions\. \* To the best of my knowledge, none of the key personnel and /or collaborating institutions are currently banned from receiving federal funds due to debarment or engagement in research misconduct\.  \* By submitting this application, I attest that I am recognized by my institution as an official authorized to enter into contractual agreements and commit institution resources\. "
VerifyIfWebElementDoesExist2 "\*I Agree "
VerifyIfWebElementDoesExist2 "AO Approval "
VerifyIfWebElementDoesExist2 "Withdraw Application "
VerifyIfWebElementDoesExist2 "Withdraw Reasons "
VerifyIfWebElementDoesExist2 "Please scroll to the top of the page to select the “Review/Submit” button"

		else
		
verify_RAAO_approvalPageFields()
				
	End If

End Function



Function fillOut_LOI_review_comments_fields_IE()
	
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUuzEAG, Press ALT 0 for help").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "1TEST TEST - Internal comments for other reviewer"
Wait 3
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUv8EAG, Press ALT 0 for help").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "1TEST TEST - Feedback for public reply"
Wait 3
	
End Function

Function fillOut_LOI_review3_comments_fields_IE()
	
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUuzEAG, Press ALT 0 for help").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "3TEST TEST - Internal comments for other reviewer"
Wait 3
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUv8EAG, Press ALT 0 for help").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "3TEST TEST - Feedback for public reply"
Wait 3
	
End Function

Function fillOut_LOI_review4_comments_fields_IE()
	
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUuzEAG, Press ALT 0 for help").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "4TEST TEST - Internal comments for other reviewer"
Wait 3
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUv8EAG, Press ALT 0 for help").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "4TEST TEST - Feedback for public reply"
Wait 3
	
End Function

Function fillOut_LOI_review5_comments_fields_IE()
	
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUuzEAG, Press ALT 0 for help").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "5TEST TEST - Internal comments for other reviewer"
Wait 3
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUv8EAG, Press ALT 0 for help").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "5TEST TEST - Feedback for public reply"
Wait 3
	
End Function


Function fillOut_LOI_review_comments_fields_Chrome()
	
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUuzEAG").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "1TEST TEST - Internal comments for other reviewer"
Wait 2
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUv8EAG").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "1TEST TEST - Feedback for public reply"
Wait 2

End Function

Function fillOut_LOI_review3_comments_fields_Chrome()
	
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUuzEAG").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "3TEST TEST - Internal comments for other reviewer"
Wait 2
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUv8EAG").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "3TEST TEST - Feedback for public reply"
Wait 2

End Function


Function fillOut_LOI_review4_comments_fields_Chrome()
	
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUuzEAG").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "4TEST TEST - Internal comments for other reviewer"
Wait 2
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUv8EAG").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "4TEST TEST - Feedback for public reply"
Wait 2

End Function

Function fillOut_LOI_review5_comments_fields_Chrome()
	
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUuzEAG").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "5TEST TEST - Internal comments for other reviewer"
Wait 2
Browser("micclass:=Browser").Page("micclass:=Page").Frame("html tag:=IFRAME", "title:=Rich Text Editor, 00N70000003DUv8EAG").WebElement("html tag:=BODY", "visible:=True").Object.innertext = "5TEST TEST - Feedback for public reply"
Wait 2

End Function


Function fillOut_MRO_responsiveness_RAapp()
	''-------------- Cost Effectiveness
dblClick_webElement_RAapp "00N39000003LYd7_ileinner"
selectWeblist_PI_Information "Yes", "00N39000003LYd7"

''-------------- Aims/Comparators
dblClick_webElement_RAapp "00N39000003LYd0_ileinner"
selectWeblist_PI_Information "Yes", "00N39000003LYd0"

''-------------- Guidelines Proposed

Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*").WebElement("html id:=00N39000003LYdl_ileinner", "html tag:=DIV", "innerhtml:=&nbsp;", "visible:=True").DoubleClick
selectWeblist_PI_Information "No", "00N39000003LYdl"

''-------------- Responsiveness to LOI Feedback
dblClick_webElement_RAapp "00N39000003LYgB_ileinner"
selectWeblist_PI_Information "No", "00N39000003LYgB"

''-------------- Move Forward to Online Review
Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*").WebElement("html id:=00N39000003LYea_ileinner", "html tag:=DIV", "innerhtml:=&nbsp;", "visible:=True", "location:=0").DoubleClick
selectWeblist_PI_Information "Flag", "00N39000003LYea"

clk_Button_usingName saveBtn
Wait 2
End Function

Function fillOut_MROresponsiveness_comments_RAapp()

	verifyIfWebElementDoesExist2 "Error: Please enter Guidelines Comments before submitting"
Wait 2
verifyIfWebElementDoesExist2 "Error: Please enter Responsiveness Comments before submitting"
Wait 2
verifyIfWebElementDoesExist2 "Error: Please enter Move Forward Comments before submitting"
Wait 2

''-------------- Fill out comments for MRO responses No/Flag and click Save
Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*").WebElement("html id:=00N39000003LYdk_ileinner", "html tag:=DIV", "innerhtml:=&nbsp;").DoubleClick
Wait 2
setWebEditBox_HtmlID "TEST - Guidelines Comments from MRO", "00N39000003LYdk"
Wait 2
clk_Button_usingName "OK"
Wait 2

dblClick_webElement_RAapp "00N39000003LYgA_ileinner"
Wait 2
setWebEditBox_HtmlID "TEST - Responsiveness Comments from MRO", "00N39000003LYgA"
Wait 2
clk_Button_usingName "OK"
Wait 2

dblClick_webElement_RAapp "00N39000003LYeZ_ileinner"
Wait 2
setWebEditBox_HtmlID "TEST - Move Forward Comments from MRO", "00N39000003LYeZ"
Wait 2
clk_Button_usingName "OK"
Wait 2

clk_Button_usingName saveBtn
Wait 2

''-------------- Verify that comments left by MRO are saved
verifyIfWebElementDoesExist2 "TEST - Guidelines Comments from MRO"
Wait 2
verifyIfWebElementDoesExist2 "TEST - Responsiveness Comments from MRO"
Wait 2
verifyIfWebElementDoesExist2 "TEST - Move Forward Comments from MRO"
Wait 2

End Function

Function fillOut_appResponsiveness_programUser()
	''-------------- Program Decision
dblClick_webElement_RAapp "00N39000003LYfs_ileinner"
Wait 2
selectWeblist_PI_Information "Accept", "00N39000003LYfs"
Wait 2
''-------------- Program Response to CMA
dblClick_webElement_RAapp "00N39000003LYfv_ileinner"
Wait 2
setWebEditBox_HtmlID "TEST - Program Response to CMA", "00N39000003LYfv"
Wait 2
clk_Button_usingName "OK"
Wait 2
''-------------- Program Response to MRO
dblClick_webElement_RAapp "00N39000003LYfw_ileinner"
Wait 2
setWebEditBox_HtmlID "TEST - Program Response to MRO", "00N39000003LYfw"
Wait 2
clk_Button_usingName "OK"
Wait 2
''-------------- Science Program Justification
dblClick_webElement_RAapp "00N39000003LYgJ_ileinner"
Wait 2
setWebEditBox_HtmlID "TEST - Science Program Justification", "00N39000003LYgJ"
Wait 2
clk_Button_usingName "OK"
Wait 2
''-------------- Checkbox "Ready for Merit Review"
dblClick_image_RAapp "00N39000003LYg6_chkbox"
Wait 2
click_Checkbox_HtmlID "00N39000003LYg6"
Wait 2

'''Enhancement SPS-5296 ( Sep27th 19 dep)

stvalue1 = "Organizational or programmatic fit;Potential for synergy within or outside the Program portfolio;Serious methodologic flaw(s);Unconstructive duplication within or outside the Program portfolio;Mitigation of panel scoring differences;Process anomalies"
dblClick_webElement_RAapp "00N39000003LtHR_ileinner"
verifyIfweblistvalueDoesExistbyhtmlid "00N39000003LtHR_unselected",stvalue1
selectWeblistany "00N39000003LtHR_unselected", "Organizational or programmatic fit"
clk_rightarrow_Link_byClass "picklistArrowRight"
clk_Button_usingName "OK"
wait 5
clk_Button_usingName saveBtn

''-------------- Verify that comments left by Science PO are saved
verifyIfWebElementDoesExist2 "TEST - Program Response to CMA"
Wait 2
verifyIfWebElementDoesExist2 "TEST - Program Response to MRO"
Wait 2
verifyIfWebElementDoesExist2 "TEST - Move Forward Comments from MRO"
Wait 2
End Function

''Click on Refresh button
Public Function Clk_RefreshChrome ()
On error resume next 

Dim Brwser
Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
Set oDesc = Description.Create
oDesc("micclass").value = "WebButton"
oDesc("html tag").value = "INPUT"
oDesc("title").value = "Refresh"
oDesc("index").value = 0
oDesc("visible").value = true 

         Brwser.WebButton(oDesc).Highlight
         Brwser.WebButton(oDesc).Click
         
If err <>0 Then      
LogReport 4,"Error", err.number & "-" & err.description
End If

End Function


''Function to verify LOI AD Page Layout as CMA user top part
Function verify_LOI_AD_PageLayout_top_asCMA()

		verifyIfButtonDoesExist editBtn
		verifyIfButtonDoesExist deleteBtn
		verifyIfButtonDoesExist Convertbtn
		verifyIfButtonDoesExist sharing_btn 
		verifyIfButtonDoesExist generatePDF_RAloi
		
verifyIfWebElementDoesExist2 "LOI Detail"

		verifyIfWebElementDoesExist2 "Pre Screen Questionnaire:"
verify_IfWebElementDoesExist_byOutertext "Decision Aid"
verify_IfWebElementDoesExist_byOutertext "New Intervention"
verify_IfWebElementDoesExist_byOutertext "Date Draft Final Research Report Submit"
verifyIfWebElementDoesExist2 "Does it involve a foreign organization"
verify_IfWebElementDoesExist_byOutertext "Practice Guidelines"
verify_IfWebElementDoesExist_byOutertext "Cost Effective Analysis"
''verifyIfWebElementDoesExist_class_Innertext "last labelCol", "Project Personnel"


		verifyIfWebElementDoesExist2 "Contacts:"
verifyIfWebElementDoesExist2 "PI/Project Lead Name"
verifyIfWebElementDoesExist2 "Administrative Official"
verifyIfWebElementDoesExist2 "Dual PI Name"
verifyIfWebElementDoesExist2 "Dual PI Email"
verifyIfWebElementDoesExist2 "PI Designee 1"
verifyIfWebElementDoesExist2 "PI Designee 2"
verifyIfWebElementDoesExist2 "Financial Officer"

		verifyIfWebElementDoesExist2 "PI Information:"
verify_IfWebElementDoesExist_byOutertext "Primary Identification"
verifyIfWebElementDoesExist2 "Primary Group Identification\*"
verify_IfWebElementDoesExist_byOutertext "Previous involvement with PCORI\*"
verifyIfWebElementDoesExist2 "Previous involvement with PCORI - Other"
verifyIfWebElementDoesExist2 "PI Work Telephone"
verify_IfWebElementDoesExist_byOutertext "Position Title"
verifyIfWebElementDoesExist2 "Project Lead Degree\*"
verifyIfWebElementDoesExist2 "Project Lead Degree - Other"

verify_IfWebElementDoesExist_byOutertext "Relevant Exp after Terminal Degree"
verifyIfWebElementDoesExist2 "How many years of relevant experience\?"
verify_IfWebElementDoesExist_byOutertext "Field Relevant Exp"
verify_IfWebElementDoesExist_byOutertext "Grants Funded as PI"
verify_IfWebElementDoesExist_byOutertext "Largest Previous Grant/Contract Fund"
verify_IfWebElementDoesExist_byOutertext "Previous Grants/Contracts"
verify_IfWebElementDoesExist_byOutertext "Other \(please specify\)"

		verifyIfWebElementDoesExist2 "Organization Information:"
verifyIfWebElementDoesExist2 "Awardee Institution/Organization"
verifyIfWebElementDoesExist2 "D-U-N-S Number"
verifyIfWebElementDoesExist2 "Location/satellite"
verifyIfWebElementDoesExist2 "Department"
''verifyIfWebElementDoesExist2 "Congressional District"
verifyIfWebElementDoesExist2 "Congressional District\*"
verifyIfWebElementDoesExist2 "Street Address"
verifyIfWebElementDoesExist2 "City"
verifyIfWebElementDoesExist2 "State/Province"
verifyIfWebElementDoesExist2 "Zip Code"
verifyIfWebElementDoesExist2 "Country"

		verifyIfWebElementDoesExist2 "Project Information:"
''SPS-4416 - PFA field name changed to Primary Campaign Source		
''verifyIfWebElementDoesExist2 "PFA"
verifyIfWebElementDoesExist2 "Primary Campaign Source"
verifyIfWebElementDoesExist2 "Cycle"
verifyIfWebElementDoesExist2 "Program"
verifyIfWebElementDoesExist2 "PFA Type"
verify_IfWebElementDoesExist_byOutertext "IHS Award Size"
verifyIfWebElementDoesExist2 "Project Name\*"
verify_IfWebElementDoesExist_byOutertext "LOI Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Invited"
verify_IfWebElementDoesExist_byOutertext "App Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Bypass Review"
verify_IfWebElementDoesExist_byOutertext "Previous Application"

verifyIfWebElementDoesExist2 "Total Direct Costs"
verifyIfWebElementDoesExist2 "Total Indirect Costs"
verifyIfWebElementDoesExist2 "LOI Amount Requested from PCORI\*"
verify_IfWebElementDoesExist_byOutertext "Project Duration"
verify_IfWebElementDoesExist_byOutertext "PCORnet involvement"
verify_IfWebElementDoesExist_byOutertext "Original PFA"
verifyIfWebElementDoesExist2 "Key Word One"
verifyIfWebElementDoesExist2 "Key Word Two"
verifyIfWebElementDoesExist2 "Key Word Three"
verifyIfWebElementDoesExist2 "Key Word Four"
verifyIfWebElementDoesExist2 "Key Word Five"
verifyIfWebElementDoesExist2 "Key Word Six"

		verifyIfWebElementDoesExist2 "Project Focus:"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Rare Disease Focus"
verify_IfWebElementDoesExist_byOutertext "Population Focus"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Focus"
verify_IfWebElementDoesExist_byOutertext "Study Design"
verify_IfWebElementDoesExist_byOutertext "Specific Analytic Methods"
verify_IfWebElementDoesExist_byOutertext "Sample Size"

verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Population Focus Other"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Other"
verify_IfWebElementDoesExist_byOutertext "Other Study Design"
	verify_IfWebElementDoesExist_byOutertext "Spec Analytic Method Other"
verifyIfWebElementDoesExist2 "Primary Patient Partner\(s\)"
verifyIfWebElementDoesExist2 "Primary Stakeholder Partner\(s\)"

		verifyIfWebElementDoesExist2 "Budget/Time Greater Than Request \(CDR and D&I\):"

verify_IfWebElementDoesExist_byOutertext "Budget/Time Greater Than Request\?"
verify_IfWebElementDoesExist_byOutertext "Submitted to other agency before"
verify_IfWebElementDoesExist_byOutertext "Agency Name"
verify_IfWebElementDoesExist_byOutertext "Considered by other agency"
verify_IfWebElementDoesExist_byOutertext "Partial funding Agency Name"
verify_IfWebElementDoesExist_byOutertext "Program Officer Name"
verify_IfWebElementDoesExist_byOutertext "May we contact PO"
verify_IfWebElementDoesExist_byOutertext "Less costly approach explanation"
verify_IfWebElementDoesExist_byOutertext "Potential for new evidence"
verify_IfWebElementDoesExist_byOutertext "Extra Money Requested"
verify_IfWebElementDoesExist_byOutertext "Extra Money Justification"
verify_IfWebElementDoesExist_byOutertext "Extra Time Requested"
verify_IfWebElementDoesExist_byOutertext "Extra Time Justification"


		verifyIfWebElementDoesExist2 "D&I:"

verify_IfWebElementDoesExist_byOutertext "Original funded application"
verify_IfWebElementDoesExist_byOutertext "Original funded cycle"
verify_IfWebElementDoesExist_byOutertext "Eligible evidence"
verify_IfWebElementDoesExist_byOutertext "Original app PI"
verify_IfWebElementDoesExist_byOutertext "LOI resubmission \(D&I\)"
verify_IfWebElementDoesExist_byOutertext "Was LOI invited \(D&I\)"
verify_IfWebElementDoesExist_byOutertext "Project resubmitted"
verify_IfWebElementDoesExist_byOutertext "Number of submissions"
verify_IfWebElementDoesExist_byOutertext "Summary statement received"
verify_IfWebElementDoesExist_byOutertext "Dissemination of multiple PCORI projects"
verify_IfWebElementDoesExist_byOutertext "Name of PIs"
verify_IfWebElementDoesExist_byOutertext "Contract IDs"
verify_IfWebElementDoesExist_byOutertext "Passive dissemination"
verify_IfWebElementDoesExist_byOutertext "Develop new tool"
''new fields added in Cycle 2 2018
verify_IfWebElementDoesExist_byOutertext "LOI to the D&I Limited Competition PFA\?"
verify_IfWebElementDoesExist_byOutertext "D&I Limited Competition PFA LOI Accepted"
verify_IfWebElementDoesExist_byOutertext "Explain SDM implementation execution"


		verifyIfWebElementDoesExist2 "PPRN:"

verify_IfWebElementDoesExist_byOutertext "PPRN List"
verify_IfWebElementDoesExist_byOutertext "Partnerships"
verify_IfWebElementDoesExist_byOutertext "Total Funding from Partners"
verify_IfWebElementDoesExist_byOutertext "Linking Data Source"
verify_IfWebElementDoesExist_byOutertext "PCORnet collaboration groups"

		verifyIfWebElementDoesExist2 "LOI Pre Screen \(For Internal Use\):"
verifyIfWebElementDoesExist2 "CMA Owner"
verifyIfWebElementDoesExist2 "CMA Comments"
verifyIfWebElementDoesExist2 "Administrative Compliance Flag"
verifyIfWebElementDoesExist2 "Program Owner"
verifyIfWebElementDoesExist2 "Program Response"
verifyIfWebElementDoesExist2 "Program Non Responsiveness Decision"
'''verifyIfWebElementDoesExist2 "MRO Owner"
'''verifyIfWebElementDoesExist2 "MRO Response"
'''verifyIfWebElementDoesExist2 "MRO Non Responsiveness Flag"


		verifyIfWebElementDoesExist2 "Budget/Time Greater Than Request Approval \(For Internal Use\):"

verify_IfWebElementDoesExist_byOutertext "Time Approval Response"
verify_IfWebElementDoesExist_byOutertext "Time Approval"
verify_IfWebElementDoesExist_byOutertext "Budget Approval Response"
verify_IfWebElementDoesExist_byOutertext "Budget Approval"


		verifyIfWebElementDoesExist2 "Final LOI Decision \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Recommended for Alternate PFA"
verifyIfWebElementDoesExist2 "Proposed PFA"
verify_IfWebElementDoesExist_byOutertext "Alternate PFA Rationale"
verifyIfWebElementDoesExist2 "Lead Reviewer"
verifyIfWebElementDoesExist2 "Reviews Status"
verifyIfWebElementDoesExist2 "Draft Comments for Public Reply"
verifyIfWebElementDoesExist2 "Consolidated Internal Comments"
verify_IfWebElementDoesExist_byOutertext "Final Decision Status"
verifyIfWebElementDoesExist2 "Final Comments Screened"
verifyIfWebElementDoesExist2 "Final Comments for Public Reply"
verifyIfWebElementDoesExist2 "Email Proof"
verifyIfWebElementDoesExist2 "Preview Email Communications"


		verifyIfWebElementDoesExist2 "Status \(For Internal Use\):"
verifyIfWebElementDoesExist2 "LOI Number"
verifyIfWebElementDoesExist2 "LOI Status"
verifyIfWebElementDoesExist2 "External Status"
verifyIfWebElementDoesExist2 "Name"
verifyIfWebElementDoesExist2 "LOI Owner"
verifyIfWebElementDoesExist2 "LOI Review1"
verifyIfWebElementDoesExist2 "Review Roll-up 1"
verifyIfWebElementDoesExist2 "LOI Review2"
verifyIfWebElementDoesExist2 "Review Roll-up 2"
verifyIfWebElementDoesExist2 "LOI Review3"
verifyIfWebElementDoesExist2 "Review Roll-up 3"
verifyIfWebElementDoesExist2 "LOI Review4"
verifyIfWebElementDoesExist2 "Review Roll-up 4"
verifyIfWebElementDoesExist2 "Project Lead Email Address\*"
verifyIfWebElementDoesExist2 "Legal name of organization\*"
verifyIfWebElementDoesExist2 "LOI Submission Deadline Date"

'''''verifyIfWebElementDoesExist2 "Custom Links"
verifyIfWebElementDoesExist2 "Created By"
verify_IfWebElementDoesExist_byOutertext "Submitted By"
verify_IfWebElementDoesExist_byOutertext "Submission Date"
verifyIfWebElementDoesExist2 "Last Modified By"
verifyIfWebElementDoesExist2 "LOI Record Type"
verifyIfWebElementDoesExist2 "Application Start Date"


	
End Function

''Function to verify bottom part of LOI Deail Page - AD - as CMA USer
Function verify_LOI_AD_PageLayout_bottomPart_asCMA()

		verifyIfWebElementDoesExist_HtmlId "00Qc000000GSXIv_00N39000003LYa3_title"
verifyIfButtonDoesExist "New Project Personnel"

		''LOI Review related list
		verifyIfWebElementDoesExist_HtmlId "00Qc000000GSXIv_00N39000003LYZU_title"

		verify_IfWebElementDoesExist_byOutertext "Question Attachments"
verifyIfButtonDoesExist "New Question Attachment"

		verify_IfWebElementDoesExist_byOutertext "Notes & Attachments"
verifyIfButtonDoesExist "New Note"
verifyIfButtonDoesExist "Attach File"
verifyIfButtonDoesExist "View All"

verify_IfWebElementDoesExist_byOutertext "Lead History"
		verify_IfWebElementDoesExist_byOutertext "Activity History"
verifyIfButtonDoesExist "Log a Call"
verifyIfButtonDoesExist "Mail Merge"
verifyIfButtonDoesExist "Send an Email"
		verify_IfWebElementDoesExist_byOutertext "Campaign History"
verifyIfButtonDoesExist "Add to Campaign"
	
		
End Function

''Function to verify LOI Page Layout = Submitted as MRO
Function verify_LOI_AD_pageLayout_asMRO()
			verifyIfButtonDoesExist editBtn
		verifyIfButtonDoesExist deleteBtn
		verifyIfButtonDoesExist Convertbtn
		verifyIfButtonDoesExist sharing_btn 
		verifyIfButtonDoesExist generatePDF_RAloi
		
verifyIfWebElementDoesExist2 "LOI Detail"

		verifyIfWebElementDoesExist2 "Pre Screen Questionnaire:"
verify_IfWebElementDoesExist_byOutertext "Decision Aid"
verify_IfWebElementDoesExist_byOutertext "New Intervention"
verify_IfWebElementDoesExist_byOutertext "Date Draft Final Research Report Submit"
verifyIfWebElementDoesExist2 "Does it involve a foreign organization"
verify_IfWebElementDoesExist_byOutertext "Practice Guidelines"
verify_IfWebElementDoesExist_byOutertext "Cost Effective Analysis"

		verifyIfWebElementDoesExist2 "Contacts:"
verifyIfWebElementDoesExist2 "PI/Project Lead Name"
verifyIfWebElementDoesExist2 "Administrative Official"
verifyIfWebElementDoesExist2 "Dual PI Name"
verifyIfWebElementDoesExist2 "Dual PI Email"
verifyIfWebElementDoesExist2 "PI Designee 1"
verifyIfWebElementDoesExist2 "PI Designee 2"
verifyIfWebElementDoesExist2 "Financial Officer"

		verifyIfWebElementDoesExist2 "PI Information:"
verify_IfWebElementDoesExist_byOutertext "Primary Identification"
verifyIfWebElementDoesExist2 "Primary Group Identification\*"
verify_IfWebElementDoesExist_byOutertext "Previous involvement with PCORI\*"
verifyIfWebElementDoesExist2 "Previous involvement with PCORI - Other"
verifyIfWebElementDoesExist2 "PI Work Telephone"
verify_IfWebElementDoesExist_byOutertext "Position Title"
verifyIfWebElementDoesExist2 "Project Lead Degree\*"
verifyIfWebElementDoesExist2 "Project Lead Degree - Other"
verify_IfWebElementDoesExist_byOutertext "Relevant Exp after Terminal Degree"
verifyIfWebElementDoesExist2 "How many years of relevant experience\?"

verify_IfWebElementDoesExist_byOutertext "Field Relevant Exp"
verify_IfWebElementDoesExist_byOutertext "Grants Funded as PI"
verify_IfWebElementDoesExist_byOutertext "Largest Previous Grant/Contract Fund"
verify_IfWebElementDoesExist_byOutertext "Previous Grants/Contracts"
verify_IfWebElementDoesExist_byOutertext "Other \(please specify\)"

		verifyIfWebElementDoesExist2 "Organization Information:"
verifyIfWebElementDoesExist2 "Awardee Institution/Organization"
verifyIfWebElementDoesExist2 "D-U-N-S Number"
verifyIfWebElementDoesExist2 "Location/satellite"
verifyIfWebElementDoesExist2 "Department"
''verifyIfWebElementDoesExist2 "Congressional District"
''verifyIfWebElementDoesExist2 "Congressional District\*"
verifyIfWebElementDoesExist2 "Street Address"
verifyIfWebElementDoesExist2 "City"
verifyIfWebElementDoesExist2 "State/Province"
verifyIfWebElementDoesExist2 "Zip Code"
verifyIfWebElementDoesExist2 "Country"

		verifyIfWebElementDoesExist2 "Project Information:"
'''SPS-4416	
''verifyIfWebElementDoesExist2 "PFA"
verifyIfWebElementDoesExist2 "Primary Campaign Source"
verifyIfWebElementDoesExist2 "Cycle"
verifyIfWebElementDoesExist2 "Program"
verifyIfWebElementDoesExist2 "PFA Type"
verify_IfWebElementDoesExist_byOutertext "IHS Award Size"
verifyIfWebElementDoesExist2 "Project Name\*"
verify_IfWebElementDoesExist_byOutertext "LOI Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Invited"
verify_IfWebElementDoesExist_byOutertext "App Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Bypass Review"
verify_IfWebElementDoesExist_byOutertext "Previous Application"

verifyIfWebElementDoesExist2 "Total Direct Costs"
verifyIfWebElementDoesExist2 "Total Indirect Costs"
verifyIfWebElementDoesExist2 "LOI Amount Requested from PCORI\*"
verify_IfWebElementDoesExist_byOutertext "Project Duration"
verify_IfWebElementDoesExist_byOutertext "PCORnet involvement"
verify_IfWebElementDoesExist_byOutertext "Original PFA"
verifyIfWebElementDoesExist2 "Key Word One"
verifyIfWebElementDoesExist2 "Key Word Two"
verifyIfWebElementDoesExist2 "Key Word Three"
verifyIfWebElementDoesExist2 "Key Word Four"
verifyIfWebElementDoesExist2 "Key Word Five"
verifyIfWebElementDoesExist2 "Key Word Six"

		verifyIfWebElementDoesExist2 "Project Focus:"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Rare Disease Focus"
verify_IfWebElementDoesExist_byOutertext "Population Focus"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Focus"
verify_IfWebElementDoesExist_byOutertext "Study Design"
verify_IfWebElementDoesExist_byOutertext "Specific Analytic Methods"
verify_IfWebElementDoesExist_byOutertext "Sample Size"

verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Population Focus Other"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Other"
verify_IfWebElementDoesExist_byOutertext "Other Study Design"
	verify_IfWebElementDoesExist_byOutertext "Spec Analytic Method Other"
verifyIfWebElementDoesExist2 "Primary Patient Partner\(s\)"
verifyIfWebElementDoesExist2 "Primary Stakeholder Partner\(s\)"

		verifyIfWebElementDoesExist2 "Budget/Time Greater Than Request \(CDR and D&I\):"

verify_IfWebElementDoesExist_byOutertext "Budget/Time Greater Than Request\?"
verify_IfWebElementDoesExist_byOutertext "Submitted to other agency before"
verify_IfWebElementDoesExist_byOutertext "Agency Name"
verify_IfWebElementDoesExist_byOutertext "Considered by other agency"
verify_IfWebElementDoesExist_byOutertext "Partial funding Agency Name"
verify_IfWebElementDoesExist_byOutertext "Program Officer Name"
verify_IfWebElementDoesExist_byOutertext "May we contact PO"
verify_IfWebElementDoesExist_byOutertext "Less costly approach explanation"
verify_IfWebElementDoesExist_byOutertext "Potential for new evidence"
verify_IfWebElementDoesExist_byOutertext "Extra Money Requested"
verify_IfWebElementDoesExist_byOutertext "Extra Money Justification"
verify_IfWebElementDoesExist_byOutertext "Extra Time Requested"
verify_IfWebElementDoesExist_byOutertext "Extra Time Justification"


		verifyIfWebElementDoesExist2 "D&I:"

verify_IfWebElementDoesExist_byOutertext "Original funded application"
verify_IfWebElementDoesExist_byOutertext "Original funded cycle"
verify_IfWebElementDoesExist_byOutertext "Original app PI"
verify_IfWebElementDoesExist_byOutertext "LOI resubmission \(D&I\)"
verify_IfWebElementDoesExist_byOutertext "Was LOI invited \(D&I\)"
verify_IfWebElementDoesExist_byOutertext "Project resubmitted"
verify_IfWebElementDoesExist_byOutertext "Number of submissions"
verify_IfWebElementDoesExist_byOutertext "Summary statement received"
verify_IfWebElementDoesExist_byOutertext "Dissemination of multiple PCORI projects"
''verify_IfWebElementDoesExist_byOutertext "Name of PIs"
verify_IfWebElementDoesExist_byOutertext "Contract IDs"
verify_IfWebElementDoesExist_byOutertext "Passive dissemination"
verify_IfWebElementDoesExist_byOutertext "Develop new tool"
''new fields added in Cycle 2 2018
verify_IfWebElementDoesExist_byOutertext "LOI to the D&I Limited Competition PFA\?"
verify_IfWebElementDoesExist_byOutertext "D&I Limited Competition PFA LOI Accepted"
verify_IfWebElementDoesExist_byOutertext "Explain SDM implementation execution"

		verifyIfWebElementDoesExist2 "PPRN:"

verify_IfWebElementDoesExist_byOutertext "PPRN List"
verify_IfWebElementDoesExist_byOutertext "Partnerships"
verify_IfWebElementDoesExist_byOutertext "Total Funding from Partners"
verify_IfWebElementDoesExist_byOutertext "Linking Data Source"
verify_IfWebElementDoesExist_byOutertext "PCORnet collaboration groups"

		verifyIfWebElementDoesExist2 "LOI Pre Screen \(For Internal Use\):"
verifyIfWebElementDoesExist2 "CMA Owner"
verifyIfWebElementDoesExist2 "CMA Comments"
verifyIfWebElementDoesExist2 "Administrative Compliance Flag"
verifyIfWebElementDoesExist2 "Program Owner"
verifyIfWebElementDoesExist2 "Program Response"
verifyIfWebElementDoesExist2 "Program Non Responsiveness Decision"
verifyIfWebElementDoesExist2 "MRO Owner"
verifyIfWebElementDoesExist2 "MRO Response"
verifyIfWebElementDoesExist2 "MRO Non Responsiveness Flag"


		verifyIfWebElementDoesExist2 "Budget/Time Greater Than Request Approval \(For Internal Use\):"

verify_IfWebElementDoesExist_byOutertext "Time Approval Response"
verify_IfWebElementDoesExist_byOutertext "Time Approval"
verify_IfWebElementDoesExist_byOutertext "Budget Approval Response"
verify_IfWebElementDoesExist_byOutertext "Budget Approval"


		verifyIfWebElementDoesExist2 "Final LOI Decision \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Recommended for Alternate PFA"
verifyIfWebElementDoesExist2 "Proposed PFA"
verify_IfWebElementDoesExist_byOutertext "Alternate PFA Rationale"
verifyIfWebElementDoesExist2 "Lead Reviewer"
verifyIfWebElementDoesExist2 "Reviews Status"
verifyIfWebElementDoesExist2 "Draft Comments for Public Reply"
verifyIfWebElementDoesExist2 "Consolidated Internal Comments"
verify_IfWebElementDoesExist_byOutertext "Final Decision Status"
verifyIfWebElementDoesExist2 "Final Comments Screened"
verifyIfWebElementDoesExist2 "Final Comments for Public Reply"
verifyIfWebElementDoesExist2 "Email Proof"
verifyIfWebElementDoesExist2 "Preview Email Communications"


		verifyIfWebElementDoesExist2 "Status \(For Internal Use\):"
verifyIfWebElementDoesExist2 "LOI Number"
verifyIfWebElementDoesExist2 "LOI Status"
verifyIfWebElementDoesExist2 "External Status"
verifyIfWebElementDoesExist2 "Name"
verifyIfWebElementDoesExist2 "LOI Owner"
verifyIfWebElementDoesExist2 "LOI Review1"
verifyIfWebElementDoesExist2 "Review Roll-up 1"
verifyIfWebElementDoesExist2 "LOI Review2"
verifyIfWebElementDoesExist2 "Review Roll-up 2"
verifyIfWebElementDoesExist2 "LOI Review3"
verifyIfWebElementDoesExist2 "Review Roll-up 3"
verifyIfWebElementDoesExist2 "LOI Review4"
verifyIfWebElementDoesExist2 "Review Roll-up 4"
verifyIfWebElementDoesExist2 "Project Lead Email Address\*"
verifyIfWebElementDoesExist2 "Legal name of organization\*"
verifyIfWebElementDoesExist2 "LOI Submission Deadline Date"
'''verifyIfWebElementDoesExist2 "Custom Links"
verifyIfWebElementDoesExist2 "Created By"
verify_IfWebElementDoesExist_byOutertext "Submitted By"
verify_IfWebElementDoesExist_byOutertext "Submission Date"
verifyIfWebElementDoesExist2 "Last Modified By"
verifyIfWebElementDoesExist2 "LOI Record Type"
verifyIfWebElementDoesExist2 "Application Start Date"
	
	
	
''Bottom Part:
		verifyIfWebElementDoesExist_HtmlId "00Q1b000004Z843_00N39000003LYa3_title"
verifyIfButtonDoesExist "New Project Personnel"

		verifyIfWebElementDoesExist_HtmlId "00Q1b000004Z843_00N39000003LYZU_title"
verifyIfButtonDoesExist "New LOI Review"

		verify_IfWebElementDoesExist_byOutertext "Question Attachments"
verifyIfButtonDoesExist "New Question Attachment"

		verify_IfWebElementDoesExist_byOutertext "Notes & Attachments"
verifyIfButtonDoesExist "New Note"
verifyIfButtonDoesExist "Attach File"
verifyIfButtonDoesExist "View All"

verify_IfWebElementDoesExist_byOutertext "Lead History"

		verify_IfWebElementDoesExist_byOutertext "Activity History"
verifyIfButtonDoesExist "Log a Call"
verifyIfButtonDoesExist "Mail Merge"
verifyIfButtonDoesExist "Send an Email"

		verify_IfWebElementDoesExist_byOutertext "Campaign History"
verifyIfButtonDoesExist "Add to Campaign"

	
End Function


''Function to verify LOI Page Layout = Submitted as Science PO
Function verify_LOI_AD_pageLayout_asSciencePO()

		verifyIfButtonDoesExist editBtn
		verifyIfButtonDoesExist deleteBtn
		verifyIfButtonDoesExist sharing_btn 
		verifyIfButtonDoesExist generatePDF_RAloi
		
verifyIfWebElementDoesExist2 "LOI Detail"

		verifyIfWebElementDoesExist2 "Pre Screen Questionnaire:"
verify_IfWebElementDoesExist_byOutertext "Decision Aid"
verify_IfWebElementDoesExist_byOutertext "New Intervention"
verify_IfWebElementDoesExist_byOutertext "Date Draft Final Research Report Submit"
verifyIfWebElementDoesExist2 "Does it involve a foreign organization"
verify_IfWebElementDoesExist_byOutertext "Practice Guidelines"
verify_IfWebElementDoesExist_byOutertext "Cost Effective Analysis"

		verifyIfWebElementDoesExist2 "Contacts:"
verifyIfWebElementDoesExist2 "PI/Project Lead Name"
verifyIfWebElementDoesExist2 "Administrative Official"
verifyIfWebElementDoesExist2 "Dual PI Name"
verifyIfWebElementDoesExist2 "Dual PI Email"
verifyIfWebElementDoesExist2 "PI Designee 1"
verifyIfWebElementDoesExist2 "PI Designee 2"
verifyIfWebElementDoesExist2 "Financial Officer"

		verifyIfWebElementDoesExist2 "PI Information:"
verify_IfWebElementDoesExist_byOutertext "Primary Identification"
verifyIfWebElementDoesExist2 "Primary Group Identification\*"
verify_IfWebElementDoesExist_byOutertext "Previous involvement with PCORI\*"
verifyIfWebElementDoesExist2 "Previous involvement with PCORI - Other"
verifyIfWebElementDoesExist2 "PI Work Telephone"
verify_IfWebElementDoesExist_byOutertext "Position Title"
verifyIfWebElementDoesExist2 "Project Lead Degree\*"
verifyIfWebElementDoesExist2 "Project Lead Degree - Other"
verify_IfWebElementDoesExist_byOutertext "Relevant Exp after Terminal Degree"
verifyIfWebElementDoesExist2 "How many years of relevant experience\?"

verify_IfWebElementDoesExist_byOutertext "Field Relevant Exp"
verify_IfWebElementDoesExist_byOutertext "Grants Funded as PI"
verify_IfWebElementDoesExist_byOutertext "Largest Previous Grant/Contract Fund"
verify_IfWebElementDoesExist_byOutertext "Previous Grants/Contracts"
verify_IfWebElementDoesExist_byOutertext "Other \(please specify\)"

		verifyIfWebElementDoesExist2 "Organization Information:"
verifyIfWebElementDoesExist2 "Awardee Institution/Organization"
verifyIfWebElementDoesExist2 "D-U-N-S Number"
verifyIfWebElementDoesExist2 "Location/satellite"
verifyIfWebElementDoesExist2 "Department"
''verifyIfWebElementDoesExist2 "Congressional District"
verifyIfWebElementDoesExist2 "Congressional District\*"
verifyIfWebElementDoesExist2 "Street Address"
verifyIfWebElementDoesExist2 "City"
verifyIfWebElementDoesExist2 "State/Province"
verifyIfWebElementDoesExist2 "Zip Code"
verifyIfWebElementDoesExist2 "Country"

		verifyIfWebElementDoesExist2 "Project Information:"
'''SPS-4416		
''verifyIfWebElementDoesExist2 "PFA"
verifyIfWebElementDoesExist2 "Primary Campaign Source"
verifyIfWebElementDoesExist2 "Cycle"
verifyIfWebElementDoesExist2 "Program"
verifyIfWebElementDoesExist2 "PFA Type"
verify_IfWebElementDoesExist_byOutertext "IHS Award Size"
verifyIfWebElementDoesExist2 "Project Name\*"
verify_IfWebElementDoesExist_byOutertext "LOI Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Invited"
verify_IfWebElementDoesExist_byOutertext "App Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Bypass Review"
verify_IfWebElementDoesExist_byOutertext "Previous Application"

verifyIfWebElementDoesExist2 "Total Direct Costs"
verifyIfWebElementDoesExist2 "Total Indirect Costs"
verifyIfWebElementDoesExist2 "LOI Amount Requested from PCORI\*"
verify_IfWebElementDoesExist_byOutertext "Project Duration"
verify_IfWebElementDoesExist_byOutertext "PCORnet involvement"
verify_IfWebElementDoesExist_byOutertext "Original PFA"
verifyIfWebElementDoesExist2 "Key Word One"
verifyIfWebElementDoesExist2 "Key Word Two"
verifyIfWebElementDoesExist2 "Key Word Three"
verifyIfWebElementDoesExist2 "Key Word Four"
verifyIfWebElementDoesExist2 "Key Word Five"
verifyIfWebElementDoesExist2 "Key Word Six"

		verifyIfWebElementDoesExist2 "Project Focus:"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Rare Disease Focus"
verify_IfWebElementDoesExist_byOutertext "Population Focus"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Focus"
verify_IfWebElementDoesExist_byOutertext "Study Design"
verify_IfWebElementDoesExist_byOutertext "Specific Analytic Methods"
verify_IfWebElementDoesExist_byOutertext "Sample Size"

verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Population Focus Other"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Other"
verify_IfWebElementDoesExist_byOutertext "Other Study Design"
	verify_IfWebElementDoesExist_byOutertext "Spec Analytic Method Other"
verifyIfWebElementDoesExist2 "Primary Patient Partner\(s\)"
verifyIfWebElementDoesExist2 "Primary Stakeholder Partner\(s\)"

		verifyIfWebElementDoesExist2 "Budget/Time Greater Than Request \(CDR and D&I\):"

verify_IfWebElementDoesExist_byOutertext "Budget/Time Greater Than Request\?"
verify_IfWebElementDoesExist_byOutertext "Submitted to other agency before"
verify_IfWebElementDoesExist_byOutertext "Agency Name"
verify_IfWebElementDoesExist_byOutertext "Considered by other agency"
verify_IfWebElementDoesExist_byOutertext "Partial funding Agency Name"
verify_IfWebElementDoesExist_byOutertext "Program Officer Name"
verify_IfWebElementDoesExist_byOutertext "May we contact PO"
verify_IfWebElementDoesExist_byOutertext "Less costly approach explanation"
verify_IfWebElementDoesExist_byOutertext "Potential for new evidence"
verify_IfWebElementDoesExist_byOutertext "Extra Money Requested"
verify_IfWebElementDoesExist_byOutertext "Extra Money Justification"
verify_IfWebElementDoesExist_byOutertext "Extra Time Requested"
verify_IfWebElementDoesExist_byOutertext "Extra Time Justification"


		verifyIfWebElementDoesExist2 "D&I:"

verify_IfWebElementDoesExist_byOutertext "Original funded application"
verify_IfWebElementDoesExist_byOutertext "Original funded cycle"
verify_IfWebElementDoesExist_byOutertext "Eligible evidence"
verify_IfWebElementDoesExist_byOutertext "Original app PI"
verify_IfWebElementDoesExist_byOutertext "LOI resubmission \(D&I\)"
verify_IfWebElementDoesExist_byOutertext "Was LOI invited \(D&I\)"
verify_IfWebElementDoesExist_byOutertext "Project resubmitted"
verify_IfWebElementDoesExist_byOutertext "Number of submissions"
verify_IfWebElementDoesExist_byOutertext "Summary statement received"
verify_IfWebElementDoesExist_byOutertext "Dissemination of multiple PCORI projects"
verify_IfWebElementDoesExist_byOutertext "Name of PIs"
verify_IfWebElementDoesExist_byOutertext "Contract IDs"
verify_IfWebElementDoesExist_byOutertext "Passive dissemination"
verify_IfWebElementDoesExist_byOutertext "Develop new tool"
''new fields added in Cycle 2 2018
verify_IfWebElementDoesExist_byOutertext "LOI to the D&I Limited Competition PFA\?"
verify_IfWebElementDoesExist_byOutertext "D&I Limited Competition PFA LOI Accepted"
verify_IfWebElementDoesExist_byOutertext "Explain SDM implementation execution"

		verifyIfWebElementDoesExist2 "PPRN:"

verify_IfWebElementDoesExist_byOutertext "PPRN List"
verify_IfWebElementDoesExist_byOutertext "Partnerships"
verify_IfWebElementDoesExist_byOutertext "Total Funding from Partners"
verify_IfWebElementDoesExist_byOutertext "Linking Data Source"
verify_IfWebElementDoesExist_byOutertext "PCORnet collaboration groups"

		verifyIfWebElementDoesExist2 "LOI Pre Screen \(For Internal Use\):"
verifyIfWebElementDoesExist2 "CMA Owner"
verifyIfWebElementDoesExist2 "CMA Comments"
verifyIfWebElementDoesExist2 "Administrative Compliance Flag"
verifyIfWebElementDoesExist2 "Program Owner"
verifyIfWebElementDoesExist2 "Program Response"
verifyIfWebElementDoesExist2 "Program Non Responsiveness Decision"
'''verifyIfWebElementDoesExist2 "MRO Owner"
'''verifyIfWebElementDoesExist2 "MRO Response"
'''verifyIfWebElementDoesExist2 "MRO Non Responsiveness Flag"


		verifyIfWebElementDoesExist2 "Budget/Time Greater Than Request Approval \(For Internal Use\):"

verify_IfWebElementDoesExist_byOutertext "Time Approval Response"
verify_IfWebElementDoesExist_byOutertext "Time Approval"
verify_IfWebElementDoesExist_byOutertext "Budget Approval Response"
verify_IfWebElementDoesExist_byOutertext "Budget Approval"


		verifyIfWebElementDoesExist2 "Final LOI Decision \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Recommended for Alternate PFA"
verifyIfWebElementDoesExist2 "Proposed PFA"
verify_IfWebElementDoesExist_byOutertext "Alternate PFA Rationale"
verifyIfWebElementDoesExist2 "Lead Reviewer"
verifyIfWebElementDoesExist2 "Reviews Status"
verifyIfWebElementDoesExist2 "Draft Comments for Public Reply"
verifyIfWebElementDoesExist2 "Consolidated Internal Comments"
verify_IfWebElementDoesExist_byOutertext "Final Decision Status"
verifyIfWebElementDoesExist2 "Final Comments Screened"
verifyIfWebElementDoesExist2 "Final Comments for Public Reply"
verifyIfWebElementDoesExist2 "Email Proof"
verifyIfWebElementDoesExist2 "Preview Email Communications"


		verifyIfWebElementDoesExist2 "Status \(For Internal Use\):"
verifyIfWebElementDoesExist2 "LOI Number"
verifyIfWebElementDoesExist2 "LOI Status"
verifyIfWebElementDoesExist2 "External Status"
verifyIfWebElementDoesExist2 "Name"
verifyIfWebElementDoesExist2 "LOI Owner"
verifyIfWebElementDoesExist2 "LOI Review1"
verifyIfWebElementDoesExist2 "Review Roll-up 1"
verifyIfWebElementDoesExist2 "LOI Review2"
verifyIfWebElementDoesExist2 "Review Roll-up 2"
verifyIfWebElementDoesExist2 "LOI Review3"
verifyIfWebElementDoesExist2 "Review Roll-up 3"
verifyIfWebElementDoesExist2 "LOI Review4"
verifyIfWebElementDoesExist2 "Review Roll-up 4"
verifyIfWebElementDoesExist2 "Project Lead Email Address\*"
verifyIfWebElementDoesExist2 "Legal name of organization\*"
verifyIfWebElementDoesExist2 "LOI Submission Deadline Date"
'''verifyIfWebElementDoesExist2 "Custom Links"
verifyIfWebElementDoesExist2 "Created By"
verify_IfWebElementDoesExist_byOutertext "Submitted By"
verify_IfWebElementDoesExist_byOutertext "Submission Date"
verifyIfWebElementDoesExist2 "Last Modified By"
verifyIfWebElementDoesExist2 "LOI Record Type"
verifyIfWebElementDoesExist2 "Application Start Date"
	
	
	
''Bottom Part:
		verifyIfWebElementDoesExist_HtmlId "00Q1b000004Z843_00N39000003LYa3_title"
verifyIfButtonDoesExist "New Project Personnel"

		verifyIfWebElementDoesExist_HtmlId "00Q1b000004Z843_00N39000003LYZU_title"
verifyIfButtonDoesExist "New LOI Review"

		verify_IfWebElementDoesExist_byOutertext "Question Attachments"

		verify_IfWebElementDoesExist_byOutertext "Notes & Attachments"
verifyIfButtonDoesExist "New Note"
verifyIfButtonDoesExist "Attach File"
verifyIfButtonDoesExist "View All"

verify_IfWebElementDoesExist_byOutertext "Lead History"

		verify_IfWebElementDoesExist_byOutertext "Activity History"
verifyIfButtonDoesExist "Log a Call"
verifyIfButtonDoesExist "Mail Merge"
verifyIfButtonDoesExist "Send an Email"

		verify_IfWebElementDoesExist_byOutertext "Campaign History"
verifyIfButtonDoesExist "Add to Campaign"

	
End Function


''Function to Verify Application AD Page Layout as CMA
Function verify_APP_AD_pageLayout_top()

verifyIfButtonDoesExist editBtn
verifyIfButtonDoesExist "Clone"
verifyIfButtonDoesExist "Generate Online Critique"
verifyIfButtonDoesExist "Generate Summary Statement"

		verifyIfWebElementDoesExist2 project_Detail
		verifyIfWebElementDoesExist2 "LOI Pre Screen Questionnaire:"

verify_IfWebElementDoesExist_byOutertext "Decision Aid"
verify_IfWebElementDoesExist_byOutertext "New Intervention"
verify_IfWebElementDoesExist_byOutertext "Practice Guidelines"
verify_IfWebElementDoesExist_byOutertext "Cost Effective Analysis"



		verifyIfWebElementDoesExist2 "Authorized Users:"
verify_IfWebElementDoesExist_byOutertext "PI/Project Lead 1 Name"
verifyIfWebElementDoesExist2 "Administrative Official Name"
verifyIfWebElementDoesExist2 "PI/Project Lead Designee 1 Name"
verifyIfWebElementDoesExist2 "PI/Project Lead Designee 2 Name"
verifyIfWebElementDoesExist2 "Financial Contact"


		verifyIfWebElementDoesExist2 "PI Information:"
verify_IfWebElementDoesExist_byOutertext "Primary Group Identification\*"
verify_IfWebElementDoesExist_byOutertext "Previous involvement with PCORI\*"
verifyIfWebElementDoesExist2 "Previous involvement with PCORI - Other"
verifyIfWebElementDoesExist2 "PI Work Telephone"
verify_IfWebElementDoesExist_byOutertext "Position Title"
verifyIfWebElementDoesExist2 "Project Lead Degree\*"
verifyIfWebElementDoesExist2 "Project Lead Degree - Other"
verifyIfWebElementDoesExist2 "Dual PI Name"

verify_IfWebElementDoesExist_byOutertext "Relevant Exp after Terminal Degree"
verify_IfWebElementDoesExist_byOutertext "Field Relevant Exp"
verifyIfWebElementDoesExist2 "How many years of relevant experience\?"
verify_IfWebElementDoesExist_byOutertext "Grants Funded as PI"
verify_IfWebElementDoesExist_byOutertext "Largest Previous Grant/Contract Fund"
verify_IfWebElementDoesExist_byOutertext "Previous Grants/Contracts"
verify_IfWebElementDoesExist_byOutertext "Other \(please specify\)"
verifyIfWebElementDoesExist2 "Dual PI Email"

		verifyIfWebElementDoesExist2 "Organization Information:"
verifyIfWebElementDoesExist2 "Awardee Institution/Organization"
verifyIfWebElementDoesExist2 "D-U-N-S Number"
verifyIfWebElementDoesExist2 "Location/satellite"
verifyIfWebElementDoesExist2 "Department"
verifyIfWebElementDoesExist2 "Congressional District\*"
verifyIfWebElementDoesExist2 "Street Address"
verifyIfWebElementDoesExist2 "City"
verifyIfWebElementDoesExist2 "State/Province"
verifyIfWebElementDoesExist2 "Zip Code"
verifyIfWebElementDoesExist2 "Country"

		verifyIfWebElementDoesExist2 "Project Information:"
verifyIfWebElementDoesExist2 "Primary Campaign Source"
verifyIfWebElementDoesExist2 "Cycle"
'''verifyIfWebElementDoesExist2 "Program"
verifyIfWebElementDoesExist2 "Owning Program"
verifyIfWebElementDoesExist2 "PFA Type"
verifyIfWebElementDoesExist2 "Priority Area"
verify_IfWebElementDoesExist_byOutertext "IHS Award Size"
verifyIfWebElementDoesExist2 "Short Project Title"
verifyIfWebElementDoesExist2 "Project Name\*-off layout"
verifyIfWebElementDoesExist2 "Full Project Title"
verify_IfWebElementDoesExist_byOutertext "LOI Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Invited"
verify_IfWebElementDoesExist_byOutertext "App Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Bypass Review"
verify_IfWebElementDoesExist_byOutertext "Previous Application"
verifyIfWebElementDoesExist2 "Total Direct Costs"
verifyIfWebElementDoesExist2 "Total Indirect Costs"

verifyIfWebElementDoesExist2 "LOI Amount Requested from PCORI\*"
verify_IfWebElementDoesExist_byOutertext "Application Amount"
verify_IfWebElementDoesExist_byOutertext "Estimated Project Duration"
verifyIfWebElementDoesExist2 "Key Word One"
verifyIfWebElementDoesExist2 "Key Word Two"
verifyIfWebElementDoesExist2 "Key Word Three"
verifyIfWebElementDoesExist2 "Key Word Four"
verifyIfWebElementDoesExist2 "Key Word Five"
verifyIfWebElementDoesExist2 "Key Word Six"

		verifyIfWebElementDoesExist2 "Project Focus:"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Rare Disease Focus"
verify_IfWebElementDoesExist_byOutertext "Population Focus"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Focus"
verify_IfWebElementDoesExist_byOutertext "Study Design SC"
verify_IfWebElementDoesExist_byOutertext "Specific Analytic Methods"
verify_IfWebElementDoesExist_byOutertext "Sample Size"
verify_IfWebElementDoesExist_byOutertext "Does it involve a foreign organization"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Population Focus Other"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Primary Other"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Other"
verify_IfWebElementDoesExist_byOutertext "Other Study Design"
verify_IfWebElementDoesExist_byOutertext "Spec Analytic Method Other"


		verifyIfWebElementDoesExist2 "Budget/Time Greater Than Request:"

verify_IfWebElementDoesExist_byOutertext "Budget/Time Greater Than Request\?"
verify_IfWebElementDoesExist_byOutertext "Extra Money Requested"
verify_IfWebElementDoesExist_byOutertext "Extra Money Justification"
verify_IfWebElementDoesExist_byOutertext "Extra Time Requested"
verify_IfWebElementDoesExist_byOutertext "Extra Time Justification"


		verifyIfWebElementDoesExist2 "Project Narratives:"
verifyIfWebElementDoesExist2 "Contract Start Date\*"
verifyIfWebElementDoesExist2 "Contract End Date\*"
verifyIfWebElementDoesExist2 "Close Date"
verifyIfWebElementDoesExist2 "Technical Abstract"
verifyIfWebElementDoesExist2 "Public Abstract"
verifyIfWebElementDoesExist2 "Project Narrative"
verify_IfWebElementDoesExist_byOutertext "Study Comparators"
verifyIfWebElementDoesExist2 "Goals"
verify_IfWebElementDoesExist_byOutertext "Engagement Plan"
verify_IfWebElementDoesExist_byOutertext "Trial Arms"
verify_IfWebElementDoesExist_byOutertext "Trial Length"
verify_IfWebElementDoesExist_byOutertext "Primary Outcome"
verify_IfWebElementDoesExist_byOutertext "Secondary Outcome"
verify_IfWebElementDoesExist_byOutertext "Recruitment/Retention Plan"
verify_IfWebElementDoesExist_byOutertext "Health Systems Factors"
verify_IfWebElementDoesExist_byOutertext "Health Systems Factors Other"
verify_IfWebElementDoesExist_byOutertext "Data Sources"
verify_IfWebElementDoesExist_byOutertext "Data Sources Other"


		verifyIfWebElementDoesExist2 "PCORI Staff \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Program Officer"
verifyIfWebElementDoesExist2 "Program Associate"
verifyIfWebElementDoesExist2 "Contract Administrator"
verifyIfWebElementDoesExist2 "Panel MRO"
verifyIfWebElementDoesExist2 "Engagement Officer"


verifyIfWebElementDoesExist2 "Budget/Time Greater Than Request Approval \(For Internal Use\):"
		verifyIfWebElementDoesExist2 "Time Approval"
		verifyIfWebElementDoesExist2 "Budget Approval"


		verifyIfWebElementDoesExist2 "Application Responsiveness CMA \(For Internal Use\):"
verifyIfWebElementDoesExist2 "CMA Administratively Compliant"
verify_IfWebElementDoesExist_byOutertext "Scrub Project Personnel Contact"
verifyIfWebElementDoesExist2 "CMA Feedback"
verifyIfWebElementDoesExist2 "CMA Budget Feedback"

		verifyIfWebElementDoesExist2 "Application Responsiveness Program \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Program Decision"
verifyIfWebElementDoesExist2 "Ready for Merit Review"
verifyIfWebElementDoesExist2 "Program Response to CMA"
verifyIfWebElementDoesExist2 "Program Response to MRO"
verifyIfWebElementDoesExist2 "Science Program Justification"

		verifyIfWebElementDoesExist2 "Merit Review Responsiveness Screen \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Cost Effectiveness Analysis"
verifyIfWebElementDoesExist2 "Changed Aims/Comparators"
verifyIfWebElementDoesExist2 "Guidelines Proposed"
verifyIfWebElementDoesExist2 "Responsiveness to LOI Feedback"
verifyIfWebElementDoesExist2 "Move Forward to Online Review"
verifyIfWebElementDoesExist2 "Cost Effectiveness Comments"
verifyIfWebElementDoesExist2 "Aims/Comparators Comments"
verifyIfWebElementDoesExist2 "Guidelines Comments"
verifyIfWebElementDoesExist2 "Responsiveness Comments"
verifyIfWebElementDoesExist2 "Move Forward Comments"

		verifyIfWebElementDoesExist2 "PIR \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Flag for PIR"

		verifyIfWebElementDoesExist2 "Merit Review - Online Review \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Reviewer 1"
verifyIfWebElementDoesExist2 "Reviewer 2"
verifyIfWebElementDoesExist2 "Reviewer 3"
verifyIfWebElementDoesExist2 "Reviewer 4"
verifyIfWebElementDoesExist2 "Reviewer 5"
verifyIfWebElementDoesExist2 "Panel"
verifyIfWebElementDoesExist2 "Panel Finalized\?"
verifyIfWebElementDoesExist2 "Online Review Record Type"
verifyIfWebElementDoesExist2 "Online Review Deadline"
verifyIfWebElementDoesExist2 "Create Online Reviews"

		verifyIfWebElementDoesExist2 "Merit Review Summary \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Add to Discussion Line"
verifyIfWebElementDoesExist2 "Discussion Line Comments"
verify_IfWebElementDoesExist_byOutertext "Discussion Order Ranking"
verifyIfWebElementDoesExist2 "Application Download - MR"
verifyIfWebElementDoesExist2 "Combined Online Critique Download"
verifyIfWebElementDoesExist2 "Summary Statement Download"
verifyIfWebElementDoesExist2 "In-Person Discussion Notes"
verifyIfWebElementDoesExist2 "All Online Reviews Completed\?"
verifyIfWebElementDoesExist2 "Average Online Review Score"
verifyIfWebElementDoesExist2 "Average In-Person Score"
verifyIfWebElementDoesExist2 "Quartile Included in Summary Statement"
verifyIfWebElementDoesExist2 "Quartile"
verify_IfWebElementDoesExist_byOutertext "Re-Review"

		verifyIfWebElementDoesExist2 "Funding Slate:"
verifyIfWebElementDoesExist2 "Scores are Ready"
verifyIfWebElementDoesExist2 "Funding Slate"
verifyIfWebElementDoesExist2 "Amount Proposed to Board"
verifyIfWebElementDoesExist2 "Funding Slate Communications"
verifyIfWebElementDoesExist2 "Funding Slate Stage"
verifyIfWebElementDoesExist2 "Funding Slate Change Rationale"
verifyIfWebElementDoesExist2 "Exception"
verifyIfWebElementDoesExist2 "Exception Criteria"
verifyIfWebElementDoesExist2 "Exception Rationale"
verifyIfWebElementDoesExist2 "Exception Action"
verifyIfWebElementDoesExist2 "Selection Committee Notes"


verifyIfWebElementDoesExist2 "D&I \(For Internal Use\):"
		verify_IfWebElementDoesExist_byOutertext "Date final research report submitted"
		verify_IfWebElementDoesExist_byOutertext "Describe the results to be disseminated"
		verify_IfWebElementDoesExist_byOutertext "Background and Significance"
		verify_IfWebElementDoesExist_byOutertext "Project Aims"
		verify_IfWebElementDoesExist_byOutertext "Original app PI"
		verify_IfWebElementDoesExist_byOutertext "Briefly describe the proposed approach"
		verify_IfWebElementDoesExist_byOutertext "Briefly describe dissemination setting"
		verify_IfWebElementDoesExist_byOutertext "Describe the measurable outcomes process"
		verify_IfWebElementDoesExist_byOutertext "PCORnet involvement"
		verify_IfWebElementDoesExist_byOutertext "Describe problems project solves"
		verify_IfWebElementDoesExist_byOutertext "Describe results proposed for D&I"
		verify_IfWebElementDoesExist_byOutertext "Outcomes you hope to achieve"
		verify_IfWebElementDoesExist_byOutertext "Project importance to patients"
		verify_IfWebElementDoesExist_byOutertext "How will partners help project succeed\?"
		verify_IfWebElementDoesExist_byOutertext "Original funded cycle"
		verify_IfWebElementDoesExist_byOutertext "Original funded application"
		verify_IfWebElementDoesExist_byOutertext "Was LOI invited \(DI\)"
		verify_IfWebElementDoesExist_byOutertext "LOI resubmission \(DI\)"
		verify_IfWebElementDoesExist_byOutertext "Project resubmitted \(DI\)"
		verify_IfWebElementDoesExist_byOutertext "Number of submissions"
		verify_IfWebElementDoesExist_byOutertext "Summary Statement received"
		verify_IfWebElementDoesExist_byOutertext "Dissemination of multiple PCORI projects"
		verify_IfWebElementDoesExist_byOutertext "Name of PIs"
		verify_IfWebElementDoesExist_byOutertext "Contract IDs"
		verify_IfWebElementDoesExist_byOutertext "Passive dissemination as primary method"
		verify_IfWebElementDoesExist_byOutertext "Aims to develop/validate new tool/system"
		verify_IfWebElementDoesExist_byOutertext "Efficacy/Effectiveness of SDM strategies"
		verify_IfWebElementDoesExist_byOutertext "Disseminate funded CER/Methods study"
		verify_IfWebElementDoesExist_byOutertext "Translate/adapt shared decision making"
		verify_IfWebElementDoesExist_byOutertext "LOI to the D&I Limited Competition PFA\?"
		verify_IfWebElementDoesExist_byOutertext "D&I Limited Competition PFA LOI Accepted"
		verify_IfWebElementDoesExist_byOutertext "Explain SDM implementation execution"
		verify_IfWebElementDoesExist_byOutertext "How will you meet this requirement\?"
		verify_IfWebElementDoesExist_byOutertext "When will you have acceptance docs\?"
		verify_IfWebElementDoesExist_byOutertext "When will you submit your DFRR to PCORI\?"
		verify_IfWebElementDoesExist_byOutertext "D&I Project type"
		verify_IfWebElementDoesExist_byOutertext "Describe chosen SDM approach"
		verify_IfWebElementDoesExist_byOutertext "Projected reach of implementation"


		verifyIfWebElementDoesExist2 "PPRN \(For Internal Use\):"

verify_IfWebElementDoesExist_byOutertext "PPRN List"
verify_IfWebElementDoesExist_byOutertext "Partnerships"
verify_IfWebElementDoesExist_byOutertext "Total Funding from Partners"
verify_IfWebElementDoesExist_byOutertext "Linking Data Source"
verify_IfWebElementDoesExist_byOutertext "PCORnet collaboration groups"


		verifyIfWebElementDoesExist2 "System Information \(For Internal Use\):"
verifyIfWebElementDoesExist2 "Status"
verifyIfWebElementDoesExist2 "External Status"
verifyIfWebElementDoesExist2 "Application Number"
verifyIfWebElementDoesExist2 "Project Record Type"
verify_IfWebElementDoesExist_byOutertext "LOI Record Type"
verifyIfWebElementDoesExist2 "LOI"
verifyIfWebElementDoesExist2 "PI Approval"
verifyIfWebElementDoesExist2 "Project Owner"
verifyIfWebElementDoesExist2 "Created By"
verifyIfWebElementDoesExist2 "Last Modified By"
verifyIfWebElementDoesExist2 "Application Submission Deadline Date"
verifyIfWebElementDoesExist2 "AO Approval"

verifyIfWebElementDoesExist2 "LOI Final Comments for Public Reply"
verifyIfWebElementDoesExist2 "Submitted to other agency before"
verifyIfWebElementDoesExist2 "Agency Name"
verifyIfWebElementDoesExist2 "Partial funding Agency Name"
verifyIfWebElementDoesExist2 "Partial Funding PI Name"
verifyIfWebElementDoesExist2 "May we contact Partial Funding PI"


		verifyIfWebElementDoesExist2 "Custom Links:"
		verifyIfWebElementDoesExist2 "Applicant Attachments:"
	
	
End Function

Function verify_App_AD_PageLayout_bottomPart_asCMA()

verify_IfWebElementDoesExist_byOutertext "Budgets"
verify_IfWebElementDoesExist_byOutertext "COI & Expertise"
verify_IfWebElementDoesExist_byOutertext "Milestones - Deliverables"
	verifyIfButtonDoesExist "New Milestone - Deliverable"
	
'verifyIfWebElementDoesExist_HtmlId "0061b000003mVl3_00N70000003CCov_title"
	verifyIfButtonDoesExist "New Project Personnel"
	
verify_IfWebElementDoesExist_byOutertext "Open Activities"
	verifyIfButtonDoesExist "New Task"
	verifyIfButtonDoesExist "New Event"
	
verify_IfWebElementDoesExist_byOutertext "External Reviews"

verify_IfWebElementDoesExist_byOutertext "PIRs"
	verifyIfButtonDoesExist "New PIR"
	
verify_IfWebElementDoesExist_byOutertext "Question Attachments"
	'''verifyIfButtonDoesExist "New Question Attachment"
	
verify_IfWebElementDoesExist_byOutertext "Files"
	verifyIfButtonDoesExist "Upload Files"
	
verify_IfWebElementDoesExist_byOutertext "Notes & Attachments"
	verifyIfButtonDoesExist "New Note"
	verifyIfButtonDoesExist "Attach File"
	verifyIfButtonDoesExist "View All"
	
verify_IfWebElementDoesExist_byOutertext "Activity History"
verify_IfWebElementDoesExist_byOutertext "Project Team"
verify_IfWebElementDoesExist_byOutertext "Project Field History"
'' Sowmya confirmed that this related list was removed from Project Detail page layout - 2/15/19
'''verify_IfWebElementDoesExist_byOutertext "Audits"
'''	verifyIfButtonDoesExist "New Audit"
verify_IfWebElementDoesExist_byOutertext "MR Appeals Inquiries"
verify_IfWebElementDoesExist_byOutertext "CMA Application Reviews"
verifyIfButtonDoesExist "New CMA Application Review"
	
		
End Function

Public Function click_webElement_any(HtmlID,strName)
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
l("innertext").Value = strName
l("html id").value = HtmlID

Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click

End Function


Public Function capture_webList_allItems (HtmlID, HtmlTag,i)
On error resume next 
Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn =DeskTop.ChildObjects(btncalc)
strHwnd =  btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebList"
l("html id").value = HtmlID
l("html tag").value = HtmlTag

Set edObj =  getParentObject().ChildObjects(l)
Captured_V = edObj(i).GetRoProperty("all items")
Print Captured_V
capture_webList_allItems = Captured_V

End Function



Function create_myPendingLOIreviews_listView()
	
clk_link_Object2 tabLOIReviews


Dim splitArr()
allItems = capture_webList_allItems ("fcf", "SELECT",0)
Print "All items captured in list are - "&allItems

DsplitArr = split(allItems, ";")

For i = 0 To ubound(DsplitArr)

			If DsplitArr(i) = listviewMyPendingLOIReview Then
				Print "Value - "& DsplitArr(i)& " - already exist"
				
				Exit Function
			End If
	
Next

	'' If List View name is NOT found - then create one:
	clk_link_Object2 editBtn	
	Wait 2
	
	''Verify that you are in Edit mode
	verifyIfWebElementDoesExist2 "Step 1. Enter View Name"
	
	''Fill out View Name
	setWebEditBox_HtmlID listviewMyPendingLOIReview , "fname"
	'' Fill out View Unique Name
	setWebEditBox_HtmlID listviewMyPendingLOIReview_uniqueN, "devname"
	
	'' Fill out filters
	'' Cycle
	selectWeblist_PI_Information "Cycle FC", "fcol1"
	selectWeblist_PI_Information "equals", "fop1"
	setWebEditBox_HtmlID TestCycleR1, "fval1"
	 
	''Review Completed
	selectWeblist_PI_Information "Review Completed", "fcol2"
	selectWeblist_PI_Information "equals", "fop2"
	setWebEditBox_HtmlID "False", "fval2"

	clk_Button_usingName "Save As"	
		
End Function

Function create_myPendingLOIreviews_listView_Methods()
	
clk_link_Object2 tabLOIReviews


Dim splitArr()
allItems = capture_webList_allItems ("fcf", "SELECT",0)
Print "All items captured in list are - "&allItems

DsplitArr = split(allItems, ";")

For i = 0 To ubound(DsplitArr)

			If DsplitArr(i) = listviewMyPendingLOIReviewMethods Then
				Print "Value - "& DsplitArr(i)& " - already exist"
				
				Exit Function
			End If
	
Next

	'' If List View name is NOT found - then create one:
	clk_link_Object2 editBtn	
	Wait 2
	
	''Verify that you are in Edit mode
	verifyIfWebElementDoesExist2 "Step 1. Enter View Name"
	
	''Fill out View Name
	setWebEditBox_HtmlID listviewMyPendingLOIReviewMethods , "fname"
	'' Fill out View Unique Name
	setWebEditBox_HtmlID listviewMyPendingLOIReview_uniqueN_methods, "devname"
	
	'' Fill out filters
	'' Cycle
	selectWeblist_PI_Information "Cycle FC", "fcol1"
	selectWeblist_PI_Information "equals", "fop1"
	setWebEditBox_HtmlID TestCycleR1, "fval1"
	 
	''Review Completed
	selectWeblist_PI_Information "Review Completed", "fcol2"
	selectWeblist_PI_Information "equals", "fop2"
	setWebEditBox_HtmlID "False", "fval2"

	clk_Button_usingName "Save As"	
		
End Function


''Function to verify that Web Element with specified Class and Innertext as parameters Exists
Public Function verifyIfWebElementDoesExist_class_htmlId_innerText (classL, htmlId, innerText)
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
	l("class").value = classL
	l("html id").value = htmlId
	l("innertext").value = innerText
 
	strcount = l.count
	print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count & " - number of Web Elements"
lo(0).Highlight

				If lo.count = 0 Then      
				LogReport 1,"verifyIfWebElementDoesExist_class_htmlId - " & innerText, "WebElement --" & innerText & "- Doesn't Exist - Not Expected"
				else
				LogReport 0,"verifyIfWebElementDoesExist_class_htmlId - " & innerText, "WebElement --" & innerText & "- Exists - Expected"
				End If

End Function


Function populate_LOI_contacts(PI_user,AO_user)
On error resume next
		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("column names").value="\*PI/Project Lead Name.*"
		odesc("html tag").value = "TABLE"
		
		Set l_link = getParentObject().ChildObjects(odesc)
		print l_link.count
		a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 		  	       
				For x  = 1 to a     
						For j  =1 to b
							c = l_link(0).GetCellData(x,j)
							print "field name - "&c
							Select Case c
												    	
                                    Case "*PI/Project Lead Name"
                                         Set OList = l_link(0).ChildItem(x,j + 1, "WebList", 0)
                                         OList.select "Partner User"
                                 	     Set oEdit1 = l_link(0).ChildItem(x,j + 1, "WebEdit", 0)
                                         oEdit1.set PI_user
                                 	
                                     Case "*Administrative Official"
                                       Set OList = l_link(0).ChildItem(x,j + 1, "WebList", 0)
                                       OList.select "Partner User"                                 	   
                                       Set oEdit1 = l_link(0).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit1.set AO_user
						                                                                    
							End Select     
                         												
						Next						
				Next                               
				 
     End Function 
 
Function Close_Latest_Open_pop_up_Browser(Object)
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


Function populate_APP_contacts(PI_user,AO_user)
On error resume next
		Set odesc=Description.Create
		odesc("micclass").value="WebTable"
		odesc("column names").value="PI/Project Lead 1 Name.*"
		odesc("html tag").value = "TABLE"
		
		Set l_link = getParentObject().ChildObjects(odesc)
		print l_link.count
		a = l_link(0).GetROProperty("rows")  
				b = l_link(0).GetROProperty("cols") 		  	       
				For x  = 1 to a     
						For j  =1 to b
							c = l_link(0).GetCellData(x,j)
							print "field name - "&c
							Select Case c
												    	
                                    Case "PI/Project Lead 1 Name"
                                         Set OList = l_link(0).ChildItem(x,j + 1, "WebList", 0)
                                         OList.select "Partner User"
                                 	     Set oEdit1 = l_link(0).ChildItem(x,j + 1, "WebEdit", 0)
                                         oEdit1.set PI_user
                                 	
                                     Case "Administrative Official Name"
                                       Set OList = l_link(0).ChildItem(x,j + 1, "WebList", 0)
                                       OList.select "Partner User"                                 	   
                                       Set oEdit1 = l_link(0).ChildItem(x,j + 1, "WebEdit", 0)
                                       oEdit1.set AO_user
						                                                                    
							End Select     
                         												
						Next						
				Next                               
				 
     End Function
     
     
Function fill_out_RA_App_all_tabs_on_portal(name_Project,name_User1,RAAO_general_user,GroupExecuted)
     	
''-------------- LOG IN to Portal as RAPI
RAPI_loginPortal_goToLOIsandApps()
Wait 2

''-------------- Verify that Application is present in Draft Status and User is able to click Edit (find by unique Project Name)
verifyIfWebElementDoesExist2 name_Project

click_webElementPortal_Paper_Pencil_Icon ()
wait 2

''-------------- Verify all the Tabs and all the Fields inside Tabs are correct - MANUALLY!!!
Verify_RAappForm_Tabs()

''-------------- Fill out required fields that are missing on Contact Information  Tab
setWebEditBox_ContactInfo_Portal name_User1, RAAO_general_user, InstitOrg, "Test DIstrict", "Test Dept"
setWebEditBox_index "Financial Contact Smoke Test", 6
setWebEditBox_index "Test Department", 10
Wait 1

clk_Button_usingName saveAndNextBtn
Wait 3
''
''Fill out Pre-Screen Questionnaire Tab
''-------------- Select "No" for all the questions and click Save
selectWeblist "No,No,No,No,No"
Wait 2
clk_Button_usingName saveAndNextBtn
Wait 2

'''-------------- Select "Yes" for the first Question => verify User is able to answer the following questions
selectWeblist "No"
Wait 2
clk_Button_usingName saveAndNextBtn
Wait 4

'' Fill out PI Info Tab
fill_Out_allFields_PIinfoPage_LOIForm ()
Wait 2
clk_Button_usingName saveAndNextBtn
Wait 4

''-------------- Fill Out fields on Project Information Tab where needed
If runOn = "Methods" Then
		fill_out_Project_info_tab_RA_App_Methods name_Project
		Wait 2
	else
		fill_out_Project_info_tab_RA_App name_Project
		Wait 2
End If


clk_Button_usingName saveAndNextBtn
Wait 3

'' -------------- At this Point 1 Project Personnel record is already created from LOI phase. RAPI User able to Edit it first and then able to Delete.
If GroupExecuted = "Three Groups 1-2-3 or Only Group 3 Executed" Then
			clk_link_Object2 editBtn
			Wait 2
			setWebEditBox editedPPNrecordRAapp
			Wait 2
			clk_Button_usingName saveBtn
			Wait 2
			verifyIfWebElementDoesExist2 "NewNameLast"
			Wait 2
			
			clk_link_Object2 deleteBtn
			Wait 2
			
			Set oShell = CreateObject("WScript.Shell") 
			oShell.SendKeys "{ENTER}"
			Wait 5
			
			''-------------- User verify if record is NO longer present = Deleted successfully
			verify_webElement_NOT_exist "NewNameLast"
			
End If


''-------------- Create 1 record for Project Personnel [make small changes to email address every time - keep zz in front]          
clk_Button_usingName newBtn
setWebEditBox newPPNrecordRAapp
click_webElement_KeyPersonnel ()
selectWeblist projectPersonnel_fillOutwebLists

clk_Button_usingName saveBtn
Wait 4


clk_link_Object2 "Templates & Uploads"
Wait 3

Attach_many_Files_RAapp_Portal PeopleandPlaces, 0
clk_Button_usingName saveBtn
'''-------------- Try to Attach another file - should NOT be able to attach - SPS-3496
Attach_many_Files_RAapp_Portal attachmentAutomation, 0
verifyIfWebElementDoesExist2 "You may only upload one file per attachment field\."

			Attach_many_Files_RAapp_Portal BudgetJustification, 1
			clk_Button_usingName saveBtn
			'''-------------- Try to Attach another file - should NOT be able to attach - SPS-3496
			Attach_many_Files_RAapp_Portal attachmentAutomation, 1
			verifyIfWebElementDoesExist2 "You may only upload one file per attachment field\."

Attach_many_Files_RAapp_Portal attachmentLettersofSupport, 2
clk_Button_usingName saveBtn
'''-------------- Try to Attach another file - should NOT be able to attach - SPS-3496
Attach_many_Files_RAapp_Portal attachmentAutomation, 2
verifyIfWebElementDoesExist2 "You may only upload one file per attachment field\."

			Attach_many_Files_RAapp_Portal attachmentResearchPlan, 3
			clk_Button_usingName saveBtn
			'''-------------- Try to Attach another file - should NOT be able to attach - SPS-3496
			Attach_many_Files_RAapp_Portal attachmentAutomation, 3
			verifyIfWebElementDoesExist2 "You may only upload one file per attachment field\."

Attach_many_Files_RAapp_Portal MethodologyStandards, 4
clk_Button_usingName saveBtn
'''-------------- Try to Attach another file - should NOT be able to attach - SPS-3496
Attach_many_Files_RAapp_Portal attachmentAutomation, 4
verifyIfWebElementDoesExist2 "You may only upload one file per attachment field\."

			Attach_many_Files_RAapp_Portal MilestonesTemplate, 5
			clk_Button_usingName saveBtn
			'''-------------- Try to Attach another file - should NOT be able to attach - SPS-3496
			Attach_many_Files_RAapp_Portal attachmentPDF, 5
			verifyIfWebElementDoesExist2 "You may only upload one file per attachment field\."

Attach_many_Files_RAapp_Portal SubcontractorDetailedBudget, 6
clk_Button_usingName saveBtn
'''-------------- Try to Attach another file - should NOT be able to attach - SPS-3496
Attach_many_Files_RAapp_Portal attachmentPDF, 6
verifyIfWebElementDoesExist2 "You may only upload one file per attachment field\."

			Attach_many_Files_RAapp_Portal attachmentResubmission, 7
			clk_Button_usingName saveBtn
			'''-------------- Try to Attach another file - should NOT be able to attach - SPS-3496
			Attach_many_Files_RAapp_Portal attachmentPDF, 7
			verifyIfWebElementDoesExist2 "You may only upload one file per attachment field\."

clk_Button_usingName saveAndNextBtn
Wait 3

selectWeblist "Yes"
clk_Button_usingName saveBtn
Wait 2

     	
     	
     End Function
     
Function fill_out_Project_info_tab_RA_App(name_Project)
	
	'''Project Title
setWebEditBox name_Project

'''Projected Start Date
Click_WebEditBox_BYindex 1
Browser("micclass:=Browser").Page("micclass:=Page").WebElement("class:=day", "html tag:=TD","innertext:=24").click
Wait 2

'''Projected End Date
Click_WebEditBox_BYindex 2
Browser("micclass:=Browser").Page("micclass:=Page").WebElement("class:=day", "html tag:=TD","innertext:=27").click
Wait 2

''' Technical & Public abstract:
fillOut_App_abstractFields_IE()

'''Name the study comparators
setWebEditBox_index "TEST Name the study comparators", 3
Wait 2
'''Provide the primary outcome for your proposed study. (Identify only one outcome)
setWebEditBox_index "TEST primary outcome", 4
Wait 2
'''Provide the secondary outcome for your proposed study (if applicable)
setWebEditBox_index "TEST secondary outcome", 5
Wait 2

'''If your project is a health systems project, which factors that drive health system change will your proposal test? (Select all that apply)
selectWeblist_PI_Information_index "N/A- this proposal is not for a health systems project", 0
Wait 2
clk_Image_Link_Portal_BYindex 0
Wait 2
'''Please describe "Other" health systems factors
setWebEditBox_index "TEST Other health systems factors", 6
Wait 2
'''Does your proposed research use any of the following data sources? (Select all that apply)
selectWeblist_PI_Information_index "Medicaid data", 3
Wait 2
clk_Image_Link_Portal_BYindex 1
Wait 2
'''Please describe "Other" data sources
setWebEditBox_index "TEST Other data sources", 7
Wait 2

''' Does any portion of your study ...
selectWeblist_PI_Information_index "No", 6
''' Is the primary focus of ...
selectWeblist_PI_Information_index "No", 7

''' Total direct costs
setWebEditBox_index "23000", 8
''' Total indirect costs
setWebEditBox_index "45000", 9
''' Total amount requested
setWebEditBox_index "70000", 10

''' Please select your estimated project length (in months)
selectWeblist_PI_Information_index "40", 8

''' Primary disease or condition
selectWeblist_PI_Information_index "Kidney Disease", 9
''' Primary disease or condition focus
selectWeblist_PI_Information_index "Chronic Kidney Disease", 10
''' Please describe "Other" primary disease or condition
setWebEditBox_index "TEST Other primary disease or condition", 11

''' Secondary disease or condition
selectWeblist_PI_Information_index "Cancer", 11
''' Secondary disease or condition focus
selectWeblist_PI_Information_index "Bladder Cancer", 12
''' Please describe "Other" secondary disease or condition
setWebEditBox_index "TEST Other secondary disease or condition", 12


selectWeblist_PI_Information_allItems "Children 13-18", "N/A - this proposal does not focus on a population;Children 0-12;Children 13-18;Children 18-21;Adults 21-64;Adults >65;Disabled persons;Racial or ethnic minorities:;Residents of rural areas;Residents of urban areas;Veterans;Women;LGBTQ;Low income groups;Patients with low health literacy/numeracy and/or limited English proficiency;Individuals with multiple chronic conditions;Individuals with rare or genetic disease;Other \(please specify\)"
clk_Image_Link_Portal_BYindex 2
''' Please describe "Other" population focus
setWebEditBox_index "TEST Other population focus", 13

selectWeblist_PI_Information_allItems "Asian", "American Indian or Alaska Native;Asian;Black or African American;Hispanic/Latino;Native Hawaiian or Pacific Islander;White;Two or more races;Other \(please specify\)"
clk_Image_Link_Portal_BYindex 3
''' Please describe "Other" racial or ethnic minority focus
setWebEditBox_index "TEST Other racial or ethnic minority focus", 14


''' primary focus of your proposal? 
selectWeblist_PI_Information_index "Disparities: Other",19
setWebEditBox_index "TEST Other primary focus healthcare topic", 15

''' secondary focus of your proposal? 
selectWeblist_PI_Information_index "Disparities: Demographic", 20
setWebEditBox_index "TEST Other secondary focus healthcare topic", 16

''' Targeted sample size for main analysis
setWebEditBox_index "5600", 17


End Function


Function fill_out_Project_info_tab_RA_App_Methods(name_Project)
	
''Project Title
setWebEditBox name_Project

'''Projected Start Date
Click_WebEditBox_BYindex 1
Browser("micclass:=Browser").Page("micclass:=Page").WebElement("class:=day", "html tag:=TD","innertext:=24").click
Wait 2

'''Projected End Date
Click_WebEditBox_BYindex 2
Browser("micclass:=Browser").Page("micclass:=Page").WebElement("class:=day", "html tag:=TD","innertext:=27").click
Wait 2

''' Technical & Public abstract:
fillOut_App_abstractFields_IE()

'''Name the study comparators
setWebEditBox_index "TEST Name the study comparators", 3
Wait 2
'''Provide the primary outcome for your proposed study. (Identify only one outcome)
setWebEditBox_index "TEST primary outcome", 4
Wait 2
'''Provide the secondary outcome for your proposed study (if applicable)
setWebEditBox_index "TEST secondary outcome", 5
Wait 2

'''If your project is a health systems project, which factors that drive health system change will your proposal test? (Select all that apply)
selectWeblist_PI_Information_index "IT (New or revamped information technologies)", 0
Wait 2
clk_Image_Link_Portal_BYindex 0
Wait 2
'''Please describe "Other" health systems factors
setWebEditBox_index "TEST Other health systems factors", 6
Wait 2
'''Does your proposed research use any of the following data sources? (Select all that apply)
selectWeblist_PI_Information_index "Medical record data", 3
Wait 2
clk_Image_Link_Portal_BYindex 1
Wait 2
'''Please describe "Other" data sources
setWebEditBox_index "TEST Other data sources", 7
Wait 2

''' Does any portion of your study include collaborations with existing PCORnet entities ...
selectWeblist_PI_Information_index "No", 6


''' Please select the research area(s) of interest listed in the Methods Cycle 3 2018 PFA
selectWeblist_PI_Information_allItems "Methods Related to Ethics & Human Subjects Protections", "Methods Related to Ethics & Human Subjects Protections;Methods to Improve Study Design;Methods to Support Data Research Networks;Methods to Improve the Use of NLP;Methods - Other"
clk_Image_Link_Portal_BYindex 2

''' All the methods fields:
setWebEditBox ",,,,,,,,TEST - Methods used- One,TEST - Methods used- Two,TEST - Methods used- Three,TEST - Methods used- Four,TEST - Methods used- Five,TEST - Methods used- Six,TEST - Methods advanced - One,TEST - Methods advanced - Two,TEST - Methods advanced - Three,TEST - Methods advanced - Four"

''' Is the primary focus of your study on a rare disease? ...
selectWeblist_PI_Information_index "No", 10

''' Total direct costs
setWebEditBox_index "22000", 18
''' Total indirect costs
setWebEditBox_index "41000", 19
''' Total amount requested
setWebEditBox_index "77000", 20

''' Please select your estimated project length (in months)
selectWeblist_PI_Information_index "40", 11

''' Primary disease or condition
selectWeblist_PI_Information_index "Kidney Disease", 12
''' Primary disease or condition focus
selectWeblist_PI_Information_index "Chronic Kidney Disease", 13
''' Please describe "Other" primary disease or condition
setWebEditBox_index "TEST Other primary disease or condition", 21

''' Secondary disease or condition
selectWeblist_PI_Information_index "Cancer", 14
''' Secondary disease or condition focus
selectWeblist_PI_Information_index "Bladder Cancer", 15
''' Please describe "Other" secondary disease or condition
setWebEditBox_index "TEST Other secondary disease or condition", 22


selectWeblist_PI_Information_allItems "Adults 21-64", "N/A - this proposal does not focus on a population;Children 0-12;Children 13-18;Children 18-21;Adults 21-64;Adults >65;Disabled persons;Racial or ethnic minorities:;Residents of rural areas;Residents of urban areas;Veterans;Women;LGBTQ;Low income groups;Patients with low health literacy/numeracy and/or limited English proficiency;Individuals with multiple chronic conditions;Individuals with rare or genetic disease;Other \(please specify\)"
clk_Image_Link_Portal_BYindex 3
''' Please describe "Other" population focus
setWebEditBox_index "TEST Other population focus", 23

selectWeblist_PI_Information_allItems "American Indian or Alaska Native", "American Indian or Alaska Native;Asian;Black or African American;Hispanic/Latino;Native Hawaiian or Pacific Islander;White;Two or more races;Other \(please specify\)"
clk_Image_Link_Portal_BYindex 4
''' Please describe "Other" racial or ethnic minority focus
setWebEditBox_index "TEST Other racial or ethnic minority focus", 24


'' primary focus of your proposal? 
selectWeblist_PI_Information_index "Disparities: Other",22
setWebEditBox_index "TEST Other primary focus healthcare topic", 25

'' secondary focus of your proposal? 
selectWeblist_PI_Information_index "Disparities: Demographic", 23
setWebEditBox_index "TEST Other secondary focus healthcare topic", 26

''' Targeted sample size for main analysis
setWebEditBox_index "7600", 27
	
	
	
End Function

Function create_list_view_RA_app_methods_under_review()

' Create RA App: Under Review Methods listview with current Cycle:
''-------------- Click on Projects Tab
clk_link_Object2 tabProjects
'-------------- Select specified ListView 
selectWeblist_LOI listviewRAappUnderReviewMethods
Wait 2

'''------------------
	clk_link_Object2 editBtn	
	Wait 2
	
	''Verify that you are in Edit mode
	verifyIfWebElementDoesExist2 "Step 1. Enter View Name"
	
	''Fill out View Name
	setWebEditBox_HtmlID list_view_RA_App_Methods_Program_Review , "fname"
	
	'' Fill out View Unique Name
	setWebEditBox_HtmlID list_view_RA_App_Methods_Program_Review_u, "devname"
	
	'' Cycle
	setWebEditBox_HtmlID TestCycleR1, "fval3"
	 
	clk_Button_usingName "Save As"
	
End Function


Function create_list_view_RA_app_under_review()
	
	' Create RA App: Under Review listview with current Cycle:
''-------------- Click on Projects Tab
clk_link_Object2 tabProjects
'-------------- Select specified ListView 
selectWeblist_LOI listviewRAappUnderReviewMethods
Wait 2

'''------------------
	clk_link_Object2 editBtn	
	Wait 2
	
	''Verify that you are in Edit mode
	verifyIfWebElementDoesExist2 "Step 1. Enter View Name"
	
	''Fill out View Name
	setWebEditBox_HtmlID list_view_RA_App_Program_Review , "fname"
	
	'' Fill out View Unique Name
	setWebEditBox_HtmlID list_view_RA_App_Program_Review_u, "devname"
	
	'' Cycle
	setWebEditBox_HtmlID TestCycleR1, "fval3"
	 
	clk_Button_usingName "Save As"
	
End Function


'Function to Capture innertext of the Link based on Table Column and Index - Project Personnel Name on Application
''k -  starts with 2 Always, it is LOI Record line
''m - is Column in LOI Review Table based on PI Last Name FC = 0,  "Minus" m = Column # we want
Public Function capture_LinkText_PP_number(k, m)
On error resume next 
Set odesc=Description.Create
odesc("micclass").value="WebTable"
odesc("column names").value = "Action;Project Personnel Number;First Name;Last Name;Status;Role;Key Personnel Flag;Email;Telephone;Institution/Org"
odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
''print l_link.count


		strName = l_link(i).GetROProperty("column names")
		''print "Table Column Names: "& strName
		
		strArr = split(strName,";")
		If Not(strName = "") Then
			''print strArr(2)
		
		
		If trim(strArr(1)) = "Project Personnel Number" Then
				
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 
		''Print "b - "&b
		
		For x  = 1 to k     
				
				c = l_link(i).GetCellData(x,b-m)
				print c
				
		Next
			End if
		End If
capture_LinkText_PP_number = c

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
     
     Public Function clk_rightarrow_Link_byClass(stClass)
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
l("class").value = stClass
''l("alt").value = "Add"
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

'''''''''''''''''''''''''''''''New Function''''''''''''''''''''''''
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
''''''''''''''''''''''''''''On LOI's Open Item Tab if a specific LOi number is searched then the Index to click on the pencil and paper is = 2'''''''''''
Public Function clk_Image_PencilPapericon (stName)
Err.Clear
On error resume next

			wait 2
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
			l("image type").value = "Image Link"
			l("name").value = stName
			Set lo =  getParentObject().ChildObjects(l)
			
			lo(2).highlight
			lo(2).click
			
			If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"clk_Image_Link - " & AltImage,"Failed to click the link" & "-" & AltImage
			else
			  	LogReport 0,"clk_Image_Link - " & AltImage,"The Link" & "-" & AltImage & "-" & "is clicked successfully"
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

Public Function verify_IfLinkDoesExist_ByURL(strName, stURL)
Err.Clear
On error resume next
Wait 3

	Set btncalc = Description.Create()  
	btncalc("micclass").value = "Browser"
	
	Set btn =DeskTop.ChildObjects(btncalc)
	
	print strHwnd
	
	Set pa = Description.Create
	pa("micclass").value = "Page"
	
	Set l = Description.Create
	l("micclass").value = "Link"
	l("html tag").value = "A"
	l("url").value = stURL
	l("innertext").value = strName

Set lo =  getParentObject().ChildObjects(l)
lo(0).Highlight

				If lo.count > 0 Then   	
					LogReport 0,"verify_IfLinkDoesExist_ByURL - "&strName, "The Link --" & stURL & "- exist _ Expected"
				else
					LogReport 1,"verify_IfLinkDoesExist_ByURL - "&strName, "The Link --" & stURL & "- NOT Exists - NOT Expected"
				End If

End Function
Function selectWeblist_Resubmission2(strData)
Err.Clear
On error resume next
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

edObj(i+1).select trim(strArr(i))   

End If

Next

End Function

Function selectWeblist_Resubmission3(strData)
Err.Clear
On error resume next
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

edObj(i+2).select trim(strArr(i+1))   

End If

Next

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

Public Function verifyIfWebElementDoesExist_class_title (classL, stTitle)
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
	l("class").value = classL
	l("title").value = stTitle
 
	strcount = l.count
	print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

				If lo.count = 0 Then      
				LogReport 1,"verifyIfWebElementDoesExist_class_Innertext - " & stTitle, "WebElement --" & stTitle & "- Doesn't Exist - Not Expected"
				else
				LogReport 0,"verifyIfWebElementDoesExist_class_Innertext - " & stTitle, "WebElement --" & stTitle & "- Exists - Expected"
				End If

End Function


Function setSFLEditBox_by_Name2(Stname,strData)
     	  On error resume next
  	       	    
Set odesc=Description.Create
odesc("micclass").value="SFLEdit"
odesc("visible").value= True

odesc("name").value= Stname 
'odesc("innertext").value= strLabelName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y        

Set od =Description.Create
od("micclass").value="SFLEdit"
od("abs_x").value= x 
od("abs_y").value= y 
'           

Set Os =  getParentObject().ChildObjects(od)

print Os.count

Os(0).set strData
If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"setSFLEditBox_by_Name - " & strData,"Webedit" & "-" & strData & "-" & "is not set successfully"
			else
			  	LogReport 0,"setSFLEditBox_by_Name - " & strData,"Webedit" & "-" & strDatae & "-" & "is set successfully"
			End If
     End Function
     
     Public Function click_webEdit_box_By_placeholder2(stHolder)
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
edObj(0).Highlight
edObj(0).click
End Function

''Generic function to Capture WebElement Value by x and y coordination''''
Public Function capture_webElement_valueby_xy (stx, sty)
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
			l("x").value = stx
			l("y").value = sty
			
			Set edObj =  getParentObject().ChildObjects(l)
			Captured_V = edObj(0).GetRoProperty("innertext")
			Print Captured_V
			capture_webElement_value = Captured_V

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

Function setWebEditBoxany_by_Placeholder(stHolder ,strData)
     	  On error resume next
  	       	    
Set odesc=Description.Create
odesc("micclass").value="WebEdit"
odesc("visible").value= True

odesc("placeholder").value= stHolder 
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
     
Public Function click_webEdit_box_By_placeholder3(stHolder,i)
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

Function selectWeblistBy_accname(stName,strData)
On error resume next
wait 3

Dim Brwser
Set Brwser = Browser("micclass:=browser","creationtime:=0","Title:=.*").Page("micclass:=page","creationtime:=0","Title:=.*")
Set oDesc = Description.Create  
oDesc("micclass").value = "WebList"
'oDesc("html tag").value = "UL"
oDesc("visible").value = "True"
'oDesc("index").value = index
'oDesc("role").value = "listbox"
oDesc("acc_name").value = stName
                Brwser.WebList(oDesc).Highlight
                Brwser.WebList(oDesc).Select strData
                
Print " Select items in Weblist of  " & "-" & strData & "-" & " done successfully " 
                
                If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1," WebList select item " & strData," Failed to select items in Weblist of " & "-" & strData
else
LogReport 0,"WebList select item " & strData," Select items in Weblist of  " & "-" & strData & "-" & " done successfully"
End If
                
End Function

Public Function clk_Button_using_accname(stAccname)
    
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

Function fillOut_LOIprescreenSection_Lightning()
''''''''''''''''''This is for CMA Owner field''''
click_webEdit_box_By_placeholder "Search People\.\.\."
'''''''setWebEditBoxany_by_Placeholder "Search People\.\.\.", "Sabiha Ahad"
'''''''set sk = CreateObject("WScript.shell")
''''sk.SendKeys("{NUMLOCK}")
''''sk.SendKeys("{NUMLOCK}")
''''sk.SendKeys("{DOWN}")
''''sk.SendKeys("{ENTER}")

Browser("Sfdc Page").Page("PI Smoke Test DB | LOI").WebEdit("Search People...").Set "Sabiha Ahad"
Browser("Sfdc Page").Page("PI Smoke Test DB | LOI").WebList("dropdown-element-2574").Select "UserSabiha AhadQ/A Automation Specialist"

wait 2
''''''''''''''''''''This is for Program Owner Field''''''''
click_webEdit_box_By_placeholder "Search People\.\.\."
Browser("Sfdc Page").Page("PI Smoke Test DB | LOI").WebEdit("Search People...").Set "Sabiha Ahad"
Browser("Sfdc Page").Page("PI Smoke Test DB | LOI").WebList("dropdown-element-2585").Select "UserSabiha AhadQ/A Automation Specialist"

 ''''''''''''''''AI Automation Test'''''''''''''''''''''''''''
wait 2
'''''''''''For CMA and Program comments field''''''''''
AIUtil.SetContext Browser("creationtime:=0")
AIUtil("text_box", "CMA Comments").Click
AIUtil("text_box", "CMA Comments").Type "Test CMA Comments"

Browser("Sfdc Page").Page("PI Smoke Test DB | LOI").WebList("Administrative Compliance").Click 5,5
Browser("Sfdc Page").Page("PI Smoke Test DB | LOI").WebList("dropdown-element-6032").Select "Compliant"

AIUtil("down_triangle", micAnyText, micFromBottom, 1).Click
AIUtil("down_triangle", micAnyText, micFromBottom, 1).Click
AIUtil("down_triangle", micAnyText, micFromBottom, 1).Click
AIUtil("down_triangle", micAnyText, micFromBottom, 1).Click
AIUtil("down_triangle", micAnyText, micFromBottom, 1).Click

wait 2
AIUtil("text_box", "Program Response").CheckExists True
AIUtil("text_box", "Program Response").Type "Test Program Response"

Browser("Sfdc Page").Page("PI Smoke Test DB | LOI").WebList("Program Non Responsiveness").Click 5,5
Browser("Sfdc Page").Page("PI Smoke Test DB | LOI").WebList("dropdown-element-6042").Select "Responsive"

End Function

Function Verify_internalPage_Broad_LOI()
	'''''''''''''''''''Pre screen questionnaire Section''''''''''''''''''''
verify_IfWebElementDoesExist_byOutertext "Decision Aid Help Decision Aid No Edit Decision Aid"
verify_IfWebElementDoesExist_byOutertext "New Intervention Help New Intervention No Edit New Intervention"
verify_IfWebElementDoesExist_byOutertext "Practice Guidelines Help Practice Guidelines No Edit Practice Guidelines"
verify_IfWebElementDoesExist_byOutertext "Cost Effective Analysis Help Cost Effective Analysis No Edit Cost Effective Analysis"
verify_IfWebElementDoesExist_byOutertext "Does it involve a foreign organization No Edit Does it involve a foreign organization"


''''''''''''''''Contacts section''''''''''''''''''
'''verify_IfWebElementDoesExist_byOutertext "PI Smoke Test DB Open PI Smoke Test DB Preview Edit PI/Project Lead Name"
'''verify_IfWebElementDoesExist_byOutertext "AO Smoke Test DB Open AO Smoke Test DB Preview Edit Administrative Official"
'''verify_IfWebElementDoesExist_byOutertext "Test Dual PI Name Edit Dual PI Name"
'''verify_IfWebElementDoesExist_byOutertext "dualpitest@nomail\.com Edit Dual PI Email"
'''verify_IfWebElementDoesExist_byOutertext "PID 1 Smoke Test new DB Open PID 1 Smoke Test new DB Preview Edit PI Designee 1"
'''verify_IfWebElementDoesExist_byOutertext "PID 2 Smoke Test DB Open PID 2 Smoke Test DB Preview Edit PI Designee 2"
'''verify_IfWebElementDoesExist_byOutertext "Financial Contact Smoke Test DB Open Financial Contact Smoke Test DB Preview Edit Financial Officer"

''''''''''''''''Below Line should be uncommented when testing in prod since the contacts will be different'''''''''''''''
verify_IfWebElementDoesExist_byOutertext "PI/Project Lead Name RAPI Smoke Test 2 Open RAPI Smoke Test 2 Preview Edit PI/Project Lead Name"
verify_IfWebElementDoesExist_byOutertext "RAAO Smoke Test Open RAAO Smoke Test Preview Edit Administrative Official"
verify_IfWebElementDoesExist_byOutertext "Test Dual PI Name Edit Dual PI Name"
verify_IfWebElementDoesExist_byOutertext "dualpitest@nomail\.com Edit Dual PI Email"
verify_IfWebElementDoesExist_byOutertext "RAPID 1 Smoke Test Open RAPID 1 Smoke Test Preview Edit PI Designee 1"
verify_IfWebElementDoesExist_byOutertext "RAPID 2 Smoke Test Open RAPID 2 Smoke Test Preview Edit PI Designee 2"
verify_IfWebElementDoesExist_byOutertext "Financial Contact Smoke Test Open Financial Contact Smoke Test Preview Edit Financial Officer"


''''''''''''''''''''PI Information Section''''''''''''''''''
verify_IfWebElementDoesExist_byOutertext "Clinician Edit Primary Group Identification\*"
verify_IfWebElementDoesExist_byOutertext "Previous involvement with PCORI\* Help Previous involvement with PCORI\* Visited PCORI’s website Edit Previous involvement with PCORI\*"
verify_IfWebElementDoesExist_byOutertext "Previous involvement with PCORI - Other TEST - ""Other"" previous interactions with PCORI Edit Previous involvement with PCORI - Other"
verify_IfWebElementDoesExist_byOutertext "\(571\) 436-9999 Edit PI Work Telephone"
verify_IfWebElementDoesExist_byOutertext "TEST - Position Title Edit Position Title"
verify_IfWebElementDoesExist_byOutertext "BCH Edit Project Lead Degree\*"
verify_IfWebElementDoesExist_byOutertext "Project Lead Degree - Other TEST - other Degree Edit Project Lead Degree - Other"
verify_IfWebElementDoesExist_byOutertext "Relevant Exp after Terminal Degree Help Relevant Exp after Terminal Degree 0-4 years Edit Relevant Exp after Terminal Degree"
verify_IfWebElementDoesExist_byOutertext "How many years of relevant experience\? 3–5 years Edit How many years of relevant experience\?"
verify_IfWebElementDoesExist_byOutertext "Grants Funded as PI Help Grants Funded as PI 6-10 Edit Grants Funded as PI"
verify_IfWebElementDoesExist_byOutertext "Largest Previous Grant/Contract Fund Help Largest Previous Grant/Contract Fund \$500,000 - 1 million Edit Largest Previous Grant/Contract Fund"
verify_IfWebElementDoesExist_byOutertext "Previous Grants/Contracts Help Previous Grants/Contracts AHRQ Edit Previous Grants/Contracts"
verify_IfWebElementDoesExist_byOutertext "Other \(please specify\) Help Other \(please specify\) TEST - other organizations Edit Other \(please specify\)"

'''''''''''''''''Organization information section''''''''''''''''''''
verify_IfWebElementDoesExist_byOutertext "RTP - Test Open RTP - Test Preview Edit Awardee Institution/Organization"
verify_IfWebElementDoesExist_byOutertext "Test Dept Edit Department"
verify_IfWebElementDoesExist_byOutertext "Congressional District\* Test DIstrict Edit Congressional District\*"

'''''''''''''''''Project Information Section'''''''''''''''
verify_IfWebElementDoesExist_byOutertext "Cycle Smoke Test Cycle"
verify_IfWebElementDoesExist_byOutertext "Program Broad Pragmatic Studies"
verify_IfWebElementDoesExist_byOutertext "PFA Type Broad"
verify_IfWebElementDoesExist_byOutertext "Project Name\*"
verify_IfWebElementDoesExist_byOutertext "Yes Edit LOI Resubmission"
verify_IfWebElementDoesExist_byOutertext "Previous LOI Invited Help Previous LOI Invited Yes Edit Previous LOI Invited"
verify_IfWebElementDoesExist_byOutertext "App Resubmission Help App Resubmission Yes Edit App Resubmission"
verify_IfWebElementDoesExist_byOutertext "Yes Edit Previous LOI Bypass Review"
verify_IfWebElementDoesExist_byOutertext "Previous Application Help Previous Application 44488 Edit Previous Application"
verify_IfWebElementDoesExist_byOutertext "Total Direct Costs \$23,000 Edit Total Direct Costs"
verify_IfWebElementDoesExist_byOutertext "\$45,000 Edit Total Indirect Costs"
verify_IfWebElementDoesExist_byOutertext "LOI Amount Requested from PCORI\* \$70,000 Edit LOI Amount Requested from PCORI\*"
verify_IfWebElementDoesExist_byOutertext "Patient Care Costs No Edit Patient Care Costs"
verify_IfWebElementDoesExist_byOutertext "\$100,000 Edit Total Patient Care Costs"
verify_IfWebElementDoesExist_byOutertext "Project Duration Help Project Duration 30 Edit Project Duration"
verify_IfWebElementDoesExist_byOutertext "Yes Edit PCORnet involvement"
verify_IfWebElementDoesExist_byOutertext "PCORI Research Priority Area/Key Topic 4\) Rare Disease Edit PCORI Research Priority Area/Key Topic"
verify_IfWebElementDoesExist_byOutertext "Achieve Health Equity Edit National Priorities primary BPS"
verify_IfWebElementDoesExist_byOutertext "Increase Evidence for Existing Interventions and Emerging Innovations in Health Edit National Priorities secondary BPS"
verify_IfWebElementDoesExist_byOutertext "Accelerate Progress Toward an Integrated Learning Health System Edit National Priorities tertiary BPS"
verify_IfWebElementDoesExist_byOutertext "Promoting health for older adults Edit Topic Themes Primary"
verify_IfWebElementDoesExist_byOutertext "Promoting healthy children and youth Edit Topic Themes Secondary"
verify_IfWebElementDoesExist_byOutertext "Improving cardiovascular health Edit Topic Themes Tertiary"
verify_IfWebElementDoesExist_byOutertext "BPS Categories Category 3 \(PCORnet® Study\) Edit BPS Categories"
verify_IfWebElementDoesExist_byOutertext "No Edit PCORnet Front Door"
verify_IfWebElementDoesExist_byOutertext "GPC Edit PCORnet ID Network"
verify_IfWebElementDoesExist_byOutertext "Yes Edit Address SAE"
verify_IfWebElementDoesExist_byOutertext "C3 BPS: Long COVID Edit Which SAE"

''''''''''''''Project Focus section'''''''''''''''''

verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Help Primary Disease or Condition Kidney Disease Edit Primary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Focus Help Primary Disease or Condition Focus Other Edit Primary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Cancer Edit Secondary Disease or Condition"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Focus Help Secondary Disease or Condition Focus Bladder Cancer Edit Secondary Disease or Condition Focus"
verify_IfWebElementDoesExist_byOutertext "Rare Disease Focus Help Rare Disease Focus No Edit Rare Disease Focus"
verify_IfWebElementDoesExist_byOutertext "Children 13-18 Edit Population Focus"
verify_IfWebElementDoesExist_byOutertext "Racial/Ethnic Minority Focus Help Racial/Ethnic Minority Focus Asian Edit Racial/Ethnic Minority Focus"
verify_IfWebElementDoesExist_byOutertext "Disparities: Other Edit Healthcare Primary Focus"
verify_IfWebElementDoesExist_byOutertext "Healthcare Secondary Focus Help Healthcare Secondary Focus Disparities: Demographic Edit Healthcare Secondary Focus"
verify_IfWebElementDoesExist_byOutertext "2,500 Edit Sample Size"
verify_IfWebElementDoesExist_byOutertext "Primary Disease or Condition Other Help Primary Disease or Condition Other Test Other primary disease or condition Edit Primary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Secondary Disease or Condition Other Help Secondary Disease or Condition Other Test Other secondary secondary disease or condition Edit Secondary Disease or Condition Other"
verify_IfWebElementDoesExist_byOutertext "Test Other population population focus Edit Population Focus Other"
verify_IfWebElementDoesExist_byOutertext "Test Other racial racial or ethnic minority focus Edit Racial/Ethnic Minority Other"
verify_IfWebElementDoesExist_byOutertext "Test Other primary focus healthcare topic Edit Healthcare Primary Other"
verify_IfWebElementDoesExist_byOutertext "Test Other secondary focus healthcare topic Edit Healthcare Secondary Other"
End Function

Public Function clk_SFLButton_usingName2(strName)
    
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
     print editObject.count

			If editObject.count = 0 Then
				LogReport 1,"clk_Button_usingName - " & strName,"Button with name" & "-" & strName &"NOT found - NOT expected"
			else
				LogReport 0,"clk_Button_usingName - " & strName,"Button with name" & "-" & strName & "-" & "is FOUND succesfully"
			End If

editObject(0).click
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
'edObj(i).Highlight
edObj(i).click

  If err <> 0 Then
  	  LogReport 4,"Error", err.number & "-" & err.description
  	  LogReport 1,"cliking the button" & stText,"Failed to click the button" & "-" & stText
  	  else
  	  LogReport 0,"cliking the button" & stText,"The button" & "-" & stText & "-" & "is clicked succesfully"
    End If
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

Public Function verifyIfWebElementDoesExist_byIndex(innertext,stTag,i)
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
	l("innertext").value = innertext
	l("html tag").value = stTag
	l("visible").value = True 
	
Set lo = getParentObject().ChildObjects(l)
Print lo.count
lo(i).Highlight

				If lo.count = 0 Then      
					LogReport 1,"verifyIfWebElementDoesExist_byIndex - " & innertext, "WebElement --" & innertext & "- Doesn't Exist - Not Expected"
				else
					LogReport 0,"verifyIfWebElementDoesExist_byIndex - " & innertext, "WebElement --" & innertext & "- Exists - Expected"
				End If

End Function

Function setSFLEditbox_by_index(strData, i)
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

editO("micclass").value = "SFLEdit"
editO("visible").value = True  
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

Function setSFLEditBox_by_Name(Stname,strData)
     	  On error resume next
  	       	    
Set odesc=Description.Create
odesc("micclass").value="SFLEdit"
odesc("visible").value= True

odesc("name").value= Stname 
'odesc("innertext").value= strLabelName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y        

Set od =Description.Create
od("micclass").value="SFLEdit"
od("abs_x").value= x 
od("abs_y").value= y 
'           

Set Os =  getParentObject().ChildObjects(od)

print Os.count

Os(0).set strData
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


Function set_SFLEditBoxany_by_Placeholder(stHolder ,strData)
     	  On error resume next
  	       	    
Set odesc=Description.Create
odesc("micclass").value="SFLEdit"
odesc("visible").value= True

odesc("placeholder").value= stHolder 
'odesc("innertext").value= strLabelName

Set O =  getParentObject().ChildObjects(odesc)
print O.count
x = O(0).getRoProperty("abs_x")
print x 
y= O(0).getRoProperty("abs_y")
print y        

Set od =Description.Create
od("micclass").value="SFLEdit"
od("abs_x").value= x 
od("abs_y").value= y 
'           

Set Os =  getParentObject().ChildObjects(od)

print Os.count

Os(0).set strData
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
Function setSFLEditbox_by_index_andTag(strData,stTag, i)
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

editO("micclass").value = "SFLEdit"
editO("class").value = "slds-textarea"
editO("visible").value = True  
editO("html tag").value = stTag
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 


edObj(i).highlight
edObj(i).set strData 
'oShell.SendKeys strSearchText
'Print " Entering the data " & "-" & strData & "-" & "is successful "				            

If err <> 0 Then
LogReport 4,"Error", err.number & "-" & err.description
LogReport 1," WebEdit box by Index value " & strData," Failed to enter the data" & "-" & strData
else
LogReport 0,"WebEdit box by Index value " & strData," Entering the data " & "-" & strData & "-" & "is successful"
End If
End Function

Public Function click_on_AppLauncher (strName)
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
l("class").value = "slds-icon-waffle"
'l("outerhtml").value = "<span class=""listTitle"">Reviews<span class=""count"">\[5\]</span></span>"
l("visible").value = True
l("html tag").value = "DIV"
l("innertext").Value = strName
'l("html id").value = HtmlID
Set edObj =  getParentObject().ChildObjects(l)
edObj(i).Highlight
edObj(i).click
End Function

Function ClosepreviousOpenedBrowser()
On error resume next
Dim oDescription
Dim BrowserObjectList
Dim oLatestBrowserIndex

Set oDescription=Description.Create
oDescription("micclass").value="Browser"
Set BrowserObjectList=Desktop.ChildObjects(oDescription)
''''''''Note for myself''''''''''oLatestBrowserIndex=BrowserObjectList.count 

Browser("creationtime:=0").close

Set oDescription=Nothing
Set BrowserObjectList=Nothing

End  Function                                       	
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
     
Public Function clk_weblist_using_data_Index(Htmlid,i)
    
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
     editO("html id").value = Htmlid
     Set editObject =  getParentObject().ChildObjects(editO) 
     editObject(i).highlight
     editObject(i).click
    
    If err <> 0 Then
  	  LogReport 4,"Error", err.number & "-" & err.description
  	  LogReport 1,"cliking the button" & strName,"Failed to click the button" & "-" & strName
  	  else
  	  LogReport 0,"cliking the button" & strName,"The button" & "-" & strName & "-" & "is clicked succesfully"
    End If
End Function

Public Function verifyIfWebEdit_Exists_default_Value(stValue, stTag)
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
	l("micclass").value = "WebEdit"
	l("default value").value = stValue
	l("html tag").value = stTag
 
	strcount = l.count
	print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight

				If lo.count = 0 Then      
				LogReport 1,"verifyIfWebEdit_Exists_default_Value - " & stValue, "WebElement --" & stValue & "- Doesn't Exist - Not Expected"
				else
				LogReport 0,"verifyIfWebEdit_Exists_default_Value - " & stValue, "WebElement --" & stValue & "- Exists - Expected"
				End If

End Function

Public Function verify_WebListDoes_Exist_by_Default_Value (HtmlID,stValue,i)
Err.Clear
On error resume next

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn = DeskTop.ChildObjects(btncalc)
strHwnd = btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebList"
l("html tag").value = "SELECT"
l("default value").value = stValue 
l("html id").value = HtmlID

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(i).Highlight

If lo.count = 0 Then
LogReport 1,"verify_WebListDoes_Exist_by_Value - " & HtmlID, "The Value --" & stValue & "- does NOT exist - Not Expected"
else
LogReport 0,"verify_WebListDoes_Exist_by_Value - " & HtmlID,  "The Value --" & stValue & "-Exists - Expected"
End If
End Function
Public Function verify_WebListDoes_Exist_by_Value (HtmlID,stValue,i)
Err.Clear
On error resume next

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn = DeskTop.ChildObjects(btncalc)
strHwnd = btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebList"
l("html tag").value = "SELECT"
l("value").value = stValue 
l("html id").value = HtmlID

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(i).Highlight

If lo.count = 0 Then
LogReport 1,"verify_WebListDoes_Exist_by_Value - " & HtmlID, stValue & "The List --" & HtmlID & "- does NOT exist - Not Expected"
else
LogReport 0,"verify_WebListDoes_Exist_by_Value - " & HtmlID, stValue &  "The List --" & HtmlID & "-Exists - Expected"
End If
End Function
''''''''''''''num 1 = disabled and num 0 = not disabled
Public Function verify_WebList_is_Disabled (HtmlID,num0or1)
Err.Clear
On error resume next

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn = DeskTop.ChildObjects(btncalc)
strHwnd = btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebList"
l("html tag").value = "SELECT"
l("disabled").value = num0or1
l("html id").value = HtmlID

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(0).Highlight
If lo.count = 0 Then
LogReport 1,"verify_WebListDoes_disabled - " & HtmlID, "The List --" & HtmlID & "- disabled  - not Expected"
else
LogReport 0,"verify_WebListDoes__disabled - " & HtmlID, "The List --" & HtmlID & "-disabled  - is Expected"
End If
End Function
Function Verify_allFields_PIinfoPage_LOIForm_DI()
verifyIfWebElementDoesExist2 Instruction_Text1
verifyIfWebElementDoesExist2 Instruction_Text2
verifyWebElement_By_Outertext PI_Tab_InstQs_1_DI
verify_IfLinkDoesExist_ByURL "here", Link1URL
verify_IfLinkDoesExist_ByURL "here", Link2URL
verifyWebElement_By_Outertext "\* PI Work Telephone " ''''''PI_Work_Telephone
verifyWebElement_By_Outertext Primary_Group_IdentificationDI
verifyWebElement_By_Outertext Previous_involvement_PCORI_DI
verifyIfWebElementDoesExist2 "Please describe ""Other"" previous interactions with PCORI"
verifyWebElement_By_Outertext Position_Title
verifyWebElement_By_Outertext Degree_DI 
verifyWebElement_By_Outertext Degree_Other_DI
verifyWebElement_By_Outertext Relevant_Exp_Terminal_Degree_DI
verifyWebElement_By_Outertext Years_of_relevant_experience_DI
verifyWebElement_By_Outertext Grants_Funded_as_PI_DI
verifyWebElement_By_Outertext Contract_Fund 
verifyWebElement_By_Outertext Previous_Grants_Contracts_DI 
verifyIfWebElementDoesExist2 "Please describe ""Other"" organizations from which you have received grants/contracts"
End Function

Function create_newProjectPersonnel_App_DI()

	'''-------------- Create 1 record for Project Personnel
clk_Button_usingName newBtn

''''-------------- Verify Fields on New Project Personnel Page
wait 2
verifyIfWebElementDoesExist2 Instruction_Text1
verifyIfWebElementDoesExist2 Instruction_Text2
verifyWebElement_By_Outertext Pro_Personnel_InText_app_DI_childrecord
verifyWebElement_By_OuterHTML "B", Pro_Personnel_BoldedText_app_DI_childrecord

verify_IfLinkDoesExist_ByURL "here", Link1URL
verify_IfLinkDoesExist_ByURL "here", Link2URL

clk_link_byName_Index "here", 0
wait 2
CloseLatestOpenedBrowser()
wait 2
clk_link_byName_Index "here", 1
wait 2
CloseLatestOpenedBrowser()
verifyWebElement_By_Outertext "\* First Name "
verifyWebElement_By_Outertext "\* Last Name "
verifyWebElement_By_Outertext "Institution or Organization "
verifyWebElement_By_Outertext "\* Primary perspective on the project team  --None-- Patient Stakeholder Scientific "
verifyWebElement_By_Outertext "\* Project Role --None-- PI/Project Lead 1 Name Dual PI Co-PI Co-Investigator Stakeholder Partner Patient Partner Project Manager Other "
verifyWebElement_By_Outertext "Please describe ""Other"" role "
verifyWebElement_By_Outertext "\* For the purposes of this project, which of the following patient or stakeholder communities reflects this person's primary affiliation\? --None-- Caregiver/Family member of patient Patient/Caregiver Advocacy Organization Clinician Clinic/Hospital/Health System Purchaser Payer Industry Research Policy Maker Training Institution N/A "
verifyWebElement_By_Outertext "Degrees AAS AB APRN BA BC BCH BCHIR BM BMBC BMEDS BPHAR BS BSC BSN CHB CHM DA DBA DC DCH DDS DES DM DMD DMH DMP DMS DNSC DO DPH DPHIL DSC DVM EDD HS JD LPN MA MB MBA MBBC MBCHB MCHIR MD MED MIS MLIS MN MPA MPH MPHIL MS MSN MSURG MSW ND NN NP Other \(please specify\) PA PHARMD PHD PTA RN SB SCD "
verifyWebElement_By_Outertext "Please describe other degree 0 of 255 Characters "
verifyWebElement_By_Outertext "\* Phone "
verifyWebElement_By_Outertext "\* Email "
verifyWebElement_By_Outertext "Key Personnel Yes No "
wait 2
setWebEditBox projectPersonnel_firstRecord
click_webElement_KeyPersonnel ()
selectWeblist projectPersonnel_fillOutwebLists

'' Click on arrow to the right to ADD Degree
clk_Image_Link_Portal_BYindex 0
Wait 1
clk_Button_usingName saveBtn
Wait 2

End Function

Public Function verify_BudgetTab_weblist_Exist()
Err.Clear
On error resume next

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn = DeskTop.ChildObjects(btncalc)
strHwnd = btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebList"
l("html tag").value = "SELECT"
l("all items").value = "Project Personnel;Consultant Cost;Supplies;Programmatic Travel;Other Expenses;Equipment;Subcontractor Direct;Subcontractor Indirect;Total Prime Indirect;Budget Summary"
l("name").value ="j_id0:budgetForm:j_id10"

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count

If lo.count = 0 Then
LogReport 1,"verify_WebListDoes_Exist_by_Value - " & StValue & "The List --" & HtmlID & "- does NOT exist - Not Expected"
else
LogReport 0,"verify_WebListDoes_Exist_by_Value - " & StValue &  "The List --" & HtmlID & "-Exists - Expected"
End If
End Function

Public Function clk_Image_PencilPapericon_AO_Ammendent_record(stName,stfilename,Index)

Err.Clear
On error resume next

			wait 2
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
			l("image type").value = "Image Link"
			l("name").value = stName
			l("file name").value = stfilename
			Set lo =  getParentObject().ChildObjects(l)
			
			lo(i).highlight
			lo(i).click
			
			If err <> 0 Then
				LogReport 4,"Error", err.number & "-" & err.description
				LogReport 1,"clk_Image_Link - " & AltImage,"Failed to click the link" & "-" & AltImage
			else
			  	LogReport 0,"clk_Image_Link - " & AltImage,"The Link" & "-" & AltImage & "-" & "is clicked successfully"
			End If
End Function 

Sub click_magnifyinglass_app(strStatus)
Err.Clear
On error resume next
Wait 3

	Set odesc=Description.Create
	odesc("micclass").value="WebElement"
	odesc("visible").value= True
	odesc("html tag").value= "TD"
	odesc("innertext").value= strStatus
	
	Set O =  getParentObject().ChildObjects(odesc)
	print O.count
	x = O(0).getRoProperty("abs_x")
	print x + 1
	y= O(0).getRoProperty("abs_y")
	print y + 1  
	
	Set od=Description.Create
	od("micclass").value="WebElement"
	od("class").value = "slds-truncate"
	od("abs_y").value= y

Set Ob =  getParentObject().ChildObjects(od)
print Ob.count

				If Ob.count = 0 Then
					LogReport 1,"click_magnifyinglass "&"magnifying glass", "magnifying glass" & "NOT Found - NOT Expected"
				else
					LogReport 0,"click_magnifyinglass "&"magnifying glass","magnifying glass" & "Found - Expected"				
				End If

Ob(0).click

End Sub

Function waitForwebElement(strName)
Err.Clear
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

editO("micclass").value = "WebElement"
editO("visible").value = true  
editO("Innertext").value = strText
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

Print "Waiting for Element " & "-" & strText& "-" &"Time elapsed is " & TimeDiff
'Logger Information: 
'                                                     ' "INFO","WaitForNextSpanObject","Waiting for " & "-" & strWebName & "-" &"Time elapsed is " &TimeDiff
'exit the loop if the object is not found for more than a minute
If  TimeDiff> 20 Then
'Log the results
Print "WebElement" & "-" & strText & "was NOT found"

Exit Do
End If    

If status >  0 Then
'Log the results   
wait 1                
print "WebElement was found"                    
status = 1

End If
Loop  
waitForwebElement =  TimeDiff       
End function
Public Function verifyWebElement_By_OuterHTML_And_Index(stTag,Stouterhtml,i)
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
lo(i).Highlight

				If lo.count = 0 Then      
					LogReport 1,"verifyWebElement_By_OuterHtml by index - " & Stouterhtml, "WebElement --" & Stouterhtml & "- Doesn't Exist - Not Expected"
				else
					LogReport 0,"verifyWebElement_By_OuterHtml by index " & Stouterhtml, "WebElement --" & Stouterhtml & "- Exists - Expected"
				End If

End Function

Public Function capture_webEdit_value (HtmlID)
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
			l("html id").value = HtmlID
			'l("html tag").value = HtmlTag
			
			Set edObj =  getParentObject().ChildObjects(l)
			Captured_V = edObj(0).GetRoProperty("value")
			Print Captured_V
			capture_webEdit_value = Captured_V

End Function

Function setWebEditBox_byHtId(HtmlID, strData)
On error resume next

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
editO("html id").value = HtmlID
''editO("html tag").value = "TEXTAREA"

Set edObj = getParentObject().ChildObjects(editO) 

edObj(0).set strData 

If err <> 0 Then
err.clear   

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
Public Function navigateAndLoginToSalesForce_Prod(strUserName,strPassword)
Err.Clear
On error resume next
Wait 3

refresh_Chrome_browser()
Open_SalesForce_Application_Prod
wait 2

'Enter valid ID and password.
''login_intoSalesForce_Application strUserName, strPassword 
set_WebEditBox_by_Name "username",UserAdmin
setWebEditBox_by_Name "pw",strPassword
'Click “Log in” button.
clk_Button_usingName logInButton_prod

				If err <> 0 Then
					LogReport 4,"Error", err.number & "-" & err.description
					LogReport 1,"navigateAndLoginToSalesForce","Was not able to log in"
				else
					LogReport 0,"navigateAndLoginToSalesForce","Was able to log in"
				End If

End Function 

Function capture_LOIreview_numbers_SOE()

captureLinkText_LOIReview_Number 2,9
ReviewNum1 = captureLinkText_LOIReview_Number(2,9)
Print "Here is the review number 1: " & ReviewNum1
write_ToAfile_anyFilePath_fileName TextfilePathForReviewer1_Number_SOE, ReviewNum1
Wait 3

captureLinkText_LOIReview_Number 3,9
ReviewNum2 = captureLinkText_LOIReview_Number(3,9)
Print "Here is the review number 2: " & ReviewNum2
write_ToAfile_anyFilePath_fileName TextfilePathForReviewer2_Number_SOE, ReviewNum2
Wait 3

End Function

Function create_newProjectPersonnel_App_Broard_Targeted()

	'''''''''''''' Create 1 record for Project Personnel
clk_Button_usingName newBtn

''''-------------- Verify Fields on New Project Personnel Page
wait 2
verifyIfWebElementDoesExist2 Instruction_Text1
verifyIfWebElementDoesExist2 Instruction_Text2
verifyWebElement_By_Outertext Pro_Personnel_InText_app_DI_childrecord
verifyWebElement_By_OuterHTML "B", Pro_Personnel_BoldedText_app_DI_childrecord

verify_IfLinkDoesExist_ByURL "here", Link1URL
verify_IfLinkDoesExist_ByURL "here", Link2URL

clk_link_byName_Index "here", 0
wait 2
CloseLatestOpenedBrowser()
wait 2
clk_link_byName_Index "here", 1
wait 2
CloseLatestOpenedBrowser()
wait 2
verifyWebElement_By_Outertext "\* First Name "
verifyWebElement_By_Outertext "\* Last Name "
verifyWebElement_By_Outertext "Institution or Organization "
verifyWebElement_By_Outertext "\* Primary perspective on the research team  --None-- Patient Stakeholder Scientific "
verifyWebElement_By_Outertext "\* Project Role --None-- PI/Project Lead 1 Name Dual PI Co-PI Co-Investigator Stakeholder Partner Patient Partner Project Manager Other "
verifyWebElement_By_Outertext "Please describe ""Other"" role "
verifyWebElement_By_Outertext "\* For the purposes of this project, which of the following patient or stakeholder communities reflects this person's primary affiliation\? --None-- Caregiver/Family member of patient Patient/Caregiver Advocacy Organization Clinician Clinic/Hospital/Health System Purchaser Payer Industry Research Policy Maker Training Institution N/A "
verifyWebElement_By_Outertext "Degrees AAS AB APRN BA BC BCH BCHIR BM BMBC BMEDS BPHAR BS BSC BSN CHB CHM DA DBA DC DCH DDS DES DM DMD DMH DMP DMS DNSC DO DPH DPHIL DSC DVM EDD HS JD LPN MA MB MBA MBBC MBCHB MCHIR MD MED MIS MLIS MN MPA MPH MPHIL MS MSN MSURG MSW ND NN NP Other \(please specify\) PA PHARMD PHD PTA RN SB SCD "
verifyWebElement_By_Outertext "Please describe other degree 0 of 255 Characters "
verifyWebElement_By_Outertext "\* Phone "
verifyWebElement_By_Outertext "\* Email "
verifyWebElement_By_Outertext "Key Personnel Yes No "
wait 2
setWebEditBox projectPersonnel_firstRecord
click_webElement_KeyPersonnel ()
selectWeblist projectPersonnel_fillOutwebLists

'' Click on arrow to the right to ADD Degree
clk_Image_Link_Portal_BYindex 0
Wait 1
clk_Button_usingName saveBtn
Wait 2

End Function

Function Verify_LOIForm_Tabs_EA_DI()

verifyIfLinkDoesExist1 "Project Name & Contact Information"
verifyIfLinkDoesExist1 "Pre-Screen Questionnaire"
verifyIfLinkDoesExist1 "Organization & Project Lead Details"
verifyIfLinkDoesExist1 "Project Summary"
verifyIfLinkDoesExist1 "Additional Project Information"
verifyIfLinkDoesExist1 "Using PCORI-funded Evidence & Tools"
verifyIfLinkDoesExist1 "Attachments"
verifyIfLinkDoesExist1 "Authorizations"

End Function

Function Verify_LOIForm_Tabs_EA_SCS_CB()

verifyIfLinkDoesExist1 "Project Name & Contact Information"
verifyIfLinkDoesExist1 "Pre-Screen Questionnaire"
verifyIfLinkDoesExist1 "Organization & Project Lead Details"
verifyIfLinkDoesExist1 "Project Summary"
verifyIfLinkDoesExist1 "Additional Project Information"
verifyIfLinkDoesExist1 "Using PCORI-funded Evidence & Tools"
'''''''''verifyIfLinkDoesExist1 "Attachments"             '''''''''Out of scope for SCS and CB PFA'''''''''
verifyIfLinkDoesExist1 "Authorizations"

End Function

Function setWebEditBox_ContactInfo_Portal_EA_LOI(Org, RAPI, RAAO, RAPI2, PID1, PID2)
On error resume next

wait 3
strArr = split(RAPI,",")

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
'editO("html id").value = "j_id0:mainForm:j_id360:5:inputFieldId" 
'editO("type").value = "text" 
Set edObj =  getParentObject().ChildObjects(editO) 

for i = 0 To uBound(strArr)
'id =   edObj(i).GetRoProperty("html id")
items = edObj(i).GetRoProperty("all items")
print items

print RAPI(i)
If Not(RAPI(i) = "")Then

'edObj(i).set strArr(i) 
''edObj(i).set 
edObj(i+1).set Org
edObj(i+2).set RAPI
edObj(i+3).set RAAO
edObj(i+4).set RAPI2
edObj(i+5).set PID1
edObj(i+6).set PID2
'edObj(i+6).set FCofficer
'edObj(i+7).set Org
'edObj(i+8).set Distr
'edObj(i+9).set Dept

oShell.SendKeys strSearchText				            
End If

Next
If err <> 0 Then
err.clear   

End If
End Function
Function fill_Out_allFields_PIinfoPage_HSII_LOIForm ()
selectWeblist "Yes"
'''selectWeblist "Clinician"
selectWeblist_PI_Information "Payer", "j_id0:j_id2:j_id3:mainForm:j_id713:4:inputFieldId"
'''selectWeblist_PI_Information_index "Payer", 2
'''selectWeblist_PI_Information_allItems "0-4 years","--None--;0-4 years;5-9 years;10\+ years"

selectWeblist_PI_Information_allItems "3–5 years","--None--;0–2 years;3–5 years;6–9 years;10–15 years;16 years \+"

selectWeblist_PI_Information_allItems "6-10","--None--;0;1-5;6-10;11-15;16-20;21-25;26 or greater"

'''''selectWeblist_PI_Information_allItems "$500,000 - 1 million","--None--;N/A;Less than \$500,000;\$500,000 - 1 million;\$1\.1 - 5 million;\$5\.1 - 10 million;Greater than \$10 million"

setWebEditBox "571-436-9999,Test - Position Title,Test - Other Degree,Test - Other Organizations"
Wait 2
''''selectWeblist_PI_Information_allItems "Visited PCORI’s website", "Joined a PCORI email list;Visited PCORI’s website;Participated in applicant training;Watched a PCORI webinar;Attended PCORI sponsored event in-person;Attended event where PCORI was featured;Met with PCORI staff;Met with a PCORI Ambassador;Applied to review PCORI funding app;Applied for PCORI funding;Received PCORI funding;Served as a PCORI Merit reviewer;Participated in a PCORI Advisory Panel;Other \(please specify\);None of the above"

'''clk_Image_Link_Portal_BYindex 0

selectWeblist_PI_Information_allItems "BCH", "AAS;AB;APRN;BA;BC;BCH;BCHIR;BM;BMBC;BMEDS;BPHAR;BS;BSC;BSN;CHB;CHM;DA;DBA;DC;DCH;DDS;DES;DM;DMD;DMH;DMS;DNSC;DO;DPH;DPHIL;DMP;DSC;DVM;EDD;HS;JD;LPN;MA;MB;MBA;MBBC;MBCHB;MCHIR;MD;MED;MIS;MLIS;MN;MPA;MPH;MPHIL;MS;MSN;MSURG;MSW;ND;NN;NP;PA;PHD;PHRMD;PTA;RN;SB;SCD;Other \(please specify\)"
clk_Image_Link_Portal_BYindex 0

selectWeblist_PI_Information_allItems "AHRQ", "PCORI;AHRQ;CDC;NIH;RWJF;Other \(please specify\);None of the above"

clk_Image_Link_Portal_BYindex 1
Wait 2

End Function

Public Function captureLinkText_QuestionAttachment_AppEA(m)
On error resume next 

Set odesc=Description.Create
odesc("micclass").value="WebTable"
'odesc("column names").value = "Action;Description;Question Attachment Name"
odesc("column names").value =";Choose a RowSelect All;Sort by:Question Attachment NameSorted AscendingShow actions;Sort by:Question TextSorted: NoneShow actions;"
odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count


		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
			print strArr(2)
		'If trim(strArr(3)) = "Question Attachment Name" Then
		If trim(strArr(2)) = "Sort by:Question Attachment NameSorted AscendingShow actions" Then
			l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows") 
print a		
		b = l_link(i).GetROProperty("cols") 		
print b		
		For x  = 1 to a     
				For j  = 1 to b
				c = l_link(i).GetCellData(x,j-m)
				print c
				Arr1 = split(c, "Open ")
Print Arr1(0)
Print Arr1(1)
New_QsName = Arr1(0)
Print New_QsName
				
				Next
		Next
			End if
		End If
captureLinkText_QuestionAttachment_AppEA = New_QsName

End Function

Function Verify_AppForm_Tabs_EA_DI()

verifyIfLinkDoesExist1 "Project Name & Contact Information"
verifyIfLinkDoesExist1 "Pre-Screen Questionnaire"
verifyIfLinkDoesExist1 "Organization & Project Lead Details"
verifyIfLinkDoesExist1 "Project Summary"
verifyIfLinkDoesExist1 "Additional Project Information"
verifyIfLinkDoesExist1 "Using PCORI-funded Evidence & Tools"
verifyIfLinkDoesExist1 "Key Personnel"
verifyIfLinkDoesExist1 "Attachments"
verifyIfLinkDoesExist1 "Budget"
verifyIfLinkDoesExist1 "Authorizations"

End Function

Public Function verify_WebEditDoes_Exist_by_Default_Value_andIndex (stValue,HtmlTag,i)
Err.Clear
On error resume next

Set btncalc = Description.Create()  
btncalc("micclass").value = "Browser"

Set btn = DeskTop.ChildObjects(btncalc)
strHwnd = btn(btn.count - 1).GetRoProperty("hwnd")
print strHwnd

Set pa = Description.Create
pa("micclass").value = "Page"


Set l = Description.Create
l("micclass").value = "WebEdit"
l("html tag").value = HtmlTag
l("default value").value = stValue 
''''l("html id").value = HtmlID

strcount = l.count
print strcount

Set lo =  getParentObject().ChildObjects(l)
Print lo.count
lo(i).Highlight

If lo.count = 0 Then
LogReport 1,"verify_WebEditDoes_Exist_by_Default_Value_andIndex - " & i, "The Value --" & stValue & "- does NOT exist - Not Expected"
else
LogReport 0,"verify_WebEditDoes_Exist_by_Default_Value_andIndex - " & i,  "The Value --" & stValue & "-Exists - Expected"
End If
End Function

Public Function captureLinkText_LOIReview_Number_EA(k, m)
On error resume next 
Set odesc=Description.Create
odesc("micclass").value="WebTable"
''odesc("column names").value = "Action;LOI Accepted or Denied\?;SF LOI Review Number;Reviewer Name;Reviewer Label;COI;Program;Institution Name FC;PI First Name FC;PI Last Name FC"

odesc("column names").value = ";Choose a RowSelect All;Sort by:Review IDSorted: NoneShow Review ID column actions;Sort by:StatusSorted: NoneShow Status column actions;Sort by:ReviewerSorted: NoneShow Reviewer column actions;Sort by:Reviewer DispositionSorted: NoneShow Reviewer Disposition column actions;Sort by:Created BySorted: NoneShow Created By column actions;"
odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count


		strName = l_link(i).GetROProperty("column names")
		print "Table Column Names: "& strName
		
		strArr = split(strName,";")
		If Not(strName = "") Then
			'print strArr(2)
		
		
		'If trim(strArr(2)) = "SF LOI Review Number" Then
		If trim(strArr(2)) = "Sort by:Review IDSorted: NoneShow Review ID column actions" Then
				
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 
		Print a
		print b
		For x  = 1 to k    
print x		
				For j  = 1 to b
				print j
				c = l_link(i).GetCellData(x,b-m)
								print c
				Next
		Next
			End if
		End If
captureLinkText_LOIReview_Number_EA = c

End Function
Public Function captureLinkText_EA_LOI_Review(m)
On error resume next 

Set odesc=Description.Create
odesc("micclass").value="WebTable"
'odesc("column names").value = "Action;Description;Question Attachment Name"
odesc("column names").value =";Choose a RowSelect All;Sort by:Review IDSorted: NoneShow Review ID column actions;Sort by:StatusSorted: NoneShow Status column actions;Sort by:ReviewerSorted: NoneShow Reviewer column actions;Sort by:Reviewer DispositionSorted: NoneShow Reviewer Disposition column actions;Sort by:Created BySorted: NoneShow Created By column actions;"
odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count


		strName = l_link(i).GetROProperty("column names")
		print strName
		strArr = split(strName,";")
		If Not(strName = "") Then
			print strArr(2)
		'If trim(strArr(3)) = "Question Attachment Name" Then
		If trim(strArr(2)) = "Sort by:Review IDSorted: NoneShow Review ID column actions" Then
			l_link(i).GetROProperty("rows") 
		
		a = l_link(i).GetROProperty("rows") 
print a		
		b = l_link(i).GetROProperty("cols") 		
print b		
		For x  = 1 to a     
				For j  = 1 to b
				c = l_link(i).GetCellData(x,j-m)
				print c
				Next
		Next
			End if
		End If
captureLinkText_EA_LOI_Review = c

End Function
Function Verify_LOIForm_Tabs_EA_CB_New()

verifyIfLinkDoesExist1 "Project Name & Contact Information"
verifyIfLinkDoesExist1 "Pre-Screen Questionnaire"
verifyIfLinkDoesExist1 "Organization & Project Lead Details"
verifyIfLinkDoesExist1 "Project Summary"
verifyIfLinkDoesExist1 "Additional Project Information"
verifyIfLinkDoesExist1 "Using PCORI-funded Tools & Resources"
'''''''''verifyIfLinkDoesExist1 "Attachments"             '''''''''Out of scope for SCS and CB PFA'''''''''
verifyIfLinkDoesExist1 "Authorizations"

End Function

Function Verify_AppForm_Tabs_EA_CB()

verifyIfLinkDoesExist1 "Project Name & Contact Information"
verifyIfLinkDoesExist1 "Pre-Screen Questionnaire"
verifyIfLinkDoesExist1 "Organization & Project Lead Details"
verifyIfLinkDoesExist1 "Project Summary"
verifyIfLinkDoesExist1 "Additional Project Information"
verifyIfLinkDoesExist1 "Using PCORI-funded Tools & Resources"
verifyIfLinkDoesExist1 "Key Personnel"
verifyIfLinkDoesExist1 "Attachments"
verifyIfLinkDoesExist1 "Budget"
verifyIfLinkDoesExist1 "Authorizations"

End Function

Public Function captureLinkText_ProposalReview_Number_EA(k, m)
On error resume next 
Set odesc=Description.Create
odesc("micclass").value="WebTable"
''odesc("column names").value = "Action;LOI Accepted or Denied\?;SF LOI Review Number;Reviewer Name;Reviewer Label;COI;Program;Institution Name FC;PI First Name FC;PI Last Name FC"

odesc("column names").value = ";Choose a RowSelect All;Sort by:Review IDSorted: NoneShow Review ID column actions;Sort by:StatusSorted: NoneShow Status column actions;Sort by:ReviewerSorted: NoneShow Reviewer column actions;Sort by:Preliminary DispositionSorted: NoneShow Preliminary Disposition column actions;Sort by:Created BySorted: NoneShow Created By column actions;"
odesc("html tag").value = "TABLE"

Set l_link=  getParentObject().ChildObjects(odesc)
print l_link.count


		strName = l_link(i).GetROProperty("column names")
		print "Table Column Names: "& strName
		
		strArr = split(strName,";")
		If Not(strName = "") Then
			'print strArr(2)
		
		
		'If trim(strArr(2)) = "SF LOI Review Number" Then
		If trim(strArr(2)) = "Sort by:Review IDSorted: NoneShow Review ID column actions" Then
				
		a = l_link(i).GetROProperty("rows")  
		b = l_link(i).GetROProperty("cols") 
		Print a
		print b
		For x  = 1 to k    
print x		
				For j  = 1 to b
				print j
				c = l_link(i).GetCellData(x,b-m)
								print c
				Next
		Next
			End if
		End If
captureLinkText_ProposalReview_Number_EA = c

End Function

