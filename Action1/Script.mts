'{"SC-OUT-PatientECD": "","Environment": "Dev","AdmittingAttendingPhysician": "Ahmed Faruq","Guarantor": "1","InsuranceFinancialClass": "M","SC-OUT-MRN": "","InsurancePayerPlan": "MEDICAID/MEDICAID PENDING","EncounterProvider": "EMR","PatientGender": "F","SC-OUT-PatientENC": "","PatientDOB": "12/11/1985","PatientLastName": "XYZAABB","Diagnosis": "R42","PatientFirstName": "AABBXYZ","IsTAP": "Y"} 
On Error Resume Next
'***********************************************************************************************************************************************************************
'Testscript name:NewPt_QuickCheckin_CompleteCheckin_ED
'Description: Checkin inpatient using QuickChkin_IncompletChkin_Complete_SSEmergency_Room
' This script is divided into two flows 1.QUICK_ED="YES" 2.COMPLETE_ED="YES"
'Testcases covered:PA.009_New Patient_Quick Check In_Incomplete Check-In WL_Complete Check In - SS Emergency Room
'***********************************************************************************************************************************************************************

Dim patient_lastName,patient_firstName,patient_dob,patient_SSN,patient_SSNReason,patient_gender,patient_ReasonforEncounter,patient_PrincipalAdmittingDiagnosisCod,patient_GuarantorList,patient_ApptReminderText,patient_addressStreet,patient_addressZip,patient_PhoneTypeEditComboBox,patient_PhoneNumber,patient_PreferredLanguage,patient_MaritalStatus,Insurance_PayerPlanQuickPick,Insurance_PolicyNumberInput,Insurance_SubscriptionType,patient_ScheduleAppointmentDate,patient_ScheduleDepartment,patient_ScheduleLocation,patient_Activity,NOK_FirstName,NOK_LastName,NOK_Patient_Is,NOK_Contact_Type,NOK_PreferredNum,PatientsFathersName,PatientsMothersName,Patient_ID,Patient_Email,Patient_IfNoEmail,Patient_ServedMilitary,Patient_EmailBelongs,Patient_Ethnicity,Patient_ReligiousAffilialition,Patient_Emp_Occupation,Patient_NorthwellEmployee,Patient_Race,Accident_Attornet_Populated,Patient_Homeless,Insurance_NotQuickPlan,Insurance_NotQuickPlanSearch,Patient_MaintainComments_Code,Insurance_MedicareA,Insurance_MedicareB,Insurance_MedicarePolicyNumber,Guarantor_patient_lastName,Guarantor_patient_firstName,Guarantor_Patient_Is,Employment_OrganizationName,Employment_Employer_Status,Employment_Primary,Encounter_Provider,Encounter_Location,Encounter_StartDate,Encounter_Time,CareProvider_FirstName,CareProvider_LastName,CareProvider_Role_Of_PatientCare,CareProvider_Speciality,CareProvider_Startdate,CareProvider_EndDate,Hippa_Stmt_Provider,Hippa_Date_Provided,Hippa_Patient_Signature,Encounter_Admission_Type,Encounter_AdmittingDiagnoses,Encounter_ReasonForEncounter,Encounter_AccidentRelated,Encounter_Incident_StartDate,Encounter_Incident_StartTime,Encounter_IncidentType,Encounter_Incident_NoFaultInsurance,Encounter_Incident_Description,Encounter_Incident_State,Encounter_Incident_Attorney_FirstName,Encounter_Incident_Attorney_Lastname,Encounter_Incident_Attorney_Email,Encounter_Incident_Attorney_Phone,Encounter_Incident_Attorney_Firmname,Encounter_Incident_Attorney_Street,Encounter_Incident_Attorney_City,Encounter_Incident_Attorney_Zip,Encounter_Incident_Attorney_State,Encounter_HealthProf_Admitting,Encounter_HealthProf_Attending,Encounter_HealthProf_ReferringForm,Encounter_HealthProf_Startdate,Encounter_HealthProf_Time,HealthProf_LastName,Encounter_HealthProf_Speciality,Encounter_Gaurantor,Encounter_PatientIs,Encounter_BenefitsAssignmentcertification,Encounter_ReleaseOfInformation,Encounter_Gaurantor_Date,Encounter_Additionadata_PhoneExt,Encounter_ClinicalService,HealthProf_FirstName,Patient_NameType,Encounter_Level_of_Care,Encounter_Nurse_Station,Encounter_Privacy,Encounter_Opt_OutofDelivery,Encounter_PatientInhousePhoneExt,Encounter_Admission_Source,QuickCheckin_ReasonforEncounter,QuickCheckin_ArrivalMode,Checkin_Admission_Source,Checkin_Admission_Type,Encounter_Additionadata_PatientCommConsent,patient_ForeignAddress_Country,patient_addressState,patient_addressCounty,patient_addressCity
Dim cnt,SkiponError,ScreenShotName
'Datasheet location for the script
FileName=sSourceDataFile'"X:\Automation\R&D\Soarian fw optimization\Soarian Project\TestData\NewPt_QuickCheckin_CompleteCheckin_ED.xlsx"
OutputFile=strLocalTempALMFolder&"NewPt_QuickCheckin_CompleteCheckin_ED_"&month(now)&"_"&day(now)&"_"&year(now)&hour(now)&"_"&minute(now)&second(now)&".xlsx"


''Getting data from table in case of TAP
'If sGTAPFlag = "Y" Then
'	Call fnInputJSONModification()
'	Call fnMappingDictionary(sSourceDataFile)
'	DataTable.Value("patient_dob","Global") = FormatDate(GetJsonDate(inTable("PatientDOB")),"MM/dd/yyyy")
'	Print DataTable.Value("patient_dob","Global")
'End If



DataTable.ImportSheet FileName, 1, 1 
DataTable.GetSheet("Global")
cnt=DataTable.GetRowCount
For i=1 to cnt
 For SkiponError = 1 To 1 Step 1
    DataTable.SetCurrentRow(i)
	scriptIteration=i
    FlagExecute=Trim(DataTable.Value("Flag","Global"))
    If Ucase(FlagExecute)="Y" Then
         
        'Stores the start time of the iteration and updates sheet
        Starttime=Now()
		DataTable.Value("Start_Time","Global")=Starttime
		Screenshot_Required=Trim(DataTable.Value("Screenshot_Required","Global"))
        'Values fetched from datasheet
    	patient_lastName=Trim(DataTable.Value("patient_lastName","Global"))	
    	patient_firstName=Trim(DataTable.Value("patient_firstName","Global"))	
    	patient_middleName=Trim(DataTable.Value("patient_middleName","Global"))
    	patient_dob=Trim(DataTable.Value("patient_dob","Global"))	
    	patient_dob = FormatDate(patient_dob,"MM/dd/yyyy")	
    	patient_SSN=Trim(DataTable.Value("patient_SSN","Global")) 
    	patient_SSNReason=Trim(DataTable.Value("patient_SSNReason","Global"))' just for testing!---> should NOT BE USED!!!! Reason for no ssn "refused to provide ssn" should be selected instead	
    	patient_gender=Trim(DataTable.Value("patient_gender","Global"))	
    	patient_ReasonforEncounter=Trim(DataTable.Value("patient_ReasonforEncounter","Global"))	
    	'patient_PrincipalAdmittingDiagnosisCod=Trim(DataTable.Value("patient_PrincipalAdmittingDiagnosisCod","Global"))	
    	patient_GuarantorList=Trim(DataTable.Value("patient_GuarantorList","Global"))	
    	patient_ApptReminderText=Trim(DataTable.Value("patient_ApptReminderText","Global"))	
    	patient_addressStreet=Trim(DataTable.Value("patient_addressStreet","Global"))	
    	Randomize
    	patient_addressStreet = Int((1000 - 10 + 1) * Rnd + 10) & " " & patient_addressStreet
    	
    	patient_addressZip=Trim(DataTable.Value("patient_addressZip","Global"))	
    	patient_PhoneTypeEditComboBox=Trim(DataTable.Value("patient_PhoneTypeEditComboBox","Global"))	
    	patient_PhoneNumber=Trim(DataTable.Value("patient_PhoneNumber","Global"))	
    	patient_PreferredLanguage=Trim(DataTable.Value("patient_PreferredLanguage","Global"))	
    	patient_MaritalStatus=Trim(DataTable.Value("patient_MaritalStatus","Global"))	
    	Insurance_PayerPlanQuickPick=Trim(DataTable.Value("Insurance_PayerPlanQuickPick","Global"))	
    	Insurance_PolicyNumberInput=Trim(DataTable.Value("Insurance_PolicyNumberInput","Global"))	
    	Insurance_SubscriptionType=Trim(DataTable.Value("Insurance_SubscriptionType","Global"))	
    	patient_ScheduleAppointmentDate	=month(now)&"/"&day(now)&"/"&year(now)'Trim(DataTable.Value("patient_ScheduleAppointmentDate","Global"))
    	patient_ScheduleDepartment	=Trim(DataTable.Value("patient_ScheduleDepartment","Global"))
    	patient_ScheduleLocation=Trim(DataTable.Value("patient_ScheduleLocation","Global"))	
    	patient_Activity=Trim(DataTable.Value("patient_Activity","Global"))	
    	NOK_FirstName=Trim(DataTable.Value("NOK_FirstName","Global"))	
    	NOK_LastName=Trim(DataTable.Value("NOK_LastName","Global"))	
    	NOK_Patient_Is=Trim(DataTable.Value("NOK_Patient_Is","Global"))	
    	NOK_Contact_Type=Trim(DataTable.Value("NOK_Contact_Type","Global"))	
    	NOK_PreferredNum=Trim(DataTable.Value("NOK_PreferredNum","Global"))	
		NOK_EmergNotprovided=Trim(DataTable.Value("Contacts_Emergnotprovided","Global"))  	
		
		PatientsFathersName=Trim(DataTable.Value("Patient_Fathername","Global"))
    	PatientsMothersName=Trim(DataTable.Value("Patient_Mothername","Global"))
		
		Patient_ID=Trim(DataTable.Value("Patient_ID","Global"))
		Patient_Email=Trim(DataTable.Value("Patient_Email","Global"))			
		Patient_IfNoEmail=Trim(DataTable.Value("Patient_IfNoEmail","Global"))	
		Patient_ServedMilitary=Trim(DataTable.Value("Patient_ServedMilitary","Global"))	
		Patient_EmailBelongs=Trim(DataTable.Value("Patient_EmailBelongs","Global"))	
		Patient_Ethnicity=Trim(DataTable.Value("Patient_Ethnicity","Global"))
		Patient_ReligiousAffilialition=Trim(DataTable.Value("Patient_ReligiousAffilialition","Global"))	
		Patient_Emp_Occupation=Trim(DataTable.Value("Patient_Emp_Occupation","Global"))	
		Patient_NorthwellEmployee=Trim(DataTable.Value("Patient_NorthwellEmployee","Global"))	
    	Patient_Race=Trim(DataTable.Value("Patient_Race","Global"))
    	Accident_Attornet_Populated=Trim(DataTable.Value("Accident_Attornet_Populated","Global"))
    	Patient_Homeless=Trim(DataTable.Value("Patient_Homeless","Global"))
		Insurance_NotQuickPlan=DataTable.Value("Insurance_NotQuickPlan","Global")'Donot trim the value of weblist
		Insurance_NotQuickPlanSearch=Trim(DataTable.Value("Insurance_NotQuickPlanSearch","Global"))
		Patient_MaintainComments_Code=Trim(DataTable.Value("Patient_MaintainComments_Code","Global"))
		Insurance_MedicareA=Trim(DataTable.Value("Insurance_MedicareA","Global"))
        Insurance_MedicareB=Trim(DataTable.Value("Insurance_MedicareB","Global"))
        Insurance_MedicarePolicyNumber=Trim(DataTable.Value("Insurance_MedicarePolicyNumber","Global"))
        Guarantor_patient_lastName=Trim(DataTable.Value("Guarantor_patient_lastName","Global"))
        Guarantor_patient_firstName=Trim(DataTable.Value("Guarantor_patient_firstName","Global"))
        Guarantor_Patient_Is=Trim(DataTable.Value("Guarantor_Patient_Is","Global"))
        Employment_OrganizationName=Trim(DataTable.Value("Employment_OrganizationName","Global"))
        Employment_Employer_Status=Trim(DataTable.Value("Employment_Employer_Status","Global"))
        Employment_Primary=Trim(DataTable.Value("Employment_Primary","Global"))
        Encounter_Provider=Trim(DataTable.Value("Encounter_Provider","Global"))
        Encounter_Location=Trim(DataTable.Value("Encounter_Location","Global"))
        Encounter_StartDate=Trim(DataTable.Value("Encounter_StartDate","Global"))
        Encounter_Time=Trim(DataTable.Value("Encounter_Time","Global"))
        Encounter_HealthProf_Admitting = Trim(DataTable.Value("Encounter_HealthProf_Admitting"))
        
        If inTable("CareProviders") <> "" Then
        	DataTable.Value("CareProvider_LastName","Global") = Trim(inTable("CareProviders"))
        End If
        DataTable.Value("CareProvider_LastName","Global") = Replace(DataTable.Value("CareProvider_LastName","Global"),",","")
        arrCareProvider_LastName = Split(Trim(DataTable.Value("CareProvider_LastName","Global")))
        If Trim(arrCareProvider_LastName(0)) <> "" or Trim(arrCareProvider_LastName(1))<>"" Then 'Ubound(arrCareProvider_LastName) > 0 Then
	        CareProvider_LastName=Trim(arrCareProvider_LastName(0))
	        CareProvider_FirstName=Trim(arrCareProvider_LastName(1))'Trim(DataTable.Value("CareProvider_FirstName","Global"))
	        DataTable.Value("CareProvider_FirstName","Global") = arrCareProvider_LastName(1)
	        DataTable.Value("CareProvider_LastName","Global") = arrCareProvider_LastName(0)
	    Else
	    	CareProvider_LastName = Trim(DataTable.Value("CareProvider_LastName","Global"))
	    	CareProvider_FirstName = Trim(DataTable.Value("CareProvider_FirstName","Global"))
        End If

        
              
		CareProvider_Role_Of_PatientCare=Trim(DataTable.Value("CareProvider_Role_Of_PatientCare","Global"))
		CareProvider_Speciality=Trim(DataTable.Value("CareProvider_Speciality","Global"))
		CareProvider_Startdate=Trim(DataTable.Value("CareProvider_Startdate","Global"))
		CareProvider_EndDate=Trim(DataTable.Value("CareProvider_EndDate","Global"))
		Hippa_Stmt_Provider=Trim(DataTable.Value("Hippa_Stmt_Provider","Global"))
		Hippa_Date_Provided=Trim(DataTable.Value("Hippa_Date_Provided","Global"))
		Hippa_Patient_Signature=Trim(DataTable.Value("Hippa_Patient_Signature","Global"))
		Encounter_Admission_Type=Trim(DataTable.Value("Encounter_Admission_Type","Global"))
        'Encounter_AdmittingDiagnoses=Trim(DataTable.Value("Encounter_AdmittingDiagnoses","Global"))
        Encounter_ReasonForEncounter=Trim(DataTable.Value("Encounter_ReasonForEncounter","Global"))
        Encounter_AccidentRelated=Trim(DataTable.Value("Encounter_AccidentRelated","Global"))
        Encounter_Incident_StartDate=Trim(DataTable.Value("Encounter_Incident_StartDate","Global"))        
		Encounter_Incident_StartTime=Trim(DataTable.Value("Encounter_Incident_StartTime","Global"))	
		Encounter_IncidentType=Trim(DataTable.Value("Encounter_IncidentType","Global"))	
		Encounter_Incident_NoFaultInsurance=Trim(DataTable.Value("Encounter_Incident_NoFaultInsurance","Global"))	
		Encounter_Incident_Description=Trim(DataTable.Value("Encounter_Incident_Description","Global"))	
		Encounter_Incident_State=Trim(DataTable.Value("Encounter_Incident_State","Global"))	
		Encounter_Incident_Attorney_FirstName=Trim(DataTable.Value("Encounter_Incident_Attorney_FirstName","Global"))	
		Encounter_Incident_Attorney_Lastname=Trim(DataTable.Value("Encounter_Incident_Attorney_Lastname","Global"))	
		Encounter_Incident_Attorney_Email=Trim(DataTable.Value("Encounter_Incident_Attorney_Email","Global"))	
		Encounter_Incident_Attorney_Phone=Trim(DataTable.Value("Encounter_Incident_Attorney_Phone","Global"))	
		Encounter_Incident_Attorney_Firmname=Trim(DataTable.Value("Encounter_Incident_Attorney_Firmname","Global"))	
		Encounter_Incident_Attorney_Street=Trim(DataTable.Value("Encounter_Incident_Attorney_Street","Global"))	
		Encounter_Incident_Attorney_City=Trim(DataTable.Value("Encounter_Incident_Attorney_City","Global"))	
		Encounter_Incident_Attorney_Zip=Trim(DataTable.Value("Encounter_Incident_Attorney_Zip","Global"))
		Encounter_Incident_Attorney_State=Trim(DataTable.Value("Encounter_Incident_Attorney_State","Global"))	
		
		
		'Encounter_HealthProf_ReferringForm=Trim(DataTable.Value("Encounter_HealthProf_ReferringForm","Global"))	
		Encounter_HealthProf_Startdate=Trim(DataTable.Value("Encounter_HealthProf_Startdate","Global"))
		Encounter_HealthProf_Time=Trim(DataTable.Value("Encounter_HealthProf_Time","Global"))	
		
		
		DataTable.Value("HealthProf_Admitting_LastName","Global") = Replace(DataTable.Value("HealthProf_Admitting_LastName","Global"),",","")
		arrHealthProf_Admitting_LastName = Split(Trim(DataTable.Value("HealthProf_Admitting_LastName","Global")))
		If Ubound(arrHealthProf_Admitting_LastName) > 0 Then
			HealthProf_Admitting_LastName=Trim(arrHealthProf_Admitting_LastName(0))
			HealthProf_Admitting_FirstName=Trim(arrHealthProf_Admitting_LastName(1))'Trim(DataTable.Value("HealthProf_FirstName","Global"))
			
			DataTable.Value("HealthProf_Admitting_LastName","Global") = Trim(arrHealthProf_Admitting_LastName(0))
			DataTable.Value("HealthProf_Admitting_FirstName","Global") = Trim(arrHealthProf_Admitting_LastName(1))
		Else
			HealthProf_Admitting_LastName = Trim(DataTable.Value("HealthProf_Admitting_LastName","Global"))
			HealthProf_Admitting_FirstName = Trim(DataTable.Value("HealthProf_Admitting_FirstName","Global"))
		End If
		
		DataTable.Value("Encounter_HealthProf_AttendingName","Global") = Replace(DataTable.Value("Encounter_HealthProf_AttendingName","Global"),",","")
		arrEncounter_HealthProf_AttendingName = Split(Trim(DataTable.Value("Encounter_HealthProf_AttendingName","Global")))
		If Ubound(arrEncounter_HealthProf_AttendingName) > 0 Then
			Encounter_HealthProf_AttendingName = Trim(arrEncounter_HealthProf_AttendingName(0))
		Else
			Encounter_HealthProf_AttendingName = Trim(DataTable.Value("Encounter_HealthProf_AttendingName","Global"))
		End If
		
		'Referring doctor
		DataTable.Value("Encounter_HealthProf_ReferringName","Global") = Replace(DataTable.Value("Encounter_HealthProf_ReferringName","Global"),",","")
		arrEncounter_HealthProf_ReferringName = Split(Trim(DataTable.Value("Encounter_HealthProf_ReferringName","Global")))
		If Ubound(arrEncounter_HealthProf_ReferringName) > 0 Then
			Encounter_HealthProf_ReferringName=Trim(arrEncounter_HealthProf_ReferringName(0))
			HealthProf_ReferringFirstName=Trim(arrEncounter_HealthProf_ReferringName(1))'Trim(DataTable.Value("HealthProf_FirstName","Global"))
			
			DataTable.Value("Encounter_HealthProf_ReferringName","Global") = Trim(arrEncounter_HealthProf_ReferringName(0))
			DataTable.Value("HealthProf_ReferringFirstName","Global") = Trim(arrEncounter_HealthProf_AdmittingName(1))
		Else
			Encounter_HealthProf_ReferringName = Trim(DataTable.Value("Encounter_HealthProf_ReferringName","Global"))
			HealthProf_ReferringFirstName = Trim(DataTable.Value("HealthProf_ReferringFirstName","Global"))
		End If
		
		Encounter_Referring_Speciality=Trim(DataTable.Value("Encounter_Referring_Speciality","Global"))	
		
		
		
'		HealthProf_LastName=Trim(DataTable.Value("HealthProf_LastName","Global"))	
		DataTable.Value("HealthProf_LastName","Global") = Replace(DataTable.Value("HealthProf_LastName","Global"),",","")
		arrHealthProf_LastName = Split(Trim(DataTable.Value("HealthProf_LastName","Global")))
		If Ubound(arrHealthProf_LastName) > 0 Then
			HealthProf_LastName=Trim(arrHealthProf_LastName(0))
			HealthProf_FirstName=Trim(arrHealthProf_LastName(1))'Trim(DataTable.Value("HealthProf_FirstName","Global"))
			
			DataTable.Value("HealthProf_LastName","Global") = Trim(arrHealthProf_LastName(0))
			DataTable.Value("HealthProf_FirstName","Global") = Trim(arrHealthProf_LastName(1))
		Else
			HealthProf_LastName = Trim(DataTable.Value("HealthProf_LastName","Global"))
			HealthProf_FirstName = Trim(DataTable.Value("HealthProf_FirstName","Global"))
		End If
		
		
			
		Encounter_HealthProf_Speciality=Trim(DataTable.Value("Encounter_HealthProf_Speciality","Global"))	
		'Encounter_Gaurantor=Trim(DataTable.Value("Encounter_Gaurantor","Global"))
		Encounter_PatientIs=Trim(DataTable.Value("Encounter_PatientIs","Global"))	
		Encounter_BenefitsAssignmentcertification=Trim(DataTable.Value("Encounter_BenefitsAssignmentcertification","Global"))	
		Encounter_ReleaseOfInformation=Trim(DataTable.Value("Encounter_ReleaseOfInformation","Global"))	
		Encounter_Gaurantor_Date=Trim(DataTable.Value("Encounter_Gaurantor_Date","Global"))	
		
		Encounter_Additionadata_PhoneExt=Trim(DataTable.Value("Encounter_Additionadata_PhoneExt","Global"))
		
		Encounter_ClinicalService=Trim(DataTable.Value("Encounter_ClinicalService","Global"))
		
		Select Case Encounter_ClinicalService
			Case "EMR"
				Encounter_ClinicalService = "Emergency"
				DataTable.Value("Encounter_ClinicalService","Global") = Encounter_ClinicalService
'			Case Else
'				GRC = 0
'				Reporter.ReportEvent micFail,"Clinical service verification","Clicial Service is neither EMR nor Emergency. Actual Value : " & Encounter_ClinicalService
'				DataTable.Value("Error_Message","Global") = "Clicial Service is neither EMR nor Emergency. Actual Value : " & Encounter_ClinicalService
		End Select
		
		Patient_NameType=Trim(DataTable.Value("Patient_NameType","Global"))
		Encounter_Level_of_Care=Trim(DataTable.Value("Encounter_Level_of_Care","Global"))
        Encounter_Nurse_Station=Trim(DataTable.Value("Encounter_Nurse_Station","Global"))
        Encounter_Privacy=Trim(DataTable.Value("Encounter_Privacy","Global"))
        Encounter_Opt_OutofDelivery=Trim(DataTable.Value("Encounter_Opt_OutofDelivery","Global"))
        Encounter_PatientInhousePhoneExt=Trim(DataTable.Value("Encounter_PatientInhousePhoneExt","Global"))
        Encounter_Admission_Source=Trim(DataTable.Value("Encounter_Admission_Source","Global"))
        
       QuickCheckin_ReasonforEncounter=Trim(DataTable.Value("QuickCheckin_ReasonforEncounter","Global"))
	QuickCheckin_ArrivalMode=Trim(DataTable.Value("QuickCheckin_ArrivalMode","Global"))
	'QuickCheckin_AmbulanceSquadCode = Trim(DataTable.Value("QuickCheckin_AmbulanceSquadCode","Global"))
	QuickCheckin_AmbulanceSquadCode=Trim(DataTable.Value("QuickCheckin_AmbulanceSquadCode","Global"))
	Checkin_Admission_Source=Trim(DataTable.Value("Checkin_Admission_Source","Global"))
	Checkin_Admission_Type=Trim(DataTable.Value("Checkin_Admission_Type","Global"))
	Encounter_Additionadata_PatientCommConsent=Trim(DataTable.Value("Encounter_Additionadata_PatientCommConsent","Global"))
	    
	    patient_ForeignAddress_Country=Trim(DataTable.Value("patient_ForeignAddress_Country","Global"))
        patient_addressState=Trim(DataTable.Value("patient_addressState","Global"))
        patient_addressCounty=Trim(DataTable.Value("patient_addressCounty","Global"))
        patient_addressCity=Trim(DataTable.Value("patient_addressCity","Global"))
        
        Encounter_GaurantorLastname=Trim(DataTable.Value("Encounter_GaurantorLastname","Global"))
        Encounter_GaurantorFirstname=Trim(DataTable.Value("Encounter_GaurantorFirstname","Global"))
        Encounter_GaurantorGender=Trim(DataTable.Value("Encounter_GaurantorGender","Global"))	
        Encounter_GaurantorDOB=Trim(DataTable.Value("Encounter_GaurantorDOB","Global"))
        
        Encounter_HealthProf_ReferringForm=Trim(DataTable.Value("Encounter_HealthProf_ReferringForm","Global"))
        HealthProf_Admitting_LastName=Trim(DataTable.Value("HealthProf_Admitting_LastName","Global"))
        HealthProf_Admitting_FirstName=Trim(DataTable.Value("HealthProf_Admitting_FirstName","Global"))
        Encounter_HealthProf_AdmittingSpeciality=Trim(DataTable.Value("Encounter_HealthProf_AdmittingSpeciality","Global"))
        HealthProf_ReferringForm_LastName=Trim(DataTable.Value("HealthProf_ReferringForm_LastName","Global"))	
        HealthProf_ReferringForm_FirstName=Trim(DataTable.Value("HealthProf_ReferringForm_FirstName","Global"))	
        Encounter_HealthProf_ReferringFormSpeciality=Trim(DataTable.Value("Encounter_HealthProf_ReferringFormSpeciality","Global"))
        
        
        Contacts_FirstName=Trim(DataTable.Value("Contacts_FirstName","Global"))
        Contacts_LastName=Trim(DataTable.Value("Contacts_LastName","Global"))
        Contacts_Patient_Is=Trim(DataTable.Value("Contacts_Patient_Is","Global"))
        Contacts_Contact_Type=Trim(DataTable.Value("Contacts_Contact_Type","Global"))
        Contacts_PreferredNum=Trim(DataTable.Value("Contacts_PreferredNum","Global"))
        patient_PhoneCheckbox=Trim(DataTable.Value("patient_PhoneCheckbox","Global"))
        CareProvider_Notprovided=Trim(DataTable.Value("CareProvider_Notprovided","Global"))
        
        QUICK_ED=Trim(DataTable.Value("QUICK_ED","Global"))
        COMPLETE_ED=Trim(DataTable.Value("COMPLETE_ED","Global"))
        healthix_consent=DataTable.Value("Healthix_consent")
        DataTable.Value("HealthixConsent_Date") = FormatDate(DataTable.Value("HealthixConsent_Date"),"MM/dd/yyyy")
        HealthixConsent_Date=DataTable.Value("HealthixConsent_Date")
        
		'Secondary insurance columns
		SecInsurance_PayerPlanQuickPick=Trim(DataTable.Value("SecInsurance_PayerPlanQuickPick","Global"))	
		SecInsurance_PolicyNumberInput=Trim(DataTable.Value("SecInsurance_PolicyNumberInput","Global"))	
		SecInsurance_SubscriptionType=Trim(DataTable.Value("SecInsurance_SubscriptionType","Global"))
		SecInsurance_Verifiedforencounter=Trim(DataTable.Value("SecInsurance_Verifiedforencounter","Global"))
		SecInsurance_Subscriber_Lastname=Trim(DataTable.Value("SecInsurance_Subscriber_Lastname","Global"))
		SecInsurance_Subscriber_Firstname=Trim(DataTable.Value("SecInsurance_Subscriber_Firstname","Global"))
		SecInsurance_PatientIs=Trim(DataTable.Value("SecInsurance_PatientIs","Global"))
		SecGroupNumber=Trim(DataTable.Value("SecGroupNumber","Global"))
		
		SecInsurance_Hosp_Auth_Type=Trim(DataTable.Value("SecInsurance_Hosp_Auth_Type","Global"))
		SecInsurance_Hosp_Auth_No=Trim(DataTable.Value("SecInsurance_Hosp_Auth_No","Global"))
		SecInsurance_Hosp_Auth_ApprovalStatus=Trim(DataTable.Value("SecInsurance_Hosp_Auth_ApprovalStatus","Global"))
		SecInsurance_Hosp_DurationCnt=Trim(DataTable.Value("SecInsurance_Hosp_DurationCnt","Global"))
		SecInsurance_Hosp_DurationType=Trim(DataTable.Value("SecInsurance_Hosp_DurationType","Global")) 

		If not IsNumeric(Patient_Email) Then
		   Patient_Email = ""
		End If

        'To capture screenshots
        ScreenShotName=strTestCaseName&"_"&patient_lastName&"_"&patient_firstName
        sGScrShotsFolder = strLocalTempALMFolder&ScreenShotName
        DataTable.Value("Screenshot_Name","Global")=ScreenShotName
	    
       'Launch soarian application and login>Click on Appointment Management
       strAppnUsername="cs-eduser"
       strAppnPassword="!Soarian1"
    	Call Fn_LoginSoarianApplication(strDevAppnURL,strAppnUsername,strAppnPassword)
    'Quick ED Checkin flow
    If Ucase(QUICK_ED)="YES" OR Ucase(QUICK_ED)="Y" Then
    
    		
    		
    		 'Click the " Quick Check-In" link.
	    	Browser("SoarianDEV").Page("Soarian").Frame("Soarian Financials Home Page").WebElement("Quick Check-In").Click @@ hightlight id_;_Browser("Cerner Soarian [4.1 DEV]").Page("Cerner Soarian [4.1 DEV]")_;_script infofile_;_ZIP::ssf16.xml_;_
	        Browser("SoarianDEV").Page("Soarian").Sync
			'wait(1)
			
			'Enter Encounter_Provider_Location_Date
			If Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Authorized Encounter Locations and HPOs").Exist(10) Then
			  Call Fn_Select_Encounter_Provider_Location_Date(Encounter_Provider,Encounter_Location,Encounter_StartDate,Encounter_Time)	
			End If
			
			 Browser("SoarianDEV").Page("Soarian").Frame("BPOForm").WebButton("QuickCheckin_Find Patient").Exist(10)
			 'Click the "Find Patient" button.
			 wait 1    'putting wait to let the page load it was throwing error while submitting aptient in train.
			 Browser("SoarianDEV").Page("Soarian").Frame("BPOForm").WebButton("QuickCheckin_Find Patient").Click
		
			'Search for patient and click on "Add New" button to add a new patient
     		Call Fn_AddPatient(patient_lastName,patient_firstName,patient_dob,patient_gender) @@ hightlight id_;_Browser("SoarianDEV").Page("Cerner Soarian [4.1 DEV] 3").Frame("tabWell1").WebEdit("WebEdit")_;_script infofile_;_ZIP::ssf55.xml_;_
			
			 wait(2)
	        
	        
            'Verify DOB
	        If patient_dob<>"" Then
	           If Trim(Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("QuickEDDOB").GetROProperty("value"))="/  /" then
	           Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("QuickEDDOB").SetEdit patient_dob
	           End if 
            End If
            wait 1  'putting wait to let the page load it was throwing error while submitting aptient in train.
	        Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("QuickCheckin_ReasonforEncounter").Set QuickCheckin_ReasonforEncounter   
	        Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("QuickCheckin_ClicnicalService").SetEdit Encounter_ClinicalService
			If Ucase(Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("QuickCheckin_ClicnicalService").GetROProperty("value")) <> Ucase(Encounter_ClinicalService) Then
				GRC = 0
				Reporter.ReportEvent micFail,"Clinical service Selection","Unable to select given Clicial Service " & Encounter_ClinicalService
				Call fnTakeSnapShot(Browser("SoarianDEV").Page("Soarian"),strTestCaseName,"Fn_QuickCheckin_ClicnicalService_Fail")
'				DataTable.Value("Error_Message","Global") = "Clicial Service is neither EMR nor Emergency. Actual Value : " & Encounter_ClinicalService
			End If
			Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("QuickCheckin_ArrivalMode").click 'Object.focus
	        Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("QuickCheckin_ArrivalMode").SetEdit QuickCheckin_ArrivalMode
	        If QuickCheckin_AmbulanceSquadCode<>"" AND Ucase(QuickCheckin_ArrivalMode) = "AMBULANCE" Then
	        	Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("AmbulanceSquadCode").click 'Object.focus
	        	Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("AmbulanceSquadCode").SetEdit QuickCheckin_AmbulanceSquadCode
	        End If
	        If Patient_ID<>""  Then
	        	Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("QuickCheckin_IDSEEN").click 'Object.focus
	        	Browser("SoarianDEV").Page("Soarian").Frame("QuickCheckIn_BPOForm").WebEdit("QuickCheckin_IDSEEN").SetEdit Patient_ID
	        End If
	        Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").Click
	        
	        
	        If Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Finish Check-In").WebElement("Finish Quick Check-in").Exist Then
	        	Call fnTakeSnapShot(Browser("SoarianDEV").Page("Soarian"),strTestCaseName,"Finish_QuickCheckin")
	        	Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Finish Check-In").WebButton("Yes").Click
	        	If Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Finish Check-In").WebButton("Yes").Exist(2) Then
	        		Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Finish Check-In").WebButton("Yes").Click
	        		setting.WebPackage("ReplayType")=2
	        		Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Finish Check-In").WebButton("Yes").Click
	        		setting.WebPackage("ReplayType")=1
	        	End If
	        	Reporter.ReportEvent micPass,"Verifying Patient Quick Checin","Finish Quick Check-in Popup displayed. Quick Check-in Successful."
	        End If
	        Browser("SoarianDEV").Sync
           'Click on done button
         If Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").Exist(10) Then
         
            titleProperty= Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").GetTOProperty("title")
		    nameProperty=Fn_Returns_NameProperty_ForFrame(Browser("SoarianDEV").Page("Soarian"),titleProperty)
		    Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").SetTOProperty "name",nameProperty
		    
            If Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").WebElement("MRN").Exist Then
               MRN=Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").WebElement("MRN").GetROProperty("innertext")
               Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").WebElement("ECD").Exist
               Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").WebElement("ECD").Highlight
               ECD=Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").WebElement("ECD").GetROProperty("innertext")
               ENC=Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").WebElement("ENC").GetROProperty("innertext")
               MRNTemp=Trim(Replace(MRN,"| MR#:",""))
               DataTable.Value("PatientMRN","Global")=Trim(Replace(MRNTemp,"|  View Events",""))
               DataTable.Value("PatientECD","Global")=Trim(Replace(ECD,"| ECD#:",""))
               DataTable.Value("PatientENC","Global")=Trim(Replace(ENC,"| Enc#:",""))
            End If
            Call fnTakeSnapShot(Browser("SoarianDEV"),strTestCaseName,"Script_View_MRN")
         	Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").Click
         	Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").Click
         	If DataTable.Value("PatientMRN","Global") <> "NA" AND DataTable.Value("PatientECD","Global") <> "NA" AND DataTable.Value("PatientENC","Global") <> "NA" Then
         		DataTable.Value("Status","Global")="PASS"
         	Else
         		DataTable.Value("Status","Global")="FAIL"
         	End If
'         	DataTable.Value("Status","Global")="PASS"
         Else
            DataTable.Value("Status","Global")="FAIL"
         End If
         
''	        Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").Click
	        wait(5) : Browser("SoarianDEV").Sync
	        If Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").Exist(2) AND Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").GetROProperty("visible") Then
	        	setting.WebPackage("ReplayType")=2
	        	Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").Click
	        	setting.WebPackage("ReplayType")=1
	        End If
	        Wait(2)
    End If
       
        '************************Quick Checkin ends**********************************************************************************
     If Ucase(COMPLETE_ED)="YES" OR Ucase(COMPLETE_ED)="Y" Then 'Complete ED Flow starts
'    	    Call Fn_LoginSoarianApplication(strAppnURL,strAppnUsername,strAppnPassword)
        ''Incomplete Check In WL
        Browser("SoarianDEV").Page("Soarian").Frame("Soarian Financials Home Page").WebElement("Incomplete Check-in Worklist").Click
        
		Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Patient Information Worklist Options").WebTable("Encounter Location").Highlight
		Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Patient Information Worklist Options").WebElement("Encounter Location").SetTOProperty "innertext",UCASE(Encounter_Provider)
		'checkprovider=fn_SelectProviderName(Encounter_Provider)
		setting.WebPackage("ReplayType")=2
		If Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Patient Information Worklist Options").WebElement("EncounterLocation_Checkbox").Exist(2) Then
			Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Patient Information Worklist Options").WebElement("EncounterLocation_Checkbox").Click
			'Nandini :16 jul 2020:Adding sort by MRN and Descending
			Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Patient Information Worklist Options").WebEdit("Emergency_SortBy_Type").SetEdit "MRN"
			Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Patient Information Worklist Options").WebEdit("Emergency_SortBy").SetEdit "Descending"
		ElseIf GRC <> 0 Then
				GRC = 0
				Reporter.ReportEvent micFail,"Encounter Location Verification","Given Encounter Location " & Encounter_Provider & " Not Found"
				DataTable.Value("Error_Message","Global") = "Given Encounter Location " & Encounter_Provider & " Not Found"
				DataTable.Value("Status","Global") = "FAIL"
				Call fnTakeSnapShot(Browser("SoarianDEV").Window("SoarianSearch Win"),strTestCaseName,"Fn_EncounterLocation_Fail")
		End If
'	   Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Patient Information Worklist Options").WebElement("EncounterLocation_Checkbox").Click

		Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Patient Information Worklist Options").WebButton("OK").Click
		wait 15
		Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").WebButton("Refresh").Click
		wait 5
'		'Click on the Patient name hyperlink
'		strTemp=patient_lastName&", "&patient_firstName
'		count=1
'		If Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").Image("IncompleteCheckin_nextarrow").Exist(5) Then
'			Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").Image("IncompleteCheckin_nextarrow").Click
'		End If
'		If Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").Image("IncompleteCheckin_FirstPage").Exist(5) Then
'			Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").Image("IncompleteCheckin_FirstPage").Click
'		End If
		Do
		    titleProperty= Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").GetTOProperty("title")
            nameProperty=Fn_Returns_NameProperty_ForFrame(Browser("SoarianDEV").Page("Soarian"),titleProperty)
            Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").SetTOProperty "name",nameProperty
'            Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").Highlight
		    Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").WebElement("IncompleteCheckin_PatientLink").SetTOProperty "outertext",".*"&strTemp&".*"
		    wait(5)
		    Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").WebElement("IncompleteCheckin_PatientLink").RefreshObject
			If Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").WebElement("IncompleteCheckin_PatientLink").Exist(20) Then
			   Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").WebElement("IncompleteCheckin_PatientLink").Click
			   Exit Do
			End If
			If Not Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").Image("IncompleteCheckin_nextarrow").Exist(5) Then 'or count>5
				Exit Do
			End If 
			Browser("SoarianDEV").Page("Soarian").Frame("IncompleteCheckin_BPOForm").Image("IncompleteCheckin_nextarrow").Click
			count=count+1
		Loop While True
			 
		If Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Options").WebElement("Complete Check-in").Exist(5) Then
'			Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Options").WebElement("Complete Check-in").Click
			setting.WebPackage("ReplayType")=2
			Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Options").WebElement("Complete Check-in").Click
			setting.WebPackage("ReplayType")=1
			
			If Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Options").WebElement("Complete Check-in").Exist(5) Then
				Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Options").WebElement("Complete Check-in").Click
			End If
		ElseIf GRC<>0 Then
		    Reporter.ReportEvent micFail,"Verify if the patient link displayed in incomplete checkin","Not displayed"
	 	   sErrorMessage="Patient link NOT displayed in incomplete checkin List"
	 	   print sErrorMessage
	 	   GRC=0
	 	   DataTable.Value("Error_Message","Global") = sErrorMessage
	 	   DataTable.Value("Status","Global") = "FAIL"
	 	   Wait 1
		End If
		       
		       
		       Call Fn_ClosePatient_NeedsInterview_Popup()
        'Enter PatientsFathersName,PatientsMothersName,Patient_ID
        Call Fn_Enter_Checkin_Name_Section(PatientsFathersName,PatientsMothersName,Patient_ID)
        
        'Enter   Checkin_SSN,patient_SSNReason,patient_MaritalStatus,Patient_Race    
        Call Fn_Enter_Checkin_Personal_Information_Section(patient_SSN,patient_SSNReason,patient_MaritalStatus,Patient_Race)
        Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Checkin_Ethnicity").SetEdit Patient_Ethnicity
		
		Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Checkin_PreferredLanguage").SetEdit patient_PreferredLanguage
        
        'Enter Checkin_Email,Checkin_Email_Belongs_To,Patient_IfNoEmail
        Call Fn_Enter_Checkin_Email_Details(Patient_Email,Patient_EmailBelongs,Patient_IfNoEmail)
        
        Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("NorthwellHealthCenterPatient").SetEdit DataTable.Value("NorthwellHealthCenterPatient",1)
		
		''Sandeep 05/08/2018: callling  Fn_ManageContact_EnterAPerson function to hanlde Emergency contact warning in "Check-in Summary" screen (After Soarian upgradation)
		Call Fn_ManageContact_EnterAPerson(NOK_FirstName,NOK_LastName,NOK_Patient_Is,NOK_Contact_Type,NOK_PreferredNum,NOK_EmergNotprovided)
		
		wait(5)    

        'Add_Patient_EmploymentDetails
        'Enter Military,NorthwellEmp
        Call Fn_Enter_Checkin_Summary_Military_NorthwellEmp("",Patient_NorthwellEmployee)
		Call Fn_Add_Patient_EmploymentDetails(Employment_OrganizationName,Employment_Employer_Status,Employment_Primary)
		
        'Enter Encounter_Admission_Source_Type
        Call Fn_Enter_Checkin_Summary_Encounter_Admission_Source_Type(Checkin_Admission_Source,Checkin_Admission_Type)
        
		If Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Checkin_PatientCommConsent").Exist(2) Then
			If Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Checkin_PatientCommConsent").GetROProperty("visible")=True Then
		    	Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Checkin_PatientCommConsent").SetEdit DataTable.Value("Encounter_Additionadata_PatientCommConsent")
		    End If 
		End If

		
'        Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Healthix_consent").SetEdit DataTable.Value("Healthix_consent")
'        Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Healthix Consent Date").SetEdit DataTable.Value("HealthixConsent_Date")
'        wait (2)
'        If Browser("SoarianDEV").Dialog("Message from webpage").WinButton("OK").Exist(2) Then
'        	Browser("SoarianDEV").Dialog("Message from webpage").WinButton("OK").Click
'        End If
        

        
        ' Enter "doctor emergency" in the "Attending "Health professional edit field.       
		If HealthProf_LastName<>"" Then
			Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").Image("Checkin_HealthProf_Attendingfindicon").Highlight 'Object.focus
			Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").Image("Checkin_HealthProf_Attendingfindicon").Click
			wait 2
			If Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Find a Health Professional").WebEdit("SESummary_LastName").Exist(5) Then		
				Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Find a Health Professional").WebEdit("SESummary_LastName").Set HealthProf_LastName
				If HealthProf_FirstName<>"" Then
					Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Find a Health Professional").WebEdit("SESummary_FirstName").Set HealthProf_FirstName	
				End If
				
				Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Find a Health Professional").WebEdit("Speciality").Set Encounter_HealthProf_Speciality
				Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Find a Health Professional").WebButton("Search").Click
			End  If
			wait(2)
			Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Find a Health Professional").WebButton("Select").Click
			wait(2)
			If Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Required Healthcare Professional Information Needed").WebEdit("HealthProf_Specialty").Exist(3) Then
			
				call EnterValue_InSoarianDropdown(Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Required Healthcare Professional Information Needed").WebEdit("HealthProf_Specialty"),Encounter_HealthProf_Speciality)
	'			Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Required Healthcare Professional Information Needed").WebEdit("HealthProf_Specialty").Set Encounter_HealthProf_Speciality
				Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Required Healthcare Professional Information Needed").WebButton("OK").Click
				wait(1)
			End If
		End If

        'Enter Encounter_BenefitsAssignmentcertification,Encounter_ReleaseOfInformation
        Call Fn_Enter_Checkin_Summary_Benefits_ReleaseOfInfo(Encounter_BenefitsAssignmentcertification,Encounter_ReleaseOfInformation)
'        Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Healthix_consent").SetEdit DataTable.Value("Healthix_consent")
'        Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Healthix Consent Date").SetEdit DataTable.Value("HealthixConsent_Date")
'        wait(2)
'        If Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Healthix Consent Date").GetROProperty("value") <> DataTable.Value("HealthixConsent_Date") Then
'        	Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Healthix Consent Date").SetEdit DataTable.Value("HealthixConsent_Date")
'        End If
        
		Call Fn_Checkin_Healthixconsent(Healthix_consent,HealthixConsent_Date)
		Call Fn_Checkin_EnterHippaDetails(Hippa_Stmt_Provider,Hippa_Patient_Signature,Hippa_Date_Provided)
        

'        Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Healthix_consent").SetEdit DataTable.Value("Healthix_consent")
'        Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Healthix Consent Date").SetEdit DataTable.Value("HealthixConsent_Date")
'        wait (2)
'        
'        If Browser("SoarianDEV").Dialog("Message from webpage").WinButton("OK").Exist(2) Then
'        	Browser("SoarianDEV").Dialog("Message from webpage").WinButton("OK").Click
'        	Browser("SoarianDEV").Page("Soarian").Frame("Checkin_Summary_BPOForm").WebEdit("Healthix Consent Date").Set DataTable.Value("HealthixConsent_Date") 'FormatDate(DataTable.Value("HealthixConsent_Date"),"MM/dd/yyyy")
'        End If
'		Call Fn_Checkin_Healthixconsent(Healthix_consent,HealthixConsent_Date)
'		Call Fn_Checkin_EnterHippaDetails(Hippa_Stmt_Provider,Hippa_Patient_Signature,Hippa_Date_Provided)
         If Ucase(patient_GuarantorList)<>"SELF" and Ucase(patient_GuarantorList)<>"" Then
        	Call Fn_Checkin_GaurantorSection(patient_GuarantorList,Encounter_GaurantorLastname,Encounter_GaurantorFirstname,Encounter_GaurantorGender,Encounter_GaurantorDOB,Encounter_PatientIs)
        End If
        
        	'Add_Care_Provider
	Call Fn_Patient_Add_Care_Provider(CareProvider_FirstName,CareProvider_LastName,CareProvider_Role_Of_PatientCare,CareProvider_Speciality,CareProvider_Startdate,CareProvider_EndDate,CareProvider_Notprovided)
	
        Call Fn_Enter_Accident_AttornetDetails(Accident_Attornet_Populated,Encounter_Incident_StartDate,Incident_StartTime,Encounter_IncidentType,Incident_NoFaultInsurance,Incident_Description,Encounter_Incident_State,Incident_Attorney_FirstName,Incident_Attorney_Lastname,Incident_Attorney_Email,Incident_Attorney_Phone,Incident_Attorney_Firmname,Incident_Attorney_Street,Incident_Attorney_City,Incident_Attorney_Zip,Incident_Attorney_State)
'		'**********************Demographic**************************************
        Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Patient Demographics_tab").Click
        Browser("SoarianDEV").Page("Soarian").Sync
        wait(2)
        Browser("SoarianDEV").Page("Soarian").Frame("PatientDemographicsForm").WebEdit("Patient_ServedMilitary").Exist(5)
        	
		
		wait(2)
         If Patient_ServedMilitary<>"" Then
            Call EnterValue_InSoarianDropdown(Browser("SoarianDEV").Page("Soarian").Frame("PatientDemographicsForm").WebEdit("Patient_ServedMilitary"),Patient_ServedMilitary)
       	    wait(2)
       	 End if

        Call Fn_Enter_Demographic_DOB(patient_dob)
   
       If patient_middleName<>"" Then
 	      Browser("SoarianDEV").Page("Soarian").Frame("PatientDemographicsForm").WebEdit("Patient_middleName").SetEdit patient_middleName
       End If
        'Add details:Demographic_Address_And_Phone_Section
        
        'Call Fn_Enter_Demographic_Address_And_Phone_Section(patient_addressStreet,patient_addressZip,patient_PhoneTypeEditComboBox,patient_PhoneNumber,Patient_Homeless,Patient_IfNoEmail,Patient_Email,Patient_EmailBelongs)
        Call Fn_Enter_Demographic_USA_Or_Foreign_Address(patient_ForeignAddress_Country,patient_addressStreet,patient_addressZip,patient_addressState,patient_addressCounty,patient_addressCity)
        
       
        'Add details:Demographic_Address_And_Phone_Section
        Call Fn_Enter_Demographic_Phone_Section(patient_PhoneTypeEditComboBox,patient_PhoneNumber,Patient_Homeless,Patient_IfNoEmail,Patient_Email,Patient_EmailBelongs,patient_PhoneCheckbox)
        
        If patient_ApptReminderText <>"" Then
			Browser("SoarianDEV").Page("Soarian").Frame("PatientDemographicsForm").WebEdit("ApptReminder").SetEdit patient_ApptReminderText
		End If
        
        'Add details:Personal_Information_Section
        Call Fn_Enter_Demographic_Personal_Information_Section("",patient_SSN,patient_SSNReason,patient_PreferredLanguage,patient_MaritalStatus,Patient_Race,Patient_Ethnicity,Patient_ReligiousAffilialition)
        
        ''Sandeep 05/08/2018: Commenting below function call to hanlde Emergency contact warning in "Check-in Summary" screen (After Soarian upgradation)
'        'Add NOK	
'		 call Fn_AddContact_EnterAPerson(NOK_FirstName,NOK_LastName,NOK_Patient_Is,NOK_Contact_Type,NOK_PreferredNum,NOK_EmergNotprovided)

		wait(5)    

'        'Add_Patient_EmploymentDetails
'		Call Fn_Add_Patient_EmploymentDetails(Employment_OrganizationName,Employment_Employer_Status,Employment_Primary)
			
		'Enter Patient_Emp_Occupation
		Call Fn_Enter_Demographic_Employment_Section(Patient_NorthwellEmployee,Patient_Emp_Occupation)
		
        Call Fn_Add_HIPAAPrivacyStatement(Hippa_Stmt_Provider,Hippa_Date_Provided,Hippa_Patient_Signature)
        CAll Fn_Healthix_consent(healthix_consent)
		Call Fn_Healthixconsent_Date(HealthixConsent_Date)
'		Call Fn_Add_HIPAAPrivacyStatement(Hippa_Stmt_Provider,Hippa_Date_Provided,Hippa_Patient_Signature)
		 Call Fn_EnterMedicareAB_InsuranceDetails(Insurance_MedicareA,Insurance_MedicareB,Insurance_MedicarePolicyNumber,EffectiveDate_MedicareA,EffectiveDate_MedicareB,Insurance_Verifiedforencounter,Insurance_SubscriptionType,Insurance_firstName,Insurance_lastName)
		 Call Fn_EnterInsuranceDetails(patient_firstName,patient_lastName,Insurance_PayerPlanQuickPick,Insurance_PolicyNumberInput,Insurance_SubscriptionType,Insurance_NotQuickPlanSearch,Insurance_NotQuickPlan,Insurance_Verifiedforencounter)
	     
	     
'	     If Insurance_PatientIs<>"SELF" and Insurance_PatientIs<>"" Then
'	     	Call Fn_EnterInsuranceDetails("","",Insurance_PayerPlanQuickPick,Insurance_PolicyNumberInput,Insurance_SubscriptionType,Insurance_NotQuickPlanSearch,Insurance_NotQuickPlan,Insurance_Verifiedforencounter)
'	     Else
'            Call Fn_EnterInsuranceDetails(patient_firstName,patient_lastName,Insurance_PayerPlanQuickPick,Insurance_PolicyNumberInput,Insurance_SubscriptionType,Insurance_NotQuickPlanSearch,Insurance_NotQuickPlan,Insurance_Verifiedforencounter)	     
'	     End If
	     
	     If SecInsurance_PayerPlanQuickPick <> "" Then
			wait 2
			If ucase(SecInsurance_PatientIs)<>"SELF" and SecInsurance_PatientIs<>"" Then
			    Call Fn_EnterSecInsuranceDetails("","",SecInsurance_PayerPlanQuickPick,SecInsurance_PolicyNumberInput,SecInsurance_SubscriptionType,SecInsurance_Verifiedforencounter,SecGroupNumber)
			Else
		        Call Fn_EnterSecInsuranceDetails(patient_firstName,patient_lastName,SecInsurance_PayerPlanQuickPick,SecInsurance_PolicyNumberInput,SecInsurance_SubscriptionType,SecInsurance_Verifiedforencounter,SecGroupNumber)	     
			End If
			     
			    Call Fn_Verify_InsuranceScreen()
			'Enter subscriber details
			    Call Fn_Enter_Insurance_Subscriber(SecInsurance_Subscriber_Lastname,SecInsurance_Subscriber_Firstname,SecInsurance_PatientIs)    
			    Call Fn_Enter_INSClaimAddress(InsPolicyContact_street,InsPolicyContact_city,InsPolicyContact_zip,InsPolicyContact_state)
			    Call Fn_Enter_Insurance_Authorization(SecInsurance_Hosp_Auth_Type,SecInsurance_Hosp_Auth_No,SecInsurance_Hosp_Auth_ApprovalStatus,SecInsurance_Hosp_DurationCnt,SecInsurance_Hosp_DurationType)
				  
		End If
	     Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Encounter Details Tab").Click
	     If not Browser("SoarianDEV").Page("Soarian").Frame("EncounterDetailsForm").WebButton("Encounter_Add/Edit Assignments").Exist then
	     	setting.WebPackage("ReplayType")=2
	     	Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Encounter Details Tab").Click
	     	setting.WebPackage("ReplayType")=1
	     End If
	     
	     AdmitDate=Browser("SoarianDEV").Page("Soarian").Frame("EncounterDetailsForm").WebEdit("EncounterStartDate").getroproperty("value")
		 AdmitTime=Browser("SoarianDEV").Page("Soarian").Frame("EncounterDetailsForm").WebEdit("Encounter_AdmitTime").getroproperty("value")
		 DataTable.Value("Out_EncounterDetails_AdmitDateTime")=AdmitDate&" "&AdmitTime

	 	  'Health Prof details
	     If HealthProf_Admitting_LastName<>"" Then
	     	Call Fn_Encounter_Add_HealthProfessional(Encounter_HealthProf_Admitting,"","",Encounter_HealthProf_Startdate,Encounter_HealthProf_Time,HealthProf_Admitting_LastName,Encounter_HealthProf_AdmittingSpeciality,HealthProf_Admitting_FirstName)
	     End If
	     
	     If HealthProf_ReferringForm_LastName<>"" Then
	        Call Fn_Encounter_Add_HealthProfessional("","",Encounter_HealthProf_ReferringForm,Encounter_HealthProf_Startdate,Encounter_HealthProf_Time,HealthProf_ReferringForm_LastName,Encounter_HealthProf_ReferringFormSpeciality,HealthProf_ReferringForm_FirstName)		
	     End If
	     
	     
		If Encounter_HealthProf_ReferringName<>"" and (Ucase(Trim(Encounter_HealthProf_ReferringName))<>"NA" or  Ucase(Trim(Encounter_HealthProf_ReferringName))<>"N/A") Then
	     	Call Fn_Encounter_Add_HealthProfessional("","",Encounter_HealthProf_ReferringForm,Encounter_HealthProf_Startdate,Encounter_HealthProf_Time,Encounter_HealthProf_ReferringName,Encounter_Referring_Speciality,HealthProf_ReferringFirstName)
	     End If	
	   
	     Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Next").Click
	     wait(3)
	     Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Next").Click
	     wait(3)
	    'Finishing steps
	    If NOT Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Finish Check-In").WebElement("Finish Check-in").Exist(5) Then
	       	Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Next").Click
         	wait(3)
         End IF
         Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Finish Check-In").WebElement("Finish Check-in").WaitProperty "visible",True
         If Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Finish Check-In").WebElement("Finish Check-in").Exist(15) Then
       	    Browser("SoarianDEV").Window("SoarianSearch Win").Page("Soarian").Frame("Finish Check-In").WebButton("Yes").Click
       	 Else
       	    Reporter.ReportEvent micFail,"Verify if Finish checkin popup is displayed","Finish popup not displayed"
         End If 
         
         If Browser("SoarianDEV").Window("Soarian -- Webpage Dialog_2").Page("Health_ProfessionalDetails").Frame("Final").WebButton("Yes").Exist Then
         	Browser("SoarianDEV").Window("Soarian -- Webpage Dialog_2").Page("Health_ProfessionalDetails").Frame("Final").WebButton("Yes").Click
         End If
         
		 wait (2)
           'Click on done button
         If Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").Exist(10) Then
         
            titleProperty= Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").GetTOProperty("title")
		    nameProperty=Fn_Returns_NameProperty_ForFrame(Browser("SoarianDEV").Page("Soarian"),titleProperty)
		    Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").SetTOProperty "name",nameProperty
		    
            If Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").WebElement("MRN").Exist Then
               MRN=Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").WebElement("MRN").GetROProperty("innertext")
               ECD=Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").WebElement("ECD").GetROProperty("innertext")
               ENC=Browser("SoarianDEV").Page("Soarian").Frame("PrintArtifactsBPOForm").WebElement("ENC").GetROProperty("innertext")
               MRNTemp=Trim(Replace(MRN,"| MR#:",""))
               DataTable.Value("PatientMRN","Global")=Trim(Replace(MRNTemp,"|  View Events",""))
               DataTable.Value("PatientECD","Global")=Trim(Replace(ECD,"| ECD#:",""))
               DataTable.Value("PatientENC","Global")=Trim(Replace(ENC,"| Enc#:",""))
            End If
            Call fnTakeSnapShot(Browser("SoarianDEV"),strTestCaseName,"Script_View_MRN")
         	Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").Click
         	Browser("SoarianDEV").Page("Soarian").Frame("BlueFrame").WebButton("Done").Click
         	If DataTable.Value("PatientMRN","Global") <> "NA" AND DataTable.Value("PatientECD","Global") <> "NA" AND DataTable.Value("PatientENC","Global") <> "NA" Then
         		DataTable.Value("Status","Global")="PASS"
         	Else
         		DataTable.Value("Status","Global")="FAIL"
         	End If
'         	DataTable.Value("Status","Global")="PASS"
         Else
            DataTable.Value("Status","Global")="FAIL"
         End If

      End If  

          'logout
		  call Fn_LogoutAndCLose_SoarianApplication("YES")
		
		  EndTime=Now()
		  DataTable.Value("End_Time","Global")=EndTime
		  Call Fn_Delete_File_If_Exists(OutputFile)
		  If sGTAPFlag <> "Y" Then
			  Datatable.Export OutputFile         
			  If Err.Number <> 0 Then
					Reporter.ReportEvent micWarning, "InitAction", "Runtime error: " & Err.Description
	'				InitAction = False
					Err.Clear
					Call DataTable.ExportSheet(OutputFile,"Global")
	'				ExitRun("Fail")
				End If
		  End If
	 		Call fnCreateZipFile(sGScrShotsFolder)
 			Call fnUploadExecutionScreenshot()
        
	End If
Next
Next

Call Fn_Delete_File_If_Exists(OutputFile)

Print GRC
If sGTAPFlag = "Y" Then
          Environment.Value("MRN")=DataTable.Value("PatientMRN","Global")
		  outTable("SC-OUT-MRN") = Environment.Value("MRN")
		  Environment.Value("PatientENC") =  DataTable.Value("PatientENC","Global")
		  outTable("SC-OUT-PatientENC") = Environment.Value("PatientENC")
		  Environment.Value("PatientECD") =  DataTable.Value("PatientECD","Global")
		  outTable("SC-OUT-PatientECD") = Environment.Value("PatientECD")
		  sGErrorMessage  =DataTable.Value("Error_Message","Global")
	Call fnOutputJSONModification()
'	OutputFile = fnCreateAndOutputTxtFile()
Else
	Datatable.Export OutputFile
End If
Call fnUploadToTestLab(OutputFile,strTestSetName)
Call Fn_Delete_File_If_Exists(OutputFile)


' Serialize the out params hashtable to JSON 
TestArgs("OutputJson") = DotNetFactory.CreateInstance("System.Web.Script.Serialization.JavaScriptSerializer", "System.Web.Extensions").Serialize(outTable)
Print TestArgs("OutputJson")

Reporter.ReportEvent micDone,"Output Json",TestArgs("OutputJson")
On Error GoTO 0
