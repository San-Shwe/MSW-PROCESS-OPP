dim app, wks, db, strBE, strFE

' strFE = CurrentProject.Path & "\" & process_pvc.accdb 'to use in project app

strDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
strFE = strDir & "\process_opp.accdb"
strBE = strDir & "\process_opp_be.accdb"

'strFE =  "C:\Users\Kaizaki Arata\Desktop\MSW-Process-PVC\process_pvc.accdb"
'strBE =  "\\process-pvc\e\10. Process PVC Database\process_pvc_be.accdb"
strUserAuthBE = "C:\Kaizaki\Document\MS Access\MSW-UserAuthentication\UserAuthorization_be.accdb"

set app = CreateObject("Access.Application")
set wks = app.dbEngine.createworkspace("", "admin", "", 2) 'dbUseJet
set db = wks.opendatabase(strFE)

LinkMyTable db, strBE, "BMD", "BMD"
LinkMyTable db, strBE, "FG", "FG"
LinkMyTable db, strBE, "Inspection Of LMN", "Inspection Of LMN"
LinkMyTable db, strBE, "Inspection Of Print", "Inspection Of Print"
LinkMyTable db, strBE, "OF", "OF"
LinkMyTable db, strBE, "Print", "Print"
LinkMyTable db, strBE, "LMN", "LMN"
LinkMyTable db, strBE, "SLT", "SLT"
LinkMyTable db, strBE, "tblOPPSummary", "tblOPPSummary"

LinkMyTable db, strUserAuthBE, "tblEmployeeAccess", "tblEmployeeAccess"
LinkMyTable db, strUserAuthBE, "tblEmployees", "tblEmployees"
LinkMyTable db, strUserAuthBE, "tblUser", "tblUser"

db.close
app.quit
msgbox "Your Application was updated", vbInformation, "Updated"
'cleanup 
set db=Nothing
set wks = Nothing
set app = Nothing

Sub DeleteOldTable(db, strTable)
	'Delete an old table 
	on Error Resume Next
	db.tabledefs.delete strTable
end sub

sub LinkMyTable(db, strBE, strSource, strAlias)
	'Link a table to a backend file 
	dim strConnect 
	dim tdf 
	DeleteOldTable db, strAlias
	strConnect	 = ";DATABASE=" & strBE
	Set tdf = db.createtabledef(strAlias, 0, strSource, strConnect)
	db.tabledefs.Append tdf
	set tdf = Nothing
	
End sub
