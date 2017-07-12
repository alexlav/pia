
'****************************************************************************************************************************
'Teste neue Nachweisplanung anlegen. Vertragsstandort, Produktionsstandort, Sachnummer auswählen. Kategorie erweitern. 
'Dann anlegen und aus Seitenmenu aufrufen.

'****************************************************************************************************************************

'****************************************************************************************************************************
'Global Variables
'****************************************************************************************************************************
'Name of the automation module
Const G_ModuleName = "PIA_Nachweisplanung"

'Path to folder which contains all function libraries 
'using Version Control System Github
Const G_FolderPathFunctionLibraries = "C:\github\pia1\pia\Framework" 'use this if running locally

PUBLIC WAITTIME_PER_STEP

'****************************************************************************************************************************
'MSG framework functions
'****************************************************************************************************************************
Function load_function_libraries_from_qc ()
	On Error Resume Next
	'Create an object for Quality Center
	Set qcc = QCUtil.QCConnection
	'Create an object for the Quality Center Tree Manager
	Set tm = qcc.TreeManager
	'Create an object for the folder in Quality Center
	Set fld = tm.NodebyPath(G_FolderPathFunctionLibraries)

	LoadAttachmentsOf_QC_Node fld, tm
	
	If Err.Number <> 0 Then
		Reporter.ReportEvent micFail, "Execute Function Library", "Not possible to execute function libraries! Error Description: " & Err.Description
		Reporter.ReportEvent micFail, "Execute Function Library", "Stopping test..."
		ExitTest
	End If
End Function

Sub LoadAttachmentsOf_QC_Node (fld, tm)
	On Error Resume Next

	'Complete path to the folder in Quality Center ("Subject/.../.../folder")
	fld_path = fld.Path
	'Create an object for the attachments in the folder
	Set atts = fld.Attachments
	'Get all attachments of the folder
	Set alist = atts.NewList("")

	'Iterate through attachments
	For Each att In alist
		'Get name of attachment
		attname = att.Name
		'Split name		
		namearray = Split (attname, "_", 4)
		'Get just the real name of the attachment like it is in Quality Center
		name = namearray(UBound(namearray))
	    ExecuteFile "[QualityCenter] " & fld_path & "\" & name
	    
	    Reporter.ReportEvent micDone, "ExecuteFile", "Executed file " & fld_path & "\" & name
	Next
	
	'Get subfolders of the current folder
	Set objFolders = tm.NodeByPath(fld_path).NewList()
	
	'For every subfolder iterate through its subfolders and execute the function libraries in it
	For intIndex = 1 To objFolders.Count
		Set objFolder = objFolders.Item(intIndex)
		Set regfld = tm.NodebyPath(fld_path & "\" & objFolder.Name)
		LoadAttachmentsOf_QC_Node regfld, tm
	Next
	
	If Err.Number <> 0 Then
		Reporter.ReportEvent micFail, "Execute Function Library", "Not possible to execute function libraries! Error Description: " & Err.Description
		Reporter.ReportEvent micFail, "Execute Function Library", "Stopping test..."
		ExitTest
	End If
End Sub

'Load all function libraries from the "lib" folder
Sub ExecuteAttachmentsOfFolder (folder)
	On Error Resume Next

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objSuperFolder = objFSO.GetFolder(folder)
	
	If Err.Number <> 0 Then
		Reporter.ReportEvent micFail, "Execute Function Library", "Not possible to execute function libraries in folder " & objSuperFolder.Path & "! Error Description: " & Err.Description
		Reporter.ReportEvent micFail, "Execute Function Library", "Stopping test..."
		ExitTest
	End If	
	
	ExecuteSubfolders (objSuperFolder)
	
	Set objFSO = Nothing
	On Error GoTo 0
End Sub


'Recursively look up the subfolders of G_FolderPathFunctionLibraries
Sub ExecuteSubfolders (folder)
	On Error Resume Next
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
    Set objFolder = objFSO.GetFolder(folder.Path)
    Set colFiles = objFolder.Files
    For Each objFile in colFiles
        If UCase(objFSO.GetExtensionName(objFile.name)) = "QFL" Then
            ExecuteFile objFolder & "\" & objFile.name
            'Reporter.ReportEvent micDone, "Function ExecuteFile", "Executed file '" & objFolder & "\" & objFile.name & "'"
        End If
        
		If Err.Number <> 0 Then
			Reporter.ReportEvent micFail, "Execute Function Library", "Not possible to execute function library  " & objFile.name & "! Error Description: " & Err.Description
			Reporter.ReportEvent micFail, "Execute Function Library", "Stopping test..."
			ExitTest
		End If
    Next

    For Each Subfolder in folder.SubFolders
        ExecuteSubfolders (Subfolder)
    Next
    
    Set objFSO = Nothing
    On Error GoTo 0
End Sub


'****************************************************************************************************************************
'Main Function to interprete the "Action" Column in the "Process" Sheet
'****************************************************************************************************************************
Function ExecuteTest ()
  Dim  CurrentRow : CurrentRow = 0
  Dim  RetVal

  RetVal = InitTest (G_ModuleName)
  If RetVal <> 0 Then: Exit Function: End If

  Do While True
    msgfw_ReportAction "Process"

    Select Case msgfw_EvaluatedData  ("Action", "Process")
      Case "Run PIA"  				 				RetVal = RunPIA ("Process")
      Case "Set step time"							RetVal = SetWaittime("Process")
      'Case "Neue Nachweisplanung"            	    RetVal = NeueNachweisplanung ("Process")
      Case "Call Neue Nachweisplanung"				RetVal = ComputeSheetNeueNachweisplanung ("NeueNachweisplanung")
      Case "Close PIA"								RetVal = ClosePIA ()
      Case "->"										RetVal = 0
      Case Else                     				RetVal = 1
    End Select
    
    Select Case RetVal
      Case 0                         'No Error; continue execution
      Case 1                         Exit Function    'Function 'ComputeStandardActions' has detected Action 'End'
      Case -1                        Exit Function    'Error while executing a function
    End Select
    
    msgfw_ComputeCurrentRow CurrentRow, "Process"
  Loop
End Function

'****************************************************************************************************************************
'PIA specific functions
'****************************************************************************************************************************

Function RunPIA (SheetName)
	Dim url
	'url2 = msgfw_EvaluatedData ("Data", SheetName)
	url = Parameter.Item("Param1")
	SystemUtil.Run "iexplore.exe",url, "", "" 
	
	wait WAITTIME_PER_STEP

	wait WAITTIME_PER_STEP
		Set wshShell = CreateObject("WScript.Shell")
		Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").Sync
'Need something here to set focus on the browser(.click or something)
		wait WAITTIME_PER_STEP
wshShell.SendKeys "% ", True  'the "alt" % and the character it is modifying should be in same string
		wait WAITTIME_PER_STEP
		wshShell.SendKeys "x"
		wait WAITTIME_PER_STEP
'		wshShell.SendKeys "%x", True  'the "alt" % and the character it is modifying should be in same string
'wshShell.SendKeys "{DOWN 2}"  ' this is a separate second command
'wshShell.SendKeys "{RIGHT}"  ' this is a separate second command
'		wait WAITTIME_PER_STEP
'		wait WAITTIME_PER_STEP
'		wait WAITTIME_PER_STEP
'	wshShell.SendKeys "{UP 4}"  ' this is a separate second command
'	wshShell.SendKeys "{ENTER}"  ' this is a separate second command
'		wait WAITTIME_PER_STEP
	'wshShell.SendKeys "%", True
	'		wait WAITTIME_PER_STEP
	'wshShell.SendKeys "a", True
	'	wait WAITTIME_PER_STEP
	'wshShell.SendKeys "z", True
	'	wait WAITTIME_PER_STEP
	'wshShell.SendKeys "{UP 4}", True 
	'	wait WAITTIME_PER_STEP
	'wshShell.SendKeys "{ENTER}" 
End Function

Function ClosePIA ()
	OpenNeueNachweisplanung	()
	wait WAITTIME_PER_STEP
	WarningDismissOK()
		wait WAITTIME_PER_STEP
	Browser("Parts Inspection and Approval").Close
End Function

Function SetWaittime(SheetName)
	WAITTIME_PER_STEP = msgfw_EvaluatedData ("Data", SheetName)	
End Function

'ComputeSheetNeueNachweisplanung
Function ComputeSheetNeueNachweisplanung (SheetName)
  Dim CurrentGroup, CurrentRow, RetVal

  CurrentGroup = msgfw_EvaluatedData  ("Group",      "process")
  CurrentRow   = SetFirstRowByGroupID (CurrentGroup, SheetName)
  
  Do While msgfw_EvaluatedData ("Group", SheetName) = CurrentGroup
    msgfw_ReportAction SheetName
    
    Select Case msgfw_EvaluatedData ("Action", SheetName)
      Case "Click Neue Nachweisplanung"  				RetVal = OpenNeueNachweisplanung ()
      Case "Click Adresse(Vertragsstandort)"			RetVal = ClickButtonAdresseVS()
      Case "Search Adresse(Vertragsstandort)"			RetVal = SearchAdresseVS()
      Case "Select Adresse(Vertragsstandort)"			RetVal = SelectAdresseVS(msgfw_EvaluatedData ("Index", SheetName))
      Case "Select different Adresse(Vertragsstandort)"	RetVal = SelectAdresseVS(msgfw_EvaluatedData ("Index", SheetName))
      Case "Clear Adresse(Vertragsstandort)"			RetVal = ClearAdresseVS()
      Case "Abort Adresse(Vertragsstandort)"		    RetVal = AbortAdresseVS()
      Case "Send Adresse(Vertragsstandort)"				RetVal = SendAdresseVS()
      Case "Check Adresse(Vertragsstandort)"			RetVal = CheckAdresseVS(SheetName)
      Case "Click Adresse(Produktionsstandort)"         RetVal = ClickButtonAdressePS()
      Case "Search Adresse(Produktionsstandort)"		RetVal = SearchAdressePS()
      Case "Select Adresse(Produktionsstandort)"		RetVal = SelectAdressePS()
      Case "Abort Adresse(Produktionsstandort)"			RetVal = AbortAdressePS()
      Case "Send Adresse(Produktionsstandort)"			RetVal = SendAdressePS()
      Case "Check Adresse(Produktionsstandort)"			RetVal = CheckAdressePS(SheetName)
      Case "Click Teilebenennung"						RetVal = ClickButtonTeilebenennung()
      Case "Search Teilebenennung"						RetVal = SearchTeilebenennung()
      Case "Select Teilebenennung"						RetVal = SelectTeilebenennung()
      Case "Abort Teilebenennung"						RetVal = AbortTeilebenennung()
      Case "Send Teilebenennung"						RetVal = SendTeilebenennung()
      Case "Check Teilebenennung"						RetVal = CheckTeilebenennung(SheetName)
      Case "Click Anlegen"								RetVal = SubmitNeueNachweisplanung ()
      Case "Errormessage-OK Missing Input"				RetVal = HandleErrorMissingInput ()
      Case "Warning-OK Dismiss Changes"					RetVal = WarningDismissOK ()
      Case "Warning-Abort Dismiss Changes"				RetVal = WarningDismissAbort ()
      Case Else                Msgbox "The testsheet contains an unknown action"  
    End Select
    wait WAITTIME_PER_STEP
    If RetVal <> 0 Then
       ComputeSheetNotifications = -1
       QTReport "ERROR", "Function 'ComputeSheetNotifications' was aborted due to an error."
       Exit Function
    End If

    msgfw_ComputeCurrentRow CurrentRow, SheetName
  Loop
End Function

Function OpenNeueNachweisplanung ()
Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("NP_NeueNachweisplanung").Click
End Function


Function SubmitNeueNachweisplanung ()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("NP_Anlegen").Click
End Function

Function HandleErrorMissingInput ()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_6").Click
 End Function
 
 Function WarningDismissOK ()
	If(	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_7").Exist) Then
		Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_7").Click
		End If

 End Function
 
 Function WarningDismissAbort ()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_8").Click

 End Function
'Vertragsstandort
Function ClickButtonAdresseVS()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_0").Click

End Function

Function SearchAdresseVS()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_21").Click

End Function

'Function SelectAdresseVS()
'
'Set oExcptnDetail = Description.Create
'oExcptnDetail("micclass").value = "WebElement"
'oExcptnDetail("html tag").value = "TD"
'Set chobj=Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebTable("suppliersSelectGrid").ChildObjects(oExcptnDetail)
'chobj(0).click
'
''MsgBox chobj(1).GetROProperty("innerhtml") , VBOKOnly, "Test"
''chobj(0).GetROProperty("innerhtml").msgfw_Edit           "Standort",          SheetName
'
'End Function

Function SelectAdresseVS(Index)

Set oExcptnDetail = Description.Create
oExcptnDetail("micclass").value = "WebElement"
oExcptnDetail("html tag").value = "TD"
Set chobj=Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebTable("suppliersSelectGrid").ChildObjects(oExcptnDetail)
'chobj(0).click
For i = 0 To chobj.Count-1 Step 1
	If(chobj(i).GetROProperty("innerhtml")=Index) Then
	chobj(i).click
	Exit For
	End If
Next
End Function

Function checkAdresseVS(SheetName)
'MsgBox Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("contractorLocation.number").GetROProperty("value") , VBOKOnly, "Test"
'Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("contractorLocation.number").ToString
CheckEqualStrings "Standort", msgfw_EvaluatedData ("Standort", SheetName), Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("contractorLocation.number").GetROProperty("value")
CheckEqualStrings "Index", msgfw_EvaluatedData ("Index", SheetName), Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("supplierIndex").GetROProperty("value")
CheckEqualStrings "Adresse", msgfw_EvaluatedData ("Adresse", SheetName), Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("address").GetROProperty("value")
End Function

Function CheckEqualStrings (displayName, expectedValue, actualValue)
	If CStr(expectedValue) = CStr(actualValue) Then
		QTReport "OK", "Check " & displayName & " = '" & actualValue & "'"
	Else 
		QTReport "WARN", "Check " & displayName & " = '" & actualValue & "', expected='" & expectedValue & "'."
		
	End If
End Function

Function ClearAdresseVS()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebElement("Spaltenfilter aktiv").Click

End Function

Function AbortAdresseVS()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_20").Click

End Function

Function SendAdresseVS()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_19").Click

End Function

'Produktstandort
Function ClickButtonAdressePS()

	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_1").Click

End Function

Function SearchAdressePS()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_21").Click

End Function

Function SelectAdressePS()

Set oExcptnDetail = Description.Create
oExcptnDetail("micclass").value = "WebElement"
oExcptnDetail("html tag").value = "TD"
Set chobj=Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebTable("suppliersSelectGrid").ChildObjects(oExcptnDetail)
chobj(0).click
End Function

Function checkAdressePS(SheetName)
'MsgBox Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("contractorLocation.number").GetROProperty("value") , VBOKOnly, "Test"
'Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("contractorLocation.number").ToString
CheckEqualStrings "Standort", msgfw_EvaluatedData ("Standort", SheetName), Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("productionLocation.number").GetROProperty("value")
CheckEqualStrings "Index", msgfw_EvaluatedData ("Index", SheetName), Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("supplierIndex_2").GetROProperty("value")
CheckEqualStrings "Name", msgfw_EvaluatedData ("Adresse", SheetName), Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("address_2").GetROProperty("value")
End Function


Function AbortAdressePS()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_20").Click

End Function

Function SendAdressePS()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_19").Click

End Function

'Teilebennung
Function ClickButtonTeilebenennung()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("part.button").Click

End Function

Function SearchTeilebenennung()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_21").Click

End Function

Function SelectTeilebenennung()

Set oExcptnDetail = Description.Create
oExcptnDetail("micclass").value = "WebElement"
oExcptnDetail("html tag").value = "TD"
Set chobj=Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebTable("suppliersSelectGrid").ChildObjects(oExcptnDetail)
chobj(0).click

End Function

Function checkTeilebenennung(SheetName)
'MsgBox Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("contractorLocation.number").GetROProperty("value") , VBOKOnly, "Test"
'Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("contractorLocation.number").ToString
CheckEqualStrings "Sachnummer", msgfw_EvaluatedData ("Sachnummer", SheetName), Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("partStub").GetROProperty("value")
CheckEqualStrings "ES1", msgfw_EvaluatedData ("ES1", SheetName), Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("es1").GetROProperty("value")
CheckEqualStrings "Benennung", msgfw_EvaluatedData ("Benennung", SheetName), Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebEdit("name").GetROProperty("value")
End Function


Function AbortTeilebenennung()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_20").Click

End Function

Function SendTeilebenennung()
	Browser("Parts Inspection and Approval").Page("Parts Inspection and Approval").WebButton("dijit_form_Button_19").Click

End Function

Function ComputeStandardActions_Application (SheetName, ByRef CurrentRow)
  ComputeStandardActions_Application = 0

  thisAction = msgfw_EvaluatedData ("Action", SheetName)

  If thisAction <> "" Then: G_EmptyRows = 0: End If
  
  Select Case thisAction
    Case "End"   RetVal = 1

    Case Else             RetVal = -1
  End Select
End Function

Set objFSO = CreateObject("Scripting.FileSystemObject")
	
If objFSO.FolderExists(G_FolderPathFunctionLibraries) Then
	'Execute all Function Libraries
	Reporter.ReportEvent micDone, "Loading framework", "from local cache in '" & G_FolderPathFunctionLibraries & "'."
	ExecuteAttachmentsOfFolder(G_FolderPathFunctionLibraries)
Else
	Reporter.ReportEvent micDone, "Loading framework", "from Quality Center."
	load_function_libraries_from_qc ()
End If
Set objFSO = Nothing

Reporter.ReportEvent micDone, "Loading framework", "finished."

 ExecuteTest ()
