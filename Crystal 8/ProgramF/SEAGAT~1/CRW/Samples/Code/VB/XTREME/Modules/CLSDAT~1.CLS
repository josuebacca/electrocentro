VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Handle the adding/deleting of records by an ADO control.  Since the code
'          for all the forms that allow adding and deleting of records
'          is essentially the same, all the event handling code is located
'          here in order to make debugging and code reading easier
'

Option Explicit

Dim Data1 As Adodc                              ' The "virtual" data control that is assigned
                                                ' a reference to the equivalent ADO data control
                                                ' in the form that uses this class

' *************************************************************
' Set the local Data control to point to the data control on the form
' that is using this class
'
Public Sub SetData(m_Data As Adodc)
    Set Data1 = m_Data
End Sub

' *************************************************************
' Add a new record
'
Public Sub cmdAdd_Click()
    Data1.Recordset.AddNew
End Sub

' *************************************************************
' Delete the currently selected record
'
Public Sub cmdDelete_Click()
    If Data1.Recordset.RecordCount = 0 Then Exit Sub
    With Data1.Recordset
        .Delete
        .MoveNext
        If .EOF Then .MoveLast
    End With
End Sub

' *************************************************************
' Refresh the data set - only really needed for multi user situation
'
Public Sub cmdRefresh_Click()
    Data1.Refresh
End Sub

' *************************************************************
' Save the user's changes
'
Public Sub cmdUpdate_Click()
    On Error GoTo Problem
    Data1.Recordset.Update
    If Err.Number = 0 Then Exit Sub
Problem:
    MsgBox "Unable to update database " & vbNewLine & _
           "Error: " & Err.Description, vbInformation
End Sub

' *************************************************************
' Show a grid version of the same data that is in the form
'
Public Sub cmdGrid_Click()
    On Error GoTo cmdGrid_ClickErr

    g_RecordSource = Data1.RecordSource
    Set g_RecordSet = Data1.Recordset
    frmDataGrid.Show vbModal

    Exit Sub
cmdGrid_ClickErr:
End Sub

' *************************************************************
' Handle errors and allow the user to continue
'
Public Sub Data1_HandleError(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    MsgBox "Data error event hit err:" & Error$(ErrorNumber)
    fCancelDisplay = True     ' Throw away the error
End Sub

' *************************************************************
' When the user changes records, update the caption on the data
' control
'
Private Sub Data1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Data1.Caption = Data1.Recordset.AbsolutePosition
End Sub

