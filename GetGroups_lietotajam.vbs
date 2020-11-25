'==========================================================================
' AUTHOR: Steve Knight
' DATE  : 8/1/2014
'
'==========================================================================
' Define constant information and global variables
'==========================================================================
Const ADS_SCOPE_SUBTREE = 2
Dim objConnection, objCommand, rootDSE, AD_Domain
Dim objUser,objAllGroups,objGroup
Dim objExcel, objWorkbook, objSheet, row


'==========================================================================
' Ask user questions
'==========================================================================

strClockCard = trim(InputBox("Enter User login name to find their groups:"))
if strClockCard="" then wscript.quit

'==========================================================================
' Get Active Directory Domain
'==========================================================================

Set rootDSE = GetObject("LDAP://rootDSE")
AD_Domain = rootDSE.Get("defaultNamingContext")

'==========================================================================
' Connect to AD to find requested user
'==========================================================================

dtStart = TimeValue(Now())
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")

objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "<LDAP://" & AD_Domain & ">;(&(objectCategory=User)" & "(samAccountName=" & strClockCard & "));samAccountName;subtree"

Set objRecordSet = objCommand.Execute

If objRecordset.RecordCount = 0 Then 'User doesn't exist
  MsgBox "User " & strClockCard & " doesn't exist in Active Directory"
  wscript.quit

Else
'==========================================================================
' This is an existing user to search for and work with
'==========================================================================

  objCommand.CommandText = "SELECT distinguishedName,cn FROM 'LDAP://" & AD_Domain & "' WHERE objectClass='user' " & "and sAMAccountName='" & strClockCard & "'"  
  objCommand.Properties("Page Size") = 1000
  objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
  Set objRecordSet = objCommand.Execute
  objRecordSet.MoveFirst
  
  objDistinguishedName = objRecordSet.Fields("distinguishedName").Value
  objFullName = objRecordSet.Fields("cn").Value
  Set objUser = GetObject("LDAP://" & objDistinguishedName)

'==========================================================================
' Create new Excel sheet with user details and list of their groups
' Auto-fit the columns and sort by the group name
'==========================================================================

  Set objExcel = CreateObject("Excel.Application")
  Set objWorkbook = objExcel.WorkBooks.Add()
  Set objSheet = objWorkBook.WorkSheets(1)

  objWorkBook.Application.Width=800
  objWorkBook.Application.Height=700

  objExcel.Visible = True

  With objSheet
    with .Cells(1,1)
      .value="Groups report for " & objUser.Fullname & " (" & strClockCard & ")" & vbCrLf
      .WrapText=False
      .Font.FontStyle = "Bold"
      .Font.Size = 12
      .Font.Underline = 2
    end with

    .Cells(3,1) = "Date checked:"
    .Cells(3,2) = now()

    .Cells(4,1) = "User:"
    .Cells(4,2) = strClockCard
  
  
    .Cells(5,1) = "Full Name:"
    .Cells(5,2) = objUser.Fullname
  
    .Cells(6,1) = "Description:"
    .Cells(6,2) = objUser.Description
 
    .Cells(8,1) = "Member of these groups:"

    row=10
    GetGroups

    .Columns("A:B").Entirecolumn.AutoFit
    .Columns("A:B").Entirecolumn.HorizontalAlignment = -4131 'xlLeft

    .Range("A10:B" & row).Sort .Range("A10"), 1


    With .PageSetup
      .PrintArea = "$A$1:$B$" & row
      .Zoom = False
      .FitToPagesWide = 1
      .FitToPagesTall = 2
    End With

  End With 

End If

objConnection.Close
wscript.quit


'==========================================================================
' Get group memberships for this user
'==========================================================================
Sub GetGroups

  set objGroups=objUser.Groups

  if IsEmpty(objGroups) then
    ' No groups for this user
  elseif TypeName(objGroups)="String" then
    ' Only one group
    objSheet.cells(row,1)=objGroups.CN
    objSheet.cells(row,2)=objgroups.description
  else
    For Each objGroup in objGroups
      objSheet.cells(row,1)=objGroup.CN
      objSheet.cells(row,2)=objgroup.description
      row=row+1
    Next
  end if

End Sub