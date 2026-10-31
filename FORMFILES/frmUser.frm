VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmUser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "User Access"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flxAddMenu 
      Height          =   1455
      Left            =   6720
      TabIndex        =   16
      Top             =   6600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxEmp 
      Height          =   1455
      Left            =   2520
      TabIndex        =   15
      Top             =   6960
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2566
      _Version        =   393216
      FixedRows       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxAccess 
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Tag             =   "xyf"
      Top             =   2520
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   6
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      FormatString    =   "     |<TRANSACTIONS                 |^FULL ACCESS        |^VIEW ONLY            |^DISABLED           |^"
   End
   Begin VB.Frame frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7935
      Begin VB.TextBox txtCancelID 
         Height          =   285
         Left            =   5880
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkAdmin 
         Alignment       =   1  'Right Justify
         Caption         =   "Administrator"
         Height          =   375
         Left            =   5640
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtConfirmPW 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   8
         Tag             =   "xyf"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtPW 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   7
         Tag             =   "xyf"
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtFullName 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label lblCancel 
         Caption         =   "CANCEL ID:"
         Height          =   255
         Left            =   4800
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   975
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "CONFIRM PASSWORD:"
         Height          =   495
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "PASSWORD:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   975
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "USER NAME:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   615
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "FULL NAME:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   255
         Width           =   1695
      End
   End
   Begin lvButton.lvButtons_H cmdAddNew 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Add New"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmUser.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   7605
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmUser.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   1600
      TabIndex        =   21
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Edit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmUser.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   4605
      TabIndex        =   22
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "E&xit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmUser.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   375
      Left            =   3100
      TabIndex        =   18
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Find"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmUser.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmUser.frx":1C02
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   240
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   8415
      m_MDI_Height    =   6885
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddNew      As Boolean
Dim ActRow      As Integer, ActCol As Integer
Dim CustSrch    As Boolean, CustBln As Boolean
Dim rsEmployee  As New ADODB.Recordset


Public Sub menuTrans()
    Dim rsMenu As New ADODB.Recordset
    
    Qry = "SELECT menu.MENUID, Menu.MENUNAME From Menu order by menu.MENUID"
  
    With rsMenu
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxAccess.Rows = 1
                Do While Not .EOF
                    
                    flxAccess.AddItem vbTab & !MENUNAME & vbTab & vbTab & vbTab & vbTab & !MENUID
                    .MoveNext
                Loop
        End If
        .Close
                
    End With
End Sub

Public Sub ShowRecord(MyID As Integer)
    Dim rsUser      As New ADODB.Recordset
    Dim rsAcc       As New ADODB.Recordset
    Dim rsAddMenu   As New ADODB.Recordset
    Dim Qry         As String
    Dim maxMID      As Integer
    Dim i           As Integer
    
    'Qry = "SELECT MAX(UserAccess.MENUID) As maxID From UserAccess Where UserAccess.EmpId = " & MyID & ";"
'    maxMID = Val(GetFieldValue("SELECT MAX(UserAccess.MENUID) As maxID From UserAccess Where UserAccess.EmpId = " & MyID & ";"))
'    Qry = "SELECT  menu.MENUID, Menu.MENUNAME From Menu Where menu.MENUID > " & maxMID & " ORDER BY MENUID;"
'    With rsAddMenu
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            flxAddMenu.Rows = 1
'            Do While Not .EOF
'                flxAddMenu.AddItem !MENUID & vbTab & !menuname
'            .MoveNext
'            Loop
'        End If
'        .Close
'    End With
'    'If flxAddMenu.TextMatrix > 1 Then
'        With flxAddMenu
'            For i = 1 To .Rows - 1
'                If .TextMatrix(i, 0) <> "" And .TextMatrix(i, 1) <> "" Then
'                    LYRcon.Execute "INSERT INTO useraccess(EMPID, MENUID, ACCESSTYPE) " & _
'                                    "VALUES(" & MyID & ", " & .TextMatrix(i, 0) & ", 3 );"
'                End If
'            Next i
'            .Rows = 1
'            .Rows = 2
'        End With
'    'End If

    Qry = "SELECT employees.EMPID, concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPLOYEE," & _
            "employees.PASSWORD, employees.USERNAME " & _
          "FROM EMPLOYEES WHERE employees.EMPID=" & MyID & ";"

    With rsUser
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtFullName.Text = !EMPLOYEE
            txtUser.Text = !UserName
            txtPW.Text = !Password
            txtConfirmPW.Text = !Password
            txtUserID.Text = !EmpId
        Else
            .Close
            Exit Sub
        End If
        .Close
    End With
    
    Qry = "SELECT menu.MENUID, menu.MENUNAME, useraccess.EMPID, UserAccess.ACCESSTYPE " & _
        "From Menu INNER JOIN useraccess ON (menu.MENUID = useraccess.MENUID) Where UserAccess.EmpId = " & MyID & " order by menu.menuid"
        
    With flxAccess
        rsAcc.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsAcc.RecordCount > 0 Then
            rsAcc.MoveFirst
            .Rows = 1
            Call menuTrans
            Do While Not rsAcc.EOF
                With flxAccess
                    For i = 1 To .Rows - 1
                        If rsAcc!ACCESSTYPE = 1 And .TextMatrix(i, 1) = rsAcc!MENUNAME Then
                            .row = i: .Col = 2: .CellFontName = "Webdings"
                            .TextMatrix(i, 2) = "a"
                        ElseIf rsAcc!ACCESSTYPE = 2 And .TextMatrix(i, 1) = rsAcc!MENUNAME Then
                            .row = i: .Col = 3: .CellFontName = "Webdings"
                            .TextMatrix(i, 3) = "a"
                        ElseIf rsAcc!ACCESSTYPE = 3 And .TextMatrix(i, 1) = rsAcc!MENUNAME Then
                            .row = i: .Col = 4: .CellFontName = "Webdings"
                            .TextMatrix(i, 4) = "a"
                        End If
                    Next i
                End With
                rsAcc.MoveNext
            Loop
        Else
            Call menuTrans
             With flxAccess
                .row = 1: .Col = 2
                .RowSel = .Rows - 1: .Col = .Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontName = "Webdings"
                    For i = 1 To .Rows - 1
                        .Col = 3: .row = i
                        .CellFontName = "Webdings"
                        .TextMatrix(i, 3) = "a" 'Label2.Caption
                    Next i
                .row = 1: .Col = 0
            End With
        End If
        rsAcc.Close
    End With

    
    Frame1.Enabled = False
    flxAccess.Enabled = False
End Sub

Private Sub ClearEntries()
    Dim i As Integer, j As Integer
    
    txtFullName.Text = ""
    txtUser.Text = ""
    txtPW.Text = ""
    txtConfirmPW.Text = ""
    chkAdmin.Value = 0
    txtCancelID.Text = ""
    txtUserID.Text = ""
    With flxAccess
        For i = 1 To .Rows - 1
            For j = 2 To 4
                .TextMatrix(i, j) = ""
            Next j
        Next i
    End With
    
End Sub

Private Sub BtnSaveCancel(Bln As Boolean)
    If Bln Then
        cmdSave.Left = cmdAddNew.Left
        cmdCancel.Left = cmdEdit.Left
        cmdSave.Visible = True
        cmdCancel.Visible = True
        cmdAddNew.Visible = False
        cmdEdit.Visible = False
        cmdFind.Visible = False
        cmdExit.Visible = False
    Else
        cmdSave.Visible = False
        cmdCancel.Visible = False
        cmdAddNew.Visible = True
        cmdEdit.Visible = True
        cmdFind.Visible = True
        cmdExit.Visible = True
    End If
End Sub


Private Sub chkAdmin_Click()
    If chkAdmin.Value = 1 Then
        lblCancel.Visible = True
        txtCancelID.Visible = True
    Else
        lblCancel.Visible = False
        txtCancelID.Visible = False
    End If
End Sub

Private Sub cmdAddNew_Click()
    Dim i As Integer, Qry As String
    
'    Qry = "SELECT useraccess.ACCESSTYPE From useraccess Where useraccess.EMPID = " & UserID & " AND useraccess.TRANSACTION = 'Users';"
'    If Val(GetFieldValue(Qry)) <> 1 Then
'        MsgBox "You don't have Full Access to this transaction.", vbExclamation + vbOKOnly, "Access Denied"
'        Exit Sub
'    End If

     Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUID = 2;"

    MyAcc = GetFieldValue(Qry)
    If MyAcc <> 1 Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    Call menuTrans
    Call ClearEntries
    Frame1.Enabled = True
    flxAccess.Enabled = True
    AddNew = True
    With flxAccess
        .row = 1: .Col = 2
        .RowSel = .Rows - 1: .Col = .Cols - 1
        .FillStyle = flexFillRepeat
        .CellFontName = "Webdings"
        For i = 1 To .Rows - 1
            .Col = 3: .row = i
            .CellFontName = "Webdings"
            .TextMatrix(i, 3) = "a" 'Label2.Caption
        Next i
        .row = 1: .Col = 0
    End With
'    Qry = GetFieldValue("SELECT settings.VALUE1 From settings Where settings.DESCRIPTION = 'Administrator'")
'    If Qry = "" Then
'        chkAdmin.Enabled = True
'    Else
'        chkAdmin.Enabled = False
'    End If
    lblCancel.Visible = False
    txtCancelID.Visible = False
    Call BtnSaveCancel(True)
    txtFullName.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    Dim ans As Integer
    
    ans = MsgBox("Do you want to Cancel user entry?", vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    Call BtnSaveCancel(False)
    Call ClearEntries
    Frame1.Enabled = False
    flxAccess.Enabled = False
End Sub

Private Sub cmdEdit_Click()
    Dim Qry As String
  
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUID = 2;"

    MyAcc = GetFieldValue(Qry)
    If MyAcc <> 1 Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    If txtUserID.Text = "" Then
        MsgBox "You dont have data to edit", vbCritical, "EDIT"
        Exit Sub
    End If
    
    Call BtnSaveCancel(True)
    Frame1.Enabled = True
    flxAccess.Enabled = True
    AddNew = False
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdFind_Click()
    Set frmSearch.FrmSource = frmUser

    'frmSearch.QrySearch = "SELECT pcuser.EMPID, pcuser.FULLNAME From pcuser Order By pcuser.FULLNAME;"
    frmSearch.QrySearch = "SELECT EMPLOYEES.EMPID, concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPLOYEE " & _
                "From EMPLOYEES WHERE employees.USERNAME IS NOT NULL AND employees.PASSWORD IS NOT NULL " & _
                " ORDER BY concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.')));"
                
             
    frmSearch.lblCap.Caption = "Search for User"
    frmSearch.Show vbModal

End Sub

Private Sub cmdSave_Click()
    Dim Qry As String, vMsg As String, LstUpdt As String
    Dim MyAcc As Integer, ans As Integer, flagUser As String
    
    'Validate Entry
    vMsg = ""
    If txtFullName.Text = "" Then
        vMsg = "Incomplete entry on field Full Name"
    End If
    If txtUser.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field User Name"
    Else
        If InStr(txtUser.Text, "'") > 0 Then txtUser.Text = Replace(txtUser.Text, "'", "`")
    End If
    If txtPW.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field Password"
    Else
        If txtPW.Text <> txtConfirmPW.Text Then vMsg = vMsg & vbCrLf & "Invalid Password, make sure you enter same Password"
    End If
    If chkAdmin.Value = 1 And txtCancelID.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Password"
    
    If vMsg <> "" Then
        MsgBox vMsg, vbExclamation + vbOKOnly, "Warning"
        Exit Sub
    End If
    ans = MsgBox("Do you want to Save User Access?", vbQuestion + vbYesNo, "Save")
    If ans = vbNo Then Exit Sub
    LstUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
    If AddNew Then
'        LYRcon.Execute "INSERT INTO pcuser(FULLNAME, USERNAME, PASSWORD, UPDATEDBY, LASTUPDATE) " & _
'        "VALUES('" & txtFullName.Text & "', '" & txtUser.Text & "', '" & txtPW.Text & "', " & UserID & ", '" & LstUpdt & "');"
'        Qry = "SELECT pcuser.EMPID From pcuser Where pcuser.FULLNAME = '" & txtFullName.Text & "' AND pcuser.LASTUPDATE = '" & LstUpdt & "'"
'        txtUserID.Text = GetFieldValue(Qry)
            LYRcon.Execute "Update Employees SET USERNAME = '" & txtUser.Text & "', Password = '" & txtPW.Text & "' " & _
                            "Where Employees.EmpId = '" & Val(txtUserID.Text) & "';"
    
    Else
'        LYRcon.Execute "Update pcuser SET FULLNAME = '" & txtFullName.Text & "', USERNAME = '" & txtUser.Text & "', PW = '" & txtPW.Text & "', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & LstUpdt & "' " & _
'        "Where pcuser.EMPID = " & Val(txtUserID.Text) & ";"
            LYRcon.Execute "Update Employees SET USERNAME = '" & txtUser.Text & "', Password = '" & txtPW.Text & "' " & _
                            "Where Employees.EmpId = '" & Val(txtUserID.Text) & "';"
    
    End If
        
        Qry = "SELECT DISTINCT UserAccess.EmpId From UserAccess Where UserAccess.EmpId = '" & Val(txtUserID.Text) & "' Group By EmpId"
        flagUser = GetFieldValue(Qry)
        If flagUser <> "" Then
            LYRcon.Execute "DELETE From UserAccess Where UserAccess.EmpId = '" & Val(txtUserID.Text) & "';"
        End If
    With flxAccess
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) <> "" Then
                MyAcc = 1
            ElseIf .TextMatrix(i, 3) <> "" Then
                MyAcc = 2
            ElseIf .TextMatrix(i, 4) <> "" Then
                MyAcc = 3
            End If
            
            LYRcon.Execute "INSERT INTO useraccess(EMPID, MENUID, ACCESSTYPE) " & _
               "VALUES(" & Val(txtUserID.Text) & ", '" & .TextMatrix(i, 5) & "', " & MyAcc & ");"



'            If AddNew Then
''                LYRcon.Execute "INSERT INTO useraccess(EMPID, TRANSACTION, ACCESSTYPE, SEQNO) " & _
''                "VALUES(" & Val(txtUserID.Text) & ", '" & .TextMatrix(i, 1) & "', " & MyAcc & ", " & i & ");"
'                LYRcon.Execute "INSERT INTO useraccess(EMPID, MENUID, ACCESSTYPE) " & _
'                "VALUES(" & Val(txtUserID.Text) & ", '" & .TextMatrix(i, 5) & "', " & MyAcc & ");"
'            Else
'                LYRcon.Execute "UPDATE useraccess SET ACCESSTYPE = " & MyAcc & " Where useraccess.EMPID = " & Val(txtUserID.Text) & " AND useraccess.menuid = '" & .TextMatrix(i, 5) & "';"
'            End If
'            If Not AddNew And UserID = Val(txtUserID.Text) Then
'                Select Case .TextMatrix(i, 1)
'                    Case "Receiving"
'                        If MyAcc = 3 Then Materials.mnuRec.Enabled = False Else Materials.mnuRec.Enabled = True
'                    Case "Outgoing"
'                        If MyAcc = 3 Then Materials.mnuOut.Enabled = False Else Materials.mnuOut.Enabled = True
'                    Case "Inventory Report"
'                        If MyAcc = 3 Then Materials.mnuInventory.Enabled = False Else Materials.mnuInventory.Enabled = True
'                    Case "Detailed Receiving"
'                        If MyAcc = 3 Then Materials.mnuDtlRec.Enabled = False Else Materials.mnuDtlRec.Enabled = True
'                    Case "Detailed Outgoing"
'                        If MyAcc = 3 Then Materials.mnuDtlOut.Enabled = False Else Materials.mnuDtlOut.Enabled = True
'                    Case "Item Data Entry"
'                        If MyAcc = 3 Then Materials.mnuItem.Enabled = False Else Materials.mnuItem.Enabled = True
'                    Case "Pricing"
'                        If MyAcc = 3 Then Materials.mnuCost.Enabled = False Else Materials.mnuCost.Enabled = True
'                    Case "Users"
'                        If MyAcc = 3 Then Materials.mnuUser.Enabled = False Else Materials.mnuUser.Enabled = True
'                    Case "Suppliers"
'                        If MyAcc = 3 Then Materials.mnuSupplier.Enabled = False Else Materials.mnuSupplier.Enabled = True
'                    Case "Branch"
'                        If MyAcc = 3 Then Materials.mnuBranch.Enabled = False Else Materials.mnuBranch.Enabled = True
'                End Select
'            End If
        Next i
    End With
    'Update Administrator
    If chkAdmin.Value = 1 Then
       LYRcon.Execute "UPDATE settings SET VALUE = '" & txtUserID.Text & "', REMARKS = '" & txtCancelID.Text & "', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & LstUpdt & "' " & _
                      "Where settings.DESCRIPTION = 'ADMINISTRATOR';"
       LYRcon.Execute "UPDATE employees SET USERTYPE = 'Admin' Where Employees.EmpId = '" & Val(txtUserID.Text) & "'"
    End If
    Call BtnSaveCancel(False)
    Frame1.Enabled = False
    flxAccess.Enabled = False
    'RESET ACCEESS
    With Warehouse
        .mnuEmp.Enabled = False
        .mnuUser.Enabled = False
        .mnuCustomer.Enabled = False
        .mnuSupplier.Enabled = False
        .mnuBranch.Enabled = False
        .mnuPOsupplier.Enabled = False
        .mnuPOBranch.Enabled = False
        .mnuReceiving.Enabled = False
        .mnuSalesReturn.Enabled = False
        .mnuStockTransfer.Enabled = False
        .mnuPurchaseReturn.Enabled = False
        .mnuMerchandise.Enabled = False
        .mnuCardex.Enabled = False
        .mnuCardexRed.Enabled = False
        .mnuBackOrder.Enabled = False
        .mnuPrice.Enabled = False
        .mnuPartsProp.Enabled = False
        .mnuReceipts.Enabled = False
        .mnuPr.Enabled = False
        .mnuChecks.Enabled = False
        .mnuCashflow.Enabled = False
        .mnuAFL.Enabled = False
        .mnuStatement.Enabled = False
        .mnuAging.Enabled = False
        .mnuDMCM.Enabled = False
        .mnuInvoice.Enabled = False
        .mnuPCV.Enabled = False
        .mnuCV.Enabled = False
        .mnuDailySales.Enabled = False
        .mnuDailyCash.Enabled = False
        .mnuDailyExpense.Enabled = False
        .mnuMonSales.Enabled = False
        .mnuMonColl.Enabled = False
        .mnuMonExpense.Enabled = False
        .mnuReceivingSum.Enabled = False
        .mnuTrasferSum.Enabled = False
        .mnuSaleSellPrice.Enabled = False
        .mnuSalesCostPrice.Enabled = False
        .mnuSalRetSum.Enabled = False
        .mnuPurchRetSum.Enabled = False
        .mnuInvAdjSum.Enabled = False
        .mnuSalesDistrib.Enabled = False
        .mnuSalesDB.Enabled = False
        .mnuBackup.Enabled = False
        .mnuScripting.Enabled = False
        .mnuPrintPref.Enabled = False
        .mnuOptions.Enabled = False
        .mnuSettings.Enabled = False
        .mnuMonthlyInv.Enabled = False
        .mnuAdjReq.Enabled = False
    End With
    Call LoadUserAccess
    MsgBox "User information successfully saved!", vbInformation + vbOKOnly, "Save"
    
End Sub
    
    

Private Sub cmdTrans_Click(Index As Integer)
End Sub

Private Sub flxAccess_DblClick()
    With flxAccess
        If .Col > 1 And .row > 0 Then
            If .Col = 2 Then
                .CellFontName = "Webdings"
                .TextMatrix(.row, .Col) = "a"
                .TextMatrix(.row, 3) = ""
                .TextMatrix(.row, 4) = ""
            ElseIf .Col = 3 Then
                .CellFontName = "Webdings"
                .TextMatrix(.row, .Col) = "a"
                .TextMatrix(.row, 2) = ""
                .TextMatrix(.row, 4) = ""
            ElseIf .Col = 4 Then
                .CellFontName = "Webdings"
                .TextMatrix(.row, .Col) = "a"
                .TextMatrix(.row, 3) = ""
                .TextMatrix(.row, 2) = ""
            End If
        End If
    End With
            
End Sub

Private Sub flxAccess_EnterCell()
    ActRow = flxAccess.row
    ActCol = flxAccess.Col
End Sub

Private Sub flxAccess_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        With flxAccess
            If .Col > 1 And .row > 0 Then
                If .Col = 2 Then
                    .CellFontName = "Webdings"
                    .TextMatrix(.row, .Col) = "a"
                    .TextMatrix(.row, 3) = ""
                    .TextMatrix(.row, 4) = ""
                ElseIf .Col = 3 Then
                    .CellFontName = "Webdings"
                    .TextMatrix(.row, .Col) = "a"
                    .TextMatrix(.row, 2) = ""
                    .TextMatrix(.row, 4) = ""
                ElseIf .Col = 4 Then
                    .CellFontName = "Webdings"
                    .TextMatrix(.row, .Col) = "a"
                    .TextMatrix(.row, 3) = ""
                    .TextMatrix(.row, 2) = ""
                End If
            End If
        End With
    End If
End Sub

Private Sub flxEmp_Click()
    txtFullName.Text = flxEmp.TextMatrix(flxEmp.row, 1)
End Sub

Private Sub flxEmp_EnterCell()
    CustSrch = True
    txtFullName.Text = flxEmp.TextMatrix(flxEmp.row, 1)
End Sub

Private Sub flxEmp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxEmp.Visible = False
        txtFullName.SetFocus
    End If
End Sub

Private Sub flxEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxEmp.Visible = False
        txtUser.SetFocus
    End If
End Sub

Private Sub Form_Load()
    
    Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPLOYEE, EMPLOYEES.EMPID From EMPLOYEES " & _
          "WHERE employees.USERNAME IS NULL AND employees.PASSWORD IS NULL " & _
          "ORDER BY concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(LEFT(EMPLOYEES.MIDDLENAME,1)='','',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.')));"


    rsEmployee.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxEmp.DataSource = rsEmployee
    With flxEmp
        .Width = txtFullName.Width
        .ColWidth(0) = 0
        .ColWidth(1) = txtFullName.Width - 300
        .ColWidth(2) = 0
    End With
    
    Frame1.Enabled = False
    flxAccess.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    rsEmployee.Close
End Sub


Private Sub RXResizer1_ResizeComplete()
    With flxAccess
        .ColWidth(0) = 300
        .ColWidth(1) = .Width * 0.3
        .ColWidth(2) = .Width * 0.19
        .ColWidth(3) = .Width * 0.19
        .ColWidth(4) = .Width * 0.19
        '.ColWidth(5) = .Width * 0.19 MENUID
    End With
End Sub

Private Sub txtConfirmPW_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then flxAccess.SetFocus
End Sub

Private Sub txtFullName_Change()
    With flxEmp
        If CustSrch Or Not AddNew Then
            CustSrch = False
            Exit Sub
        End If
        
        If txtFullName.Text = "" Then .Visible = False: Exit Sub
            rsEmployee.Find "[EMPLOYEE] LIKE '" + txtFullName.Text + "%'", , adSearchForward, 1
                If Not rsEmployee.EOF Then
                    .Top = Frame1.Top + txtFullName.Top + txtFullName.Height + 30
                    .Left = Frame1.Left + txtFullName.Left
                    .Width = txtFullName.Width
                    .Visible = True
                If Not CustBln Then
                    txtFullName.Tag = Len(txtFullName.Text)
                    CustBln = True
                End If
                    .TopRow = rsEmployee.AbsolutePosition - 1
                    .row = rsEmployee.AbsolutePosition - 1
                    .Col = 0: .ColSel = 2
                    txtFullName.Text = rsEmployee!EMPLOYEE
                    
                    'Insert Additional retrieval here
                    txtUserID.Text = flxEmp.TextMatrix(.row, 2)
                    
                    txtFullName.SelStart = Val(txtFullName.Tag)
                    txtFullName.SelLength = Len(txtFullName.Text)
                    CustBln = False
        Else
            'rsEmployee.Requery
            txtUserID.Text = ""
            .Visible = False
        End If
        
    End With
End Sub

Private Sub txtFullName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If flxEmp.Visible Then
                flxEmp.SetFocus
            Else
                txtUser.SetFocus
            End If
        Case vbKeyEscape
            txtFullName.Text = ""
            flxEmp.Visible = False
        Case vbKeyDelete
            CustSrch = True
        Case vbKeyBack
            CustSrch = True
    End Select
End Sub

Private Sub txtFullName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxEmp.Visible = False
        txtUser.SetFocus
        
    End If
End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtConfirmPW.SetFocus
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPW.SetFocus
End Sub
