VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SPRequest 
   Caption         =   "SP Request Update"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   12795
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBranchID 
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetails 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1508
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.TextBox txtBranchName 
      Height          =   375
      Left            =   9480
      TabIndex        =   5
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtBranch2 
      Height          =   285
      Left            =   9480
      TabIndex        =   4
      Text            =   "MOTO-R ULAS"
      Top             =   720
      Width           =   3375
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtBranch1 
      Height          =   285
      Left            =   9480
      TabIndex        =   2
      Text            =   "MCP DEPOT TORIL"
      Top             =   360
      Width           =   3375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSavePO 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin MSFlexGridLib.MSFlexGrid flxPO 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
   End
End
Attribute VB_Name = "SPRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUPO As New ADODB.Recordset

Sub UpdatePO()
    Dim rsCup As New ADODB.Recordset
    Dim iRw As Long, uRw As Long, POTotAmount As Currency
    Dim rsDtl As New ADODB.Recordset
    
    With flxSavePO
        pBar.Value = 0
        pBar.Max = .Rows - 1
        
        For iRw = 1 To .Rows - 1
            
            If .TextMatrix(iRw, 0) <> "" Then
                flxDetails.Clear
                flxDetails.Rows = 1
                LYRcon.Execute "Insert Into sporderbranch(ORDERDATE, BRANCHNAME, REFERENCE, ORDERTYPE, BRANCHID)VALUES('" & Format(.TextMatrix(iRw, 1), "yyyy-mm-dd") & "', '" & .TextMatrix(iRw, 2) & "', '" & .TextMatrix(iRw, 3) & "', '" & .TextMatrix(iRw, 4) & "', " & cmbBranch.ItemData(cmbBranch.ListIndex) & ");"
                DoEvents
                Qry = "SELECT spitr.REFERENCE, spitr.SPAREPARTID, spitr.TRANSTYPE, spitr.`ORDER`, spitr.PRICE " & _
                    "From spitr Where spitr.REFERENCE = '" & .TextMatrix(iRw, 3) & "' AND spitr.TRANSTYPE <> 'CANCELLED';"
                If rsDtl.State = adStateOpen Then rsDtl.Close
                   
                    'Select Branches to open
                    If cmbBranch.Text = "MOTO-R ULAS" Then
                        If UlasConnect Then
                            rsDtl.Open Qry, conUlas, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "YAMAHA MOTORHUB" Then
                        If YamahaConnect Then
                            rsDtl.Open Qry, conYM, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "MC PARTS DEPOT" Then
                        If TorilConnect Then
                            rsDtl.Open Qry, conToril, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "MRC PANABO" Then
                        If PanaboConnect Then
                            rsDtl.Open Qry, conPanabo, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "BINOS POLOMOLOK" Then
                        If BinosPolConnect Then
                            rsDtl.Open Qry, conPol, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf cmbBranch.Text = "MRC1 SOBRECARY" Then
                        If MRC1Connect Then
                            rsDtl.Open Qry, conMRC1, adOpenStatic, adLockOptimistic
                        End If
                    End If
                    
                
                Set flxDetails.Recordset = rsDtl
                DoEvents
                For uRw = 0 To flxDetails.Rows - 1
                    POTotAmount = Val(flxDetails.TextMatrix(uRw, 3)) * Val(flxDetails.TextMatrix(uRw, 4))
                    LYRcon.Execute "INSERT INTO spbranchdtl(REFERENCE, SPAREPARTID, TRANSTYPE, `ORDER`, BRANCHNAME, ORDERBAL, PRICE, TOTAMOUNT)" & _
                        "VALUES('" & flxDetails.TextMatrix(uRw, 0) & "', " & Val(flxDetails.TextMatrix(uRw, 1)) & ", '" & flxDetails.TextMatrix(uRw, 2) & "', " & _
                        "" & Val(flxDetails.TextMatrix(uRw, 3)) & ", '" & cmbBranch.Text & "', " & Val(flxDetails.TextMatrix(uRw, 3)) & ", " & Val(flxDetails.TextMatrix(uRw, 4)) & ", " & POTotAmount & ");"
                Next uRw
                DoEvents
                'Verify records if already added
                Qry = "SELECT sporderbranch.ORDERDATE, sporderbranch.BRANCHNAME, sporderbranch.REFERENCE, sporderbranch.ORDERTYPE From sporderbranch where sporderbranch.REFERENCE = '" & .TextMatrix(iRw, 3) & "' AND sporderbranch.BRANCHID = " & cmbBranch.ItemData(cmbBranch.ListIndex) & ";"
                If rsCup.State = adStateOpen Then rsCup.Close
                rsCup.Open Qry, LYRcon, adOpenStatic, adLockOptimistic
                If rsCup.RecordCount > 0 Then
                     rsCup.MoveFirst
                     If flexUpdatePO.Rows = 2 And flexUpdatePO.TextMatrix(1, 1) = "" Then
                      ' flexUpdatePO.Rows = 1: Call FormatFlxPO
                     End If
                     flexUpdatePO.AddItem rsCup!ORDERDATE & vbTab & rsCup!BRANCHNAME & vbTab & rsCup!REFERENCE & vbTab & rsCup!ORDERTYPE
                     flexUpdatePO.FixedRows = 1
                End If
                
            End If
                pBar.Value = pBar.Value + 1
                'PrcntRec = PrgBar.Value / PrgBar.Max
                'lblp.Caption = Format(PrcntRec, "###%")
        Next iRw
        'lblp.Caption = "100%"
        MsgBox ("P.O Successfully Update"), vbInformation, "P.O"
    End With
End Sub



Public Sub LoadBranchPO()
    Dim rsChk As New ADODB.Recordset
    Dim rsDtl As New ADODB.Recordset
    Dim rsCup As New ADODB.Recordset
    Dim iRw As Long, uRw As Long, POTotAmount As Currency
    Dim Qry As String
    'Dim iRw As Long
    'Branch-ID      Branch Name
        '12             MRC1 SOBRECARY
        '13             MOTO-R ULAS
        '15             MC PARTS DEPOT
        '19             MRC PANABO
        '23             MRC GENSAN
    
    'connect to mcp depot
    With rsUPO
        'Select Order table from Branches
        Qry = "SELECT sporder.REFERENCE, sporder.ORDERTYPE, sporder.ORDERDATE, sporder.CONFIRMNO, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER " & _
            "From sporder Where sporder.`STATUS` = 'NEW' AND sporder.CONFIRMNO IS NULL AND sporder.ORDERDATE > '2011/07/01';"
             
        If .State = adStateOpen Then .Close
        'Select Branches to open
        If txtBranch1.Text = "MCP DEPOT TORIL" Then
            txtBranchName.Text = txtBranch1.Text
            txtBranchID.Text = 15
            If ConnectToril Then
                TorilConnect = True
                    If TorilConnect Then
                        .Open Qry, conToril, adOpenStatic, adLockOptimistic
                    End If
            Else
                TorilConnect = False
            End If
         End If
        
        If .RecordCount = 0 Then: Exit Sub
        .MoveFirst
        flxPO.Rows = 1
        'flxSavePO.Rows = 1
        'Add P.O record to flex
        pBar.Visible = True
        pBar.Value = 0
        pBar.Max = .RecordCount
        .MoveFirst
        For iRw = 1 To .RecordCount
                'flxMain.AddItem !ORDERDATE & vbTab & cmbBranch.Text & vbTab & !REFERENCE & vbTab & !ORDERTYPE
                flxPO.AddItem txtBranchName.Text & vbTab & !REFERENCE & vbTab & !Status & vbTab & !ORDERDATE & vbTab & !QTYPN & vbTab & !QTYORDER
                'Verify records if already added
                'Qry = "SELECT * From sporderbranch where sporderbranch.REFERENCE = '" & flxPO.TextMatrix(iRw, 2) & "' and sporderbranch.BRANCHID = " & txtBranchID.Text & ";"
                Qry = "SELECT * From sporderbranch where sporderbranch.REFERENCE = '" & flxPO.TextMatrix(iRw, 2) & "' AND sporderbranch.BRANCHID = " & txtBranchID.Text & ";"
                If rsChk.State = adStateOpen Then rsChk.Close
                rsChk.Open Qry, LYRcon, adOpenStatic, adLockOptimistic
                If rsChk.RecordCount = 0 Then
                     'Add record if not Addedd
                     flxSavePO.AddItem txtBranchName.Text & vbTab & !REFERENCE & vbTab & !Status & vbTab & !ORDERDATE & vbTab & !QTYPN & vbTab & !QTYORDER & vbTab & txtBranchID.Text
                End If
             .MoveNext
              pBar.Value = pBar.Value + 1
              
        Next iRw
             pBar.Visible = False
        
    End With
    
    'connect to mrc ulas
    With rsUPO
        'Select Order table from Branches
        Qry = "SELECT sporder.REFERENCE, sporder.ORDERTYPE, sporder.ORDERDATE, sporder.CONFIRMNO, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER " & _
            "From sporder Where sporder.`STATUS` = 'NEW' AND sporder.CONFIRMNO IS NULL AND sporder.ORDERDATE > '2011/07/01';"
             
        If .State = adStateOpen Then .Close
        'Select Branches to open
        If txtBranch2.Text = "MOTO-R ULAS" Then
            txtBranchName.Text = txtBranch2.Text
            txtBranchID.Text = 13
            If ConnectUlas Then
                UlasConnect = True
                    If UlasConnect Then
                        .Open Qry, conUlas, adOpenStatic, adLockOptimistic
                    End If
            Else
                UlasConnect = False
            End If
        End If
        
        If .RecordCount = 0 Then: Exit Sub
        .MoveFirst
        flxPO.Rows = 1
        'flxSavePO.Rows = 1
        'Add P.O record to flex
        pBar.Visible = True
        pBar.Value = 0
        pBar.Max = .RecordCount
        .MoveFirst
        For iRw = 1 To .RecordCount
                'flxMain.AddItem !ORDERDATE & vbTab & cmbBranch.Text & vbTab & !REFERENCE & vbTab & !ORDERTYPE
                flxPO.AddItem txtBranchName.Text & vbTab & !REFERENCE & vbTab & !Status & vbTab & !ORDERDATE & vbTab & !QTYPN & vbTab & !QTYORDER
                'Verify records if already added
                'Qry = "SELECT * From sporderbranch where sporderbranch.REFERENCE = '" & flxPO.TextMatrix(iRw, 2) & "' and sporderbranch.BRANCHID = " & txtBranchID.Text & ";"
                Qry = "SELECT * From sporderbranch where sporderbranch.REFERENCE = '" & flxPO.TextMatrix(iRw, 2) & "' AND sporderbranch.BRANCHID = " & txtBranchID.Text & ";"
                If rsChk.State = adStateOpen Then rsChk.Close
                rsChk.Open Qry, LYRcon, adOpenStatic, adLockOptimistic
                If rsChk.RecordCount = 0 Then
                     'Add record if not Addedd
                     flxSavePO.AddItem txtBranchName.Text & vbTab & !REFERENCE & vbTab & !Status & vbTab & !ORDERDATE & vbTab & !QTYPN & vbTab & !QTYORDER & vbTab & txtBranchID.Text
                End If
             .MoveNext
              pBar.Value = pBar.Value + 1
              
        Next iRw
             pBar.Visible = False
        
    End With
    
     'Saving for sporderbranch
    With flxSavePO
        pBar.Value = 0
        pBar.Max = .Rows - 1
        
        For iRw = 1 To .Rows - 1
            If .TextMatrix(iRw, 0) <> "" Then
                'flxDetails.Clear
                'flxDetails.Rows = 1
                'LYRcon.Execute "Insert Into sporderbranch(ORDERDATE, BRANCHNAME, REFERENCE, ORDERTYPE, BRANCHID)VALUES('" & Format(.TextMatrix(iRw, 1), "yyyy-mm-dd") & "', '" & .TextMatrix(iRw, 2) & "', '" & .TextMatrix(iRw, 3) & "', '" & .TextMatrix(iRw, 4) & "', " & cmbBranch.ItemData(cmbBranch.ListIndex) & ");"
                LYRcon.Execute "INSERT INTO sporderbranch(BRANCHNAME, REFERENCE, `STATUS`, ORDERDATE, QTYPN, QTYORDER, BRANCHID) " & _
                "VALUES('" & .TextMatrix(iRw, 0) & "', '" & .TextMatrix(iRw, 1) & "', '" & .TextMatrix(iRw, 2) & "', '" & Format(.TextMatrix(iRw, 3), "yyyy-mm-dd") & "', '" & .TextMatrix(iRw, 4) & "' , '" & .TextMatrix(iRw, 5) & "', " & .TextMatrix(iRw, 6) & ")"
                
                DoEvents
                Qry = "SELECT spitr.REFERENCE, spitr.SPAREPARTID, spitr.TRANSTYPE, spitr.`ORDER`, spitr.PRICE " & _
                    "From spitr Where spitr.REFERENCE = '" & .TextMatrix(iRw, 2) & "' AND spitr.TRANSTYPE <> 'CANCELLED';"
                If rsDtl.State = adStateOpen Then rsDtl.Close
                   
                    'Select Branches to open
                     'Branch-ID      Branch Name
                        '12             MRC1 SOBRECARY
                        '13             MOTO-R ULAS
                        '15             MC PARTS DEPOT
                        '19             MRC PANABO
                        '23             MRC GENSAN
                        
                    If .TextMatrix(iRw, 6) = 15 Then
                        If TorilConnect Then
                            rsDtl.Open Qry, conToril, adOpenStatic, adLockOptimistic
                        End If
                    ElseIf .TextMatrix(iRw, 6) = 13 Then
                        If UlasConnect Then
                            rsDtl.Open Qry, conUlas, adOpenStatic, adLockOptimistic
                        End If
                    End If
                    
                
                Set flxDetails.Recordset = rsDtl
                DoEvents
                For uRw = 0 To flxDetails.Rows - 1
                    POTotAmount = Val(flxDetails.TextMatrix(uRw, 3)) * Val(flxDetails.TextMatrix(uRw, 4))
                    LYRcon.Execute "INSERT INTO spbranchdtl(REFERENCE, SPAREPARTID, TRANSTYPE, `ORDER`, BRANCHNAME, ORDERBAL, PRICE, TOTAMOUNT)" & _
                        "VALUES('" & flxDetails.TextMatrix(uRw, 0) & "', " & Val(flxDetails.TextMatrix(uRw, 1)) & ", '" & flxDetails.TextMatrix(uRw, 2) & "', " & _
                        "" & Val(flxDetails.TextMatrix(uRw, 3)) & ", '" & cmbBranch.Text & "', " & Val(flxDetails.TextMatrix(uRw, 3)) & ", " & Val(flxDetails.TextMatrix(uRw, 4)) & ", " & POTotAmount & ");"
                Next uRw
                DoEvents
                'Verify records if already added
                'Qry = "SELECT sporderbranch.ORDERDATE, sporderbranch.BRANCHNAME, sporderbranch.REFERENCE, sporderbranch.ORDERTYPE From sporderbranch where sporderbranch.REFERENCE = '" & .TextMatrix(iRw, 3) & "' AND sporderbranch.BRANCHID = " & cmbBranch.ItemData(cmbBranch.ListIndex) & ";"
                'If rsCup.State = adStateOpen Then rsCup.Close
                'rsCup.Open Qry, LYRcon, adOpenStatic, adLockOptimistic
                'If rsCup.RecordCount > 0 Then
                '     rsCup.MoveFirst
                '     If flexUpdatePO.Rows = 2 And flexUpdatePO.TextMatrix(1, 1) = "" Then
                      ' flexUpdatePO.Rows = 1: Call FormatFlxPO
                '     End If
                '     flexUpdatePO.AddItem rsCup!ORDERDATE & vbTab & rsCup!BRANCHNAME & vbTab & rsCup!REFERENCE & vbTab & rsCup!ORDERTYPE
                '     flexUpdatePO.FixedRows = 1
                'End If
                
            End If
                pBar.Value = pBar.Value + 1
                'PrcntRec = PrgBar.Value / PrgBar.Max
                'lblp.Caption = Format(PrcntRec, "###%")
        Next iRw
        'lblp.Caption = "100%"
        MsgBox ("P.O Successfully Update"), vbInformation, "P.O"
    End With
    
   
    
End Sub

Private Sub Form_Load()
    Call LoadBranchPO
End Sub
