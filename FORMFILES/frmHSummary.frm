VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHSummary 
   Caption         =   "Summary"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15060
   ControlBox      =   0   'False
   Icon            =   "frmHSummary.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   15060
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   14040
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   7560
      TabIndex        =   12
      Tag             =   "xyf"
      Top             =   480
      Width           =   7335
      Begin VB.TextBox txtFind 
         Height          =   375
         Left            =   5040
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   240
         Width           =   2175
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   390
         Left            =   3720
         TabIndex        =   13
         Tag             =   "xyf"
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         Caption         =   "Find Next"
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
         Image           =   "frmHSummary.frx":058A
         cBack           =   -2147483633
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSum 
      Height          =   8175
      Left            =   120
      TabIndex        =   5
      Tag             =   "xyf"
      Top             =   1320
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   14420
      _Version        =   393216
      RowHeightMin    =   300
      FocusRect       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Summary"
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
      Image           =   "frmHSummary.frx":0F9C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Print Preview"
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
      Image           =   "frmHSummary.frx":19AE
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Tag             =   "xyf"
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin MSComCtl2.DTPicker Dfrom 
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   101515265
         CurrentDate     =   42107
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   390
         Left            =   6120
         TabIndex        =   8
         Tag             =   "xyf"
         Top             =   195
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         ImgAlign        =   4
         Image           =   "frmHSummary.frx":23C0
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker Dto 
         Height          =   300
         Left            =   3840
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   101515265
         CurrentDate     =   42107
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   390
         Left            =   5707
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   195
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         ImgAlign        =   4
         Image           =   "frmHSummary.frx":2DD2
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   390
         Left            =   3360
         TabIndex        =   11
         Tag             =   "xyf"
         Top             =   195
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         ImgAlign        =   4
         Image           =   "frmHSummary.frx":37E4
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtYear 
         Height          =   315
         Left            =   3135
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   240
         Width           =   720
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
      End
      Begin MSComCtl2.UpDown UpdYear 
         Height          =   315
         Left            =   3840
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   2015
         OrigLeft        =   11310
         OrigTop         =   480
         OrigRight       =   11565
         OrigBottom      =   795
         Max             =   2020
         Min             =   2005
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "For the month of"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   300
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdFilter 
      Height          =   372
      Left            =   1800
      TabIndex        =   3
      Tag             =   "xyf"
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Filter Records"
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
      Image           =   "frmHSummary.frx":41F6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Tag             =   "xyf"
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Sho&w Form"
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
      Image           =   "frmHSummary.frx":4C08
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   13440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15300
      m_MDI_Height    =   10230
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Menu MNU 
      Caption         =   "1"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmHSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSummary As New ADODB.Recordset
Dim DailyRange As Boolean, MonthlyRange As Boolean, FixRange As Boolean
Dim ActRow As Integer, fRow As Long, fCol As Long, MaxCol As Integer, ShowDate As Integer
Public MySource As String
Private Sub cboMonth_Click()
    If cboMonth.ListIndex = -1 Then Exit Sub
    
'    Identify source whose monthly reports is from Day 1 to Last day of the month
    If mySrc = "CLOSED ACCOUNTS FOR AGENT" Or mySrc = "Monthly Collection" Or mySrc = "Monthly Inventory" Then
        dFrom.Value = DateSerial(Val(txtYear.Text), cboMonth.ListIndex + 1, 1)
        dTo.Value = DateAdd("M", 1, dFrom.Value)
        dTo.Value = DateAdd("D", -1, dTo.Value)
    Else    'Use for reports whose date covered is from 7th day to 6th of the following month
        dFrom.Value = DateSerial(Val(txtYear.Text), cboMonth.ListIndex + 1, 7)
        dTo.Value = DateSerial(Year(DateAdd("M", 1, dFrom.Value)), Month(DateAdd("M", 1, dFrom.Value)), 6)
    End If
    Call cmdRefresh_Click
End Sub

Private Sub cmdExit_Click()
      
    SPpricing.Visible = True
    Unload Me
    
End Sub
Private Sub cmdFind_Click()
    Dim i As Long, j As Long
    
    If txtFind.Text = "" Then Exit Sub
    With flxSum
        For i = fRow To .Rows - 1
            For j = fCol To MaxCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = MaxCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = MaxCol - 1, i + 1, i)
                    fCol = IIf(j = MaxCol - 1, 1, j + 1)
                    ActRow = i
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1
    
End Sub
Private Sub cmdNext_Click()
    cmdNext.Visible = False
    cmdPrev.Visible = True
    dTo.Visible = True
    ShowDate = 2
End Sub
Private Sub cmdPrev_Click()
    dTo.Value = dFrom.Value
    cmdNext.Left = cmdPrev.Left
    cmdNext.Visible = True
    cmdPrev.Visible = False
    dTo.Visible = False
    ShowDate = 1
    If MySource = "PARTNUMBER" Then
        dTo.Value = Format(dFrom.Value, "yyyy-mm-dd 23:59:59")
    End If
    
End Sub
Private Sub cmdPrint_Click()
    Dim Qry As String
    If MySource = "PRICING" Then
        Qry = "SELECT spareparts.PARTNUMBER,  spareparts.DESCRIPTION,  Format(spareparts.STOREPRICE,2),  DATE_FORMAT(spareparts.PRICEUPDATE,'%m-%d-%Y'), FORMAT(SUBSTR(spareparts.PRICEHIST, 12, 10),2) As PREVCOST,  DATE_FORMAT(SUBSTR(spareparts.PRICEHIST, 1, 10),'%m-%d-%Y') AS PREVUPDATE " & "From spareparts WHERE spareparts.PRICEUPDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And spareparts.PRICEUPDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY spareparts.PARTNUMBER"
        
        frmReport.MySource = "PRICING"
        frmReport.QryPrn = Qry
        frmReport.ParamValue1 = "MOTO-R REPUBLIC CORPORATION"
        frmReport.ParamValue2 = "LYR BLDG. Gante Road, Tagum City"
        frmReport.Show vbModal
        
    ElseIf MySource = "PARTNUMBER" Then
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, DATE_FORMAT(spareparts.LASTUPDATE,'%m-%d-%Y') as DATE_FORMATTED, spareparts.STOREPRICE, spareparts.REMARKS FROM spareparts WHERE spareparts.LASTUPDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd h:mm:ss") & "' And spareparts.LASTUPDATE <= '" & Format(dTo.Value, "yyyy-mm-dd h:mm:ss") & "' ORDER BY spareparts.PARTNUMBER"
        
        frmReport.MySource = "PARTNUMBER"
        frmReport.QryPrn = Qry
        frmReport.ParamValue1 = "MOTO-R REPUBLIC CORPORATION"
        frmReport.ParamValue2 = "LYR BLDG. Gante Road, Tagum City"
        frmReport.Show vbModal
    End If
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshRecord
End Sub

Public Sub RefreshRecord()
'    MsgBox rCode
'    ----------------------- refresh naa diri
    Dim Qry As String
    
    
    If MySource = "PRICING" Then
        
         Qry = "SELECT spareparts.PARTNUMBER,  spareparts.DESCRIPTION,  Format(spareparts.STOREPRICE,2),  DATE_FORMAT(spareparts.PRICEUPDATE,'%m/%d/%Y'), FORMAT(SUBSTR(spareparts.PRICEHIST, 12, 10),2) As PREVCOST," & _
        "DATE_FORMAT(SUBSTR(spareparts.PRICEHIST, 1, 10), '%m/%d/%Y') AS PREVUPDATE " & " From spareparts WHERE spareparts.PRICEUPDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And spareparts.PRICEUPDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY spareparts.PARTNUMBER"
    
    ElseIf MySource = "PARTNUMBER" Then
        
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, DATE_FORMAT(spareparts.LASTUPDATE,'%m/%d/%Y') as DATE_FORMATTED, Format(spareparts.STOREPRICE,2), spareparts.REMARKS FROM spareparts WHERE spareparts.LASTUPDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd h:mm:ss") & "' And spareparts.LASTUPDATE <= '" & Format(dTo.Value, "yyyy-mm-dd 23:59:59") & "' ORDER BY spareparts.PARTNUMBER"
'        MsgBox Qry
'        format(SUBSTR(spareparts.LASTUPDATE, 1, 10),'%m,%d%y')
    End If
    
'    Display Query in Flx
    If rsSummary.State = adStateOpen Then rsSummary.Close
    rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSummary.RecordCount = 0 Then   'No records to show
        flxSum.Clear
        flxSum.Rows = 2
        flxSum.Cols = MaxCol
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSum.DataSource = rsSummary   'With records
    ActRow = 1
    Call RXResizer1_ResizeComplete
End Sub

Private Sub cmdShow_Click()
    If MySource = "PRODUCTS" Then
        Call frmproductEnt.ShowRecord(Val(flxSum.TextMatrix(ActRow, 8)))
        frmproductEnt.Show vbModal
    End If
End Sub
Private Sub Dfrom_Change()
    If ShowDate = 1 Then dTo.Value = dFrom.Value
    If MySource = "PARTNUMBER" Then
        
        If ShowDate = 1 Then
            dTo.Value = Format(dFrom.Value, "yyyy-mm-dd 23:59:59")
            dFrom.Value = Format(dFrom.Value, "yyyy-mm-dd 00:00:00")
            
        End If
    End If
End Sub

Private Sub flxSum_Click()
    ActRow = flxSum.Row
End Sub

Private Sub flxSum_EnterCell()
    ActRow = flxSum.Row
End Sub

Private Sub Form_Load()
    Me.Caption = MySource & " SUMMARY"
    
    If MySource = "PRODUCTS" Then
        FixRange = True
        DailyRange = False
        MonthlyRange = False
        MaxCol = 8
    ElseIf MySource = "PRICING" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        MaxCol = 7
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDate = 2
    ElseIf MySource = "PARTNUMBER" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        MaxCol = 6
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDate = 2
    ElseIf MySource = "CLOSED ACCOUNTS FOR AGENT" Then
        FixRange = False
        DailyRange = False
        MonthlyRange = True
        MaxCol = 7
    End If
    
    
    If DailyRange Then
        Frame1.Visible = True
        Frame3.Visible = False
    ElseIf MonthlyRange Then
        Frame1.Visible = False
        Frame3.Visible = True
        For i = 1 To 12
            cboMonth.AddItem MonthName(i)
        Next i
        UpdYear.Value = Year(Date)
        txtYear.Text = Year(Date)
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = DateAdd("d", -1, DateAdd("m", 1, dFrom.Value))
        cboMonth.ListIndex = Month(dFrom.Value) - 1
    ElseIf FixRange Then
        Frame1.Visible = False
        Frame3.Visible = False
    End If
    
    
    
    Timer1.Enabled = True
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxSum
        If MySource = "PRICING" Then
            .FormatString = "  |<PARTNUMBER|^DESCRIPTION|^STOREPRICE|^DATE UPDATED|<PREV-PRICE|<PREV-UPDATE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.18
            .ColWidth(2) = (.Width - 300) * 0.27
            .ColWidth(3) = (.Width - 300) * 0.13
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.13
            .ColWidth(6) = (.Width - 300) * 0.13
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            
        ElseIf MySource = "PARTNUMBER" Then
            .FormatString = "  |<PARTNUMBER|^DESCRIPTION|^DATE CREATED|^STOREPRICE|<REMARKS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.18
            .ColWidth(2) = (.Width - 300) * 0.32
            .ColWidth(3) = (.Width - 300) * 0.13
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.2
        End If
    End With
    
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    Dim Qry As String
    
'    Qry = "SELECT spareparts.PARTNUMBER,  spareparts.DESCRIPTION,  Format(spareparts.STOREPRICE,2),  spareparts.PRICEUPDATE,  Format(SUBSTR(spareparts.PRICEHIST, 12, 10),2) As PREVCOST,  SUBSTR(spareparts.PRICEHIST, 1, 10) AS PREVUPDATE " &         "From spareparts WHERE spareparts.PRICEUPDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And spareparts.PRICEUPDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY spareparts.PARTNUMBER"
    
    If MySource = "PRICING" Then
        Qry = "SELECT spareparts.PARTNUMBER,  spareparts.DESCRIPTION,  Format(spareparts.STOREPRICE,2),  DATE_FORMAT(spareparts.PRICEUPDATE,'%m/%d/%Y'), FORMAT(SUBSTR(spareparts.PRICEHIST, 12, 10),2) As PREVCOST," & _
        "DATE_FORMAT(SUBSTR(spareparts.PRICEHIST, 1, 10), '%m/%d/%Y') AS PREVUPDATE " & " From spareparts WHERE spareparts.PRICEUPDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And spareparts.PRICEUPDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY spareparts.PARTNUMBER"
        
        If rsSummary.State = adStateOpen Then rsSummary.Close
        rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If rsSummary.RecordCount = 0 Then   'No records to show
            flxSum.Clear
            flxSum.Rows = 2
            flxSum.Cols = MaxCol
            Call RXResizer1_ResizeComplete
            Timer1.Enabled = False
            Exit Sub
        End If
        
        Set flxSum.DataSource = rsSummary   'With records
        rsSummary.Close
        
    ElseIf MySource = "PARTNUMBER" Then
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, DATE_FORMAT(spareparts.LASTUPDATE,'%m/%d/%Y') as DATE_FORMATTED, Format(spareparts.STOREPRICE,2), spareparts.REMARKS FROM spareparts WHERE spareparts.LASTUPDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd h:mm:ss") & "' And spareparts.LASTUPDATE <= '" & Format(dTo.Value, "yyyy-mm-dd 23:59:59") & "' ORDER BY spareparts.PARTNUMBER"
        
        If rsSummary.State = adStateOpen Then rsSummary.Close
        rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsSummary.RecordCount = 0 Then   'No records to show
            flxSum.Clear
            flxSum.Rows = 2
            flxSum.Cols = MaxCol
            Call RXResizer1_ResizeComplete
            Timer1.Enabled = False
            Exit Sub
        End If
        
        Set flxSum.DataSource = rsSummary   'With records
        rsSummary.Close
        
'        Format(SUBSTR(spareparts.LASTUPDATE, 1, 10)
    End If
    
    ActRow = 1
    
    Timer1.Enabled = False
    Call RXResizer1_ResizeComplete
    
End Sub
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    
    Dim i As Long, j As Long
  
      If Chr(KeyAscii) = "'" Then KeyAscii = 0
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
  
   If KeyAscii = 13 Then
      If txtFind.Text = "" Then Exit Sub
        With flxSum
            For i = fRow To .Rows - 1
                For j = fCol To MaxCol - 1
                    If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                        .Col = 1: .Row = i
                        .ColSel = MaxCol - 1: .RowSel = i
                        .TopRow = i
                        fRow = IIf(j = MaxCol - 1, i + 1, i)
                        fCol = IIf(j = MaxCol - 1, 1, j + 1)
                        ActRow = i
                        Exit Sub
                    End If
                Next j
                fCol = 1
            Next i
        End With
        MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
        fRow = 1
        fCol = 1
    End If
    
    
End Sub

