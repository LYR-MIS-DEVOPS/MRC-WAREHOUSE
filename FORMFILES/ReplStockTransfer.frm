VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form ReplStockTransfer 
   Caption         =   "Replenish Stock Transfer"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   19215
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabConfirm 
      Height          =   2055
      Left            =   6960
      TabIndex        =   34
      Tag             =   "XY"
      Top             =   3720
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3625
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "ReplStockTransfer.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdQuit"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdConfirm"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtChecker"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDriver"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtDriver 
         Height          =   375
         Left            =   1800
         TabIndex        =   36
         Tag             =   "xyf"
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtChecker 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   960
         Width           =   2655
      End
      Begin lvButton.lvButtons_H cmdConfirm 
         Height          =   375
         Left            =   1320
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "CONFIRM"
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
         Image           =   "ReplStockTransfer.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdQuit 
         Height          =   375
         Left            =   2760
         TabIndex        =   40
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "CANCEL"
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
         Image           =   "ReplStockTransfer.frx":05B6
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DRIVER NAME:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Tag             =   "xyf"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "CHECKER NAME:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   19440
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   19440
      TabIndex        =   28
      Text            =   "HoldBackColor"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Tag             =   "XY"
      Top             =   7680
      Width           =   18975
      Begin VB.TextBox txtDateStart 
         Height          =   285
         Left            =   3000
         TabIndex        =   27
         Tag             =   "XY"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtAppBy 
         Height          =   285
         Left            =   1560
         TabIndex        =   26
         Tag             =   "XY"
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtPrepBy 
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Tag             =   "XY"
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblSTARTREPL 
         Caption         =   "DATE START OF REPL. PROCESS:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Tag             =   "XY"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblAppby 
         Caption         =   "APPROVED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Tag             =   "XY"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblPrepBy 
         Caption         =   "PREPARED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Tag             =   "XY"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   120
      TabIndex        =   17
      Tag             =   "XY"
      Top             =   1920
      Width           =   18975
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   0
         TabIndex        =   42
         Tag             =   "XY"
         Top             =   5520
         Visible         =   0   'False
         Width           =   18975
         _ExtentX        =   33470
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxPartsTtl 
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   4440
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   1931
         _Version        =   393216
         Cols            =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flxParts 
         Height          =   3975
         Left            =   120
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   240
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   7011
         _Version        =   393216
         Cols            =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSummary 
         Caption         =   "SUMMARY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Tag             =   "XY"
         Top             =   4200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Tag             =   "XY"
      Top             =   480
      Width           =   18975
      Begin VB.CheckBox chkFinal 
         Caption         =   "FINAL"
         Height          =   255
         Left            =   4680
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Tag             =   "XY"
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtSTATUS 
         Height          =   285
         Left            =   17280
         TabIndex        =   16
         Tag             =   "XY"
         Text            =   " "
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtSTFDate 
         Height          =   285
         Left            =   17280
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtSTFNo 
         Height          =   285
         Left            =   17280
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboRefNo 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox cboTransTo 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Tag             =   "XY"
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblSTATUS 
         Alignment       =   1  'Right Justify
         Caption         =   "STATUS:"
         Height          =   255
         Left            =   16440
         TabIndex        =   13
         Tag             =   "XY"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblSTFDATE 
         Alignment       =   1  'Right Justify
         Caption         =   "STF DATE:"
         Height          =   255
         Left            =   16320
         TabIndex        =   12
         Tag             =   "XY"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblSTFNo 
         Alignment       =   1  'Right Justify
         Caption         =   "STF #:"
         Height          =   255
         Left            =   16440
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblRefNo 
         Caption         =   "REPL REF #:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Tag             =   "XY"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblAddress 
         Caption         =   "ADDRESS:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblTransfer 
         Caption         =   "TRANSFER TO:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   240
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&New Stock Transfer"
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
      Image           =   "ReplStockTransfer.frx":0B50
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   19320
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   6600
      Width           =   1920
      _ExtentX        =   3387
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
      Image           =   "ReplStockTransfer.frx":10EA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&View Summary"
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
      Image           =   "ReplStockTransfer.frx":1684
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "E&xit Stock Transfer"
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
      Image           =   "ReplStockTransfer.frx":1C1E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   4
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "Print Preview"
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
      Image           =   "ReplStockTransfer.frx":21B8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   19440
      TabIndex        =   30
      Top             =   600
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "Save"
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
      Image           =   "ReplStockTransfer.frx":2752
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   19440
      TabIndex        =   31
      Top             =   1080
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "Cancel"
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
      Image           =   "ReplStockTransfer.frx":2CEC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   33
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "Edit Stock Transfer"
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
      Image           =   "ReplStockTransfer.frx":3286
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   19560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   19455
      m_MDI_Height    =   9720
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row"
      End
   End
End
Attribute VB_Name = "ReplStockTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long, j As Long, z As Long, x As Long
Dim arrBRANCH() As String
Dim arrREFNO() As String
Dim Tmp() As String
Dim BranchCode As String
Dim strPartial As String, strTotal As String
Dim StatSrch As Boolean
Dim tmpRs As New ADODB.Recordset
Dim formAction As String
Dim Qry As String
Dim Qry2 As String
Dim strSTFNO As String
Dim matrixID As Integer, intSTFID As Integer

'************* CODE UPDATED: 3/15/17 *****************
Dim pBOOLEAN As String, strModelBrand As String
'********* END CODE UPDATED: 3/15/17 *****************

Public MySource
Dim ProperExit As Boolean, AddNew As Boolean
Dim ans As String

'************* CODE UPDATED: 3/16/17 *****************
Sub CreateTempTableFinalTransfer()
Dim mySTFNO As String, strINTRANSIT As String, SEQNO As Integer
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_FINALTRANSFER_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_FINALTRANSFER_DTLS( " & _
        "    REMARKS Varchar(20)      " & _
        "   ,SPAREPARTID Int(11)      " & _
        "   ,BRANCHSPOID Int(11)      " & _
        "   ,REFERENCE Varchar(20)      " & _
        "   ,TRANSTYPE Varchar(10)      " & _
        "   ,`OUT` Int(11)      " & _
        "   ,POREF Varchar(30)      " & _
        "   ,UPDATEDBY Int(11)      " & _
        "   ,SEQNO Int(11)      " & _
        ");"
 '*************** UPDATED CODE: 06/14/2017 ********************                                                                                                               'QTYONHAND
        Qry = "Insert into sp_partsmatrix.TMP_FINALTRANSFER_DTLS(REMARKS,SPAREPARTID,BRANCHSPOID,REFERENCE,TRANSTYPE,`OUT`,POREF,UPDATEDBY,SEQNO) VALUES "
        mySTFNO = "STF#" & txtSTFNO.Text
        strINTRANSIT = "INTRANSIT"
        SEQNO = 0
        For i = 2 To flxParts.Rows - 1
            If i = flxParts.Rows - 1 Then
                If i > 2 Then Qry = Qry & ","
                If Val(flxParts.TextMatrix(i, 7)) <> 0 Then
                    SEQNO = SEQNO + 1
                    Qry = Qry & "('" & strINTRANSIT & "'," & Val(flxParts.TextMatrix(i, 7)) & "," & matrixID & ",'" & mySTFNO & "', 'OUT' ," & Val(flxParts.TextMatrix(i, 5)) & ",'" & cboRefNo.Text & "'," & UserID & "," & SEQNO & ") "
                End If
                GoTo SkipData
            End If
            If flxParts.TextMatrix(i, 2) <> flxParts.TextMatrix(i + 1, 6) Then
                If i > 2 Then Qry = Qry & ","
                If Val(flxParts.TextMatrix(i, 7)) <> 0 Then
                    SEQNO = SEQNO + 1
                    Qry = Qry & "('" & strINTRANSIT & "'," & Val(flxParts.TextMatrix(i, 7)) & "," & matrixID & ",'" & mySTFNO & "', 'OUT' ," & Val(flxParts.TextMatrix(i, 5)) & ",'" & cboRefNo.Text & "'," & UserID & "," & SEQNO & ") "
                End If
            End If
SkipData:
        Next i
 '*********** END UPDATED CODE: 06/14/2017 ********************
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTableFinalTransfer()")
End Sub
'********* END CODE UPDATED: 3/16/17 *****************

'************* CODE UPDATED: 3/16/17 *****************
'Sub CreateTempTableSPITR()
'Dim mySTFNO As String, SEQNO As Integer
'On Error GoTo er
'        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_SPITR_DTLS;"
'
'         DoEvents
'       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_SPITR_DTLS( " & _
'        "    REFERENCE Varchar(20)      " & _
'        "   ,TRANSTYPE Varchar(10)      " & _
'        "   ,SPAREPARTID Int(11)      " & _
'        "   ,`OUT` Int(11)      " & _
'        "   ,POREF Varchar(30)      " & _
'        "   ,UPDATEDBY Int(11)      " & _
'        "   ,SEQNO Int(11)      " & _
'        ");"
'
'        Qry = "Insert into sp_partsmatrix.TMP_SPITR_DTLS(REFERENCE,TRANSTYPE,SPAREPARTID,`OUT`,POREF,UPDATEDBY,SEQNO) VALUES "
'        mySTFNO = "STF#" & txtSTFNo.Text
'        SEQNO = 0
'        For i = 2 To flxParts.Rows - 1
'            If i > 2 Then Qry = Qry & ","
'           SEQNO = SEQNO + 1
'           Qry = Qry & "('" & mySTFNO & "', 'OUT' ," & Val(flxParts.TextMatrix(i, 7)) & "," & Val(flxParts.TextMatrix(i, 5)) & ",'" & cboRefNo.Text & "'," & UserID & "," & SEQNO & ") "
'
'        Next i
'
'
'       Call LYRcon.Execute(Qry)
'
'Exit Sub
'er:
'    DoEvents
'    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTableSPITR()")
'End Sub
'
'Sub CreateTempTableSPAREPARTS()
'On Error GoTo er
'        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_SPAREPARTS_DTLS;"
'
'         DoEvents
'       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_SPAREPARTS_DTLS( " & _
'        "   QTYONHAND Int(11)      " & _
'        "   ,UPDATEDBY Int(11)      " & _
'        "   ,SPAREPARTID Int(11)      " & _
'        ");"
'
'        Qry = "Insert into sp_partsmatrix.TMP_SPAREPARTS_DTLS(QTYONHAND,UPDATEDBY,SPAREPARTID) VALUES "
'        For i = 2 To flxParts.Rows - 1
'        If i > 2 Then Qry = Qry & ","
'
'           Qry = Qry & "(" & Val(flxParts.TextMatrix(i, 5)) & ", " & UserID & ", " & Val(flxParts.TextMatrix(i, 7)) & ") "
'        End If
'        Next i
'
'
'       Call LYRcon.Execute(Qry)
'
'Exit Sub
'er:
'    DoEvents
'    Call modObjTools.ErrorMessage(Err.Description, "Error->TMP_SPAREPARTS_DTLS()")
'End Sub
'********* END CODE UPDATED: 3/16/17 *****************

Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_RSTOCKTRANSFER_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_RSTOCKTRANSFER_DTLS( " & _
        "    SPAREPARTID Int(11)      " & _
        "   ,QTY Int(11)      " & _
        "   ,UNIT Varchar(10)      " & _
        "   ,POREF Varchar(30)      " & _
        ");"
        
        Qry = "Insert into sp_partsmatrix.TMP_RSTOCKTRANSFER_DTLS(SPAREPARTID,QTY,UNIT,POREF) VALUES "
'********************** CODE UPDATED: 06/16/17 **************************
        For i = 2 To flxParts.Rows - 1
            If i = flxParts.Rows - 1 Then
                If i > 2 Then Qry = Qry & ","
                    Qry = Qry & "(" & Val(flxParts.TextMatrix(i, 7)) & "," & Val(flxParts.TextMatrix(i, 5)) & ",'PCS', '" & cboRefNo.Text & "') "
                GoTo SkipData
            End If
            If flxParts.TextMatrix(i, 2) <> flxParts.TextMatrix(i + 1, 6) Then
                If i > 2 Then Qry = Qry & ","
                    Qry = Qry & "(" & Val(flxParts.TextMatrix(i, 7)) & "," & Val(flxParts.TextMatrix(i, 5)) & ",'PCS', '" & cboRefNo.Text & "') "
            End If
SkipData:
        Next i
'****************** END CODE UPDATED: 06/16/17 **************************
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Public Sub flxColChanger()

With flxParts
    'MERGE COL =FALSE
    .TextMatrix(0, 1) = "NO."
    .TextMatrix(0, 2) = "PARTNUMBER"
    .TextMatrix(0, 3) = "DESCRIPTION"
    .TextMatrix(0, 4) = "PARTSGROUP"
    .TextMatrix(0, 5) = "QTY"
End With
    Call flxWidth
End Sub

Public Sub flxColChangerForSumm()

With flxPartsTtl
    'MERGE COL =FALSE
    .TextMatrix(0, 1) = "NO."
    .TextMatrix(0, 2) = "MODEL BRAND (PM)"
    .TextMatrix(0, 3) = "TTL PN"
    .TextMatrix(0, 4) = "TTL SRQ"
    .TextMatrix(0, 5) = "TTL ALQ"
    .TextMatrix(0, 6) = "EFFICIENCY"
End With
    
    Call flxWidthForSumm
    
End Sub



Public Sub flxWidth()

    With flxParts
        '.Clear
        .Cols = 8
        .FormatString = "|<NO |<PARTNUMBER |<DESCRIPTION |<PARTSGROUP |^QTY"
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.04  'NO
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.15 'PARTNUMBER
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.5 'DESCRIPTION
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.15 'PARTSGROUP
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.05 'QTY
        .ColWidth(6) = 0 '(.Width - .ColWidth(0)) * 0.06 'MODELBRAND
        .ColWidth(7) = 0 '(.Width - .ColWidth(0)) * 0.06 'SPAREPARTID
    End With
End Sub

Public Sub flxWidthForSumm()

    With flxPartsTtl
        '.Clear
        .Cols = 6
        .FormatString = "|<NO |<MODEL BRAND (PM) |^TTL PN |^TTL SRQ |^TTL ALQ |^EFFICIENCY"
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.04  'NO
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2 'MODEL BRAND (PM)
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'TTL PN
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'TTL SRQ
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1 'TTL ALQ
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'EFFICIENCY
       ' .ColWidth(7) = 0 '(.Width - .ColWidth(0)) * 0.06 'MODEL
       ' .ColWidth(8) = 0 '(.Width - .ColWidth(0)) * 0.06 'ID

    End With
End Sub

Private Sub cboRefNo_Click()
    
'************* CODE UPDATED: 3/14/17 *****************
    pBar.Visible = True
'********* END CODE UPDATED: 3/14/17 *****************

    matrixID = arrREFNO(cboRefNo.ListIndex, 0)
    Call ShowRecord(BranchCode, cboRefNo.Text, matrixID, "TRUE", "")
End Sub

Private Sub cboRefNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboRefNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAddress, KeyAscii, ValidChars)
End Sub

Private Sub cboTransTo_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cboTransTo
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboTransTo-Change()")
End Sub



Private Sub cboTransTo_Click()
  
    DoEvents
    BranchCode = arrBRANCH(cboTransTo.ListIndex, 2)
    txtAddress.Text = arrBRANCH(cboTransTo.ListIndex, 3)
    Call CboDisplayREFNo
   

    cboRefNo.SetFocus
End Sub

Private Sub cboTransTo_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cboTransTo.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Sub clearform()
    flxParts.Rows = 1
    flxParts.Rows = 2
    flxPartsTtl.Rows = 1
    flxPartsTtl.Rows = 2
    cboTransTo.Text = ""
    txtAddress.Text = ""
    cboRefNo.Text = ""
    txtSTFNO.Text = "<AUTO>"
    txtSTFDate.Text = Format(Date, "mm/dd/yyyy")
    txtStatus.Text = ""
    txtPrepBy.Text = ""
    txtAppBy.Text = ""
    chkFinal.Value = 0
    txtDateStart.Text = Format(Date, "mm/dd/yyyy")
 
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->SAVE
'5 ->CANCEL


    'hide all button
    For i = 0 To 4
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
    For i = 0 To 4
        cmdTrans(i).Visible = True
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
    Next i
    
    
    End If
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    cboTransTo.Locked = DataEntry
    txtAddress.Locked = DataEntry
    cboRefNo.Locked = DataEntry
    txtSTFNO.Locked = DataEntry
    txtSTFDate.Locked = DataEntry
    txtStatus.Locked = DataEntry
    txtPrepBy.Locked = DataEntry
    txtAppBy.Locked = DataEntry
    txtDateStart.Locked = DataEntry
   
    If cboTransTo.Locked Then
        cboTransTo.BackColor = txtHoldGrayBackColor.BackColor
        txtAddress.BackColor = txtHoldGrayBackColor.BackColor
        cboRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtSTFNO.BackColor = txtHoldGrayBackColor.BackColor
        txtSTFDate.BackColor = txtHoldGrayBackColor.BackColor
        txtStatus.BackColor = txtHoldGrayBackColor.BackColor
        txtPrepBy.BackColor = txtHoldGrayBackColor.BackColor
        txtAppBy.BackColor = txtHoldGrayBackColor.BackColor
        txtDateStart.BackColor = txtHoldGrayBackColor.BackColor

    Else
        cboTransTo.BackColor = txtHoldWhiteBackColor.BackColor
        txtAddress.BackColor = txtHoldWhiteBackColor.BackColor
        cboRefNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtSTFNO.Locked = True
        txtSTFDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtStatus.Locked = True
        txtPrepBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtAppBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtDateStart.BackColor = txtHoldWhiteBackColor.BackColor
        
    End If
End Sub


Private Sub cboTransTo_KeyDown(KeyCode As Integer, Shift As Integer)
'************** CODE UPDATED: 04/07/17 **************************
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
'********** END CODE UPDATED: 04/07/17 **************************
End Sub

Private Sub cboTransTo_KeyPress(KeyAscii As Integer)
'************** CODE UPDATED: 04/07/17 **************************
Dim cbolistIndex As Integer

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboTransTo, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cboTransTo, cboTransTo.Text)
        
        If cbolistIndex < 0 Then
            cboTransTo.SetFocus
        Else
            cboTransTo.ListIndex = cbolistIndex
        End If
    
        Call cboTransTo_Click
        
    End If
'********** END CODE UPDATED: 04/07/17 **************************
End Sub



Private Sub cmdConfirm_Click()
    Dim MyRef       As String
    Dim rsConfirm   As New ADODB.Recordset
    Dim boQty       As Integer
    
        ans = MsgBox("Do you want to confirm this transaction?", vbInformation + vbYesNo, "Confirm Intransit")
                If ans = vbNo Then Exit Sub
        
        On Error GoTo LocalErr
        If txtDriver.Text = "" And txtChecker.Text = "" Then
            MsgBox "Please fill the Driver Name and Truck Plate No.", vbExclamation, "Confirm Intransit"
        Else
            Call CreateTempTableFinalTransfer
            Call modPMTransfer.sp_StockTransfer_FinalTransfer(txtDriver.Text, txtChecker.Text, cboRefNo.Text, tmpRs)
            txtStatus.Text = tmpRs!Status
        End If
        MsgBox "Stock Transfer Ref#" & cboRefNo.Text & " is INTRANSIT.", vbOKOnly, "Successfully Save"
        tabConfirm.Visible = False
        
        With cmdTrans()
            For i = 0 To .Count - 1
                cmdTrans(i).Enabled = True
            Next i
        End With
        
        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
       
        Exit Sub
LocalErr:
    MsgBox Err.Description & vbCrLf & ". Please Contact MIS", vbCritical, "Confirm Stock Transfer"
End Sub

Private Sub cmdQuit_Click()
    tabConfirm.Visible = False
    With cmdTrans()
        For i = 0 To .Count - 1
            cmdTrans(i).Enabled = True
        Next i
    End With
        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
End Sub

Private Sub cmdTrans_Click(Index As Integer)
If Index = 0 Then 'Add
    formAction = "ADD"
    AddNew = True
    
    Call clearform
    Call CboDisplay
    Call LockData(True)
    Call SetButtonVisibility(True)
    txtPrepBy.Text = EmployeesName
    chkFinal.Enabled = True
ElseIf Index = 1 Then 'CONFIRM INSTRANSIT

    'Call RXResizer1_ResizeComplete
    If cboRefNo.Text = "" Then
        MsgBox "No available record", vbOKOnly + vbExclamation, "Confirm Intransit"
        Exit Sub
    End If
    If InStr(txtStatus.Text, "CANCELLED") > 0 Then
        MsgBox "Unable to Confirm " & txtSTFNO.Text & " because it was already Cancelled.", vbOKOnly + vbExclamation, "Confirm Intransit"
        Exit Sub
    End If
    If txtStatus.Text = "INTRANSIT" Then
        MsgBox "Stock Transfer " & txtSTFNO.Text & " already Confirm", vbOKOnly + vbExclamation, "Confirm Intransit"
        Exit Sub
    End If
'**************** CODE UPDATED: 04/06/2017 *************************
    If txtStatus.Text <> "PACKING" Then
        If txtStatus.Text <> "FINAL" Then
            MsgBox "Unable to Confirm Intransit Ref. No. " & txtSTFNO.Text & " because record is not final.", vbOKOnly + vbExclamation, "Edit"
            Exit Sub
        End If
    End If

    If chkFinal.Value = 1 Then
        tabConfirm.Visible = True
        tabConfirm.Caption = "CONFIRM INTRANSIT"
        tabConfirm.Left = (Me.Width - tabConfirm.Width) / 2
        With cmdTrans()
            For i = 0 To .Count - 1
                cmdTrans(i).Enabled = False
            Next i
        End With
            Frame1.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
    Else
        formAction = "EDIT"
        AddNew = True
        Call SetButtonVisibility(True)
        chkFinal.Enabled = True
    End If
'*********** END CODE UPDATED: 04/06/2017 *************************

ElseIf Index = 2 Then 'View Summary
    frmSummaryRepl.mySrc = "Replenishment StockTransfer"
    frm_OnProcess2.OwnerForm = "View Summary Repl"
    frmSummaryRepl.Show
    Unload Me
        
ElseIf Index = 3 Then 'Print
'************* CODE UPDATED: 3/17/17 *****************
    If txtStatus.Text = "PICKING" Or txtStatus.Text = "PACKING" Or txtStatus.Text = "INTRANSIT" Or txtStatus.Text = "FINAL" Then
    Qry = "call sp_partsmatrix.sp_StockTransfer_getDetails('" & sysBranchCode & "','" & BranchCode & "','" & cboRefNo.Text & "'," & matrixID & ",'" & "STF#" & txtSTFNO.Text & "','" & pBOOLEAN & "')"
    Qry2 = "call sp_partsmatrix.sp_StockTransfer_getDetailsTtl('" & sysBranchCode & "','" & BranchCode & "','" & cboRefNo.Text & "'," & matrixID & ",'" & pBOOLEAN & "','" & "STF#" & txtSTFNO.Text & "')"
    frmPrintPO.QryPrn = Qry
    frmPrintPO.QryPrn2 = Qry2
        
    frmPrintPO.ParamValue1 = SysBranch
    frmPrintPO.ParamValue2 = SysAddress
    frmPrintPO.ParamValue3 = cboTransTo.Text
    frmPrintPO.ParamValue4 = txtAddress.Text
    frmPrintPO.ParamValue5 = cboRefNo.Text
    frmPrintPO.ParamValue6 = txtSTFNO.Text
    frmPrintPO.ParamValue7 = txtSTFDate.Text
    frmPrintPO.ParamValue8 = txtStatus.Text
    frmPrintPO.ParamValue9 = txtPrepBy.Text
    frmPrintPO.ParamValue10 = txtAppBy.Text
    frmPrintPO.ParamValue11 = txtDateStart.Text
    frmPrintPO.MySource = "Replenishment StockTransfer"
    frmPrintPO.Show
    Else
        MsgBox "No Available Stock Transfer to Print", vbInformation + vbOKOnly, "Print"
    End If
'********* END CODE UPDATED: 3/17/17 *****************
        
ElseIf Index = 4 Then 'Exit
    formAction = "EXIT"
    ProperExit = True
    Call clearform
    Unload Me
ElseIf Index = 5 Then 'Save
    If txtAppBy.Text = "" Then
        MsgBox "APPROVED BY IS REQUIRED!", vbExclamation, "Validation"
        Exit Sub
    End If

    If txtStatus.Text = "RESERVED" Then
        MsgBox "REPL REF #:" & cboRefNo.Text & " IS NOT YET PICKED!", vbExclamation, "Validation"
        Exit Sub
    End If
    
    ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
    
    If ans = vbNo Then Exit Sub
            
    If formAction = "ADD" Then
        Call CreateTempTable
        DoEvents
'************* CODE UPDATED: 3/17/17 *****************
        MsgBox ("The status of REF# " + cboRefNo.Text + " change to Packing."), vbInformation, "PO Status"
        Call modPMTransfer.sp_StockTransfer_create(Format(txtSTFDate.Text, "YYYY-mm-dd"), cboTransTo.Text, txtAddress.Text, cboRefNo.Text, txtPrepBy.Text, txtAppBy.Text, UserID, strModelBrand, matrixID, Val(chkFinal.Value), tmpRs)
        MsgBox "Spareparts for Packing Successfully Saved!", vbInformation + vbOKOnly, "SAVE"
        txtSTFNO.Text = tmpRs!TRANSNO
        txtStatus.Text = tmpRs!Status
'********* END CODE UPDATED: 3/17/17 *****************
'************* CODE UPDATED: 4/03/17 *****************
    Else
        Call CreateTempTable
        DoEvents
        Call modPMTransfer.sp_StockTransfer_update(Format(txtSTFDate.Text, "YYYY-mm-dd"), cboTransTo.Text, txtAddress.Text, cboRefNo.Text, txtPrepBy.Text, txtAppBy.Text, UserID, intSTFID, Val(chkFinal.Value), tmpRs)
        If chkFinal <> 0 Then
            MsgBox "Spareparts for Finalizing Successfully Saved!", vbInformation + vbOKOnly, "SAVE"
            txtStatus.Text = tmpRs!Status
            If txtStatus.Text = "FINAL" Then
                cmdTrans(1).Caption = "Confirm Intransit"
            End If
        End If
    End If
'********* END CODE UPDATED: 4/03/17 *****************
        Call SetButtonVisibility(False)
        Call LockData(False)
        chkFinal.Enabled = False
        formAction = ""
ElseIf Index = 6 Then 'Cancel
    ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
    
    If ans = vbNo Then Exit Sub
    
    Call SetButtonVisibility(False)
    Call LockData(False)
'        Call clearform
    formAction = ""
    chkFinal.Enabled = False
End If

End Sub




Private Sub Form_Load()
    'Call flxColChanger
   ' Call flxColChangerForSumm
    Call LockData(False)
    
    chkFinal.Enabled = False
    txtSTFNO.Text = "<AUTO>"
    txtSTFDate.Text = Format(Date, "mm/dd/yyyy")
    txtDateStart.Text = Format(Date, "mm/dd/yyyy")
    '===================BUTTON POSITION===================
    
    If cmdTrans(0).Visible = False Then
            'Align Save Button
            cmdTrans(5).Left = cmdTrans(0).Left
            cmdTrans(5).Top = cmdTrans(0).Top
            'Align Cancel Button
            cmdTrans(6).Left = cmdTrans(1).Left
            cmdTrans(6).Top = cmdTrans(1).Top
     End If
            
End Sub

Public Sub CboDisplay()
On Error GoTo er

    'If formAction = "add" Then
        'Get Data from database
        DoEvents
        Call modPMTransfer.sp_get_Branch(tmpRs)
        If tmpRs.State = 1 Then
           
         Call modObjTools.fillArray(modGlobalVar.AllArrBranch, tmpRs)
          
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
        
         'Load Data to combo box
        Call SetCboListData
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub CboDisplayREFNo()
On Error GoTo er
    'Get Data from database
    DoEvents
    Call modPMTransfer.sp_getRefNo(BranchCode, tmpRs)
    If tmpRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cboRefNo, tmpRs, arrREFNO, "BRANCHSPOID", "REFNO")
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub


Sub SetCboListData()
On Error GoTo er

    If cboTransTo.Text = "" Then
        DoEvents
        Call modObjTools.addToListObjectUsingArrayWithCols(cboTransTo, arrBRANCH, modGlobalVar.AllArrBranch, 4, 1)
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->SetCboListData()")
End Sub



Private Sub Form_Resize()
Call RXResizer1_ResizeComplete
End Sub



Private Sub RXResizer1_ResizeComplete()
    Call flxColChanger
    Call flxColChangerForSumm
'************* CODE UPDATED: 3/14/17 *****************
'    'Frame width
'    Frame1.Width = Me.Width - 465
'    Frame2.Width = Me.Width - 465
'    Frame3.Width = Me.Width - 465
'
'    'Frame1 label Left
'    lblTransfer.Left = 120
'    lblAddress.Left = 120
'    lblREFNO.Left = 120
'    lblSTFNo.Left = Frame1.Width - 2655
'    lblSTFDATE.Left = Frame1.Width - 2655
'    lblStatus.Left = Frame1.Width - 2655
'
'    'Frame1 Label Width
'    lblTransfer.Width = 1575
'    lblAddress.Width = 1575
'    lblREFNO.Width = 1575
'    lblSTFNo.Width = 855
'    lblSTFDATE.Width = 855
'    lblStatus.Width = 855
'
'
'
'    'Frame1 Combobox Left
'    cboTransTo.Left = 1800
'    cboRefNo.Left = 1800
'
'    'Frame1 Combobox Width
'    cboTransTo.Width = 3855
'    cboRefNo.Width = 3855
'
'
'    'Frame1 Text Left
'    txtAddress.Left = 1800
'
'    txtSTFNo.Left = Frame1.Width - 1815
'    txtSTFDate.Left = Frame1.Width - 1815
'    txtStatus.Left = Frame1.Width - 1815
'
'    'Frame1 Text Width
'    txtAddress.Width = 3855
'    txtSTFNo.Width = 1575
'    txtSTFDate.Width = 1575
'    txtStatus.Width = 1575
'
'    'Frame2 FlexGrid Width
'    flxPARTS.Width = Frame2.Width - 240
'    flxPartsTtl.Width = Frame2.Width - 240
''
''    'Frame2 FlexGrid Height
''    flxParts.Height = Frame2.Height - 2400
''    flxPartsTtl.Height = Frame2.Height - 4560
'
'    'Frame3 Label Left
'    lblPrepBy.Left = 240
'    lblAppby.Left = 240
'    lblSTARTREPL.Left = 240
'
'    'Frame3 Label Width
'    lblPrepBy.Width = 1215
'    lblAppby.Width = 1215
'    lblSTARTREPL.Width = 2655
'
'    'Frame3 Text Left
'    txtPrepBy.Left = 1560
'    txtAppBy.Left = 1560
'    txtDateStart.Left = 3000
'
'    'Frame3 Text Width
'    txtPrepBy.Width = 3615
'    txtAppBy.Width = 3615
'    txtDateStart.Width = 2175
'********* END CODE UPDATED: 3/14/17 *****************
End Sub

Public Sub getSTFID(mySTFID As Integer, mySTFNO As String)
    intSTFID = mySTFID
    strSTFNO = mySTFNO
End Sub


Public Sub ShowRecord(myBranchCode As String, myREFNO As String, MyID As Integer, MyBoolean As String, mySTATUS As String)
Dim strSTFNoSPITR As String
    Screen.MousePointer = vbHourglass
    cboRefNo.Text = myREFNO
'************* CODE UPDATED: 6/12/17 *****************
    pBOOLEAN = MyBoolean ' use for printing
    strModelBrand = ""
    BranchCode = myBranchCode
    frm_OnProcess2.OwnerForm = "Replenish StockTransfer"
    frm_OnProcess2.Show
'********* END CODE UPDATED: 6/12/17 *****************

'************* CODE UPDATED: 3/14/17 *****************
    If MyBoolean = "FALSE" Then
        txtSTFNO.Text = strSTFNO
        strSTFNoSPITR = "STF#" & strSTFNO
    End If
'********* END CODE UPDATED: 3/14/17 *****************
    DoEvents
    Call modPMTransfer.sp_get_BranchName(myBranchCode, tmpRs)
    cboTransTo.Text = tmpRs!branchName
    txtAddress.Text = tmpRs!Address
    
'************* CODE UPDATED: 3/14/17 *****************
    If mySTATUS = "PACKING" Or mySTATUS = "" Then
        chkFinal.Value = 0
        chkFinal.Enabled = True
    ElseIf mySTATUS = "FINAL" Or mySTATUS = "INTRANSIT" Then
        chkFinal.Value = 1
        cmdTrans(1).Caption = "Confirm Intransit"
        chkFinal.Enabled = False
    End If
'********* END CODE UPDATED: 3/14/17 *****************

    DoEvents
    Call modPMTransfer.sp_StockTransfer_getDetails(myBranchCode, myREFNO, MyID, strSTFNoSPITR, MyBoolean, tmpRs)
 
    If tmpRs.State <> 1 Then Exit Sub
    
'************* CODE UPDATED: 3/14/17 *****************
    pBar.Visible = True
    pBar.Value = 0
'********* END CODE UPDATED: 3/14/17 *****************

    If tmpRs.RecordCount <= 0 Then
'************* CODE UPDATED: 3/14/17 *****************
        pBar.Visible = False
        pBar.Value = 0
        Call Unload(frm_OnProcess2)
'********* END CODE UPDATED: 3/14/17 *****************
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With tmpRs
        If .RecordCount > 0 Then
            .MoveFirst
           ' intTtlPN = .RecordCount
            flxParts.Rows = .RecordCount + 2
        '************* CODE UPDATED: 3/14/17 *****************
            pBar.Max = .RecordCount
        '********* END CODE UPDATED: 3/14/17 *****************
           j = 0
            For i = 2 To flxParts.Rows - 1
                j = j + 1
                flxParts.TextMatrix(i, 1) = j
                flxParts.TextMatrix(i, 2) = !PartNumber
                flxParts.TextMatrix(i, 3) = !Description
                flxParts.TextMatrix(i, 4) = !partsGroup
                flxParts.TextMatrix(i, 5) = !qty
                flxParts.TextMatrix(i, 6) = !MODELBRAND
                flxParts.TextMatrix(i, 7) = !SPAREPARTID
                .MoveNext
            '************* CODE UPDATED: 3/14/17 *****************
                pBar.Value = pBar.Value + 1
            '********* END CODE UPDATED: 3/14/17 *****************
            Next i
            
            
            With flxParts
        
                .TextMatrix(1, 1) = .TextMatrix(2, 6)
                .TextMatrix(1, 2) = .TextMatrix(2, 6)
                .TextMatrix(1, 3) = .TextMatrix(2, 6)
                .TextMatrix(1, 4) = .TextMatrix(2, 6)
                .TextMatrix(1, 5) = .TextMatrix(2, 6)
                .TextMatrix(1, 6) = .TextMatrix(2, 6)
                'SET MERGING
                .MergeCells = 1
                .MergeRow(1) = True
                .MergeCol(1) = True
                .MergeCol(1) = True
                For j = 1 To 6
                .MergeCol(j) = False
                Next j
            End With
'**************** CODE UPDATED: 06/14/17 ******************
            For i = 2 To flxParts.Rows - 1
                If i = flxParts.Rows - 1 Then
                    GoTo SkipData
                End If
                If flxParts.TextMatrix(i, 6) <> flxParts.TextMatrix(i + 1, 6) Then
                    
                    flxParts.Rows = flxParts.Rows + 1
                        For z = flxParts.Rows - 1 To i Step -1
                            If z = i Then
                              
                                For x = 1 To flxParts.Cols - 1
                                flxParts.TextMatrix(z + 1, x) = flxParts.TextMatrix(i + 2, 6)
                                    'SET MERGING
                                    
                                    flxParts.MergeCells = 1
                                    flxParts.MergeRow(z + 1) = True
                                    flxParts.MergeCol(0) = True
                                    flxParts.MergeCol(0) = True
                                    For j = 1 To 6
                                    flxParts.MergeCol(j) = False
                                    Next j
                                    
                                Next x
                            Else
                                For x = 1 To flxParts.Cols - 1
                                flxParts.TextMatrix(z, x) = flxParts.TextMatrix(z - 1, x)
                                Next x
                            End If
                        Next z
                End If
SkipData:
            Next i
'**************** CODE UPDATED: 06/14/17 ******************
        End If
    End With
    If tmpRs.State <> 0 Then
        Call Unload(frm_OnProcess2)
        Screen.MousePointer = vbDefault
        tmpRs.Close
    End If
    
    
    Call modPMTransfer.sp_StockTransfer_getDetailsTtl(myBranchCode, myREFNO, MyID, MyBoolean, strSTFNoSPITR, tmpRs)
    If tmpRs.State <> 1 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        pBar.Visible = False
        pBar.Value = 0
        Call Unload(frm_OnProcess2)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    With tmpRs
     If .RecordCount > 0 Then
        .MoveFirst
        '************* CODE UPDATED: 3/14/17 *****************
        If MyBoolean = "FALSE" Then
           txtPrepBy.Text = Format$(!PROCESSEDBY, "")
           txtAppBy.Text = Format$(!APRVBY, "")
        End If
        'strModelBrand = IIf(IsNull(!tmpMODELBRAND), "", !tmpMODELBRAND)
        '********* END CODE UPDATED: 3/14/17 *****************
        txtStatus.Text = Format$(!Status, "")
        txtDateStart.Text = Format$(!DATEPREPARED, "")
        flxPartsTtl.Rows = .RecordCount + 1
        
        j = 0
        For i = 1 To flxPartsTtl.Rows - 1
            j = j + 1
            flxPartsTtl.TextMatrix(i, 1) = j
            flxPartsTtl.TextMatrix(i, 2) = !MODELBRAND
            flxPartsTtl.TextMatrix(i, 3) = !TOTALPN
            flxPartsTtl.TextMatrix(i, 4) = !TOTALSRQ
            flxPartsTtl.TextMatrix(i, 5) = !TOTALALQ
            .MoveNext
        Next i
        
'        strModelBrand = ""
'        For i = 1 To flxPartsTtl.Rows - 1
'            If i > 0 Then
'                strModelBrand = strModelBrand & flxPartsTtl.TextMatrix(i, 2) & ","
'            End If
'        Next
'        strModelBrand = Replace(Left(strModelBrand, Len(strModelBrand) - 1), "HONDA", "") ' To remove the last slash
    End If
    End With
    If tmpRs.State <> 0 Then
        Call Unload(frm_OnProcess2)
        Screen.MousePointer = vbDefault
        tmpRs.Close
    End If

'************* CODE UPDATED: 3/14/17 *****************
    Call Unload(frm_OnProcess2)
    pBar.Visible = False
    pBar.Value = 0
'********* END CODE UPDATED: 3/14/17 *****************

End Sub


Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAddress, KeyAscii, ValidChars)
End Sub

Private Sub txtAppBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAppBy, KeyAscii, ValidChars)
End Sub

Private Sub txtChecker_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtChecker, KeyAscii, ValidChars)
End Sub

Private Sub txtDriver_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtDriver, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtChecker.SetFocus
End Sub

Private Sub txtPrepBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtPrepBy, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtAppBy.SetFocus
End Sub
