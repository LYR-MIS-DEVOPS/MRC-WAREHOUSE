VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form DataEntry 
   Caption         =   "Parts Matrix"
   ClientHeight    =   9810
   ClientLeft      =   120
   ClientTop       =   555
   ClientWidth     =   19155
   ControlBox      =   0   'False
   ForeColor       =   &H000000C0&
   Icon            =   "PMDataEntry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   19155
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Tag             =   "XY"
      Top             =   8880
      Width           =   18975
      Begin VB.TextBox txtUP 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Tag             =   "XY"
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblUpdateBy 
         Alignment       =   1  'Right Justify
         Caption         =   "UPDATED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   17040
      TabIndex        =   41
      Tag             =   "XY"
      Top             =   1800
      Width           =   1935
   End
   Begin lvButton.lvButtons_H btnFind 
      Height          =   315
      Left            =   15600
      TabIndex        =   40
      Tag             =   "XY"
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Find"
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
      Image           =   "PMDataEntry.frx":57E2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   600
      TabIndex        =   38
      Tag             =   "XY"
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Show Option"
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
      ImgAlign        =   3
      Image           =   "PMDataEntry.frx":5D7C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdHide 
      Height          =   375
      Left            =   600
      TabIndex        =   37
      Tag             =   "XY"
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Hide Option"
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
      ImgAlign        =   3
      Image           =   "PMDataEntry.frx":6316
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   120
      TabIndex        =   21
      Tag             =   "XY"
      Top             =   1800
      Visible         =   0   'False
      Width           =   18975
      Begin MSFlexGridLib.MSFlexGrid flxPMSum 
         Height          =   1575
         Left            =   240
         TabIndex        =   36
         Tag             =   "XY"
         Top             =   480
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   7
         BackColorBkg    =   14737632
         FormatString    =   "|^CLASSIFICATION |^TOTAL PN |^COMPLETE PARTS |^MAINTENANCE PARTS |^PLASTIC PARTS |^UNIVERSAL PARTS"
      End
      Begin MSFlexGridLib.MSFlexGrid flxClass 
         Height          =   1455
         Left            =   3720
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   3
      End
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   19320
      TabIndex        =   32
      Tag             =   "XY"
      Top             =   7800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFlxBrand 
      Height          =   285
      Left            =   19320
      TabIndex        =   31
      Tag             =   "XY"
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtpmID 
      Height          =   285
      Left            =   19320
      TabIndex        =   30
      Tag             =   "XY"
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbPG 
      Height          =   315
      ItemData        =   "PMDataEntry.frx":68B0
      Left            =   19320
      List            =   "PMDataEntry.frx":68C0
      TabIndex        =   26
      Tag             =   "XY"
      Top             =   4320
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox txtCountrP 
      Height          =   285
      Left            =   19320
      MaxLength       =   50
      TabIndex        =   25
      Tag             =   "XY"
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtRP 
      Height          =   285
      Left            =   19320
      MaxLength       =   11
      TabIndex        =   24
      Tag             =   "XY"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSPECMODEL 
      Height          =   285
      Left            =   19320
      MaxLength       =   50
      TabIndex        =   23
      Tag             =   "XY"
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPARTNO 
      Height          =   285
      Left            =   19320
      MaxLength       =   30
      TabIndex        =   22
      Tag             =   "XY"
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   19800
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   19320
      Top             =   480
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   19320
      TabIndex        =   18
      Text            =   "HoldBackColor"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   19320
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Tag             =   "XY"
      Top             =   480
      Width           =   18975
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   720
         Width           =   6135
      End
      Begin MSMask.MaskEdBox txtDCreated 
         Height          =   375
         Left            =   16920
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtTOTALPN 
         Height          =   285
         Left            =   16920
         TabIndex        =   8
         Tag             =   "XY"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtModel 
         Height          =   285
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "XY"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtBrand 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "XY"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         Caption         =   "MC DESCRIPTION:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Tag             =   "XY"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblTPN 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL PN:"
         Height          =   255
         Left            =   15480
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblDCreated 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE CREATED:"
         Height          =   255
         Left            =   15480
         TabIndex        =   13
         Tag             =   "XY"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblBrand 
         Alignment       =   1  'Right Justify
         Caption         =   "MC BRAND:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Tag             =   "XY"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblModel 
         Alignment       =   1  'Right Justify
         Caption         =   "MC MODELS:"
         Height          =   255
         Left            =   4080
         TabIndex        =   11
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
      Tag             =   "XY"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&New Parts Matrix"
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
      Image           =   "PMDataEntry.frx":68EF
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Tag             =   "XY"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&Edit Parts Matrix"
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
      Image           =   "PMDataEntry.frx":6E89
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Tag             =   "XY"
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
      Image           =   "PMDataEntry.frx":7423
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   19320
      TabIndex        =   19
      Top             =   960
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
      Image           =   "PMDataEntry.frx":79BD
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   19320
      TabIndex        =   20
      Top             =   1440
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
      Image           =   "PMDataEntry.frx":7F57
      cBack           =   -2147483633
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPARTNO 
      Height          =   1935
      Left            =   19320
      TabIndex        =   27
      Tag             =   "XXYY"
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCounterPart 
      Height          =   1935
      Left            =   19800
      TabIndex        =   28
      Tag             =   "XXYY"
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   7800
      TabIndex        =   3
      Tag             =   "XY"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "E&xit Parts Matrix"
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
      Image           =   "PMDataEntry.frx":84F1
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxPARTS 
      Height          =   6735
      Left            =   120
      TabIndex        =   33
      Tag             =   "XY"
      Top             =   2160
      Width           =   18975
      _ExtentX        =   33470
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   17
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   14737632
      ScrollBars      =   2
      AllowUserResizing=   3
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
   Begin VB.CommandButton cmdAltD 
      Caption         =   "&D"
      Height          =   255
      Left            =   1680
      TabIndex        =   34
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cmdAltA 
      Caption         =   "&A"
      Height          =   255
      Left            =   1200
      TabIndex        =   35
      Top             =   5160
      Width           =   375
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   42
      Tag             =   "XY"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
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
      Image           =   "PMDataEntry.frx":8A8B
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   19320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   19395
      m_MDI_Height    =   10380
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Row     Alt A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Row Alt D"
      End
   End
End
Attribute VB_Name = "DataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'String Variables
Dim Tmp() As String
Dim Ans As String
Dim formAction As String
Dim Qry As String
Dim flxfindQry As String
Dim MyErr As String
Dim myMsg As String
Dim strPartial As String, strTotal As String
Dim cmbstr As String
Dim SPid As String, CPID As String, GENID As String, CPPNO As String
Dim txtstr As String


'Recordset Variables
Dim mainRs As New ADODB.Recordset
Dim cpRs As New ADODB.Recordset
Dim flxfindRs As New ADODB.Recordset
Dim optRow As Integer

'Integer Variables

Dim i As Long, j As Long
Dim MyID As Integer
Dim textval As Integer
Dim numval As Integer
Dim intLoopCount As Integer
Dim intCP As Integer, intMP As Integer, intUP As Integer, intPP As Integer
Dim ActRow As Integer, ActCol As Integer, RCol As Integer, TCol As Integer, CCol As Integer, intRow As Integer
Dim fRow As Integer, fCol As Integer, chckDupCount As Integer, HoldLastRow As Integer

'Boolean Variables
Dim StatSrch As Boolean
Dim PrtBln As Boolean, PrtSrch As Boolean, PrtAdd As Boolean, rowEmpty As Boolean
Dim AddNew As Boolean, blnFound As Boolean, ProperExit As Boolean, blnOpt As Boolean



Public Sub ShowRecord(summID As Long)
On Error GoTo er

    txtPMID.Text = summID 'Used for Edit Verification
    
    
'    Qry = " SELECT partsmatrix_dtls.SPAREPARTID " & _
'          " ,spareparts.PARTNUMBER " & _
'          " ,spareparts.DESCRIPTION " & _
'          " ,partsmatrix_dtls.PMSPBRAND " & _
'          " ,partsmatrix_dtls.PMSIZE " & _
'          " ,partsmatrix_dtls.PMCLASS " & _
'          " ,partsmatrix_dtls.PMAPPMODEL " & _
'          " ,partsmatrix_dtls.PARTSGROUP " & _
'          " ,partsmatrix_dtls.PMREPLCODE " & _
'          " ,(SELECT spareparts.PARTNUMBER From mrc2retail.spareparts Where partsmatrix_dtls.COUNTERPARTID = spareparts.SPAREPARTID) AS COUNTERPART " & _
'          " ,partsmatrix_dtls.COUNTERPARTID " & _
'          " ,partsmatrix_dtls.GENIUNEID " & _
'          " From mrc2retail.partsmatrix_dtls " & _
'          " INNER JOIN mrc2retail.spareparts ON partsmatrix_dtls.SPAREPARTID = spareparts.SPAREPARTID where partsmatrix_dtls.PMATRIXID = '" & txtPMID.Text & "' "
'   Set mainRs = dbCon.Execute(Qry)
    
    Screen.MousePointer = vbHourglass
    
    DoEvents
    Call modPMDataEntry.sp_dataentry_getDetails(txtPMID.Text, mainRs)
    
    If mainRs.State <> 1 Then Exit Sub
    
    If mainRs.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With mainRs
        If .RecordCount > 0 Then
            .MoveFirst
            flxPARTS.Rows = .RecordCount + 1
            For i = 1 To flxPARTS.Rows - 1
            
                'if the value is 0 then Empty or NULL
                SPid = Format$(!SPAREPARTID, 0)
                CPID = Format$(!COUNTERPARTID, 0)
                GENID = Format$(!GENIUNEID, 0)
                
                If SPid = 0 Then
                    SPid = ""
                End If
                If CPID = 0 Then
                    CPID = ""
                End If
                If GENID = 0 Then
                    GENID = ""
                End If
                
            
                flxPARTS.TextMatrix(i, 1) = !PartNumber
                flxPARTS.TextMatrix(i, 2) = !Description
                flxPARTS.TextMatrix(i, 3) = !PMCLASS
                flxPARTS.TextMatrix(i, 4) = !PMSPBRAND
                flxPARTS.TextMatrix(i, 5) = !PMAPPMODEL
                flxPARTS.TextMatrix(i, 6) = !PMSIZE
                
                flxPARTS.TextMatrix(i, 7) = !partsGroup
                flxPARTS.TextMatrix(i, 8) = !PMREPLCODE
                flxPARTS.TextMatrix(i, 9) = Format$(!COUNTERPART, 0)
                flxPARTS.TextMatrix(i, 10) = SPid
                flxPARTS.TextMatrix(i, 11) = CPID
                flxPARTS.TextMatrix(i, 12) = GENID
                .MoveNext
            
            
                'change color of a row
                If Val(flxPARTS.TextMatrix(i, 12)) > 0 Then
                    For j = 1 To 12
                        flxPARTS.Row = i
                        flxPARTS.Col = j
                        flxPARTS.CellBackColor = &HC0C0C0
                        
                    Next j
                End If
            Next i
             
        End If
    End With
    
    If mainRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        mainRs.Close
    End If



   DoEvents
    Call modPMDataEntry.sp_dataentry_getDetails_total(txtPMID.Text, flxfindRs)
    
    If flxfindRs.State <> 1 Then Exit Sub
    
    If flxfindRs.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With flxfindRs
        If .RecordCount > 0 Then
                .MoveFirst
            txtBrand.Text = !PMBRAND
            txtModel.Text = !PMMODEL
            txtDescription.Text = !PMDESCRIPTION
            txtDCreated.Text = Format$(!DateCreated, "mm/dd/yyyy")
            txtTOTALPN.Text = !PMTOTALPN
            txtUP.Text = !UPDATEDBY
        End If
    End With
  If flxfindRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        flxfindRs.Close
    End If


Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->dataEntryShowRecord()")
End Sub



'===================TEXT LOCK===================
Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    txtBrand.Locked = DataEntry
    txtModel.Locked = DataEntry
   ' txtDCreated.Locked = DataEntry
    txtTOTALPN.Locked = DataEntry
    txtPARTNO.Locked = DataEntry
    txtUP.Locked = DataEntry
   
    txtDescription.Locked = DataEntry
   ' txtFind.Locked = DataEntry
    If txtBrand.Locked Then
       ' flxPARTS.Enabled = False
        txtPARTNO.Enabled = False
        txtSPECMODEL.Enabled = False
        cmbPG.Enabled = False
        txtRP.Enabled = False
        txtCountrP.Enabled = False
        cmdAltA.Enabled = False
        cmdAltD.Enabled = False
        txtBrand.BackColor = txtHoldGrayBackColor.BackColor
        txtModel.BackColor = txtHoldGrayBackColor.BackColor
        txtDCreated.BackColor = txtHoldGrayBackColor.BackColor
        txtUP.BackColor = txtHoldGrayBackColor.BackColor
        'txtFind.BackColor = txtHoldGrayBackColor.BackColor
        txtTOTALPN.BackColor = txtHoldGrayBackColor.BackColor
        txtDescription.BackColor = txtHoldGrayBackColor.BackColor
        
        Else
        
        'flxPARTS.Enabled = True
        txtPARTNO.Enabled = True
        txtSPECMODEL.Enabled = True
        cmbPG.Enabled = True
        txtRP.Enabled = True
        txtCountrP.Enabled = True
        cmdAltA.Enabled = True
        cmdAltD.Enabled = True
        txtBrand.BackColor = txtHoldWhiteBackColor.BackColor
        txtModel.BackColor = txtHoldWhiteBackColor.BackColor
        txtDCreated.BackColor = txtHoldWhiteBackColor.BackColor
        txtUP.BackColor = txtHoldWhiteBackColor.BackColor
        txtDescription.BackColor = txtHoldWhiteBackColor.BackColor
        
    End If
End Sub


Public Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_PARTSMATRIX_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_PARTSMATRIX_DTLS( " & _
        "   PMCLASS Varchar(30),         " & _
        "   PMAPPMODEL Varchar(50),      " & _
        "   PMSPBRAND Varchar(50),       " & _
        "   PMSIZE Varchar(50),          " & _
        "   PARTSGROUP Varchar(30),      " & _
        "   PMREPLCODE Int(11),          " & _
        "   SPAREPARTID Int(11),         " & _
        "   COUNTERPARTID Int(11),       " & _
        "   GENIUNEID Int(11),           " & _
        "   SEQNO int(11)                " & _
        ");"

      ' "   PMATRIXID Int(11),           "
        Qry = "insert into sp_partsmatrix.TMP_PARTSMATRIX_DTLS(PMCLASS,PMSPBRAND,PMAPPMODEL,PMSIZE,PARTSGROUP,PMREPLCODE,SPAREPARTID,COUNTERPARTID,GENIUNEID,SEQNO) VALUES "
        
        For i = 1 To flxPARTS.Rows - 1
            If i > 1 Then Qry = Qry & ","
        
           Qry = Qry & "('" & flxPARTS.TextMatrix(i, 3) & "', '" & flxPARTS.TextMatrix(i, 4) & "', '" & flxPARTS.TextMatrix(i, 5) & "', '" & flxPARTS.TextMatrix(i, 6) & "','" & flxPARTS.TextMatrix(i, 7) & "','" & Val(flxPARTS.TextMatrix(i, 8)) & "','" & Val(flxPARTS.TextMatrix(i, 10)) & "','" & Val(flxPARTS.TextMatrix(i, 11)) & "','" & Val(flxPARTS.TextMatrix(i, 12)) & "', " & i & ")"
    
        Next
        
       Call LYRcon.Execute(Qry)

Exit Sub
er:
 Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub


'===================BUTTON VISIBILITY===================
Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->SAVE
'5 ->CANCEL

    Dim i As Integer

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


Private Sub btnFind_Click()
    
    If txtFind.Text = "" Then Exit Sub
    With flxPARTS
        For i = fRow To .Rows - 1
            For j = fCol To TCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 12:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = TCol - 1, i + 1, i)
                    fCol = IIf(j = TCol - 1, 1, j + 1)
                    ActRow = 1
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


Private Sub cmbPG_Click()

    Call cmbPG_Change
    flxPARTS.TextMatrix(ActRow, 7) = cmbPG
        
End Sub



Function ValidateEntry() As Boolean
  ValidateEntry = False
    
    If Trim(txtBrand.Text) = "" Then
        MsgBox "BRAND IS REQUIRED!", vbExclamation, "Validation"
        txtBrand.SetFocus
        Exit Function
    End If
    
    If Trim(txtModel.Text) = "" Then
        MsgBox "MODEL IS REQUIRED!", vbExclamation, "Validation"
        txtModel.SetFocus
        Exit Function
    End If
        
    
    If Trim(txtUP.Text) = "" Then
        MsgBox "DESCRIPTION IS REQUIRED!", vbExclamation, "Validation"
        txtUP.SetFocus
        Exit Function
    End If
    
    ValidateEntry = True

End Function



Function ValidationPartsMatrix() As String
On Error GoTo er
   
    MyErr = ""
    If flxPARTS.TextMatrix(flxPARTS.Rows - 1, 1) = "" Then
        
        If flxPARTS.Rows = 2 Then
            flxPARTS.Rows = 1
            flxPARTS.Rows = 2
            
        Else
        
            flxPARTS.RemoveItem flxPARTS.Rows - 1
        
            For i = 1 To flxPARTS.Rows - 1
                If flxPARTS.TextMatrix(i, 1) <> "" Then
                    txtTOTALPN.Text = i
                End If
            Next i
        
        End If
        
    End If
    
    If flxPARTS.TextMatrix(flxPARTS.Rows - 1, 1) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row of Parts Matrix"


    For i = 1 To flxPARTS.Rows - 1
        If flxPARTS.TextMatrix(i, 7) = "" Then MyErr = MyErr & vbCrLf & "There is no Parts Group @ Row " & i & "!"
    Next i

    If Val(flxPARTS.TextMatrix(flxPARTS.Row, 8)) > 0 Then

        If flxPARTS.TextMatrix(flxPARTS.Row, 9) = "" Then MyErr = MyErr & vbCrLf & "There is no Counter Part @ Row " & i & "!"

    Else
        'do nothing
    End If

    
    ValidationPartsMatrix = MyErr


er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ValidationPartsMatrix()")
End Function


'Public Sub SavingData()
'
'    If formAction = "add" Then
'
'        'Insertion of Data
'        dbCon.Execute "Insert Into mrc2retail.partsmatrix_dtls(PMATRIXID, PMCLASS, PMSPBRAND, PMAPPMODEL, PMSIZE, PARTSGROUP, PMREPLCODE, SPAREPARTID, COUNTERPARTID, GENIUNEID, PMSTDQ, PMLACKING, PMEXCESS, PMSO)" & _
'        "Select PMATRIXID,PMCLASS,PMSPBRAND,PMAPPMODEL,PMSIZE,PARTSGROUP,PMREPLCODE,SPAREPARTID,COUNTERPARTID,GENIUNEID, PMSTDQ, PMLACKING, PMEXCESS, PMSO From sp_partsmatrix.TMP_PARTSMATRIX_DTLS"
'
'
'    ElseIf formAction = "edit" Then
'
'        'Deleting data
'        dbCon.Execute "DELETE From mrc2retail.partsmatrix_dtls " & _
'        "WHERE partsmatrix_dtls.PMATRIXID=" & txtPMID.Text & ";"
'
'        'Updating data
'        dbCon.Execute "UPDATE partsmatrix SET PMBRAND = '" & txtBrand.Text & "', PMMODEL = '" & txtModel.Text & "', PMDESCRIPTION = '" & txtDescription.Text & "', DATECREATED = '" & Format(txtDCreated.Text, "yyyy-mm-dd") & "', UPDATEDBY = '" & txtUP.Text & "' where PMATRIXID = '" & txtPMID.Text & "';"
'
'        'Insertion of data
'       dbCon.Execute "Insert Into mrc2retail.partsmatrix_dtls(PMATRIXID, PMCLASS, PMSPBRAND, PMAPPMODEL, PMSIZE, PARTSGROUP, PMREPLCODE, SPAREPARTID, COUNTERPARTID, GENIUNEID, PMSTDQ, PMLACKING, PMEXCESS, PMSO)" & _
'        "(Select PMATRIXID,PMCLASS,PMSPBRAND,PMAPPMODEL,PMSIZE,PARTSGROUP,PMREPLCODE,SPAREPARTID,COUNTERPARTID,GENIUNEID,PMSTDQ,PMLACKING,PMEXCESS,PMSO From sp_partsmatrix.TMP_PARTSMATRIX_DTLS)"
'
'
'    End If
'End Sub





Private Sub cmdHide_Click()
    Frame3.Visible = False
    cmdHide.Left = cmdShow.Left
    
    flxPARTS.Top = cmdShow.Top + cmdShow.Height
    flxPARTS.Height = Me.Height - 3918
    btnFind.Top = cmdShow.Top
    txtFind.Top = cmdShow.Top
    cmdShow.Visible = True
    cmdHide.Visible = False
    blnOpt = False
    
    'if cmdHide is click while textbox is visible
    txtPARTNO.Visible = False
    txtFlxBrand.Visible = False
    txtSPECMODEL.Visible = False
    txtSize.Visible = False
    txtRP.Visible = False
    txtCountrP.Visible = False
End Sub

Private Sub cmdShow_Click()
    
    If flxPARTS.TextMatrix(1, 1) <> "" Then
        Screen.MousePointer = vbHourglass
    
        Frame3.Visible = True
        flxPARTS.Top = Frame3.Top + Frame3.Height
        flxPARTS.Height = Me.Height - Frame3.Height - 3918 + cmdShow.Height
        btnFind.Top = cmdShow.Top + cmdShow.Height
        txtFind.Top = cmdShow.Top + cmdShow.Height
        cmdHide.Visible = True
        cmdShow.Visible = False
        blnOpt = True
        
        'if cmdShow is click while textbox is visible
        txtPARTNO.Visible = False
        txtFlxBrand.Visible = False
        txtSPECMODEL.Visible = False
        txtSize.Visible = False
        txtRP.Visible = False
        txtCountrP.Visible = False
        
        Call Optshow
    
        Screen.MousePointer = vbDefault
    End If
End Sub

Public Sub Optshow()
    flxPMSum.Rows = 1
    flxPMSum.Rows = 2
    flxClass.Rows = 1
    flxClass.Rows = 2
    

    Call passToFlxClass
    Call distinctClass
    Call passToFlxPMSum
    Call countingPartsGroup
    
    
End Sub





Public Sub passToFlxClass()

 'pass all the data from flxparts to flxclass
 Dim strLoopClass As String, strLoopClassCount As String, intPN As Integer
    intPN = 1
   
    For i = 1 To flxPARTS.Rows - 1
        If flxPARTS.TextMatrix(i, 3) <> "" Then
            strLoopClass = flxPARTS.TextMatrix(i, 3)
            
    
            If strLoopClass = "" Then
                flxClass.RemoveItem (i)
                GoTo SKIP
            End If
            
            For j = 2 To flxPARTS.Rows - 1
                strLoopClassCount = flxPARTS.TextMatrix(j, 3)
                    If strLoopClassCount = strLoopClass Then
                        intPN = intPN + 1
                    End If
            Next j
            
            If i >= flxClass.Rows Then
                flxClass.AddItem ""
            End If
            
            flxClass.TextMatrix(i, 1) = strLoopClass
            flxClass.TextMatrix(i, 2) = intPN
            intPN = 0
        Else
            GoTo SKIP
        End If
SKIP:
    Next i
    
End Sub

'===================BUTTON FUNCTION===================
Private Sub cmdTrans_Click(Index As Integer)
'Index Legend
' 0 -> add
' 1-> edit
' 2 -> view
' 3 -> exit
' 4 - > save
' 5 - > cancel



    If Index = 0 Then 'add
        formAction = "add"
        AddNew = True
        Call LockData(True)
        Call SetButtonVisibility(True)
        Call clearform
        flxPARTS.TextMatrix(flxPARTS.Row, 8) = 0
        txtUP.Text = EmployeesName
        
    ElseIf Index = 1 Then 'edit

        If txtPMID.Text = "" Then
    
            MsgBox "No available Parts Matrix to edit", vbOKOnly + vbExclamation, "Edit"
    
        Else
            If mainRs.State = adStateOpen Then mainRs.Close
            Timer1.Enabled = True
            
            If cpRs.State = adStateOpen Then cpRs.Close
            Timer2.Enabled = True
            formAction = "edit"
            Call LockData(True)
            Call SetButtonVisibility(True)
            AddNew = False
            txtUP.Text = EmployeesName
        End If
        
    ElseIf Index = 2 Then 'view summary
        
        frmSummaryRepl.mySrc = "Data Entry"
        frmSummaryRepl.Show
        flxPARTS.Rows = 1
        flxPARTS.Rows = 2
        Unload Me
        
    ElseIf Index = 3 Then 'print preview
        formAction = "PRINT"
        
        If flxPARTS.TextMatrix(1, 1) <> "" Then
            Qry = "call sp_partsmatrix.sp_dataentry_getDetails('" & sysBranchCode & "', " & txtPMID & ")"
            frmPrn.QryPrn = Qry
         
            
            frmPrn.ParamValue1 = SysBranch
            frmPrn.ParamValue2 = Format(SysAddress, "")
            frmPrn.ParamValue3 = Format(SysTel, "")
            frmPrn.ParamValue4 = Format(txtBrand, "")
            frmPrn.ParamValue5 = Format(txtModel, "")
            frmPrn.ParamValue34 = Format(txtDCreated, "mm-dd-yyyy")
            frmPrn.ParamValue36 = Format(txtTOTALPN, "")
            frmPrn.ParamValue6 = Format(txtUP, "")
            frmPrn.MySource = "Data Entry"
            frmPrn.Show
        End If
    ElseIf Index = 4 Then 'exit
        formAction = "EXIT"
        ProperExit = True
        Call clearform
        Unload Me
        blnOpt = False
        
    ElseIf Index = 5 Then 'save
  
        If ValidateEntry() = False Then Exit Sub
        
            myMsg = ValidationPartsMatrix
            
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
        
        Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        
        If Ans = vbNo Then Exit Sub
        
      
        
        If formAction = "add" Then
        
            Call CreateTempTable
            
            DoEvents
            Call modPMDataEntry.sp_dataentry_create(txtBrand.Text, txtModel.Text, txtDescription.Text, Format(txtDCreated.Text, "YYYY-mm-dd"), UserID, UserID, mainRs)

        ElseIf formAction = "edit" Then
            Call CreateTempTable
            
            DoEvents
            Call modPMDataEntry.sp_dataentry_update(txtBrand.Text, txtModel.Text, txtDescription.Text, Format(txtDCreated.Text, "YYYY-mm-dd"), UserID, txtPMID.Text, mainRs)
        
        End If
              
        Ans = MsgBox("Successfully Saved!" & vbCrLf & vbCrLf & "Process Another Entry?", vbQuestion + vbYesNo, "Save")
        
        If Ans = vbNo Then
            Call SetButtonVisibility(False)
            Call LockData(False)
            'Call clearform
            formAction = ""
            Timer1.Enabled = True
            flxPARTS.TextMatrix(flxPARTS.Row, 8) = 0
            
            DoEvents
            Call modPMDataEntry.sp_get_PartsmatrixID(mainRs)
            txtPMID.Text = mainRs!PMATRIXID
            
        Else
            Call SetButtonVisibility(True)
            Call clearform
            formAction = "add"
            AddNew = True
            Timer1.Enabled = True
            flxPARTS.TextMatrix(flxPARTS.Row, 8) = 0
        End If
        
        For i = 1 To flxPARTS.Rows - 1
          If flxPARTS.TextMatrix(i, 1) <> "" Then
            txtTOTALPN.Text = i
            End If
        Next i
        
    ElseIf Index = 6 Then 'cancel
            Ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
            If Ans = vbNo Then Exit Sub
            
            
            Call SetButtonVisibility(False)
            Call LockData(False)
            'Call clearform
            flxPARTS.TextMatrix(flxPARTS.Row, 8) = 0
            flxPARTNO.Visible = False
            cmbPG.Visible = False
            txtPARTNO.Text = ""
            txtPARTNO.Visible = False
           ' txtBrand.SetFocus
            formAction = ""
            
    End If

End Sub


Private Sub flxCounterPart_LostFocus()
Call CounterPartValidation
End Sub

Private Sub flxPARTNO_LostFocus()
Call txtPartNo_LostFocus
End Sub

Private Sub flxParts_Click()
On Error GoTo er
    ActRow = flxPARTS.Row
    ActCol = flxPARTS.Col


    If formAction = "edit" Then 'clickable all
    
        If ActCol = 0 Or ActRow = 0 Then Exit Sub
        Select Case ActCol
        Case 1
            With txtPARTNO
                If flxPARTS.TextMatrix(ActRow, 1) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 1) & "|" & flxPARTS.TextMatrix(ActRow, 2) & "|" & flxPARTS.TextMatrix(ActRow, 3) & "|" & flxPARTS.TextMatrix(ActRow, 4) & "|" & flxPARTS.TextMatrix(ActRow, 5) & "|" & flxPARTS.TextMatrix(ActRow, 6) & "|" & flxPARTS.TextMatrix(ActRow, 7) & "|" & flxPARTS.TextMatrix(ActRow, 8) & "|" & flxPARTS.TextMatrix(ActRow, 9)
                .Top = flxPARTS.CellTop + flxPARTS.Top
                .Left = flxPARTS.CellLeft + flxPARTS.Left
                .Height = flxPARTS.CellHeight
                .Width = flxPARTS.CellWidth
                .Text = flxPARTS.Text
                .Visible = True
                .Enabled = True
                txtSPECMODEL.Enabled = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
            
        Case 4
        'If flxPARTS.TextMatrix(actrow, 12) = "" Then
            With txtFlxBrand
                If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 4) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 4)
                    .Top = flxPARTS.CellTop + flxPARTS.Top
                    .Left = flxPARTS.CellLeft + flxPARTS.Left
                    .Height = flxPARTS.CellHeight
                    .Width = flxPARTS.CellWidth
                    .Text = flxPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End If
            End With
        'End If
        Case 5
            With txtSPECMODEL
                If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 5) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 5)
                    .Top = flxPARTS.CellTop + flxPARTS.Top
                    .Left = flxPARTS.CellLeft + flxPARTS.Left
                    .Height = flxPARTS.CellHeight
                    .Width = flxPARTS.CellWidth
                    .Text = flxPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End If
            End With
        Case 6
        'If flxPARTS.TextMatrix(actrow, 12) = "" Then
            With txtSize
                If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 6) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 6)
                    .Top = flxPARTS.CellTop + flxPARTS.Top
                    .Left = flxPARTS.CellLeft + flxPARTS.Left
                    .Height = flxPARTS.CellHeight
                    .Width = flxPARTS.CellWidth
                    .Text = flxPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End If
            End With
        'End If
        
        Case 7
            With cmbPG
                If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 7) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 7)
                    .Top = flxPARTS.CellTop + flxPARTS.Top
                    .Left = flxPARTS.CellLeft + flxPARTS.Left
                    .Width = flxPARTS.CellWidth
                    .Text = flxPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End If
            End With
    
        Case 8
            With txtRP
                If flxPARTS.TextMatrix(ActRow, 8) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 8) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 8)
                    .Top = flxPARTS.CellTop + flxPARTS.Top
                    .Left = flxPARTS.CellLeft + flxPARTS.Left
                    .Height = flxPARTS.CellHeight
                    .Width = flxPARTS.CellWidth
                    .Text = flxPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                ElseIf flxPARTS.TextMatrix(ActRow, 1) = "" Then
                    MsgBox "No Part Number Selected", vbExclamation + vbOKOnly, "Warning"
                    cmbPG.Visible = False
                    flxPARTS.SetFocus
                    flxPARTS.Col = 1
                End If
            End With
            
            
    Case 9
        With txtCountrP
            If Val(flxPARTS.TextMatrix(ActRow, 8)) = 1 Then
            If flxPARTS.TextMatrix(ActRow, 9) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 9)
                .Top = flxPARTS.CellTop + flxPARTS.Top
                .Left = flxPARTS.CellLeft + flxPARTS.Left
                .Height = flxPARTS.CellHeight
                .Width = flxPARTS.CellWidth
                .Text = flxPARTS.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            Else
                MsgBox "No CounterPart available", vbExclamation + vbOKOnly, "Warning"
                txtCountrP.Visible = False
                flxPARTS.SetFocus
                flxPARTS.Col = 1
            End If
        End With
    End Select
    
    
    Exit Sub
    End If
    
    
    
    If formAction = "add" Then 'Clickable Optional
    
    If txtBrand.Text <> "" And txtModel.Text <> "" Then
    
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
    Select Case ActCol
    Case 1
        If flxPARTS.TextMatrix(ActRow, 12) = "" Then
            With txtPARTNO
                If flxPARTS.TextMatrix(ActRow, 1) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 1) & "|" & flxPARTS.TextMatrix(ActRow, 2) & "|" & flxPARTS.TextMatrix(ActRow, 3) & "|" & flxPARTS.TextMatrix(ActRow, 4) & "|" & flxPARTS.TextMatrix(ActRow, 5) & "|" & flxPARTS.TextMatrix(ActRow, 6) & "|" & flxPARTS.TextMatrix(ActRow, 7) & "|" & flxPARTS.TextMatrix(ActRow, 8) & "|" & flxPARTS.TextMatrix(ActRow, 9)
                .Top = flxPARTS.CellTop + flxPARTS.Top
                .Left = flxPARTS.CellLeft + flxPARTS.Left
                .Height = flxPARTS.CellHeight
                .Width = flxPARTS.CellWidth
                .Text = flxPARTS.Text
                .Visible = True
                .Enabled = True
                txtSPECMODEL.Enabled = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        End If
        
    Case 4
    
        'If flxPARTS.TextMatrix(actrow, 12) = "" Then
            With txtFlxBrand
                If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 4) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 4)
                    .Top = flxPARTS.CellTop + flxPARTS.Top
                    .Left = flxPARTS.CellLeft + flxPARTS.Left
                    .Height = flxPARTS.CellHeight
                    .Width = flxPARTS.CellWidth
                    .Text = flxPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End If
            End With
        'End If
        
    Case 5
            With txtSPECMODEL
                If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 5) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 5)
                    .Top = flxPARTS.CellTop + flxPARTS.Top
                    .Left = flxPARTS.CellLeft + flxPARTS.Left
                    .Height = flxPARTS.CellHeight
                    .Width = flxPARTS.CellWidth
                    .Text = flxPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End If
            End With
            
    Case 6
        'If flxPARTS.TextMatrix(actrow, 12) = "" Then
            With txtSize
                If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 6) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 6)
                    .Top = flxPARTS.CellTop + flxPARTS.Top
                    .Left = flxPARTS.CellLeft + flxPARTS.Left
                    .Height = flxPARTS.CellHeight
                    .Width = flxPARTS.CellWidth
                    .Text = flxPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End If
            End With
        'End If
        
    Case 7
       ' If flxParts.TextMatrix(ActRow, 12) = "" Then
            With cmbPG
                If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 7) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 7)
                    .Top = flxPARTS.CellTop + flxPARTS.Top
                    .Left = flxPARTS.CellLeft + flxPARTS.Left
                    .Width = flxPARTS.CellWidth
                    .Text = flxPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End If
            End With
        'End If
        
    Case 8
        With txtRP
            If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
            If flxPARTS.TextMatrix(ActRow, 8) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 8)
                .Top = flxPARTS.CellTop + flxPARTS.Top
                .Left = flxPARTS.CellLeft + flxPARTS.Left
                .Height = flxPARTS.CellHeight
                .Width = flxPARTS.CellWidth
                .Text = flxPARTS.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            ElseIf flxPARTS.TextMatrix(ActRow, 1) = "" Then
                MsgBox "No Part Number Selected", vbExclamation + vbOKOnly, "Warning"
                cmbPG.Visible = False
                flxPARTS.SetFocus
                flxPARTS.Col = 1
            End If
        End With
    Case 9
    
        With txtCountrP
            If Val(flxPARTS.TextMatrix(ActRow, 8)) = 1 Then
            If flxPARTS.TextMatrix(ActRow, 9) <> "" Then .Tag = flxPARTS.TextMatrix(ActRow, 9)
                .Top = flxPARTS.CellTop + flxPARTS.Top
                .Left = flxPARTS.CellLeft + flxPARTS.Left
                .Height = flxPARTS.CellHeight
                .Width = flxPARTS.CellWidth
                .Text = flxPARTS.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            Else
                MsgBox "No CounterPart available", vbExclamation + vbOKOnly, "Warning"
                txtCountrP.Visible = False
                flxPARTS.SetFocus
                flxPARTS.Col = 1
            End If
        End With
    End Select
    
    Else
        If txtBrand.Text = "" Then
            MsgBox "Please Enter a Brand first!", vbInformation + vbOKOnly, "Parts Matrix"
            txtBrand.SetFocus
        Else
            MsgBox "Please Enter a Model first!", vbInformation + vbOKOnly, "Parts Matrix"
            txtModel.SetFocus
        End If
    End If
    
    Exit Sub
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxPARTSCLICK()")
End Sub





Private Sub flxPARTS_Scroll()
txtCountrP.Visible = False
txtPARTNO.Visible = False
txtFlxBrand.Visible = False
txtSPECMODEL.Visible = False
txtSize.Visible = False
txtRP.Visible = False
End Sub



Private Sub Form_Resize()
    Call RXResizer1_ResizeComplete
End Sub



Public Sub flxChnger()

    With flxPARTS
        '.Clear
        .Cols = 17
        .FormatString = "|^PARTS NUMBER|<SP DESCRIPTION |<SP CLASSIFICATION |<SP BRAND |<MC MODEL |<SP SIZE|<PARTS GROUP |<RP |^COUNTERPART"
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15  'PARTS NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.28 'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.11 'CLASSIFICATION
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.07 'BRAND
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.07 'SP MODEL
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.06 'SIZE
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.08 'PARTS GROUP
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.03 'REPLACEMENT(RP)
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.13 'COUNTERPART
        .ColWidth(10) = 0 '(.Width - .ColWidth(0)) * 0.05  ' SPID
        .ColWidth(11) = 0 '(.Width - .ColWidth(0)) * 0.05  ' CPID
        .ColWidth(12) = 0 '(.Width - .ColWidth(0)) * 0.05 ' GENID
        .ColWidth(13) = 0 ' DESCRIPTION
        .ColWidth(14) = 0 ' CLASSIFICATION
         
        .ColWidth(15) = 0 ' DESCRIPTION
        .ColWidth(16) = 0 ' CLASSIFICATION
         
        flxPARTNO.ColWidth(1) = .ColWidth(1)
        flxCounterPart.ColWidth(7) = .ColWidth(7)
        
        
    End With
    
    With flxPARTNO
        .Top = txtPARTNO.Height + txtPARTNO.Top
        .Left = txtPARTNO.Left
        .Width = txtPARTNO.Width
    End With
    
    With flxCounterPart
        .Top = txtCountrP.Height + txtCountrP.Top
        .Left = txtCountrP.Left
        .Width = txtCountrP.Width
    End With
    
   
    With flxPMSum
        .Cols = 7
        .FormatString = "|^CLASSIFICATION |^TOTAL PN |^COMPLETE PARTS |^MAINTENANCE PARTS |^PLASTIC PARTS |^UNIVERSAL PARTS"
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.2  'CLASSIFICATION
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.15 'TOTAL PN
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.17 'COMPLETE PARTS
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.2 'MAINTENANCE PARTS
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.16 'PLASTIC PARTS
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.17 'UNIVERSAL PARTS
    End With
    

End Sub






Sub InsertRow(intRow As Integer)

    flxPARTS.Rows = flxPARTS.Rows + 1

        For i = flxPARTS.Rows - 1 To intRow Step -1
            If i = intRow Then
                For j = 0 To flxPARTS.Cols - 1
                flxPARTS.TextMatrix(i + 1, j) = ""

                Next j
            Else
                For j = 0 To flxPARTS.Cols - 1
                flxPARTS.TextMatrix(i, j) = flxPARTS.TextMatrix(i - 1, j)
                Next j
            End If
        Next i

End Sub











'================FORM RESIZER============================
Private Sub RXResizer1_ResizeComplete()


    Call flxChnger

    'Frame width
    Frame1.Width = Me.Width - 465
    Frame2.Width = Me.Width - 465
    Frame3.Width = Me.Width - 465
    
    'Frame1 label width
    lblBrand.Width = 1455
    lblModel.Width = 1215
    lblDCreated.Width = 1335
    lblTPN.Width = 1335
    lblDescription.Width = 1455
    
    'frame1 label left
    lblBrand.Left = 240
    lblModel.Left = 4080
    lblDCreated.Left = Frame1.Width - 3495
    lblTPN.Left = Frame1.Width - 3495
    lblDescription.Left = 240

    'frame1 text width
    txtBrand.Width = 2295
    txtModel.Width = 2535
    txtDCreated.Width = 1935
    txtTOTALPN.Width = 1935
    txtDescription.Width = 6135

    'frame1 text left
    txtBrand.Left = 1800
    txtModel.Left = 5400
    txtDCreated.Left = Frame1.Width - 2055
    txtTOTALPN.Left = Frame1.Width - 2055
    txtDescription.Left = 1800

    'Frame2 label width
    lblUpdateBy.Width = 1215
    
    'frame2 label left
    lblUpdateBy.Left = 240

    'frame2 text left
    txtUP.Left = 1560

    'frame2 text width
    txtUP.Width = 3495


    'Frame3 text left
    txtFind.Left = Me.Width - 2385
    
    'frame3 button left
    btnFind.Left = Me.Width - 3825
    cmdShow.Left = 480
    cmdHide.Left = 480
    
    'frame3 button width
    btnFind.Width = 1335
    cmdShow.Width = 1455
    cmdHide.Width = 1455
    
    'frame3 text with
    txtFind.Width = 1935
    
   ' flxPARTS.Height = flxPARTS.Height + 100
    If blnOpt = True Then
        flxPARTS.Top = Frame3.Top + Frame3.Height
        flxPARTS.Height = Me.Height - Frame3.Height - 3918 + cmdShow.Height
        btnFind.Top = cmdShow.Top + cmdShow.Height
        txtFind.Top = cmdShow.Top + cmdShow.Height
    ElseIf blnOpt = False Then
        flxPARTS.Top = cmdShow.Top + cmdShow.Height
        flxPARTS.Height = Me.Height - 3918
        btnFind.Top = cmdShow.Top
        txtFind.Top = cmdShow.Top
    End If
    
    flxPMSum.Width = Frame3.Width - 9720
End Sub



Private Sub Form_Load()
    Call LockData(False)
    TCol = 12
    ActRow = 0
    ActCol = 0
    optRow = 1
    'Call flxChnger
    Frame3.Visible = False
    
   ' Call modCon.Connection
    txtDCreated.Text = Format(Date, "mm/dd/yyyy") 'date
    
    flxPARTS.TextMatrix(flxPARTS.Row, 8) = 0
    
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





Sub clearform()

    flxPARTS.Rows = 1
    flxPARTS.Rows = 2
   
    txtBrand.Text = ""
    txtModel.Text = ""
    txtTOTALPN.Text = ""
    txtFind.Text = ""
    txtDCreated.Text = Format(Date, "mm/dd/yyyy") 'date
   
    txtDescription.Text = ""
    txtPMID.Text = ""
   
End Sub





Private Sub Timer1_Timer()
On Error GoTo er
'    Qry = "Select spareparts.PARTNUMBER,spareparts.DESCRIPTION, spareparts.REMARKS, spareparts.SPAREPARTID From mrc2retail.spareparts " & _
'          "WHERE spareparts.DESCRIPTION NOT LIKE '%INACTIVE%' ORDER BY spareparts.PARTNUMBER"
'    Set mainRs = dbCon.Execute(Qry)
  
    DoEvents
    Call modPMDataEntry.sp_get_Partnumber(mainRs)
    
    Set flxPARTNO.DataSource = mainRs
    flxPARTNO.Refresh
    With flxPARTNO
        .Width = flxPARTS.ColWidth(1)
        .ColWidth(0) = 0
        .ColWidth(1) = flxPARTS.ColWidth(1)
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
    End With
    
    Timer1.Enabled = False
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadPartNumber()")
End Sub


Private Sub Timer2_Timer()
On Error GoTo er
'    Qry = "Select spareparts.PARTNUMBER,spareparts.DESCRIPTION, spareparts.REMARKS, spareparts.SPAREPARTID From mrc2retail.spareparts " & _
'            "WHERE spareparts.DESCRIPTION NOT LIKE '%INACTIVE%' ORDER BY spareparts.PARTNUMBER;"
'
'     Set mainRs = dbCon.Execute(Qry)

    DoEvents
    Call modPMDataEntry.sp_get_Partnumber(cpRs)

     Set flxCounterPart.DataSource = cpRs
     flxCounterPart.Refresh
    With flxCounterPart
        .Width = flxPARTS.ColWidth(9)
        .ColWidth(0) = 0
        .ColWidth(1) = flxPARTS.ColWidth(9)
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
    End With

    Timer2.Enabled = False
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadCounterPartNumber()")
End Sub



Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxParts_Click
    End If
End Sub



Private Sub txtBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtBrand, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtModel.SetFocus
End Sub







Private Sub txtDCreated_GotFocus()

    txtDCreated.SelStart = 0
    txtDCreated.SelLength = Len(txtDCreated.Text)
End Sub



Private Sub txtDCreated_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtDCreated, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        flxPARTS.SetFocus
        flxPARTS.Col = 1
    End If
End Sub



Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtBrand, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtDCreated.SetFocus
End Sub



Private Sub txtFlxBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtBrand, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
            CCol = 2
            flxPARTS.TextMatrix(ActRow, 4) = txtFlxBrand.Text
            Call ColCheckForEmpty
        End If
    End If
End Sub

Private Sub txtFlxBrand_LostFocus()
    If txtFlxBrand.Tag = "skiplostfocus" Then
        txtFlxBrand.Tag = ""
        Exit Sub
    End If
    
    If txtFlxBrand.Text = "" Then
        txtFlxBrand.Visible = False
    End If
    
       ' CCol = 2
        flxPARTS.TextMatrix(ActRow, 4) = txtFlxBrand.Text
        txtFlxBrand.Visible = False
        'Call ColCheckForEmpty
End Sub




Private Sub txtSize_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtBrand, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
            CCol = 4
            flxPARTS.TextMatrix(ActRow, 6) = txtSize.Text
            txtSize.Visible = False
            Call ColCheckForEmpty
        End If
    End If
End Sub



Private Sub txtSize_LostFocus()
    If txtSize.Tag = "skiplostfocus" Then
        txtSize.Tag = ""
        Exit Sub
    End If
    If txtSize.Text = "" Then
        txtSize.Visible = False
    End If
        'CCol = 4
        flxPARTS.TextMatrix(ActRow, 6) = txtSize.Text
        txtSize.Visible = False
       ' Call ColCheckForEmpty
End Sub

Private Sub txtSPECMODEL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then
    End If
End Sub

Private Sub txtSPECMODEL_LostFocus()
    If txtSPECMODEL.Tag = "skiplostfocus" Then
        txtSPECMODEL.Tag = ""
        Exit Sub
    End If
    
    If txtSPECMODEL.Text = "" Then
        txtSPECMODEL.Visible = False
    End If
    'CCol = 3
    flxPARTS.TextMatrix(ActRow, 5) = txtSPECMODEL.Text
    txtSPECMODEL.Visible = False
    'Call ColCheckForEmpty
    
End Sub

Private Sub txtTOTALPN_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtTOTALPN, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then flxPARTS.SetFocus
End Sub



Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFind, KeyAscii, ValidChars)
    If KeyAscii = 13 Then 'txtUP.SetFocus
        Call btnFind_Click
    End If
End Sub


Private Sub txtUP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtUP, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        txtBrand.SetFocus
    End If
End Sub



Private Sub txtModel_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
        Call ObjKeyTrapping.Validation(txtModel, KeyAscii, ValidChars)
        If KeyAscii = 13 Then
            txtDescription.SetFocus
            
        End If
End Sub



Private Sub txtPartNo_Change()

On Error GoTo LocalErr
    With flxPARTNO
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPARTNO.Height + txtPARTNO.Top
        .Left = txtPARTNO.Left
         .Width = txtPARTNO.Width
        .Visible = True
        .ZOrder
            
        If txtPARTNO.Text = "" Then .Visible = False: Exit Sub
        mainRs.Find "[PARTNUMBER] LIKE '" + txtPARTNO.Text + "%'", , adSearchForward, 1
        If Not mainRs.EOF Then
            If Not PrtBln Then
                txtPARTNO.Tag = Len(txtPARTNO.Text)
                PrtBln = True
            End If
            
            .TopRow = mainRs.AbsolutePosition - 1
            .Row = mainRs.AbsolutePosition - 1
            .Col = 0: .ColSel = 1
            txtPARTNO.Text = mainRs!PartNumber
            flxPARTS.TextMatrix(flxPARTS.Row, 2) = mainRs!Description
            flxPARTS.TextMatrix(flxPARTS.Row, 3) = mainRs!Remarks
            flxPARTS.TextMatrix(flxPARTS.Row, 10) = mainRs!SPAREPARTID
           ' flxPARTS.TextMatrix(flxPARTS.Row, 8) = 0
            'flxPARTS.TextMatrix(flxPARTS.Row, 7) = ""
            'flxPARTS.TextMatrix(flxPARTS.Row, 9) = ""
            'FLXPARTS.TEXTMATRIX(FLXPARTS.ROW,4) = ""
            txtPARTNO.SelStart = Val(txtPARTNO.Tag)
            txtPARTNO.SelLength = Len(txtPARTNO.Text)
            PrtBln = False
                 
                 
        Else
    
            flxPARTS.TextMatrix(ActRow, 1) = ""
            flxPARTS.TextMatrix(ActRow, 2) = ""
            flxPARTS.TextMatrix(ActRow, 3) = ""
            flxPARTS.TextMatrix(ActRow, 4) = ""
            flxPARTS.TextMatrix(ActRow, 5) = ""
            flxPARTS.TextMatrix(ActRow, 6) = ""
            flxPARTS.TextMatrix(ActRow, 7) = ""
            flxPARTS.TextMatrix(ActRow, 8) = 0
            flxPARTS.TextMatrix(ActRow, 9) = ""
            flxPARTS.TextMatrix(ActRow, 10) = ""
            .Visible = False
        End If
   End With
   
   Exit Sub

LocalErr:
    If Err.Number = 30009 Then
        flxPARTNO.Visible = False
        txtPARTNO.Visible = False
        If mainRs.State = adStateOpen Then mainRs.Close
        Timer1.Enabled = True
        MsgBox "Program is reloading Part Number" & vbCrLf & "Please wait for a few seconds...", vbInformation + vbOKOnly, "Reloading"
    Else
        flxPARTNO.Visible = False
        txtPARTNO.Visible = False
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
    
End Sub




Private Sub txtPartNo_GotFocus()
    If mainRs.State = adStateClosed Then Timer1.Enabled = True
    With flxPARTNO
        .Top = txtPARTNO.Height + txtPARTNO.Top
        .Left = txtPARTNO.Left
        .Width = txtPARTNO.Width
    End With
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ProperExit Then
        If mainRs.State = adStateOpen Then mainRs.Close
        If cpRs.State = adStateOpen Then cpRs.Close
        Call cmdTrans_Click(4)
        ProperExit = False
    Else
        Cancel = True
    End If
End Sub

Private Sub txtPartNo_LostFocus()
On Error GoTo er
'    Call CheckDuplicate
    If txtPARTNO.Tag = "skiplostfocus" Then
        txtPARTNO.Tag = ""
        Exit Sub
    End If
    
'    If chckDupCount < 1 Then
    If txtPARTNO.Text <> "" Then
        If flxPARTS.TextMatrix(ActRow, 10) = "" Or txtPARTNO.Text = "" Then
            MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
            txtPARTNO.Visible = False
            flxPARTNO.Visible = False
            flxPARTS.TextMatrix(ActRow, 1) = ""
            flxPARTS.TextMatrix(ActRow, 2) = ""
            flxPARTS.TextMatrix(ActRow, 3) = ""
            flxPARTS.TextMatrix(ActRow, 4) = ""
            flxPARTS.TextMatrix(ActRow, 5) = ""
            flxPARTS.TextMatrix(ActRow, 6) = ""
            flxPARTS.TextMatrix(ActRow, 7) = ""
            flxPARTS.TextMatrix(ActRow, 8) = 0
            flxPARTS.TextMatrix(ActRow, 9) = ""
            flxPARTS.TextMatrix(ActRow, 10) = ""
            flxPARTS.SetFocus
            flxPARTS.Col = 1
        Exit Sub
        End If
    End If
        
        'CCol = 1
        flxPARTS.TextMatrix(ActRow, 1) = txtPARTNO.Text
        txtPARTNO.Visible = False
        flxPARTNO.Visible = False
        'Call ColCheckForEmpty
        
'    ElseIf chckDupCount > 0 Then
'        flxPARTS.SetFocus
'        flxPARTS.Col = 1
'        chckDupCount = 0
'        Exit Sub
'    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->PartNumber-LostFocus()")
End Sub

Sub RowCheckForEmpty()

    For i = 1 To flxPARTS.Rows - 1
        If flxPARTS.TextMatrix(i, 1) = "" Then
            rowEmpty = True
        End If
    Next i
End Sub

Sub ColCheckForEmpty()
On Error GoTo er
        If CCol = 1 Then 'FROM COLUMN 1 JUMP TO COLUMN 4
            If flxPARTS.TextMatrix(ActRow, 2) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 4) <> "" Then
                    flxPARTS.SetFocus
                    flxPARTS.Col = 5 'SP MODEL
                Else
                    flxPARTS.SetFocus
                    flxPARTS.Col = 4 'SP BRAND
                   
                End If
                Exit Sub
            End If
            
            If flxPARTS.TextMatrix(ActRow, 1) = "" Then
                flxPARTS.SetFocus
                flxPARTS.Col = 1 'PARTNUMBER
            End If
        ElseIf CCol = 2 Then 'FROM COLUMN 4 JUMP TO COLUMN 5
            If flxPARTS.TextMatrix(ActRow, 4) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 5) <> "" Then
                    flxPARTS.SetFocus
                    flxPARTS.Col = 6 'SP SIZE
                Else
                    flxPARTS.SetFocus
                    flxPARTS.Col = 5 ' SP MODEL
                End If
            Else
                flxPARTS.SetFocus
                flxPARTS.Col = 5
            End If
        ElseIf CCol = 3 Then 'FROM COLUMN 5 JUMP TO COLUMN 6
            If flxPARTS.TextMatrix(ActRow, 5) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 6) <> "" Then
                    flxPARTS.SetFocus
                    flxPARTS.Col = 8 'PARTS GROUP
                Else
                    flxPARTS.SetFocus
                    flxPARTS.Col = 6 'SP SIZE
                End If
            Else
                flxPARTS.SetFocus
                flxPARTS.Col = 6 'SP SIZE
            End If
        ElseIf CCol = 4 Then 'FROM COLUMN 6 JUMP TO COLUMN 8
            If flxPARTS.TextMatrix(ActRow, 6) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 7) <> "" Then
                    flxPARTS.SetFocus
                    flxPARTS.Col = 9 'COUNTERPART
                Else
                    flxPARTS.SetFocus
                    flxPARTS.Col = 8 'REPLACEMENT (RP)
                End If
            Else
                flxPARTS.SetFocus
                flxPARTS.Col = 8 'REPLACEMENT (RP)
            End If
        ElseIf CCol = 5 Then 'FROM COLUMN 7 JUMP TO COLUMN 9
            If flxPARTS.TextMatrix(ActRow, 7) <> "" Then
                If flxPARTS.TextMatrix(ActRow, 8) <> "" Then
                    flxPARTS.SetFocus
                    flxPARTS.Col = 9 'COUNTERPART
                Else
                    flxPARTS.SetFocus
                    flxPARTS.Col = 8 'REPLACEMENT (RP)
                End If
            Else
                flxPARTS.SetFocus
                flxPARTS.Col = 8 'REPLACEMENT (RP)
            End If
        ElseIf flxPARTS.TextMatrix(ActRow, 9) = "" Then
            flxPARTS.SetFocus
            flxPARTS.Col = 9 'COUNTERPART
        End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->ColCheckForEmpty()")
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
On Error GoTo er
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtPARTNO, KeyAscii, ValidChars)
    
    If mainRs.State = adStateClosed Then Timer1.Enabled = True
     
    
    If KeyAscii = 13 Then
       ' flxParts.TextMatrix(flxParts.Row, 5) = txtModel.Text
        Call CheckDuplicate
        
        If chckDupCount < 1 Then
            CCol = 1
            
            Call ColCheckForEmpty
        ElseIf chckDupCount > 0 Then
            flxPARTS.SetFocus
            flxPARTS.Col = 1
            chckDupCount = 0
            Exit Sub
        End If
        
       
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtPARTNO_KeyPress()")
End Sub

Public Sub distinctClass()
'diria kay tangalun ang mga repeated classification
Dim lCounter As Long
Dim lCounter2 As Long
Dim lTemp As Long
Dim lAccumulator As Long

lTemp = flxClass.Rows - 1
lAccumulator = 0

For lCounter = 1 To lTemp
    For lCounter2 = lCounter + 1 To lTemp
        ' now compare every item to current item from main loop
        If flxClass.TextMatrix(lCounter, 1) = flxClass.TextMatrix(lCounter2 - lAccumulator, 1) Then
            flxClass.RemoveItem (lCounter2 - lAccumulator)
            lTemp = lTemp - 1
            lAccumulator = lAccumulator + 1
        End If
    Next lCounter2
   lAccumulator = 0
Next lCounter
End Sub

Public Sub countingPartsGroup()
    
    intCP = 1
    intMP = 1
    intPP = 1
    intUP = 1
    For i = 1 To flxPARTS.Rows - 1
    For j = 1 To flxPMSum.Rows - 1
    If flxPARTS.TextMatrix(i, 3) = flxClass.TextMatrix(j, 1) Then
            If flxPARTS.TextMatrix(i, 7) = "COMPLETE" Then
                intCP = intCP + Val(flxPMSum.TextMatrix(j, 3))
                flxPMSum.TextMatrix(j, 3) = intCP
                flxPMSum.TextMatrix(j, 2) = flxClass.TextMatrix(j, 2)
                
            ElseIf flxPARTS.TextMatrix(i, 7) = "MAINTENANCE" Then
                intMP = intMP + Val(flxPMSum.TextMatrix(j, 4))
                flxPMSum.TextMatrix(j, 4) = intMP
                flxPMSum.TextMatrix(j, 2) = flxClass.TextMatrix(j, 2)
                
            ElseIf flxPARTS.TextMatrix(i, 7) = "PLASTIC" Then
                intPP = intPP + Val(flxPMSum.TextMatrix(j, 5))
                flxPMSum.TextMatrix(j, 5) = intPP
                flxPMSum.TextMatrix(j, 2) = flxClass.TextMatrix(j, 2)
                
            ElseIf flxPARTS.TextMatrix(i, 7) = "UNIVERSAL" Then
                intUP = intUP + Val(flxPMSum.TextMatrix(j, 6))
                flxPMSum.TextMatrix(j, 6) = intUP
                flxPMSum.TextMatrix(j, 2) = flxClass.TextMatrix(j, 2)
                
            End If
            
    End If
    
    intCP = 1
    intMP = 1
    intPP = 1
    intUP = 1
    Next j
   
    Next i
    
    
              
End Sub

Public Sub passToFlxPMSum()
    
    'pass all the data from flxclass to flxpmsum
    flxPMSum.Rows = flxClass.Rows
    For i = 1 To flxClass.Rows - 1
        flxPMSum.TextMatrix(i, 1) = flxClass.TextMatrix(i, 1)
    Next i
         
End Sub


Public Sub CheckDuplicate()
On Error GoTo er
'Verify the Part Number if already exist
        If txtPARTNO.Text = flxPARTS.TextMatrix(flxPARTS.Row, 1) Then
            Exit Sub
        End If
    
    With flxPARTS
    
        For intLoopCount = 1 To .Rows - 1
        txtstr = flxPARTS.TextMatrix(intLoopCount, 1)
        
            If flxPARTS.TextMatrix(intLoopCount, 12) <> "" Then
                GoTo SKIP
            End If
        If txtstr = txtPARTNO.Text Then 'flxPARTS.TextMatrix(actrow, 1) Then
                MsgBox "Part Number is already exist in row " & " " & intLoopCount & "" & Err.Description, vbExclamation + vbOKOnly, "Warning"
                .TextMatrix(.Row, 1) = ""
                .TextMatrix(.Row, 2) = ""
                .TextMatrix(.Row, 3) = ""
                .TextMatrix(.Row, 4) = ""
                .TextMatrix(.Row, 5) = ""
                .TextMatrix(.Row, 6) = ""
                .TextMatrix(.Row, 7) = ""
                '.TextMatrix(.Row, 8) = ""
                .TextMatrix(.Row, 9) = ""
                .TextMatrix(.Row, 10) = ""
                txtPARTNO.Text = ""
                txtPARTNO.Visible = False
                flxPARTNO.Visible = False
                chckDupCount = chckDupCount + 1
            Exit Sub
        End If
SKIP:

        Next intLoopCount
    
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CheckDuplicate()")
End Sub


Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
            txtPARTNO.Tag = "skiplostfocus"
            flxPARTNO.Visible = True
            flxPARTNO.SetFocus
            
    ElseIf KeyCode = vbKeyEscape Then
          '  txtPARTNO.Text = Empty
           ' Tmp = Split(txtPARTNO.Tag, "|")
            
'            For i = LBound(Tmp) To UBound(Tmp)
'                flxPARTNO.TextMatrix(ActRow, i + 1) = Tmp(i)
'            Next i
'
            'txtPARTNO.Tag = "skiplostfocus"
            flxPARTS.SetFocus
            flxPARTS.Col = 1
            'txtPARTNO.Text = ""
            txtPARTNO.Visible = False
            flxPARTNO.Visible = False
'            flxPARTS.TextMatrix(flxPARTS.Row, 2) = ""
'            flxPARTS.TextMatrix(flxPARTS.Row, 3) = ""
'            flxPARTS.TextMatrix(flxPARTS.Row, 6) = ""
'            'flxPARTS.TextMatrix(flxPARTS.Row, 8) = ""
            
    ElseIf KeyCode = vbKeyDelete Then
            PrtSrch = True
            mainRs.MoveFirst
            flxPARTS.TextMatrix(flxPARTS.Row, 8) = ""
    ElseIf KeyCode = vbKeyBack Then
        
        PrtSrch = True
        mainRs.MoveFirst
        flxPARTS.TextMatrix(flxPARTS.Row, 2) = ""
        flxPARTS.TextMatrix(flxPARTS.Row, 3) = ""
        flxPARTS.TextMatrix(flxPARTS.Row, 6) = ""
        'flxPARTS.TextMatrix(flxPARTS.Row, 8) = ""
        flxPARTNO.Visible = False
        
    End If
End Sub
    




Private Sub flxPARTNO_Click()
On Error GoTo er
'    With flxPARTS
'
'        For intLoopCount = 1 To .Rows - 1
'
'            txtstr = flxPARTS.TextMatrix(intLoopCount, 1)
'
'            If flxPARTS.TextMatrix(intLoopCount, 12) <> "" Then
'                GoTo skip
'            End If
'
'            If txtstr = flxPARTNO.TextMatrix(flxPARTNO.Row, 1) Then 'flxPARTS.TextMatrix(actrow, 1) Then
'                    MsgBox "Part Number is already exist in row " & " " & intLoopCount & "" & Err.Description, vbExclamation + vbOKOnly, "Warning"
'                    .TextMatrix(.Row, 1) = ""
'                    .TextMatrix(.Row, 2) = ""
'                    .TextMatrix(.Row, 3) = ""
'                    .TextMatrix(.Row, 4) = ""
'                    .TextMatrix(.Row, 5) = ""
'                    .TextMatrix(.Row, 6) = ""
'                    .TextMatrix(.Row, 7) = ""
'                    .TextMatrix(.Row, 8) = ""
'                    .TextMatrix(.Row, 9) = ""
'                    .TextMatrix(.Row, 10) = ""
'                    txtPARTNO.Visible = False
'                    flxPARTNO.Visible = False
'                    flxPARTS.SetFocus
'                    flxPARTS.Col = 1
'                    Exit Sub
'            End If
'skip:
'
'        Next intLoopCount
'
'    End With
        
    If chckDupCount < 1 Then
        With flxPARTNO
            txtPARTNO.Visible = True
            txtPARTNO.ZOrder
            PrtSrch = True
            
            txtPARTNO.Text = flxPARTNO.TextMatrix(flxPARTNO.Row, 1)
            
            flxPARTS.TextMatrix(flxPARTS.Row, 2) = flxPARTNO.TextMatrix(flxPARTNO.Row, 2)
            flxPARTS.TextMatrix(flxPARTS.Row, 3) = flxPARTNO.TextMatrix(flxPARTNO.Row, 3)
            flxPARTS.TextMatrix(flxPARTS.Row, 10) = flxPARTNO.TextMatrix(flxPARTNO.Row, 4)
            'flxParts.TextMatrix(flxParts.Row, 5) = txtModel.Text
            txtPARTNO.SelStart = Val(txtPARTNO.Tag)
            txtPARTNO.SelLength = Len(txtPARTNO.Text)
                
        End With
            
        Call txtPartNo_LostFocus
        flxPARTNO.Visible = False
        flxPARTS.SetFocus
        CCol = 1
        Call ColCheckForEmpty
        
    ElseIf chckDupCount > 1 Then
    
        flxPARTS.SetFocus
        flxPARTS.Col = 1
        chckDupCount = 0
        
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->FlxPartNumber()")
End Sub

Private Sub flxPARTNO_EnterCell()
On Error GoTo er
    With flxPARTNO
        txtPARTNO.Visible = True
        txtPARTNO.ZOrder
        PrtSrch = True
        txtPARTNO.Text = flxPARTNO.TextMatrix(flxPARTNO.Row, 1)
        flxPARTS.TextMatrix(flxPARTS.Row, 2) = flxPARTNO.TextMatrix(flxPARTNO.Row, 2)
        flxPARTS.TextMatrix(flxPARTS.Row, 3) = flxPARTNO.TextMatrix(flxPARTNO.Row, 3)
        flxPARTS.TextMatrix(flxPARTS.Row, 10) = flxPARTNO.TextMatrix(flxPARTNO.Row, 4)
        'flxPARTS.TextMatrix(flxPARTS.Row, 8) = 0
       ' flxPARTS.TextMatrix(flxPARTS.Rows - 1, 4) = txtModel.Text
        txtPARTNO.SelStart = Val(txtPARTNO.Tag)
        txtPARTNO.SelLength = Len(txtPARTNO.Text)
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->FlxPartNumber-EnterCell()")
End Sub


Private Sub flxPARTNO_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = vbKeyEscape Then
       Tmp = Split(txtPARTNO.Tag, "|")
       For i = LBound(Tmp) To UBound(Tmp)
           flxPARTNO.TextMatrix(ActRow, i + 1) = Tmp(i)
       Next i
       txtPARTNO.Visible = False
       flxPARTNO.Visible = False
       flxPARTS.SetFocus
       flxPARTS.Col = 9
       
    End If
End Sub



Private Sub flxPARTNO_KeyPress(KeyAscii As Integer)
On Error GoTo er
   If KeyAscii = 13 Then
       
        
'        With flxPARTS
'
'            For intLoopCount = 1 To .Rows - 1
'
'                txtstr = flxPARTS.TextMatrix(intLoopCount, 1)
'
'                If flxPARTS.TextMatrix(intLoopCount, 12) <> "" Then
'                    GoTo skip
'                End If
'
'                If txtstr = flxPARTNO.TextMatrix(flxPARTNO.Row, 1) Then 'flxPARTS.TextMatrix(actrow, 1) Then
'                    MsgBox "Part Number is already exist in row " & " " & intLoopCount & "" & Err.Description, vbExclamation + vbOKOnly, "Warning"
'                    .TextMatrix(.Row, 1) = ""
'                    .TextMatrix(.Row, 2) = ""
'                    .TextMatrix(.Row, 3) = ""
'                    .TextMatrix(.Row, 4) = ""
'                    .TextMatrix(.Row, 5) = ""
'                    .TextMatrix(.Row, 6) = ""
'                    .TextMatrix(.Row, 7) = ""
'                    .TextMatrix(.Row, 8) = ""
'                    .TextMatrix(.Row, 9) = ""
'                    .TextMatrix(.Row, 10) = ""
'                    txtPARTNO.Visible = False
'                    flxPARTNO.Visible = False
'                    flxPARTS.SetFocus
'                    flxPARTS.Col = 1
'                    Exit Sub
'                End If
'skip:
'
'             Next intLoopCount
'
'        End With
        
    
'        flxPARTS.SetFocus
'        flxPARTS.Col = 4

       ' flxParts.TextMatrix(flxParts.Row, 5) = txtModel.Text
        Call CheckDuplicate
        
        If chckDupCount < 1 Then
            Call txtPartNo_LostFocus
            flxPARTNO.Visible = False
            flxPARTS.SetFocus
            
        ElseIf chckDupCount > 0 Then
            flxPARTS.SetFocus
            flxPARTS.Col = 1
            chckDupCount = 0
            Exit Sub
        End If
        CCol = 1
        Call ColCheckForEmpty
                
  End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->FlxPartNumber-Keypress()")
End Sub




Private Sub flxParts_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

    If Button = 2 Then
        If formAction = "add" Then
            PopupMenu mnuStatus, vbPopupMenuRightButton
        ElseIf formAction = "edit" Then
            PopupMenu mnuStatus, vbPopupMenuRightButton
        ElseIf formAction = "summary" Then
            Exit Sub
        End If
    End If
    
End Sub




Private Sub mnuAdd_Click()
On Error GoTo er
    ActRow = flxPARTS.Rows - 1
    If Val(flxPARTS.TextMatrix(ActRow, 8)) = 1 Then
        
        If flxPARTS.TextMatrix(ActRow, 1) <> flxPARTS.TextMatrix(flxPARTS.Rows - 1, 1) Then
            flxPARTS.SetFocus
            flxPARTS.Col = 1
            PrtAdd = False
            Exit Sub
        End If

        If flxPARTS.TextMatrix(flxPARTS.Rows - 1, 7) = "" Or flxPARTS.TextMatrix(flxPARTS.Rows - 1, 8) = "" Or flxPARTS.TextMatrix(flxPARTS.Rows - 1, 9) = "" Then
            MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
            Call textVisible(False)
                If flxPARTS.TextMatrix(flxPARTS.Rows - 1, 1) = "" Then
                    flxPARTS.SetFocus
                    flxPARTS.Col = 1
                ElseIf flxPARTS.TextMatrix(flxPARTS.Rows - 1, 7) = "" Then
                    flxPARTS.SetFocus
                    flxPARTS.Col = 7
                ElseIf flxPARTS.TextMatrix(flxPARTS.Rows - 1, 8) = "" Then
                    flxPARTS.SetFocus
                    flxPARTS.Col = 8
                ElseIf flxPARTS.TextMatrix(flxPARTS.Rows - 1, 9) = "" Then
                    flxPARTS.SetFocus
                    flxPARTS.Col = 9
                End If
            
            'PrtAdd = False
            Exit Sub
        End If

        flxPARTS.AddItem ""
        flxPARTS.SetFocus
        flxPARTS.Row = flxPARTS.Rows - 1
        flxPARTS.Col = 1
        PrtAdd = True
        flxPARTS.TextMatrix(flxPARTS.Row, 8) = 0
        

    ElseIf Val(flxPARTS.TextMatrix(ActRow, 8)) = 0 Then
        
        If flxPARTS.TextMatrix(flxPARTS.Rows - 1, 7) = "" Then 'Or flxPARTS.TextMatrix(flxPARTS.Rows - 1, 4) = "" Then
            MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
                Call textVisible(False)
                Call RowCheckForEmpty
                If rowEmpty = True Then
                    'Goto to last row of the table
                    CCol = 1
                    Call ColCheckForEmpty
                    Exit Sub
                Else
                    If flxPARTS.TextMatrix(flxPARTS.Row, 7) = "" Then
                        flxPARTS.SetFocus
                        flxPARTS.Col = 7
                    ElseIf flxPARTS.TextMatrix(flxPARTS.Row, 8) = "" Then
                        flxPARTS.SetFocus
                        flxPARTS.Col = 8
                    ElseIf flxPARTS.TextMatrix(flxPARTS.Row, 9) = "" Then
                        flxPARTS.SetFocus
                        flxPARTS.Col = 9
                    End If
                End If
            Exit Sub
        End If
       
            flxPARTS.AddItem ""
            flxPARTS.SetFocus
            flxPARTS.Row = flxPARTS.Rows - 1
            flxPARTS.Col = 1
            flxPARTS.TextMatrix(flxPARTS.Row, 8) = 0
            
        End If
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuAdd()")
    
End Sub




Private Sub mnuDelete_Click()
On Error GoTo er

    'actrow = flxPARTS.Rows - 1
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxPARTS.Rows = 2 Then
            flxPARTS.Rows = 1
            flxPARTS.Rows = 2
        Else
            flxPARTS.RemoveItem ActRow

            For i = 1 To flxPARTS.Rows - 1
                If flxPARTS.TextMatrix(i, 1) <> "" Then
                    txtTOTALPN.Text = i
                End If
             Next i
        End If
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuDelete()")
End Sub




Private Sub txtRP_Change()
On Error GoTo er
    textval = txtRP.Text
    If IsNumeric(textval) Then
        numval = textval
        If numval > 1 Then
            MsgBox "Unrecognized Replacement Value!" & vbCrLf & "Please choose Replacement Value from 0 and 1" & vbCrLf & "0 - No RP, 1 - w/ RP", vbExclamation + vbOKOnly, "Warning"
            txtRP.Visible = False
            flxPARTS.TextMatrix(ActRow, 8) = txtRP.Tag
            flxPARTS.SetFocus
            flxPARTS.Col = 8
                
            Exit Sub
        ElseIf numval = 1 Or numval = 0 Then
            flxPARTS.TextMatrix(flxPARTS.Row, 8) = txtRP
            flxPARTS.TextMatrix(flxPARTS.Row, 9) = ""
            txtRP.Visible = False
            flxPARTS.SetFocus
            flxPARTS.Col = 9
        End If
    End If
    
    If flxPARTS.TextMatrix(flxPARTS.Row, 7) = "" Then
        txtRP.Visible = False
        flxPARTS.SetFocus
        flxPARTS.Col = 7
    
    ElseIf flxPARTS.TextMatrix(flxPARTS.Row, 9) = "" Then
        txtRP.Visible = False
        flxPARTS.SetFocus
        flxPARTS.Col = 9
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtRP()")
End Sub



Private Sub txtRP_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyBack Then
        txtRP.Text = ""
        flxPARTS.TextMatrix(flxPARTS.Row, 8) = ""
    End If
End Sub



Private Sub txtRP_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRP, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        flxPARTS.TextMatrix(ActRow, 8) = txtRP.Tag
        txtRP.Visible = False
            If flxPARTS.TextMatrix(ActRow, 8) = 0 Then
                flxPARTS.SetFocus
                flxPARTS.Col = 7
            ElseIf flxPARTS.TextMatrix(ActRow, 8) = 1 Then
                flxPARTS.SetFocus
                flxPARTS.Col = 9
            End If
            
    End If
    
End Sub



Private Sub txtRP_LostFocus()
On Error GoTo er
    If txtRP.Tag = "skiplostfocus" Then
        txtRP.Tag = ""
        Exit Sub
    End If
    
    If txtRP.Text = "" Then
        MsgBox "Unrecognized Replacement Value!" & vbCrLf & "Please choose Replacement Value from 0 to 2" & vbCrLf & "0 - No RP, 1 - w/ RP, 2 - w/ Genuine CP", vbExclamation + vbOKOnly, "Warning"
        txtRP.Visible = False
        flxPARTS.TextMatrix(ActRow, 8) = txtRP.Tag
        flxPARTS.SetFocus
        flxPARTS.Col = 8
    End If
    
    txtRP.Visible = False
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtRP-LostFocus()")
End Sub



Private Sub txtCountrP_GotFocus()
On Error GoTo er
    If cpRs.State = adStateClosed Then Timer2.Enabled = True
    With flxCounterPart
        .Top = txtCountrP.Height + txtCountrP.Top
        .Left = txtCountrP.Left
        .Width = txtCountrP.Width
    End With

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtCountrP-GotFocus()")

End Sub


Private Sub txtCountrP_Change()
On Error GoTo er
    With flxCounterPart
    
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        
        .Top = txtCountrP.Height + txtCountrP.Top
        .Left = txtCountrP.Left
        .Width = txtCountrP.Width
        .Visible = True
        .ZOrder
        
        If txtCountrP.Text = "" Then .Visible = False: Exit Sub
        cpRs.Find "[PARTNUMBER] LIKE '" + txtCountrP.Text + "%'", , adSearchForward, 1
        If Not cpRs.EOF Then
            If Not PrtBln Then
                txtCountrP.Tag = Len(txtCountrP.Text)
                PrtBln = True
            End If
            
           .TopRow = cpRs.AbsolutePosition - 1
           .Row = cpRs.AbsolutePosition - 1
           .Col = 0: .ColSel = 1
           txtCountrP.Text = cpRs!PartNumber
           flxPARTS.TextMatrix(flxPARTS.Row, 11) = cpRs!SPAREPARTID
           flxPARTS.TextMatrix(flxPARTS.Row, 13) = cpRs!Description
           flxPARTS.TextMatrix(flxPARTS.Row, 14) = cpRs!Remarks
        
           txtCountrP.SelStart = Val(txtCountrP.Tag)
           txtCountrP.SelLength = Len(txtCountrP.Text)
           PrtBln = False
         
        Else
     
            flxPARTS.TextMatrix(flxPARTS.Row, 9) = ""
            flxPARTS.TextMatrix(flxPARTS.Row, 11) = ""
            .Visible = False
        End If
    End With

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtCountrP-change()")
'LocalErr:
'    If Err.Number = 30009 Then
'        flxPARTNO.Visible = False
'        txtPARTNO.Visible = False
'        If cpRs.State = adStateOpen Then cpRs.Close
'        Timer2.Enabled = True
'        MsgBox "Program is reloading Part Number" & vbCrLf & "Please wait for a few seconds...", vbInformation + vbOKOnly, "Reloading"
'    Else
'        flxPARTNO.Visible = False
'        txtPARTNO.Visible = False
'        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
'    End If
    
End Sub
    
    
Private Sub txtCountrP_KeyDown(KeyCode As Integer, Shift As Integer)
     
    If KeyCode = vbKeyDown Then
        txtCountrP.Tag = "skiplostfocus"
        flxCounterPart.Visible = True
        flxCounterPart.SetFocus
        
    ElseIf KeyCode = vbKeyEscape Then
        txtCountrP.Text = Empty
        Tmp = Split(txtCountrP.Tag, "|")
        
        For i = LBound(Tmp) To UBound(Tmp)
            flxCounterPart.TextMatrix(ActRow, i + 1) = Tmp(i)
        Next i
        
        txtCountrP.Tag = "skiplostfocus"
        flxPARTS.SetFocus
        flxPARTS.Col = 9
        txtCountrP.Visible = False
        flxCounterPart.Visible = False
        flxPARTS.TextMatrix(flxPARTS.Row, 11) = ""
    
    ElseIf KeyCode = vbKeyDelete Then
        PrtSrch = True
        cpRs.MoveFirst
        flxPARTS.TextMatrix(flxPARTS.Row, 11) = ""
    
    ElseIf KeyCode = vbKeyBack Then
        PrtSrch = True
        cpRs.MoveFirst
        flxPARTS.TextMatrix(flxPARTS.Row, 11) = ""
        flxCounterPart.Visible = False
    End If
    
End Sub

Public Sub CounterPartValidation()
On Error GoTo er
    If flxPARTS.TextMatrix(ActRow, 9) = "" Then
        flxPARTS.SetFocus
        flxPARTS.Col = 9
    End If
    
    
    If flxPARTS.TextMatrix(ActRow, 7) <> "" And flxPARTS.TextMatrix(ActRow, 11) <> "" Then
        
        
      
        If ActRow = flxPARTS.Rows - 1 Then
                
                flxPARTS.AddItem ""
                flxPARTS.SetFocus
                flxPARTS.Row = flxPARTS.Rows - 1
                'flxPARTS.Col = 4
                txtCountrP.Visible = False
                flxCounterPart.Visible = False
                PrtAdd = True
                Call CounterPartValue
        Else
            
            txtCountrP.Visible = False
            flxCounterPart.Visible = False
            flxPARTS.SetFocus
            flxPARTS.Col = 4
            If flxPARTS.TextMatrix(ActRow + 1, 1) = "" Then
            
                PrtAdd = False
                Call CounterPartValue
                
            ElseIf flxPARTS.TextMatrix(ActRow, 10) <> flxPARTS.TextMatrix(ActRow + 1, 12) Or flxPARTS.TextMatrix(ActRow, 11) <> flxPARTS.TextMatrix(ActRow + 1, 10) Then
            
                Ans = MsgBox("Do You Want To Insert New Row For New CounterPart?", vbQuestion + vbYesNo, "Information")
                If Ans = vbYes Then
                    PrtAdd = False
                    InsertRow (flxPARTS.Row)
                    Call CounterPartValue
                Else
                Ans = MsgBox("The Data from the Next row will be OverWrite for the New CounterPart" & vbCrLf & "Do You Want to Continue", vbQuestion + vbYesNo, "Information")
                    If Ans = vbYes Then
                        PrtAdd = False
                        Call CounterPartValue
                    Else
                        
                        flxPARTS.TextMatrix(ActRow, 9) = flxPARTS.TextMatrix(ActRow + 1, 1)
                        flxPARTS.TextMatrix(ActRow, 11) = flxPARTS.TextMatrix(ActRow + 1, 10)
                        Exit Sub
                    End If
                End If
                
            Else
                'Goto to last row of the table
                ActRow = flxPARTS.Rows - 1
                flxPARTS.TopRow = ActRow
                flxPARTS.Row = ActRow
                flxPARTS.RowSel = ActRow
                PrtAdd = True
            End If
            
        End If
    
    ElseIf flxPARTS.TextMatrix(ActRow, 7) = "" Then
    
        flxPARTS.SetFocus
        flxPARTS.Col = 7
        
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CounterPartValidation()")
End Sub

 Public Sub CounterPartValue()
On Error GoTo er
 'if new row is empty
    If PrtAdd = True Then
        For i = 1 To 12
            flxPARTS.Col = i
            flxPARTS.Row = ActRow + 1
            flxPARTS.CellBackColor = &HC0C0C0
        Next i
        flxPARTS.TextMatrix(ActRow, 9) = txtCountrP.Text
        flxPARTS.TextMatrix(flxPARTS.Rows - 1, 1) = txtCountrP.Text
        flxPARTS.TextMatrix(flxPARTS.Rows - 1, 2) = flxPARTS.TextMatrix(ActRow, 13)
        flxPARTS.TextMatrix(flxPARTS.Rows - 1, 3) = flxPARTS.TextMatrix(ActRow, 14)
        flxPARTS.TextMatrix(flxPARTS.Rows - 1, 10) = flxPARTS.TextMatrix(ActRow, 11)
        flxPARTS.TextMatrix(flxPARTS.Rows - 1, 12) = flxPARTS.TextMatrix(ActRow, 10)
        'flxParts.TextMatrix(flxParts.Rows - 1, 5) = txtModel.Text
        flxPARTS.TextMatrix(flxPARTS.Rows - 1, 7) = flxPARTS.TextMatrix(ActRow, 7)
        flxPARTS.TextMatrix(flxPARTS.Rows - 1, 8) = 0
        txtPARTNO.Enabled = False
        flxPARTS.SetFocus
        flxPARTS.Col = 4
    'if new row is not empty
    ElseIf PrtAdd = False Then
        flxPARTS.TextMatrix(ActRow, 9) = txtCountrP.Text
        If flxPARTS.TextMatrix(ActRow, 9) <> flxPARTS.TextMatrix(ActRow + 1, 1) Then
            flxPARTS.TextMatrix(ActRow + 1, 1) = txtCountrP.Text
            flxPARTS.TextMatrix(ActRow + 1, 2) = flxPARTS.TextMatrix(ActRow, 13)
            flxPARTS.TextMatrix(ActRow + 1, 3) = flxPARTS.TextMatrix(ActRow, 14)
            flxPARTS.TextMatrix(ActRow + 1, 10) = flxPARTS.TextMatrix(ActRow, 11)
            flxPARTS.TextMatrix(ActRow + 1, 12) = flxPARTS.TextMatrix(ActRow, 10)
           ' flxParts.TextMatrix(ActRow + 1, 5) = txtModel.Text
            flxPARTS.TextMatrix(ActRow + 1, 7) = flxPARTS.TextMatrix(ActRow, 7)
            'flxPARTS.TextMatrix(actrow + 1, 7) = ""
            'flxPARTS.TextMatrix(actrow + 1, 9) = ""
            flxPARTS.TextMatrix(ActRow + 1, 8) = 0
'            With flxPARTS
'            If Val(.TextMatrix(i, 11)) = 1 Then 'Final for Older Taking
'                      .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = &HC0C0C0
'                      .Col = 2
'            End With
            For i = 1 To 12
                flxPARTS.Col = i
                flxPARTS.Row = ActRow + 1
                flxPARTS.CellBackColor = &HC0C0C0
            Next i
            flxPARTS.Row = ActRow + 1
            flxPARTS.SetFocus
            flxPARTS.Col = 4
        Else
            flxPARTS.Row = ActRow + 1
            flxPARTS.SetFocus
            flxPARTS.Col = 4
        End If

    End If
    
    If flxPARTS.TextMatrix(flxPARTS.Row, 1) <> "" Then
        txtTOTALPN.Text = flxPARTS.Rows - 1
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CounterPartValue()")
End Sub

Private Sub txtCountrP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtCountrP, KeyAscii, ValidChars)
    If cpRs.State = adStateClosed Then Timer2.Enabled = True
    
    If KeyAscii = 13 Then
            Call CounterPartValidation
            Call Optshow
    End If
    
End Sub



Private Sub txtCountrP_LostFocus()
On Error GoTo er
   HoldLastRow = flxPARTS.Rows - 1
   
    If txtCountrP.Tag = "skiplostfocus" Then
        txtCountrP.Tag = ""
        Exit Sub
    End If
    
    If txtCountrP.Text <> "" Then
        If flxPARTS.TextMatrix(ActRow, 11) = "" Then
    
        MsgBox "Unrecognized Counter Part #!" & vbCrLf & "Please select a CounterPart # from list.", vbExclamation + vbOKOnly, "Warning"
        txtCountrP.Visible = False
        flxCounterPart.Visible = False
        
        If ActRow = HoldLastRow Then
            If flxPARTS.TextMatrix(ActRow, 11) = "" Then
                flxPARTS.TextMatrix(ActRow, 9) = ""
                flxPARTS.TextMatrix(ActRow, 11) = ""
                flxPARTS.TextMatrix(ActRow, 8) = 0
                flxPARTS.SetFocus
                flxPARTS.Col = 9
            End If

        ElseIf flxPARTS.TextMatrix(ActRow + 1, 1) <> "" Then
            flxPARTS.TextMatrix(ActRow, 9) = flxPARTS.TextMatrix(ActRow + 1, 1)
            flxPARTS.TextMatrix(ActRow, 11) = flxPARTS.TextMatrix(ActRow + 1, 10)
            flxPARTS.SetFocus
            flxPARTS.Col = 9
            
        ElseIf flxPARTS.TextMatrix(ActRow + 1, 1) = "" Then
            flxPARTS.TextMatrix(ActRow, 9) = ""
            flxPARTS.TextMatrix(ActRow, 11) = ""
            flxPARTS.TextMatrix(ActRow, 8) = 0
            flxPARTS.SetFocus
            flxPARTS.Col = 9
            
        End If
        Exit Sub
        End If
    End If
    
    'flxPARTS.TextMatrix(actrow, 9) = txtCountrP.Tag
    txtCountrP.Visible = False
    flxCounterPart.Visible = False
    'flxPARTS.SetFocus

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtCountrP-LostFocus()")
End Sub


Private Sub flxCounterPart_Click()
    Call CounterPartValidation
    Call Optshow
End Sub



Private Sub flxCounterPart_EnterCell()
On Error GoTo er
    With flxCounterPart
            txtCountrP.Visible = True
            txtCountrP.ZOrder
            PrtSrch = True
            txtCountrP.Text = flxCounterPart.TextMatrix(flxCounterPart.Row, 1)
            flxPARTS.TextMatrix(ActRow, 11) = flxCounterPart.TextMatrix(flxCounterPart.Row, 4)
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxCounterPart-EnterCell()")
End Sub


Private Sub flxCounterPart_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyEscape Then
        Tmp = Split(txtCountrP.Tag, "|")
        
        For i = LBound(Tmp) To UBound(Tmp)
                flxCounterPart.TextMatrix(ActRow, i + 1) = Tmp(i)
        Next i
        
        txtCountrP.Visible = False
        flxCounterPart.Visible = False
        flxPARTS.SetFocus
        flxPARTS.Col = 7
    End If
    
End Sub



Private Sub flxCounterPart_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call CounterPartValidation
        Call Optshow
    End If

End Sub



Private Sub cmbPG_Change()

    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If
    With cmbPG
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
 
End Sub


Private Sub cmbPG_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cmbPG.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbPG_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        StatSrch = True
    End If
 
End Sub

Private Sub cmbPG_KeyPress(KeyAscii As Integer)

On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cmbPG, KeyAscii, ValidChars)
    
    
    If KeyAscii = 13 Then
        flxPARTS.SetFocus
         
        
        If Val(flxPARTS.TextMatrix(ActRow, 8)) = 0 Then
            cmbstr = cmbPG.Text
        
            For i = 0 To cmbPG.ListCount - 1
                If UCase$(cmbPG.List(i)) = UCase$(cmbstr) Then
                    cmbPG.ListIndex = i
                    blnFound = True
                    Exit For
                End If
            Next i
    
            If Not blnFound Then
                MsgBox "Item was not found in the list.", _
                vbExclamation, _
                "Warning"
                cmbPG.Visible = False
                flxPARTS.SetFocus
                flxPARTS.Col = 7
                Exit Sub
            End If
    
            
            flxPARTS.TextMatrix(ActRow, 7) = cmbPG
            Call Optshow
            CCol = 5
            Call ColCheckForEmpty
            
            If ActRow = flxPARTS.Rows - 1 Then flxPARTS.AddItem ""
                flxPARTS.SetFocus
                flxPARTS.Row = flxPARTS.Rows - 1
                flxPARTS.Col = 1
            flxPARTS.TextMatrix(flxPARTS.Row, 8) = 0
            
            
            If flxPARTS.TextMatrix(ActRow, 1) <> "" Then
                txtTOTALPN.Text = ActRow 'flxPARTS.Rows - 1
            End If
            blnFound = False
           
            Exit Sub
        End If
    
    
    If Val(flxPARTS.TextMatrix(ActRow, 8)) = 1 Then
        cmbstr = cmbPG.Text
        
        For i = 0 To cmbPG.ListCount - 1
        
            If UCase$(cmbPG.List(i)) = UCase$(cmbstr) Then
                cmbPG.ListIndex = i
                blnFound = True
                flxPARTS.TextMatrix(ActRow, 7) = cmbPG
                Exit For
            End If
        Next i
    
    
        If Not blnFound Then
            MsgBox "Item was not found in the list.", _
            vbExclamation, _
            "Warning"
            cmbPG.Visible = False
            flxPARTS.SetFocus
            flxPARTS.Col = 7
            Exit Sub
        End If
        


        CCol = 5
        Call ColCheckForEmpty
        If flxPARTS.TextMatrix(ActRow, 9) <> "" Then
        
         
            flxPARTS.TextMatrix(ActRow, 7) = cmbPG
            Call CounterPartValidation

        Else
            Exit Sub 'do nothing
        End If
    
        ActRow = flxPARTS.Rows - 1
        flxPARTS.TopRow = ActRow
        flxPARTS.Row = ActRow
        flxPARTS.RowSel = ActRow
   
    End If
    
       
End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmbPG-KeyPress()")
End Sub

Private Sub cmbPG_LostFocus()
    cmbPG.Visible = False
End Sub




Private Sub txtSPECMODEL_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSPECMODEL, KeyAscii, ValidChars)
        If KeyAscii = 13 Then
        
            flxPARTS.SetFocus
'            If Val(flxPARTS.TextMatrix(actrow, 8)) = 0 Then
'                CCol = 2
'                flxPARTS.TextMatrix(actrow, 4) = txtSPECMODEL
'                Call ColCheckForEmpty
'                If flxPARTS.TextMatrix(actrow, 7) <> "" Then
'
'                    txtSPECMODEL.Visible = False
'                    If actrow = flxPARTS.Rows - 1 Then flxPARTS.AddItem ""
'
'                        flxPARTS.SetFocus
'                        flxPARTS.Row = flxPARTS.Rows - 1
'                        flxPARTS.Col = 1
'                    Exit Sub
'
'                End If
'                Exit Sub
'            End If
        CCol = 3
        flxPARTS.TextMatrix(ActRow, 5) = txtSPECMODEL.Text
        'txtSPECMODEL.Visible = False
        
        Call ColCheckForEmpty
        
'          'Goto to last row of the table
'        actrow = flxPARTS.Rows - 1
'        flxPARTS.TopRow = actrow
'        flxPARTS.Row = actrow
'        flxPARTS.RowSel = actrow
'        flxPARTS.Col = 1
'
       
    End If
     
End Sub

Public Sub textVisible(textBoxVisible As Boolean)

    txtPARTNO.Visible = textBoxVisible
    flxPARTNO.Visible = textBoxVisible
    txtFlxBrand.Visible = textBoxVisible
    txtSPECMODEL.Visible = textBoxVisible
    txtSize.Visible = textBoxVisible
    cmbPG.Visible = textBoxVisible
    txtRP.Visible = textBoxVisible
    txtCountrP.Visible = textBoxVisible
    flxCounterPart.Visible = textBoxVisible
    
    textBoxVisible = True

End Sub

Private Sub cmdAltA_Click()
    Call mnuAdd_Click
End Sub



Private Sub cmdAltD_Click()
    Call mnuDelete_Click
End Sub


