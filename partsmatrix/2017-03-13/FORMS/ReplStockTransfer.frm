VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form ReplStockTransfer 
   Caption         =   "Replenish Stock Transfer"
   ClientHeight    =   10290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   19185
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabConfirm 
      Height          =   2055
      Left            =   6240
      TabIndex        =   34
      Tag             =   "xy"
      Top             =   3360
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
      Tab(0).Control(2)=   "txtChecker"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdQuit"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdConfirm"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDriver"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtDriver 
         Height          =   375
         Left            =   1800
         TabIndex        =   38
         Tag             =   "xyf"
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "CONFIRM"
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
         Left            =   1920
         TabIndex        =   37
         Tag             =   "xyf"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "CANCEL"
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
         Left            =   3120
         TabIndex        =   36
         Tag             =   "xyf"
         Top             =   1440
         Width           =   1095
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
         TabIndex        =   40
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
         TabIndex        =   39
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
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Tag             =   "XY"
      Top             =   8520
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
      Height          =   6615
      Left            =   120
      TabIndex        =   17
      Tag             =   "XY"
      Top             =   2040
      Width           =   18975
      Begin MSFlexGridLib.MSFlexGrid flxPartsTtl 
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   5640
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   1508
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
         Height          =   5175
         Left            =   120
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   240
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   9128
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
         Top             =   5400
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
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         Tag             =   "XY"
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtSTATUS 
         Height          =   285
         Left            =   17160
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtSTFDate 
         Height          =   285
         Left            =   17160
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtSTFNo 
         Height          =   285
         Left            =   17160
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboRefNo 
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   960
         Width           =   3855
      End
      Begin VB.ComboBox cboTransTo 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Tag             =   "XY"
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblSTATUS 
         Caption         =   "STATUS"
         Height          =   255
         Left            =   16320
         TabIndex        =   13
         Tag             =   "XY"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblSTFDATE 
         Caption         =   "STF DATE"
         Height          =   255
         Left            =   16320
         TabIndex        =   12
         Tag             =   "XY"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblSTFNo 
         Caption         =   "STF #"
         Height          =   255
         Left            =   16320
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblRefNo 
         Caption         =   "REPLENISH REF #"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Tag             =   "XY"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblAddress 
         Caption         =   "ADDRESS"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblTransfer 
         Caption         =   "TRANSFER TO"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   240
         Width           =   1575
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&Add "
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
      Image           =   "ReplStockTransfer.frx":05B6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Tag             =   "XYF"
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
      Image           =   "ReplStockTransfer.frx":0B50
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   7800
      TabIndex        =   3
      Tag             =   "XYF"
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
      Image           =   "ReplStockTransfer.frx":10EA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   4
      Tag             =   "XYF"
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
      Image           =   "ReplStockTransfer.frx":1684
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
      Image           =   "ReplStockTransfer.frx":1C1E
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
      Image           =   "ReplStockTransfer.frx":21B8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "Confirm &Intransit"
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
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   19560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   19425
      m_MDI_Height    =   10860
      OriginalBorder  =   10
      OriginalCaption =   31
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

Dim i As Long, j As Long, z As Long, X As Long
Dim arrBRANCH() As String
Dim arrREFNO() As String
Dim Tmp() As String
Dim branchCode As String
Dim strPartial As String, strTotal As String
Dim StatSrch As Boolean
Dim tmpRs As New ADODB.Recordset
Dim formAction As String
Dim Qry As String
Dim matrixID As Integer, intSTFNO As Integer

Public MySource
Dim ProperExit As Boolean, AddNew As Boolean
Dim Ans As String

Sub CreateTempTableSPBRANCHDTL()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_SPBRANCHDTL_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_SPBRANCHDTL_DTLS( " & _
        "    REMARKS Varchar(20)      " & _
        "   ,SPAREPARTID Int(11)      " & _
        "   ,BRANCHSPOID Int(11)      " & _
        ");"
        
        Qry = "Insert into sp_partsmatrix.TMP_SPBRANCHDTL_DTLS(REMARKS,SPAREPARTID,BRANCHSPOID) VALUES "
       
        For i = 2 To flxParts.Rows - 1
            If i > 2 Then Qry = Qry & ","
           
           Qry = Qry & "('" & txtSTATUS.Text & "'," & Val(flxParts.TextMatrix(i, 7)) & "," & matrixID & ") "
    
        Next i
        
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTableSPBRANCHDTL()")
End Sub


Sub CreateTempTableSPITR()
Dim mySTFNO As String, SEQNO As Integer
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_SPITR_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_SPITR_DTLS( " & _
        "    REFERENCE Varchar(20)      " & _
        "   ,TRANSTYPE Varchar(10)      " & _
        "   ,SPAREPARTID Int(11)      " & _
        "   ,`OUT` Int(11)      " & _
        "   ,POREF Varchar(30)      " & _
        "   ,UPDATEDBY Int(11)      " & _
        "   ,SEQNO Int(11)      " & _
        ");"
        
        Qry = "Insert into sp_partsmatrix.TMP_SPITR_DTLS(REFERENCE,TRANSTYPE,SPAREPARTID,`OUT`,POREF,UPDATEDBY,SEQNO) VALUES "
        mySTFNO = "STF#" & txtSTFNo.Text
        SEQNO = 0
        For i = 2 To flxParts.Rows - 1
            If i > 2 Then Qry = Qry & ","
           SEQNO = SEQNO + 1
           Qry = Qry & "('" & mySTFNO & "', 'OUT' ," & Val(flxParts.TextMatrix(i, 7)) & "," & Val(flxParts.TextMatrix(i, 5)) & ",'" & cboRefNo.Text & "'," & UserID & "," & SEQNO & ") "
    
        Next i
        
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTableSPITR()")
End Sub

Sub CreateTempTableSPAREPARTS()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_SPAREPARTS_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_SPAREPARTS_DTLS( " & _
        "   QTYONHAND Int(11)      " & _
        "   ,UPDATEDBY Int(11)      " & _
        "   ,SPAREPARTID Int(11)      " & _
        ");"
        
        Qry = "Insert into sp_partsmatrix.TMP_SPAREPARTS_DTLS(QTYONHAND,UPDATEDBY,SPAREPARTID) VALUES "
        For i = 2 To flxParts.Rows - 1
        If flxParts.TextMatrix(i, 5) <> 0 Then
            If i > 2 Then Qry = Qry & ","
          
           Qry = Qry & "('" & Val(flxParts.TextMatrix(i, 5)) & "', " & UserID & ", " & Val(flxParts.TextMatrix(i, 7)) & ") "
        End If
        Next i
        
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->TMP_SPAREPARTS_DTLS()")
End Sub
Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_RSTOCKTRANSFER_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_RSTOCKTRANSFER_DTLS( " & _
        "    SPAREPARTID Int(11)      " & _
        "   ,QTY Int(11)      " & _
        "   ,UNIT Varchar(10)      " & _
        ");"
        
        Qry = "Insert into sp_partsmatrix.TMP_RSTOCKTRANSFER_DTLS(SPAREPARTID,QTY,UNIT) VALUES "
       
        For i = 2 To flxParts.Rows - 1
            If i > 2 Then Qry = Qry & ","
           
           Qry = Qry & "(" & Val(flxParts.TextMatrix(i, 7)) & "," & Val(flxParts.TextMatrix(i, 5)) & ",'PCS') "
    
        Next i
        
       
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
    matrixID = arrREFNO(cboRefNo.ListIndex, 0)
    Call ShowRecord(branchCode, cboRefNo.Text, matrixID)
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
  Call cboTransTo_Change
    
    DoEvents
    
    branchCode = arrBRANCH(cboTransTo.ListIndex, 2)
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

    flxParts.Rows = 2
    flxPartsTtl.Rows = 2
    cboTransTo.Text = ""
    txtAddress.Text = ""
    cboRefNo.Text = ""
    txtSTFNo.Text = "<AUTO>"
    txtSTFDate.Text = Format(Date, "mm/dd/yyyy")
    txtSTATUS.Text = ""
    txtPrepBy.Text = ""
    txtAppBy.Text = ""
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
    txtSTFNo.Locked = DataEntry
    txtSTFDate.Locked = DataEntry
    txtSTATUS.Locked = DataEntry
    txtPrepBy.Locked = DataEntry
    txtAppBy.Locked = DataEntry
    txtDateStart.Locked = DataEntry
   
    If cboTransTo.Locked Then
        cboTransTo.BackColor = txtHoldGrayBackColor.BackColor
        txtAddress.BackColor = txtHoldGrayBackColor.BackColor
        cboRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtSTFNo.BackColor = txtHoldGrayBackColor.BackColor
        txtSTFDate.BackColor = txtHoldGrayBackColor.BackColor
        txtSTATUS.BackColor = txtHoldGrayBackColor.BackColor
        txtPrepBy.BackColor = txtHoldGrayBackColor.BackColor
        txtAppBy.BackColor = txtHoldGrayBackColor.BackColor
        txtDateStart.BackColor = txtHoldGrayBackColor.BackColor

    Else
        cboTransTo.BackColor = txtHoldWhiteBackColor.BackColor
        txtAddress.BackColor = txtHoldWhiteBackColor.BackColor
        cboRefNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtSTFNo.Locked = True
        txtSTFDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtSTATUS.Locked = True
        txtPrepBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtAppBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtDateStart.BackColor = txtHoldWhiteBackColor.BackColor
        
    End If
End Sub


Private Sub cmdConfirm_Click()
 Dim MyRef       As String
    Dim rsConfirm   As New ADODB.Recordset
    Dim boQty       As Integer
    
        Ans = MsgBox("Do you want to confirm this transaction?", vbInformation + vbYesNo, "Confirm Intransit")
                If Ans = vbNo Then Exit Sub
        
        On Error GoTo LocalErr
        If txtDriver.Text = "" And txtChecker.Text = "" Then
            MsgBox "Please fill the Driver Name and Truck Plate No.", vbExclamation, "Confirm Intransit"
        Else
            Call CreateTempTableSPBRANCHDTL
            Call CreateTempTableSPITR
            Call CreateTempTableSPAREPARTS
            Call modPMTransfer.sp_StockTransfer_FinalTransfer(txtDriver.Text, txtChecker.Text, cboRefNo.Text, tmpRs)
            txtSTATUS.Text = tmpRs!Status
        End If
        MsgBox "Stock Transfer Ref#" & cboRefNo.Text & " is INTRANSIT.", vbOKOnly, "Successfully Save"
        tabConfirm.Visible = False
        
        With cmdTrans()
            For i = 0 To .Count - 1
                cmdTrans(i).Enabled = True
            Next i
        End With
       ' Frame1.Enabled = True
       ' Frame2.Enabled = True
       ' cmdFind.Enabled = True
       ' txtFind.Enabled = True
       ' flxOrd.Enabled = True
        Exit Sub
LocalErr:
    MsgBox Err.Description & vbCrLf & ". Please Contact MIS", vbCritical, "Confirm Stock Transfer"
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
    
ElseIf Index = 1 Then 'CONFIRM INSTRANSIT

            'Call RXResizer1_ResizeComplete
            If cboRefNo.Text = "" Then
                MsgBox "No available record", vbOKOnly + vbExclamation, "Confirm Intransit"
                Exit Sub
            End If
            If InStr(txtSTATUS.Text, "CANCELLED") > 0 Then
                MsgBox "Unable to Confirm " & txtSTFNo.Text & " because it was already Cancelled.", vbOKOnly + vbExclamation, "Confirm Intransit"
                Exit Sub
            End If
            If txtSTATUS.Text = "INTRANSIT" Then
                MsgBox "Stock Transfer " & txtSTFNo.Text & " already Confirm", vbOKOnly + vbExclamation, "Confirm Intransit"
                Exit Sub
            End If
            If txtSTATUS.Text <> "PACKING" Then
                MsgBox "Unable to Confirm Intransit Ref. No. " & txtSTFNo.Text & " because record is not save.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
             
            tabConfirm.Visible = True
            tabConfirm.Caption = "CONFIRM INTRANSIT"
            'Call LockData(False)
            With cmdTrans()
                For i = 0 To .Count - 1
                    cmdTrans(i).Enabled = False
                Next i
            End With
'                Frame1.Enabled = False
'                Frame2.Enabled = False
'                cmdFind.Enabled = False
'                txtFind.Enabled = False
'                flxOrd.Enabled = False
                
ElseIf Index = 2 Then 'View Summary
        frmSummaryRepl.mySrc = "Replenishment StockTransfer"
        frmSummaryRepl.Show
        Unload Me
        
ElseIf Index = 3 Then 'Print
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
        
        If txtSTATUS.Text <> "PICKING" Then
            MsgBox "ONLY THE STATUS 'PICKING' CAN BE TRANSFER!", vbExclamation, "Validation"
            Exit Sub
        End If
        
        Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        
        If Ans = vbNo Then Exit Sub
            
            
        If formAction = "ADD" Then
                Call CreateTempTable
                DoEvents
                Call modPMTransfer.sp_StockTransfer_create(Format(txtSTFDate.Text, "YYYY-mm-dd"), cboTransTo.Text, txtAddress.Text, cboRefNo.Text, txtPrepBy.Text, txtAppBy.Text, txtSTATUS.Text, UserID, matrixID, tmpRs)
                txtSTFNo.Text = tmpRs!TRANSFERID
'        Else
'            Call CreateTempTable
'            DoEvents
'            Call modPMBranchIdentifying.sp_brnchIdentify_update(intREFID, txtRefNo.Text, Format(txtDatePrepared.Text, "YYYY-mm-dd"), Format(txtDateFinal.Text, "YYYY-mm-dd"), UserID, mainRs)
'            Call ShowRecord(branchCode, intID, strBoolean) 'to refresh
        End If
            Call SetButtonVisibility(False)
            Call LockData(False)
            formAction = ""
ElseIf Index = 6 Then 'Cancel
        Ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        
        If Ans = vbNo Then Exit Sub
        
        Call SetButtonVisibility(False)
        Call LockData(False)
'        Call clearform
        formAction = ""
        
End If

End Sub


Private Sub Form_Load()
    'Call flxColChanger
   ' Call flxColChangerForSumm
    Call LockData(False)
    
    txtSTFNo.Text = "<Auto>"
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
    Call modPMTransfer.sp_getRefNo(branchCode, tmpRs)
    If tmpRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cboRefNo, tmpRs, arrREFNO, "PMATRIXID", "REFNO")
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
    
    'Frame width
    Frame1.Width = Me.Width - 465
    Frame2.Width = Me.Width - 465
    Frame3.Width = Me.Width - 465
  
    'Frame1 label Left
    lblTransfer.Left = 120
    lblAddress.Left = 120
    lblRefNo.Left = 120
    lblSTFNo.Left = Frame1.Width - 2655
    lblSTFDATE.Left = Frame1.Width - 2655
    lblSTATUS.Left = Frame1.Width - 2655
  
    'Frame1 Label Width
    lblTransfer.Width = 1575
    lblAddress.Width = 1575
    lblRefNo.Width = 1575
    lblSTFNo.Width = 855
    lblSTFDATE.Width = 855
    lblSTATUS.Width = 855
  

    
    'Frame1 Combobox Left
    cboTransTo.Left = 1800
    cboRefNo.Left = 1800
    
    'Frame1 Combobox Width
    cboTransTo.Width = 3855
    cboRefNo.Width = 3855
   
    
    'Frame1 Text Left
    txtAddress.Left = 1800

    txtSTFNo.Left = Frame1.Width - 1815
    txtSTFDate.Left = Frame1.Width - 1815
    txtSTATUS.Left = Frame1.Width - 1815
    
    'Frame1 Text Width
    txtAddress.Width = 3855
    txtSTFNo.Width = 1575
    txtSTFDate.Width = 1575
    txtSTATUS.Width = 1575

    'Frame2 FlexGrid Width
    flxParts.Width = Frame2.Width - 240
    flxPartsTtl.Width = Frame2.Width - 240
'
'    'Frame2 FlexGrid Height
'    flxParts.Height = Frame2.Height - 2400
'    flxPartsTtl.Height = Frame2.Height - 4560
    
    'Frame3 Label Left
    lblPrepBy.Left = 240
    lblAppby.Left = 240
    lblSTARTREPL.Left = 240
    
    'Frame3 Label Width
    lblPrepBy.Width = 1215
    lblAppby.Width = 1215
    lblSTARTREPL.Width = 2655
    
    'Frame3 Text Left
    txtPrepBy.Left = 1560
    txtAppBy.Left = 1560
    txtDateStart.Left = 3000
    
    'Frame3 Text Width
    txtPrepBy.Width = 3615
    txtAppBy.Width = 3615
    txtDateStart.Width = 2175
End Sub

Public Sub getSTFID(mySTFID As Integer)
   intSTFNO = mySTFID
End Sub


Public Sub ShowRecord(MyBranchCode As String, myREFNO As String, MyID As Integer)

  Screen.MousePointer = vbHourglass
  cboRefNo.Text = myREFNO
  txtSTFNo.Text = intSTFNO
   DoEvents
  Call modPMTransfer.sp_get_BranchName(MyBranchCode, tmpRs)
    cboTransTo.Text = tmpRs!branchName
    txtAddress.Text = tmpRs!Address
    
    
  DoEvents
 Call modPMTransfer.sp_StockTransfer_getDetails(MyBranchCode, myREFNO, MyID, tmpRs)
 
    If tmpRs.State <> 1 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With tmpRs
        If .RecordCount > 0 Then
            .MoveFirst
           ' intTtlPN = .RecordCount
            flxParts.Rows = .RecordCount + 2
            
           j = 0
            For i = 2 To flxParts.Rows - 1
                j = j + 1
                flxParts.TextMatrix(i, 1) = j
                flxParts.TextMatrix(i, 2) = !PartNumber
                flxParts.TextMatrix(i, 3) = !Description
                flxParts.TextMatrix(i, 4) = !partsGroup
                flxParts.TextMatrix(i, 5) = !Qty
                flxParts.TextMatrix(i, 6) = !MODELBRAND
                flxParts.TextMatrix(i, 7) = !SPAREPARTID
                .MoveNext
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
            
            For i = 2 To flxParts.Rows - 2
                If flxParts.TextMatrix(i, 6) <> flxParts.TextMatrix(i + 1, 6) Then
                    
                    flxParts.Rows = flxParts.Rows + 1
                        For z = flxParts.Rows - 1 To i Step -1
                            If z = i Then
                              
                                For X = 1 To flxParts.Cols - 1
                                flxParts.TextMatrix(z + 1, X) = flxParts.TextMatrix(i + 2, 6)
                                    'SET MERGING
                                    
                                    flxParts.MergeCells = 1
                                    flxParts.MergeRow(z + 1) = True
                                    flxParts.MergeCol(0) = True
                                    flxParts.MergeCol(0) = True
                                    For j = 1 To 6
                                    flxParts.MergeCol(j) = False
                                    Next j
                                    
                                Next X
                            Else
                                For X = 1 To flxParts.Cols - 1
                                flxParts.TextMatrix(z, X) = flxParts.TextMatrix(z - 1, X)
                                Next X
                            End If
                        Next z
                End If
            Next i
        End If
    End With
    If tmpRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        tmpRs.Close
    End If
    
    
    Call modPMTransfer.sp_StockTransfer_getDetailsTtl(MyBranchCode, myREFNO, MyID, tmpRs)
    If tmpRs.State <> 1 Then Exit Sub

    If tmpRs.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    With tmpRs
     If .RecordCount > 0 Then
            .MoveFirst
           ' intTtlPN = .RecordCount
           txtSTATUS.Text = Format$(tmpRs!Remarks, "")
           txtPrepBy.Text = Format$(!PROCESSEDBY, "")
           txtAppBy.Text = Format$(!APRVBY, "")
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
            
    End If
    End With
    If tmpRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        tmpRs.Close
    End If
   
End Sub

