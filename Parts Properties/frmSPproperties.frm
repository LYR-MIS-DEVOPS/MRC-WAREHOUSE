VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Spareparts Properties"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   14445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   " SUBSTITUTE SPAREPARTS"
      Height          =   2535
      Left            =   240
      TabIndex        =   29
      Top             =   4560
      Width           =   14055
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1815
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   5
         Cols            =   7
         RowHeightMin    =   310
         FormatString    =   $"frmSPproperties.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   14055
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   7680
         TabIndex        =   28
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   7680
         TabIndex        =   27
         Top             =   840
         Width           =   6135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1920
         TabIndex        =   26
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1920
         TabIndex        =   25
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   7680
         TabIndex        =   20
         Top             =   3240
         Width           =   3375
      End
      Begin VB.CheckBox Check2 
         Caption         =   "SALE DATE END:"
         Height          =   255
         Left            =   5760
         TabIndex        =   19
         Top             =   3270
         Width           =   2295
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1920
         TabIndex        =   18
         Top             =   2760
         Width           =   3375
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   7680
         TabIndex        =   16
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox Text7 
         Height          =   795
         Left            =   7680
         TabIndex        =   15
         Top             =   2280
         Width           =   6135
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   7680
         TabIndex        =   14
         Top             =   360
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "SALE DATE:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "PARTS MATRIX MC"
         Height          =   255
         Index           =   6
         Left            =   5760
         TabIndex        =   24
         Top             =   1350
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "BARCODE VALUE"
         Height          =   255
         Index           =   5
         Left            =   5760
         TabIndex        =   23
         Top             =   870
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "RE-ORDER POINT"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   1350
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "CLASSIFICATION"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   870
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "SUPPLIER BRAND:"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   17
         Top             =   2790
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "SUPPLIER"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   9
         Top             =   2310
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "APPLICABLE MODEL:"
         Height          =   255
         Index           =   8
         Left            =   5760
         TabIndex        =   8
         Top             =   2310
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "PARTSMATRIX MODEL"
         Height          =   255
         Index           =   7
         Left            =   5760
         TabIndex        =   7
         Top             =   1830
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "DESCRIPTION"
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   6
         Top             =   390
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "MAXIMUM PO QTY."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1830
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "PARTNUMBER"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   390
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
