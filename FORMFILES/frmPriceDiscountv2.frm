VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPriceDiscountv2 
   Caption         =   "Price and Discount Adjustment"
   ClientHeight    =   9795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16350
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   16350
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   80
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Caption         =   "SHOW PRICE HISTORY"
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
      Image           =   "frmPriceDiscountv2.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame6 
      Height          =   2775
      Left            =   10800
      TabIndex        =   64
      Tag             =   "XYF"
      Top             =   6600
      Width           =   5055
      Begin VB.TextBox txtKAFixPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   72
         Tag             =   "XYF"
         Text            =   "0"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtKAProfMar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   71
         Tag             =   "XYF"
         Text            =   "0"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtKAMarkUp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   70
         Tag             =   "XYF"
         Text            =   "0"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtKA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   69
         Tag             =   "XYF"
         Text            =   "0"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "GROSS PROFIT MARGIN"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   68
         Tag             =   "XYF"
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "MARK-UP (%)"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   67
         Tag             =   "XYF"
         Top             =   1100
         Width           =   2295
      End
      Begin VB.TextBox txtKAMaxDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   66
         Tag             =   "XYF"
         Text            =   "0"
         Top             =   2355
         Width           =   1935
      End
      Begin VB.TextBox txtKAMarkUpAmnt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   65
         Tag             =   "XYF"
         Text            =   "0"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "KEY ACCOUNTS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   255
         Left            =   4395
         TabIndex        =   78
         Tag             =   "XYF"
         Top             =   780
         Width           =   135
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "NEW K-ACCOUNTS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Tag             =   "XYF"
         Top             =   1950
         Width           =   2295
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   255
         Left            =   4395
         TabIndex        =   76
         Tag             =   "XYF"
         Top             =   2445
         Width           =   135
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   " MAX DISCOUNT:"
         Height          =   255
         Left            =   360
         TabIndex        =   75
         Tag             =   "XYF"
         Top             =   2355
         Width           =   2055
      End
      Begin VB.Label Label27 
         Caption         =   "AMOUNT MARK-UP"
         Height          =   255
         Left            =   150
         TabIndex        =   74
         Tag             =   "XYF"
         Top             =   1480
         Width           =   2055
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   255
         Left            =   4395
         TabIndex        =   73
         Tag             =   "XYF"
         Top             =   1120
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   120
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   4920
      Width           =   16095
      Begin VB.TextBox txtNetofVat 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   13680
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtRefDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   1200
         Width           =   8295
      End
      Begin VB.TextBox txtRefNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   840
         Width           =   3135
      End
      Begin lvButton.lvButtons_H cmdTrans 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "&EDIT"
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
         Image           =   "frmPriceDiscountv2.frx":57F2
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3015
         Left            =   120
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   1560
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   5318
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "     ."
         TabPicture(0)   =   "frmPriceDiscountv2.frx":5D8C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VB.Frame Frame5 
            Height          =   2775
            Left            =   5280
            TabIndex        =   41
            Tag             =   "XYF"
            Top             =   120
            Width           =   4935
            Begin VB.TextBox txtWSMarkUpAmnt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   49
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   1440
               Width           =   1935
            End
            Begin VB.TextBox txtWSMaxDiscount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   48
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   2355
               Width           =   1935
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "MARK-UP (%)"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   47
               Tag             =   "XYF"
               Top             =   1100
               Width           =   2295
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "GROSS PROFIT MARGIN"
               Enabled         =   0   'False
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   46
               Tag             =   "XYF"
               Top             =   720
               Width           =   2295
            End
            Begin VB.TextBox txtWS 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   45
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtWSMarkUp 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   44
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   1080
               Width           =   1935
            End
            Begin VB.TextBox txtWSProfMar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   43
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   720
               Width           =   1935
            End
            Begin VB.TextBox txtWSFixPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   42
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   1920
               Width           =   1935
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   4395
               TabIndex        =   56
               Tag             =   "XYF"
               Top             =   1120
               Width           =   135
            End
            Begin VB.Label Label21 
               Caption         =   "AMOUNT MARK-UP"
               Height          =   255
               Left            =   150
               TabIndex        =   55
               Tag             =   "XYF"
               Top             =   1480
               Width           =   2055
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   " MAX DISCOUNT:"
               Height          =   255
               Left            =   360
               TabIndex        =   54
               Tag             =   "XYF"
               Top             =   2355
               Width           =   2055
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   4395
               TabIndex        =   53
               Tag             =   "XYF"
               Top             =   2445
               Width           =   135
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "NEW WSP:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Tag             =   "XYF"
               Top             =   1950
               Width           =   2295
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   4395
               TabIndex        =   51
               Tag             =   "XYF"
               Top             =   780
               Width           =   135
            End
            Begin VB.Label Label13 
               Caption         =   "BRANCH WHOLESALE:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   50
               Tag             =   "XYF"
               Top             =   240
               Width           =   2055
            End
         End
         Begin VB.Frame Frame4 
            Height          =   2775
            Left            =   120
            TabIndex        =   25
            Tag             =   "XYF"
            Top             =   120
            Width           =   4815
            Begin VB.TextBox txtSRPMarkUpAmnt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   33
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   1440
               Width           =   1815
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "MARK-UP (%)"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   32
               Tag             =   "XYF"
               Top             =   1100
               Width           =   2295
            End
            Begin VB.TextBox txtSRPMaxDis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   31
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   2355
               Width           =   1815
            End
            Begin VB.TextBox txtSRPFixPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               TabIndex        =   30
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   1920
               Width           =   1815
            End
            Begin VB.TextBox txtSRPProfMar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   29
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox txtSRPMarkUp 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   28
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox txtSRP 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   27
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "GROSS PROFIT MARGIN"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   26
               Tag             =   "XYF"
               Top             =   720
               Width           =   2295
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   4275
               TabIndex        =   40
               Tag             =   "XYF"
               Top             =   1120
               Width           =   135
            End
            Begin VB.Label Label20 
               Caption         =   "AMOUNT MARK-UP"
               Height          =   255
               Left            =   150
               TabIndex        =   39
               Tag             =   "XYF"
               Top             =   1480
               Width           =   2055
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Caption         =   "NEW SRP:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1440
               TabIndex        =   38
               Tag             =   "XYF"
               Top             =   1950
               Width           =   975
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   4275
               TabIndex        =   37
               Tag             =   "XYF"
               Top             =   780
               Width           =   135
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   4275
               TabIndex        =   36
               Tag             =   "XYF"
               Top             =   2400
               Width           =   135
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   " MAX DISCOUNT:"
               Height          =   255
               Left            =   1080
               TabIndex        =   35
               Tag             =   "XYF"
               Top             =   2355
               Width           =   1335
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "CURRENT SRP:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Tag             =   "XYF"
               Top             =   240
               Width           =   1575
            End
         End
      End
      Begin lvButton.lvButtons_H cmdTrans 
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "&CANCEL"
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
         Image           =   "frmPriceDiscountv2.frx":5DA8
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdTrans 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "&SAVE"
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
         Image           =   "frmPriceDiscountv2.frx":6342
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. DATE:"
         Height          =   255
         Left            =   12000
         TabIndex        =   63
         Tag             =   "XYF"
         Top             =   915
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. NO.:"
         Height          =   255
         Left            =   12000
         TabIndex        =   62
         Tag             =   "XYF"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "NET OF VAT:"
         Height          =   255
         Left            =   12000
         TabIndex        =   61
         Tag             =   "XYF"
         Top             =   1230
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "DESCRIPTION:"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Tag             =   "XYF"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "PART NUMBER:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Tag             =   "XYF"
         Top             =   870
         Width           =   1455
      End
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   16680
      TabIndex        =   16
      Text            =   "HoldBackColor"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   16680
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "txtHoldGrayBackColor"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   16800
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame SPpricing 
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   600
      Width           =   16095
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   2210
         TabIndex        =   9
         Top             =   270
         Width           =   255
      End
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         Left            =   12480
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   240
         Width           =   3495
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All Suppliers"
         Height          =   315
         Left            =   14760
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   12480
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtFindPart 
         Height          =   375
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
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
         Image           =   "frmPriceDiscountv2.frx":68DC
         cBack           =   -2147483633
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFind 
         Height          =   3135
         Left            =   120
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   960
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   11
         RowHeightMin    =   300
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "           "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   11
      End
      Begin VB.Label lblFindParts 
         Caption         =   "FIND PART DESCRIPTION:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   285
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "SUPPLIER:"
         Height          =   255
         Left            =   11520
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "TYPE:"
         Height          =   255
         Left            =   11520
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7200
      Top             =   0
   End
   Begin VB.Timer tmr0SRP 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8400
      Top             =   0
   End
   Begin VB.Timer tmrSMSQTYLSI 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9000
      Top             =   0
   End
   Begin lvButton.lvButtons_H cmd0SRP 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "Zero SRP"
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
      Image           =   "frmPriceDiscountv2.frx":C0CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "&REFRESH"
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
      Image           =   "frmPriceDiscountv2.frx":C668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   " E&XIT FORM"
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
      Image           =   "frmPriceDiscountv2.frx":CC02
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   16590
      m_MDI_Height    =   10365
   End
   Begin VB.Menu mnuPopUpFormat 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmPriceDiscountv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ans As VbMsgBoxResult
Dim AllArrSupplier() As String
Dim AllArrClass() As String
Dim AllArrPartNo() As String
Dim AllArrDesc() As String
Dim arrSupplier() As String
Dim arrClass() As String
Dim arrPartNo() As String
Dim arrDesc() As String
Dim rsParts As New ADODB.Recordset

Dim ActRow As Integer, fRow As Integer, fCol As Integer, TCol As Integer, AccID As Integer, ColPrn As Integer

Dim SearchcboSup As Boolean
Dim SearchcboType As Boolean
Dim SearchPartNo As Boolean
Dim SupCode As String
Dim pSupCode As String, chkSup As Boolean, PrevSup As String, pClass As String, IsSelect As Boolean
Dim didRecordEntry As Boolean
Dim i As Long, j As Long
Dim SPid As Long

Dim PartDesc As String
Dim flxrow As Integer
Dim bNone As Boolean

Sub getSRPProfMarginRatio()
On Error GoTo er

    Dim AmntMarkUp As Double
    Dim CostVat As Double
    Dim NewSRP As Double
    Dim ProfMar As Double
    Dim MarkUpPerc As Double
    
    If txtNetofVat.Text = 0 Or txtSRP.Text = 0 Then Exit Sub
    
    NewSRP = Val(Replace(Trim(txtSRP.Text), ",", ""))
    'CostVat = Format(txtNetofVat.Text, "#,#00")
    CostVat = Val(Replace(txtNetofVat, ",", ""))
    CostVat = CostVat * 1.12
    CostVat = Format(CostVat, "#,#00.00")
    
    
    AmntMarkUp = NewSRP - CostVat
    
    'Ratio = ProMar / NewSRP
    ProfMar = AmntMarkUp / NewSRP
    ProfMar = ProfMar * 100
    
    txtSRPMarkUpAmnt.Text = Format(Round(AmntMarkUp), "#,#00")
    txtSRPFixPrice.Text = Format(Round(NewSRP), "#,#00")
    txtSRPProfMar.Text = Format(Round(ProfMar), "#00")
    
    MarkUpPerc = ((CostVat + AmntMarkUp) - CostVat) / CostVat
    MarkUpPerc = MarkUpPerc * 100
    txtSRPMarkUp.Text = Format(MarkUpPerc, "#,#00.00")
        
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->getSRPProfMarginRatio()")
End Sub

Sub getWSProfMarRatio()
On Error GoTo er

    Dim NewWSP As Double
    Dim AmntMarkUp As Double
    Dim CostVat As Double
    Dim ProfMar As Double
    Dim MarkUpPerc As Double
    If txtNetofVat.Text = 0 Or txtWS.Text = 0 Then Exit Sub
    
    NewWSP = Val(Replace(Trim(txtWS.Text), ",", ""))
    'CostVat = Format(txtNetofVat.Text, "#,#00")
    CostVat = Val(Replace(txtNetofVat, ",", ""))
    CostVat = CostVat * 1.12
    CostVat = Format(CostVat, "#,#00.00")
    
    AmntMarkUp = NewWSP - CostVat
    
    'Ratio = ProMar / NewSRP
    ProfMar = AmntMarkUp / NewWSP
    ProfMar = ProfMar * 100
    
    txtWSMarkUpAmnt.Text = Format(AmntMarkUp, "#,#00.00")
    txtWSFixPrice.Text = Format(NewWSP, "#,#00.00")
    txtWSProfMar.Text = Format(ProfMar, "#00")
    
    MarkUpPerc = ((CostVat + AmntMarkUp) - CostVat) / CostVat
    MarkUpPerc = MarkUpPerc * 100
    txtWSMarkUp.Text = Format(MarkUpPerc, "#00")
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->getWSProfMarRatio()")
End Sub

Sub getKAProfMarRatio()
On Error GoTo er

    Dim NewKAP As Double
    Dim AmntMarkUp As Double
    Dim CostVat As Double
    Dim ProfMar As Double
    Dim MarkUpPerc As Double
    If txtNetofVat.Text = 0 Or txtKA.Text = 0 Then Exit Sub
    
    NewKAP = Val(Replace(Trim(txtKA.Text), ",", ""))
    'CostVat = Format(txtNetofVat.Text, "#,#00")
    CostVat = Val(Replace(txtNetofVat, ",", ""))
    CostVat = CostVat * 1.12
    CostVat = Format(CostVat, "#,#00.00")
    
    AmntMarkUp = NewKAP - CostVat
    
    'Ratio = ProMar / NewSRP
    ProfMar = AmntMarkUp / NewKAP
    ProfMar = ProfMar * 100
    
    txtKAMarkUpAmnt.Text = Format(AmntMarkUp, "#,#00.00")
    txtKAFixPrice.Text = Format(NewKAP, "#,#00.00")
    txtKAProfMar.Text = Format(ProfMar, "#00")
    
    MarkUpPerc = ((CostVat + AmntMarkUp) - CostVat) / CostVat
    MarkUpPerc = MarkUpPerc * 100
    txtKAMarkUp.Text = Format(MarkUpPerc, "#00")
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->getKAProfMarRatio()")
End Sub
Private Sub cbotype_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer


    If SearchcboType Then
        SearchcboType = False
        Exit Sub
    End If


    With cboType
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)

            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboType = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboType_Change()")
End Sub
Private Sub cboType_Click()
    
    Screen.MousePointer = vbHourglass
    tmrRefresh.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal

End Sub
Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboType = True
    End If
End Sub
Private Sub cboType_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboType, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then

        pClass = cboType.Text
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
    End If
End Sub
Private Sub cboSupplier_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If SearchcboSup Then
        SearchcboSup = False
        Exit Sub
    End If

    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
            
            SupCode = arrSupplier(i, 0)
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboSup = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboSupplier_Change()")
End Sub
Private Sub cboSupplier_Click()
    
    SupCode = arrSupplier(cboSupplier.ListIndex, 0)

        
    chkAll.Value = 0
    
    If IsSelect = True Then: IsSelect = False: Exit Sub
    
    Screen.MousePointer = vbHourglass
    tmrRefresh.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
        
    
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboSup = True
    End If
End Sub
Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboSupplier, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        If cboSupplier.Text = "" Then Exit Sub
        
        chkAll.Value = 0
        
        If IsSelect = True Then: IsSelect = False: Exit Sub
        
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        frm_OnProcess2.OwnerForm = "SRP Updater"
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
    End If
End Sub
Private Sub Check1_Click(Index As Integer)
On Error GoTo er
    
    If didRecordEntry = False Then
        Check1(0).Value = 0
        Check1(1).Value = 0
        Check1(3).Value = 0
        Check1(4).Value = 0
        Exit Sub
    End If

    If Index = 0 Then ' Check box SRP Mark-Up
        
        If Check1(0).Value = 1 Then
            txtSRPMarkUp.Locked = False
            txtSRPMarkUp.BackColor = txtHoldWhiteBackColor.BackColor
            
            If txtSRPMarkUp.Text <> "" Then
                txtSRPMarkUp.SelStart = 0
                txtSRPMarkUp.SelLength = Len(txtSRPMarkUp.Text)
                txtSRPMarkUp.SetFocus
            End If
            
            Check1(1).Value = 0

        Else
            txtSRPMarkUp.Locked = True
            txtSRPMarkUp.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
    ElseIf Index = 3 Then ' Check box WS Mark-Up

        If Check1(3).Value = 1 Then
            txtWSMarkUp.Locked = False
            txtWSMarkUp.BackColor = txtHoldWhiteBackColor.BackColor

            If txtWSMarkUp.Text <> "" Then
                txtWSMarkUp.SelStart = 0
                txtWSMarkUp.SelLength = Len(txtWSMarkUp.Text)
                txtWSMarkUp.SetFocus
            End If

            Check1(4).Value = 0
        Else
            txtWSMarkUp.Locked = True
            txtWSMarkUp.BackColor = txtHoldGrayBackColor.BackColor
        End If
    ElseIf Index = 5 Then ' Check box KA Mark-Up

        If Check1(5).Value = 1 Then
            txtKAMarkUp.Locked = False
            txtKAMarkUp.BackColor = txtHoldWhiteBackColor.BackColor

            If txtKAMarkUp.Text <> "" Then
                txtKAMarkUp.SelStart = 0
                txtKAMarkUp.SelLength = Len(txtWSMarkUp.Text)
                txtKAMarkUp.SetFocus
            End If

            Check1(4).Value = 0
        Else
            txtKAMarkUp.Locked = True
            txtKAMarkUp.BackColor = txtHoldGrayBackColor.BackColor
        End If
    
   End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->Check1_Click")
End Sub
Private Sub chkAll_Click()
    
    If chkAll.Value = 1 Then
        PrevSup = cboSupplier.Text
        chkSup = True: cboSupplier.Text = ""
        SupCode = ""
        IsSelect = False
        
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
    Else
    
        chkSup = False
        IsSelect = True
        If PrevSup <> "" Then cboSupplier.Text = PrevSup
        
        pSupCode = SupCode
        
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
    End If
    
End Sub
Private Sub FindZeroSRP()
    Dim Qry As String
    
    If rsParts.State = adStateOpen Then rsParts.Close
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION " & _
          ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS', sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT" & _
          " From SPAREPARTS as sp WHERE sp.STOREPRICE=0 ORDER BY sp.PARTNUMBER;"
          
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxFind.Clear
        flxFind.Rows = 2
        flxFind.FormatString = "|<PART NUMBER |<DESCRIPTION |^NET OF VAT |^COST WITH VAT |>SRP  |>WSP  |>KEYACCTS   |>QTY. |>SIX MSQ |>BACK ORDER |>LAST DATE IN |<APPLICABLE MODEL"
        
    Else
        Set flxFind.DataSource = rsParts
        flxFind.Refresh
        flxFind.ColAlignment(3) = flexAlignRightCenter
        flxFind.ColAlignment(4) = flexAlignRightCenter
        flxFind.ColAlignment(5) = flexAlignRightCenter
    End If
    'format cur
    Call FormatCur
    Call flxFind_Click

End Sub
Private Sub GetSMSQTYLDI(pSPID As Long)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim sortby As String
    
    sortby = "SPAREPART"
    Call modSRP.sp_get_SMSQTY_LDI(pSPID, sortby, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If


    tmpRs.MoveFirst

        While Not tmpRs.EOF
                flxFind.TextMatrix(flxFind.row, 7) = tmpRs!SMSQTY
                flxFind.TextMatrix(flxFind.row, 9) = tmpRs!TRANSDATE
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetSMSQTYLDI()")
End Sub
Private Sub cmd0SRP_Click()
    
    Screen.MousePointer = vbHourglass
    tmr0SRP.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal

End Sub
Private Sub cmdExit_Click()
    
    ans = MsgBox("ARE YOU SURE YOU WANT TO EXIT?", vbQuestion + vbYesNo, Me.Caption)
    If ans = vbNo Then Exit Sub
    
    Unload Me
End Sub
'Private Sub cmdHistory_Click()
'     frmPriceHist.Show vbModal
'End Sub
Private Sub cmdFind_Click()
On Error GoTo er
        
        
        cmdTrans(0).Enabled = False
        
'        flxSRPPriceHis.Rows = 1
'        flxSRPPriceHis.Rows = 2
'        flxWSPriceHis.Rows = 1
'        flxWSPriceHis.Rows = 2
        
        Call ClearEntryForm
        
        txtSRPMarkUp.Text = ""
        txtSRPProfMar.Text = ""
        txtSRPFixPrice.Text = ""
        txtSRPMaxDis.Text = ""
        
        txtWSMarkUp.Text = ""
        txtWSProfMar.Text = ""
        txtWSFixPrice.Text = ""
        
        txtKAMarkUp.Text = ""
        txtKAProfMar.Text = ""
        txtKAFixPrice.Text = ""
    
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cmdFind_Click")
End Sub

Private Sub cmdHistory_Click(Index As Integer)
    Call frmPriceHist.LoadSRPHist(SPid)
    frmPriceHist.Show vbModal
    
End Sub

Private Sub cmdRefresh_Click()

    If cboSupplier.Text = "" Then SupCode = 0
    
    Screen.MousePointer = vbHourglass
    tmrRefresh.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
    
End Sub
Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
'cmdTrans Legend
'0 ->EDIT
'1 ->SAVE
'2 ->CANCEL

    Dim level As String
    Dim message As String
    
    If Index = 0 Then 'Edit
        ans = MsgBox("Do you want to edit price?", vbYesNo + vbInformation, Me.Caption)
        If ans = vbNo Then Exit Sub
        
        flxrow = flxFind.row
        
        cmdTrans(0).Visible = False
        cmdTrans(1).Visible = True
        cmdTrans(2).Visible = True
                
        txtRefNo.Locked = False
        txtRefNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtRefNo.SetFocus

        txtNetofVat.Locked = False
        txtNetofVat.BackColor = txtHoldWhiteBackColor.BackColor
        
        txtRefDate.Locked = False
        txtRefDate.Text = Format(Now(), "mm / dd / yyyy")
        txtRefDate.Locked = True
        
        txtSRPMaxDis.Locked = False
        txtSRPMaxDis.BackColor = txtHoldWhiteBackColor.BackColor
 
        txtSRPFixPrice.Locked = False
        txtSRPFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
        
        txtWSMaxDiscount.Locked = False
        txtWSMaxDiscount.BackColor = txtHoldWhiteBackColor.BackColor
        
        txtWSFixPrice.Locked = False
        txtWSFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
        
        txtKAMaxDiscount.Locked = False
        txtKAMaxDiscount.BackColor = txtHoldWhiteBackColor.BackColor
        
        txtKAFixPrice.Locked = False
        txtKAFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
        
        cmdExit.Enabled = False
        cmdRefresh.Enabled = False
        cmd0SRP.Enabled = False
        SPpricing.Enabled = False
        
        didRecordEntry = True
        
    ElseIf Index = 1 Then 'SAVE
        Dim maxdis As Integer, wsmaxdis As Integer, kamaxdis As Integer
        Dim SRP As String
        Dim wsp As String
        Dim kap As String
        Dim myRem As String
        
        myRem = ""
        If ValidateEntry() <> "" Then
            MsgBox ValidateEntry, vbExclamation, Me.Caption
            Exit Sub
        End If
        
        If Val(Replace(txtNetofVat.Text, ",", "")) <> 0 Then
            If Val(txtSRPProfMar) <= 5 And Val(Replace(txtNetofVat.Text, ",", "")) <= 99 Then
                ans = MsgBox("SRP:Profit Margin is Less Than 5%. Do you want to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Validation")
                myRem = myRem & "SRP:Profit Margin is Less Than 5% "
                
                If ans = vbNo Then Exit Sub
            ElseIf Val(txtSRPProfMar) <= 10 And Val(Replace(txtNetofVat.Text, ",", "")) >= 100 Then
                ans = MsgBox("SRP:Profit Margin is Less Than 10%. Do you want to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Validation")
                myRem = myRem & "SRP:Profit Margin is Less Than 10% "
                If ans = vbNo Then Exit Sub
            End If
            
            If Val(txtWSProfMar) <= 5 And Val(Replace(txtNetofVat.Text, ",", "")) <= 99 Then
                ans = MsgBox("WSP:Profit Margin is Less Than 5%. Do you want to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Validation")
                myRem = myRem & "WSP:Profit Margin is Less Than 5% "
                If ans = vbNo Then Exit Sub
            ElseIf Val(txtWSProfMar) <= 10 And Val(Replace(txtNetofVat.Text, ",", "")) >= 100 Then
                ans = MsgBox("WSP:Profit Margin is Less Than 10%. Do you want to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Validation")
                myRem = myRem & "WSP:Profit Margin is Less Than 10% "
                If ans = vbNo Then Exit Sub
            End If
            
            If Val(txtKAProfMar) <= 5 And Val(Replace(txtNetofVat.Text, ",", "")) <= 99 Then
                ans = MsgBox("KEYACCTS:Profit Margin is Less Than 5%. Do you want to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Validation")
                myRem = myRem & "KEYACCTS:Profit Margin is Less Than 5% "
                If ans = vbNo Then Exit Sub
            ElseIf Val(txtKAProfMar) <= 10 And Val(Replace(txtNetofVat.Text, ",", "")) >= 100 Then
                ans = MsgBox("KEYACCTS:Profit Margin is Less Than 10%. Do you want to proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Validation")
                myRem = myRem & "KEYACCTS:Profit Margin is Less Than 10% "
                If ans = vbNo Then Exit Sub
            End If
        End If
        
        ans = MsgBox("DO YOU WANT TO SAVE CHANGE?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption)
        If ans = vbNo Then Exit Sub
            
        Screen.MousePointer = vbHourglass
            
        If txtSRPMaxDis.Text <> "" Then
            maxdis = txtSRPMaxDis.Text
        End If
        
        If txtWSMaxDiscount.Text <> "" Then
            wsmaxdis = txtWSMaxDiscount.Text
        End If
        
        If txtKAMaxDiscount.Text <> "" Then
            kamaxdis = txtKAMaxDiscount.Text
        End If
        
        If Val(Replace(txtSRPFixPrice.Text, ",", "")) < 1 Then ' if no srp
            SRP = Val(Replace(txtSRP.Text, ",", ""))
        Else
            SRP = Val(Replace(txtSRPFixPrice.Text, ",", ""))
        End If
        
        If Val(Replace(txtWSFixPrice.Text, ",", "")) < 1 Then 'If no wsp
            wsp = Val(Replace(txtWS.Text, ",", ""))
        Else
            wsp = Val(Replace(txtWSFixPrice.Text, ",", ""))
        End If
        
        If Val(Replace(txtKAFixPrice.Text, ",", "")) < 1 Then 'If no kap
            kap = Val(Replace(txtKA.Text, ",", ""))
        Else
            kap = Val(Replace(txtKAFixPrice.Text, ",", ""))
        End If
        
        Call modSRP.sp_priceupdate_save(Trim(SPid), _
                                        Trim(txtRefNo.Text), _
                                        Val(Replace(txtNetofVat.Text, ",", "")), _
                                        SRP, _
                                        Format(txtSRPProfMar.Text, "#,##0"), _
                                        txtSRPMarkUp.Text & "%", _
                                        Val(Replace(txtSRPMarkUpAmnt.Text, ",", "")), _
                                        wsp, _
                                        Format(txtWSProfMar.Text, "#,##0"), _
                                        txtWSMarkUp.Text & "%", _
                                        Val(Replace(txtWSMarkUpAmnt.Text, ",", "")), _
                                        kap, _
                                        Format(txtKAProfMar.Text, "#,##0"), _
                                        txtKAMarkUp.Text & "%", _
                                        Val(Replace(txtKAMarkUpAmnt.Text, ",", "")), _
                                        maxdis, _
                                        wsmaxdis, _
                                        kamaxdis, _
                                        myRem, _
                                        level, _
                                        message)

       Screen.MousePointer = vbDefault
        
        If Trim(level) = "" Then
            Exit Sub
        ElseIf level = "Granted" Or level = "Warning" Then
            If level = "Granted" Then
                MsgBox message, vbInformation, "SAVE."
            Else
                MsgBox message, vbExclamation, "SAVE!"
                Exit Sub
            End If
        Else
            MsgBox message, vbCritical, "Error!"
            Exit Sub
        End If

        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = False
        cmdTrans(2).Visible = False
        
        cmdTrans(0).Enabled = False
        
        txtSRP.Text = Format(SRP, "#,##0.00")
        txtWS.Text = Format(wsp, "#,##0.00")
        txtKA.Text = Format(kap, "#,##0.00")
        
        flxFind.TextMatrix(flxrow, 4) = Format(SRP, "#,##0.00")
        flxFind.TextMatrix(flxrow, 5) = Format(wsp, "#,##0.00")
        flxFind.TextMatrix(flxrow, 6) = Format(kap, "#,##0.00")
        
        
        Call LockedEntryObject(False)
        
        DoEvents
'        Call LoadWSHist(SPid)
'        DoEvents
'        Call LoadSRPHist(SPid)
        
        cmdExit.Enabled = True
        cmdRefresh.Enabled = True
        cmd0SRP.Enabled = True
        SPpricing.Enabled = True
        
        didRecordEntry = False
    
    ElseIf Index = 2 Then 'CANCEL
        ans = MsgBox("Cancel the currrent operation?", vbYesNo + vbQuestion, Me.Caption)
        If ans = vbNo Then Exit Sub

        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = False
        cmdTrans(2).Visible = False
        cmdTrans(0).Enabled = False
    
        Call ClearEntryForm
        Call LockedEntryObject(False)
        
        cmdExit.Enabled = True
        cmdRefresh.Enabled = True
        cmd0SRP.Enabled = True
        SPpricing.Enabled = True
        
        didRecordEntry = False
        flxrow = 0
        
    End If

Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans_Click()")
End Sub
Function ValidateEntry() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If txtRefNo.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field ref. no."
    If txtSRPFixPrice.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field SRP fix price."
    If txtWSFixPrice.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field WS fix price."
    If txtKAFixPrice.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field KeyAccts fix price."
              
    ValidateEntry = vMsg
    
End Function
Private Sub flxFind_Click()

    If didRecordEntry = True Then Exit Sub
    
    If flxFind.TextMatrix(flxFind.row, 1) = "" Then Exit Sub

    'flxFind.HighLight = flexHighlightWithFocus
    
    Screen.MousePointer = vbHourglass
    
    txtRefNo.Text = ""
    
'0       1             2            3          5       6         7          8        9        10              11              12          14              13              15                  16
'"|<PART NUMBER |<DESCRIPTION  |^NET OF VAT  |>SRP  |>WSP  |>KEYACCTS   |>QTY. |>SIX MSQ |>BACK ORDER |>LAST DATE IN |<APPLICABLE MODEL |SPAREPARTS |MAX DISCOUNT   |WS MAX DISCOUNT   |KEYACCTS MAX DISCOUNT"
    
    txtRefDate.Text = Format(Now(), "mm/dd/YYYY")
    txtPartNo.Text = flxFind.TextMatrix(flxFind.row, 1)
    txtDesc.Text = flxFind.TextMatrix(flxFind.row, 2)
    txtNetofVat.Text = Format(flxFind.TextMatrix(flxFind.row, 3), "#,##0.00")
    txtSRP.Text = Format(flxFind.TextMatrix(flxFind.row, 5), "#,##0.00")
    txtWS.Text = Format(flxFind.TextMatrix(flxFind.row, 6), "#,##0.00")
    txtKA.Text = Format(flxFind.TextMatrix(flxFind.row, 7), "#,##0.00")
    SPid = flxFind.TextMatrix(flxFind.row, 14)
    txtSRPMaxDis.Text = flxFind.TextMatrix(flxFind.row, 13)
    txtWSMaxDiscount.Text = flxFind.TextMatrix(flxFind.row, 15)
    txtKAMaxDiscount.Text = flxFind.TextMatrix(flxFind.row, 16)
    
    txtSRPMarkUp.Text = "0"
    txtSRPProfMar.Text = "0"
    txtSRPFixPrice.Text = "0"
    txtSRPMarkUpAmnt.Text = "0"
    
    txtWSMarkUp.Text = "0"
    txtWSProfMar.Text = "0"
    txtWSFixPrice.Text = "0"
    txtWSMarkUpAmnt.Text = "0"
    
    txtKAMarkUp.Text = "0"
    txtKAProfMar.Text = "0"
    txtKAFixPrice.Text = "0"
    txtKAMarkUpAmnt.Text = "0"
    
    fRow = 1
    fCol = 1
    
'    DoEvents
'    Call LoadSRPHist(SPid)
'    DoEvents
'    Call LoadWSHist(SPid)
    
    DoEvents
    Call getSRPProfMarginRatio
    Call getWSProfMarRatio
    Call getKAProfMarRatio
    
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0

    fRow = i + 1
    fCol = 2
    
    cmdTrans(0).Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub flxFind_DblClick()
    Screen.MousePointer = vbHourglass
    tmrSMSQTYLSI.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
End Sub
Private Sub flxFind_EnterCell()
    Call flxFind_Click
End Sub

Private Sub flxFind_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If didRecordEntry = False Then Exit Sub
    If Button = vbRightButton Then
        'PopupMenu mnuPopUpFormat, vbPopupMenuRightButton
    End If
End Sub
Private Sub Form_Load()
On Error GoTo er
Dim tmpRs As ADODB.Recordset

    DoEvents
    Call modSRP.sp_supplier_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Call modObjTools.fillArray(AllArrSupplier, tmpRs)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboSupplier, arrSupplier, AllArrSupplier, 1, 1)
    
    DoEvents
    Call modSRP.sp_class_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Call modObjTools.fillArray(AllArrClass, tmpRs)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close

    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboType, arrClass, AllArrClass, 0, 0)
    
    
    
    Warehouse.sBar.Panels(6).Text = "Price Adjustment"
    pSupCode = ""
    
    DoEvents
    Screen.MousePointer = vbHourglass
    tmrRefresh.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal

    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->Form_Load()")
End Sub
Private Sub mnuBSMSQTY_Click()
    frmBSMSQTY.LoadRecord (SPid)
    frmBSMSQTY.Show vbModal
End Sub

Private Sub RXResizer1_ResizeComplete()
On Error Resume Next
       
       'If flxFind.TextMatrix(1, 1) <> "" Then Exit Sub
       
        DoEvents
        With flxFind
            '.Clear
            '.Rows = 2
             TCol = 17
            .Cols = TCol
            .FixedCols = 1
            
            '               0       1             2            3           4             5       6        7            8              9        10           11          12
            .FormatString = "|<PART NUMBER |<DESCRIPTION |^NET OF VAT |^COST WITH VAT  |>SRP  |>WSP  |>KEYACCTS  |>QTY ON HAND |>SIX MSQ |>BACK ORDER |>LAST DATE IN |<APPLICABLE MODEL"
            
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.1  'PART NUMBER
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.28  'DESCRIPTION
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.07  'NET OF VAT
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.07  'COST W/VAT
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.07  'SRP
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.07  'WSP
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.07  'KAP
            
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.07  'QTY ON HAND
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.07  'SIX MSQ
            .ColWidth(10) = (.Width - .ColWidth(0)) * 0.07  'BACKORDER
            .ColWidth(11) = (.Width - .ColWidth(0)) * 0.07  'LAST DATE IN
            .ColWidth(12) = (.Width - .ColWidth(0)) * 0.1 'APPLICABLE MODEL
            .ColWidth(13) = 0 'MAX DISCOUNT
            .ColWidth(14) = 0 'SPAREPARTID
            .ColWidth(15) = 0 'WS MAX DISCOUNT
            .ColWidth(16) = 0 'KEYACCTS MAX DISCOUNT
        End With
        
    
End Sub
Private Sub tmr0SRP_Timer()
' On Error GoTo er

    tmr0SRP.Enabled = False
    Screen.MousePointer = vbHourglass
    frmPriceDiscountv2.Enabled = False
    
    DoEvents
    DoEvents
    Call FindZeroSRP
    DoEvents
    DoEvents
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    frmPriceDiscountv2.Enabled = True
    Screen.MousePointer = vbDefault
    
    'flxFind.HighLight = flexHighlightNever
    
    didRecordEntry = False
    
Exit Sub

End Sub
Private Sub tmrSMSQTYLSI_Timer()
On Error GoTo er

    tmrSMSQTYLSI.Enabled = False
    Screen.MousePointer = vbHourglass
    frmPriceDiscountv2.Enabled = False
    
    DoEvents
    Call GetSMSQTYLDI(SPid)
    
    DoEvents
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    frmPriceDiscountv2.Enabled = True
    Screen.MousePointer = vbDefault
    
    'flxFind.HighLight = flexHighlightNever
    
    didRecordEntry = False
    
Exit Sub
er:
    frmPriceDiscountv2.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrSMSQTYLSI_Timer()")
End Sub
Private Sub tmrRefresh_Timer()
On Error GoTo er

    tmrRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    frmPriceDiscountv2.Enabled = False
    
    DoEvents
    DoEvents
    'Call LoadRecordBySup(SupCode)
    If lblFindParts.Caption = "FIND PART DESCRIPTION:" Then
        Call FindDesc
    Else
        Call FindPart
    End If
    DoEvents
    
    DoEvents
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    frmPriceDiscountv2.Enabled = True
    Screen.MousePointer = vbDefault
    
    flxFind.Highlight = flexHighlightAlways
    
    didRecordEntry = False
    
Exit Sub
er:
    frmPriceDiscountv2.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrRefresh_Timer()")
End Sub
Private Sub txtFindPart_Change()
On Error GoTo er
    Dim Qry As String
    
    If txtFindPart.Text = "" Then Exit Sub
    If lblFindParts.Caption = "FIND PART DESCRIPTION" Then Exit Sub
    
    If bNone Then
        bNone = False
        Exit Sub
    End If
    
    If rsParts.State = adStateOpen Then rsParts.Close
    
    Qry = "SELECT PARTNUMBER FROM mrcwarehouse.spareparts WHERE PARTNUMBER LIKE '" + txtFindPart.Text + "%';"
    'InputBox Qry, Qry, Qry
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic

    If rsParts.RecordCount > 0 Then
        If Not bNone Then
            txtTag.Text = Len(txtFindPart.Text)
            bNone = True
        End If
        
        txtFindPart.Text = rsParts!PartNumber
        txtFindPart.SelStart = Val(txtTag.Text)
        txtFindPart.SelLength = Len(txtFindPart.Text)
    End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtFindPart_Change()")
End Sub
Private Sub txtFindPart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        bNone = True
    End If
End Sub
Private Sub txtFindPart_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub
Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Check1(0).SetFocus
End Sub
Private Sub txtSRPFixPrice_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPFixPrice, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
            Dim ProMar As Double
            Dim CostVat As Double
            Dim NewSRP As Double
            Dim MarkUpPerc As Double
            Dim AmntMarkUp As Double
            
            NewSRP = Val(Replace(Trim(txtSRPFixPrice.Text), ",", ""))
            CostVat = Val(Replace(txtNetofVat, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00.00")
            
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
'            If NewSRP < CostVat Then
'                MsgBox "New SRP is less than new cost!", vbExclamation, Me.Caption
'                Exit Sub
'            Else
            
            AmntMarkUp = NewSRP - CostVat
            txtSRPMarkUpAmnt.Text = Format(Round(AmntMarkUp), "#,#00.00")
            
            ProMar = (AmntMarkUp / NewSRP) * 100
            txtSRPProfMar.Text = Format(Round(ProMar), "#,#00")
            
            MarkUpPerc = (((CostVat + AmntMarkUp) - CostVat) / CostVat) * 100
            txtSRPMarkUp.Text = Format(MarkUpPerc, "#,#00.00")
    
    End If

Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSRPFixPrice_KeyPress()")
End Sub
Private Sub txtSRPFixPrice_LostFocus()
    If txtSRPFixPrice.Text < 1 Or txtSRPFixPrice.Text = "" Then Exit Sub
    Call txtSRPFixPrice_KeyPress(13)
End Sub
Private Sub txtSRPMarkUp_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPMarkUp, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
            Dim CostVat As Double
            Dim NewSRP As Double
            Dim GrossProfMar As Double
            Dim AmntMarkUp As Double
            Dim MarkUpPerc As Double
            
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat.Text, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00.00")
            
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            AmntMarkUp = Val(Replace(txtSRPMarkUp.Text, ",", "")) / 100
            AmntMarkUp = CostVat * AmntMarkUp
            txtSRPMarkUpAmnt.Text = Format(Round(AmntMarkUp), "#,#00.00")
            
            NewSRP = CostVat + AmntMarkUp
            txtSRPFixPrice.Text = Format(Round(NewSRP), "#,###")
            
            GrossProfMar = AmntMarkUp / NewSRP
            GrossProfMar = GrossProfMar * 100
            txtSRPProfMar.Text = Format(Round(GrossProfMar), "#,#00")

    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSRPMarkUp_KeyPress()")
End Sub
Private Sub txtSRPMarkUp_LostFocus()
    If txtSRPMarkUp.Text < 1 Or txtSRPMarkUp.Text = "" Then Exit Sub
    Call txtSRPMarkUp_KeyPress(13)
End Sub
Private Sub txtSRPMaxDis_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPMaxDis, KeyAscii, NumbersOnly)
End Sub
Private Sub txtSRPProfMar_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPProfMar, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim CostVat As Double
            Dim NewSRP As Double
            Dim GrossProfMar As Double
            Dim AmntMarkUp As Double
            Dim MarkUpPerc As Double
            
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat.Text, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00.00")
            
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            NewSRP = txtSRPProfMar.Text / 100
            NewSRP = CostVat + NewSRP
            txtSRPFixPrice.Text = Format(Round(NewSRP), "#,#00")

            AmntMarkUp = NewSRP - CostVat
            txtSRPMarkUpAmnt.Text = Format(Round(AmntMarkUp), "#,#00.00")
            
            GrossProfMar = AmntMarkUp / NewSRP
            GrossProfMar = GrossProfMar * 100
            txtSRPProfMar.Text = Format(Round(GrossProfMar), "#,#00")
            
            MarkUpPerc = AmntMarkUp / CostVat
            MarkUpPerc = MarkUpPerc * 100
            txtSRPMarkUp.Text = Format(MarkUpPerc, "#,#00.00")
            
    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSRPProfMar_KeyPress()")
End Sub
Private Sub txtSRPProfMar_LostFocus()
    If txtSRPProfMar.Text < 1 Or txtSRPProfMar.Text = "" Then Exit Sub
    Call txtSRPProfMar_KeyPress(13)
End Sub
Private Sub txtWSFixPrice_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtWSFixPrice, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim ProMar As Double
            Dim CostVat As Double
            Dim NewWSP As Double
            Dim MarkUpPerc As Double
            Dim AmntMarkUp As Double
            
            
            NewWSP = Val(Replace(Trim(txtWSFixPrice.Text), ",", ""))
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00")
            
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
'            If NewWSP < CostVat Then
'                MsgBox "New WSP is less than new cost!", vbExclamation, Me.Caption
'                Exit Sub
'            Else
            AmntMarkUp = NewWSP - CostVat
            txtWSMarkUpAmnt.Text = Format(AmntMarkUp, "#,#00.00")
            
            ProMar = (AmntMarkUp / NewWSP) * 100
            txtWSProfMar.Text = Format(ProMar, "#,#00")
            
            MarkUpPerc = (((CostVat + AmntMarkUp) - CostVat) / CostVat) * 100
            txtWSProfMar.Text = Format(MarkUpPerc, "#,#00")
    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtWSFixPrice_KeyPress()")
End Sub
Private Sub txtWSFixPrice_LostFocus()
    If txtWSFixPrice.Text < 1 Or txtWSFixPrice.Text = "" Then Exit Sub
    Call txtWSFixPrice_KeyPress(13)
End Sub
Private Sub txtWSMarkUp_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtWSMarkUp, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim CostVat As Double
            Dim NewWSP As Double
            Dim GrossProfMar As Double
            Dim AmntMarkUp As Double
            Dim MarkUpPerc As Double
            
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat.Text, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00.00")
            
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            AmntMarkUp = Val(Replace(txtWSMarkUp.Text, ",", "")) / 100
            AmntMarkUp = CostVat * AmntMarkUp
            txtWSMarkUpAmnt.Text = Format(AmntMarkUp, "#,#00.00")
            
            NewWSP = CostVat + AmntMarkUp
            txtWSFixPrice.Text = Format(NewWSP, "#,###.00")
            
            GrossProfMar = AmntMarkUp / NewWSP
            GrossProfMar = GrossProfMar * 100
            txtWSProfMar.Text = Format(GrossProfMar, "#,#00")

    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtWSMarkUp_KeyPress()")
End Sub
Private Sub txtWSMarkUp_LostFocus()
    If txtWSMarkUp.Text < 1 Or txtWSMarkUp.Text = "" Then Exit Sub
    Call txtWSMarkUp_KeyPress(13)
End Sub
Private Sub txtWSProfMar_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtWSProfMar, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim CostVat As Double
            Dim NewWSP As Double
            Dim Ratio As Double
            
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat.Text, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00")
    
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            NewWSP = CostVat + Val(Replace(txtWSProfMar.Text, ",", ""))
            
            'Ratio = txtWSProfMar.Text / NewWSP
            Ratio = Val(Replace(txtWSProfMar.Text, ",", "")) / CostVat
            Ratio = Ratio * 100
            
            txtWSFixPrice.Text = Format(NewWSP, "#,#00")
            txtWSProfMar.Text = Format(txtWSProfMar, "#,#00")
            
'            If Ratio > 100 Then
'                Ratio = 100
'            End If
            
          '  txtWSProfRatio.Text = Format(Ratio, "#,#00.00")
    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtWSProfMar_KeyPress()")
End Sub
Private Sub txtWSProfMar_LostFocus()
    If txtWSProfMar.Text < 1 Or txtWSProfMar.Text = "" Then Exit Sub
    Call txtWSProfMar_KeyPress(13)
End Sub
Private Sub txtKAFixPrice_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtKAFixPrice, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim ProMar As Double
            Dim CostVat As Double
            Dim NewKAP As Double
            Dim MarkUpPerc As Double
            Dim AmntMarkUp As Double
            
            
            NewKAP = Val(Replace(Trim(txtKAFixPrice.Text), ",", ""))
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00")
            
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
'            If NewWSP < CostVat Then
'                MsgBox "New WSP is less than new cost!", vbExclamation, Me.Caption
'                Exit Sub
'            Else
            AmntMarkUp = NewKAP - CostVat
            txtKAMarkUpAmnt.Text = Format(AmntMarkUp, "#,#00.00")
            
            ProMar = (AmntMarkUp / NewKAP) * 100
            txtKAProfMar.Text = Format(ProMar, "#,#00")
            
            MarkUpPerc = (((CostVat + AmntMarkUp) - CostVat) / CostVat) * 100
            txtKAProfMar.Text = Format(MarkUpPerc, "#,#00")
    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtKAFixPrice_KeyPress()")
End Sub
Private Sub txtKAFixPrice_LostFocus()
    If txtKAFixPrice.Text < 1 Or txtKAFixPrice.Text = "" Then Exit Sub
    Call txtKAFixPrice_KeyPress(13)
End Sub
Private Sub txtKAProfMar_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtKAProfMar, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim CostVat As Double
            Dim NewKAP As Double
            Dim Ratio As Double
            
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat.Text, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00")
    
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            NewKAP = CostVat + Val(Replace(txtKAProfMar.Text, ",", ""))
            
            'Ratio = txtWSProfMar.Text / NewWSP
            Ratio = Val(Replace(txtKAProfMar.Text, ",", "")) / CostVat
            Ratio = Ratio * 100
            
            txtKAFixPrice.Text = Format(NewKAP, "#,#00")
            txtKAProfMar.Text = Format(txtKAProfMar, "#,#00")
            
'            If Ratio > 100 Then
'                Ratio = 100
'            End If
            
          '  txtWSProfRatio.Text = Format(Ratio, "#,#00.00")
    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtKAProfMar_KeyPress()")
End Sub
Private Sub txtKAProfMar_LostFocus()
    If txtKAProfMar.Text < 1 Or txtKAProfMar.Text = "" Then Exit Sub
    Call txtKAProfMar_KeyPress(13)
End Sub
Private Sub txtKAMarkUp_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtWSMarkUp, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim CostVat As Double
            Dim NewKAP As Double
            Dim GrossProfMar As Double
            Dim AmntMarkUp As Double
            Dim MarkUpPerc As Double
            
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat.Text, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00.00")
            
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            AmntMarkUp = Val(Replace(txtKAMarkUp.Text, ",", "")) / 100
            AmntMarkUp = CostVat * AmntMarkUp
            txtKAMarkUpAmnt.Text = Format(AmntMarkUp, "#,#00.00")
            
            NewKAP = CostVat + AmntMarkUp
            txtKAFixPrice.Text = Format(NewKAP, "#,###.00")
            
            GrossProfMar = AmntMarkUp / NewKAP
            GrossProfMar = GrossProfMar * 100
            txtKAProfMar.Text = Format(GrossProfMar, "#,#00")

    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtKAMarkUp_KeyPress()")
End Sub
Private Sub txtKAMarkUp_LostFocus()
    If txtKAMarkUp.Text < 1 Or txtKAMarkUp.Text = "" Then Exit Sub
    Call txtKAMarkUp_KeyPress(13)
End Sub
Private Sub VScroll1_Change()
    Dim tmpRs As ADODB.Recordset
    VScroll1.Max = 2
    
    If VScroll1.Value = 0 Then
        lblFindParts.Caption = "FIND PART DESCRIPTION:"
        
    ElseIf VScroll1.Value = 1 Then
        lblFindParts.Caption = "FIND PART NUMBER:"
        
    ElseIf VScroll1.Value = 2 Then
        VScroll1.Value = 0
    End If
    
End Sub
Sub flxFindColsChanger()
On Error Resume Next
    With flxFind
        .Clear
        .Rows = 2
         TCol = 12
        .Cols = TCol
        .FixedCols = 1
        
        '               0       1             2            3           4      5       6             7        8              9       10
        .FormatString = "|<PART NUMBER |<DESCRIPTION  |^NET OF VAT  |>SRP  |>WSP  |>KEYACCTS    |>QTY. |>SIX MSQ |>BACK ORDER |<APPLICABLE MODEL"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15  'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.27  'DESCRIPTION
        
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.08  'NET OF VAT
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.08  'SRP
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.08  'WSP
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.08  'KAP
        
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.07  'QTY.
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.07  'SIX MSQ
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.07  'BACKORDER
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.1   'APPLICABLE MODEL
        .ColWidth(11) = 0 'SPAREPARTID
        .ColWidth(12) = 0 'MAX DISCOUNT
    End With
    
End Sub
Public Sub FindDesc()
On Error GoTo er
    Dim MySP() As String
    Dim Qry As String
    Dim myW As String, MySel As String
    

    MySP = Split(txtFindPart.Text, " ")
    myW = ""
    For i = LBound(MySP) To UBound(MySP)
        If myW = "" Then
            myW = " sp.DESCRIPTION Like '%" & MySP(i) & "%'"
        Else
            myW = myW & " AND sp.DESCRIPTION Like '%" & MySP(i) & "%'"
        End If
    Next i
    
    If lblFindParts.Caption = "FIND PART NUMBER:" Then
        txtFindPart.Text = ""
        VScroll1.Value = 0
    End If
    
    If rsParts.State = adStateOpen Then rsParts.Close
    

    If txtFindPart.Text = "" And chkAll.Value = 0 And SupCode = "" And cboType.Text = "" Then Exit Sub
    
    flxFind.Clear
    
'    If chkSup = True Then
'        SupCode = pSupCode
'    End If
    If txtFindPart.Text <> "" And chkAll.Value = 0 And cboSupplier.Text <> "" And cboType.Text = "" Then
    
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS', sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT, sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp WHERE sp.SUPPLIERCODE= '" & SupCode & "' AND " & myW & " ORDER BY sp.PARTNUMBER;"

        
    ElseIf chkAll.Value = 0 And txtFindPart.Text = "" And cboSupplier.Text <> "" And cboType.Text = "" Then
    
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp WHERE sp.SUPPLIERCODE= '" & SupCode & "' ORDER BY sp.PARTNUMBER;"

            
    ElseIf chkAll.Value = 1 And txtFindPart.Text <> "" Then
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp WHERE " & myW & " ORDER BY sp.PARTNUMBER;"

    ElseIf chkAll.Value = 1 And txtFindPart.Text = "" And cboSupplier.Text = "" And cboType.Text = "" Then
    
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp ORDER BY sp.PARTNUMBER limit 50;"

            
    'InputBox Qry, Qry, Qry
    ElseIf txtFindPart.Text <> "" And chkAll.Value = 0 And SupCode = "" Then
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp WHERE " & myW & " ORDER BY sp.PARTNUMBER;"

            
    ElseIf txtFindPart.Text <> "" And chkAll.Value = 0 And cboSupplier.Text = "" Then
    
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp WHERE " & myW & " ORDER BY sp.PARTNUMBER;"

     
            
    ElseIf txtFindPart.Text = "" And chkAll.Value = 0 And cboSupplier.Text = "" And cboType.Text = "" Then
    
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp ORDER BY sp.PARTNUMBER limit 30;"

         
            
    ElseIf txtFindPart.Text = "" And chkAll.Value = 1 And cboSupplier.Text = "" And cboType.Text <> "" Then
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"

         
        
    ElseIf txtFindPart.Text <> "" And chkAll.Value = 0 And cboSupplier.Text <> "" And cboType.Text <> "" Then
  Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"

         

    ElseIf chkAll.Value = 0 And cboSupplier.Text <> "" And cboType.Text <> "" Then
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp WHERE sp.SUPPLIERCODE= '" & SupCode & "' AND remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"

      
            
    ElseIf txtFindPart.Text <> "" And chkAll.Value = 1 And cboSupplier.Text = "" And cboType.Text <> "" Then
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp WHERE " & myW & " AND remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"

       
        
    ElseIf txtFindPart.Text = "" And chkAll.Value = 0 And cboSupplier.Text = "" And cboType.Text <> "" Then
    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
            "From SPAREPARTS as sp WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"

         
            
    End If
   
    If Qry = "" Then Exit Sub
    
   'InputBox Qry, Qry, Qry
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxFind.Clear
        flxFind.Rows = 2
        Call RXResizer1_ResizeComplete
    Else
        Set flxFind.DataSource = rsParts
        flxFind.Refresh
        flxFind.ColAlignment(3) = flexAlignRightCenter
        flxFind.ColAlignment(4) = flexAlignRightCenter
        flxFind.ColAlignment(5) = flexAlignRightCenter
        Call RXResizer1_ResizeComplete
        Call flxFind_Click
    End If
    
    Call FormatCur
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->FindDesc()")
End Sub
Public Sub FindPart()
On Error GoTo er
    Dim Qry As String
    Dim myW As String, MySel As String
    
    
    If rsParts.State = adStateOpen Then rsParts.Close
    If txtFindPart.Text = "" Then Exit Sub
    
    flxFind.Clear

    Qry = "SELECT sp.PARTNUMBER,Ifnull(TRIM(sp.DESCRIPTION), '') AS DESCRIPTION" & _
            ",sp.COST as 'NET OF VAT',(sp.COST*1.12) as 'COST WITH VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.KEYACCTS as 'KEYACCTS',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
        ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID,sp.WSMAXDISCOUNT,sp.KAMAXDISCOUNT " & _
        "From SPAREPARTS as sp WHERE sp.PARTNUMBER LIKE '%" & Trim(txtFindPart.Text) & "%' ORDER BY sp.PARTNUMBER;"
   
     
   
    If Qry = "" Then Exit Sub
    
   'InputBox Qry, Qry, Qry
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxFind.Clear
        flxFind.Rows = 2
        flxFind.FormatString = "|<PART NUMBER |<DESCRIPTION |^NET OF VAT |^COST WITH VAT |>SRP  |>WSP  |>KEYACCTS   |>QTY. |>SIX MSQ |>BACK ORDER |>LAST DATE IN |<APPLICABLE MODEL"
        Call flxFindColsChanger
    Else
        Set flxFind.DataSource = rsParts
        flxFind.Refresh
        flxFind.ColAlignment(3) = flexAlignRightCenter
        flxFind.ColAlignment(4) = flexAlignRightCenter
        flxFind.ColAlignment(5) = flexAlignRightCenter
        Call flxFind_Click
    End If
    
    Call FormatCur
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->FindPart()")
End Sub
Sub FormatCur()
    Dim rw As Integer
    For rw = 1 To flxFind.Rows - 1
        flxFind.TextMatrix(rw, 3) = Format(flxFind.TextMatrix(rw, 3), "##,###,###.#0")
        flxFind.TextMatrix(rw, 4) = Format(flxFind.TextMatrix(rw, 4), "##,###,###.#0")
        flxFind.TextMatrix(rw, 5) = Format(flxFind.TextMatrix(rw, 5), "##,###,###.#0")
    Next rw
End Sub
Sub ClearEntryForm()
   
    txtRefDate.Text = ""
    txtNetofVat.Text = ""
   
    txtPartNo.Text = ""
    txtDesc.Text = ""
    
    txtSRP.Text = ""
    txtSRPMarkUp.Text = ""
    txtSRPProfMar.Text = ""
    txtSRPFixPrice.Text = ""
    txtSRPMaxDis.Text = ""
    txtSRPMarkUpAmnt.Text = ""
    
    txtWSMarkUp.Text = ""
    txtWSProfMar.Text = ""
    txtWS.Text = ""
    txtWSFixPrice.Text = ""
    txtWSMarkUpAmnt.Text = ""
    
    txtKAMarkUp.Text = ""
    txtKAProfMar.Text = ""
    txtKA.Text = ""
    txtKAFixPrice.Text = ""
    txtKAMarkUpAmnt.Text = ""
    
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
                              
End Sub
Sub LockedEntryObject(DataEntry As Boolean)

    DataEntry = Not DataEntry
        
        txtRefNo.Locked = DataEntry
        txtSRPMaxDis.Locked = DataEntry
        txtRefDate.Locked = DataEntry
        txtNetofVat.Locked = DataEntry
        txtPartNo.Locked = DataEntry
        txtDesc.Locked = DataEntry
        txtSRP.Locked = DataEntry
        txtSRPProfMar.Locked = DataEntry
        txtSRPMarkUp.Locked = DataEntry
        txtSRPFixPrice.Locked = DataEntry
        txtSRPMaxDis.Locked = DataEntry
        txtWSProfMar.Locked = DataEntry
        txtWSMarkUp.Locked = DataEntry
        txtWSFixPrice.Locked = DataEntry
        txtWSMaxDiscount.Locked = DataEntry
        txtKAProfMar.Locked = DataEntry
        txtKAMarkUp.Locked = DataEntry
        txtKAFixPrice.Locked = DataEntry
        txtKAMaxDiscount.Locked = DataEntry
        
    If txtRefNo.Locked Then
        
        txtRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldGrayBackColor.BackColor
        txtRefDate.BackColor = txtHoldGrayBackColor.BackColor
        txtNetofVat.BackColor = txtHoldGrayBackColor.BackColor
        txtPartNo.BackColor = txtHoldGrayBackColor.BackColor
        txtDesc.BackColor = txtHoldGrayBackColor.BackColor
        txtSRP.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPProfMar.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPMarkUp.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldGrayBackColor.BackColor
        txtWSProfMar.BackColor = txtHoldGrayBackColor.BackColor
        txtWSMarkUp.BackColor = txtHoldGrayBackColor.BackColor
        txtWSFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        txtWSMaxDiscount.BackColor = txtHoldGrayBackColor.BackColor
        txtKAProfMar.BackColor = txtHoldGrayBackColor.BackColor
        txtKAMarkUp.BackColor = txtHoldGrayBackColor.BackColor
        txtKAFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        txtKAMaxDiscount.BackColor = txtHoldGrayBackColor.BackColor
    Else
        
        txtRefNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldWhiteBackColor.BackColor
        txtRefDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtNetofVat.BackColor = txtHoldWhiteBackColor.BackColor
        txtPartNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtDesc.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRP.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPProfMar.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPMarkUp.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldWhiteBackColor.BackColor
        txtWSProfMar.BackColor = txtHoldWhiteBackColor.BackColor
        txtWSMarkUp.BackColor = txtHoldWhiteBackColor.BackColor
        txtWSFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
        txtWSMaxDiscount.BackColor = txtHoldWhiteBackColor.BackColor
        txtKAProfMar.BackColor = txtHoldWhiteBackColor.BackColor
        txtKAMarkUp.BackColor = txtHoldWhiteBackColor.BackColor
        txtKAFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
        txtKAMaxDiscount.BackColor = txtHoldWhiteBackColor.BackColor
    End If

End Sub

