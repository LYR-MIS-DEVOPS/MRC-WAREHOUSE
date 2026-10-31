VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm Warehouse 
   BackColor       =   &H008080FF&
   Caption         =   "MRC Warehouse System"
   ClientHeight    =   9750
   ClientLeft      =   225
   ClientTop       =   -570
   ClientWidth     =   15465
   Icon            =   "WarehouseMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "WarehouseMDI.frx":0A02
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cDlgBkUp 
      Left            =   960
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9375
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "WarehouseMDI.frx":190F8
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "User Name"
            TextSave        =   "User Name"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "WarehouseMDI.frx":19CEA
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1940
            MinWidth        =   1940
            Text            =   "Transactions"
            TextSave        =   "Transactions"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13282
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "WarehouseMDI.frx":1A5C4
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "Date"
            TextSave        =   "Date"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuAcct 
      Caption         =   "&Account Settings"
      Begin VB.Menu mnuEmp 
         Caption         =   "Employees"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUser 
         Caption         =   "User Profile"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCustomer 
         Caption         =   "Customer"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSupplier 
         Caption         =   "Supplier"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBranch 
         Caption         =   "Branch"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu menuInventory 
      Caption         =   "&Inventory"
      Begin VB.Menu MenuOrdering 
         Caption         =   "Ordering"
         Begin VB.Menu mnuPOsupplier 
            Caption         =   "P.O. to Supplier"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPOBranch 
            Caption         =   "P.O. from Branch"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuBBOPO 
            Caption         =   "Branch Back Order P.O."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBBOsumm 
            Caption         =   "Branch Back Order Summary"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSBOsumm 
            Caption         =   "Supplier Back Order Summary"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MenuRecStocks 
         Caption         =   "Receiving Stocks"
         Begin VB.Menu mnuReceiving 
            Caption         =   "Receiving Receipts"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSalesReturn 
            Caption         =   "Sales Return"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuAllocBBO 
            Caption         =   "Allocation Branch Back Order"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MenuOutStocks 
         Caption         =   "Outgoing Stocks"
         Begin VB.Menu mnuStockTransfer 
            Caption         =   "Stock Transfer"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPurchaseReturn 
            Caption         =   "Purchase Return"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MenuMonitoring 
         Caption         =   "Monitoring"
         Begin VB.Menu mnuMerchandise 
            Caption         =   "Merchandise Inventory"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCardex 
            Caption         =   "Cardex"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCardexRed 
            Caption         =   "Cardex (RED)"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuMonthlyInv 
            Caption         =   "Monthly Merchandise Inventory"
         End
         Begin VB.Menu mnuBackOrder 
            Caption         =   "Back Order"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuBranchRecDis 
            Caption         =   "Branch Receiving Discrepancy"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MenuAdjustment 
         Caption         =   "Adjustments"
         Begin VB.Menu mnuPrice 
            Caption         =   "Price Adjustment"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPartsProp 
            Caption         =   "Parts Properties"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuInvAdj 
            Caption         =   "Inventory Adjustment"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu MenuFinance 
      Caption         =   "&Financing"
      Begin VB.Menu MenuCash 
         Caption         =   "Cash"
         Begin VB.Menu mnuReceipts 
            Caption         =   "Receipts"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPr 
            Caption         =   "Provisional"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuChecks 
            Caption         =   "Checks"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCashflow 
            Caption         =   "Cashflow"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuAFL 
            Caption         =   "Adv. for Liquidation"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MenuAR 
         Caption         =   "Accounts Receivable"
         Begin VB.Menu mnuStatement 
            Caption         =   "Statement of Accounts"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuAging 
            Caption         =   "Aging of Accounts"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDMCM 
            Caption         =   "DMCM"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MenuAP 
         Caption         =   "Accounts Payable"
         Begin VB.Menu mnuBilling 
            Caption         =   "Billing"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MenuDisbursement 
         Caption         =   "Disbursement"
         Begin VB.Menu mnuPCV 
            Caption         =   "Petty Cash Vouchers"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCV 
            Caption         =   "Cash Voucher"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu MenuReports 
      Caption         =   "&Reports"
      Begin VB.Menu MenuSalesCollection 
         Caption         =   "Sales & Collection"
         Begin VB.Menu mnuDailySales 
            Caption         =   "Daily Sales"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDailyCash 
            Caption         =   "Daily Cash"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDailyExpense 
            Caption         =   "Daily Expenses"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuMonSales 
            Caption         =   "Monthly Sales"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuMonColl 
            Caption         =   "Monthly Collection"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuMonExpense 
            Caption         =   "Monthly Expenses"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MenuInvReport 
         Caption         =   "Inventory"
         Begin VB.Menu mnuReceivingSum 
            Caption         =   "Detailed Receiving Report"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTrasferSum 
            Caption         =   "Detailed Outgoing Report"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSaleSellPrice 
            Caption         =   "Sales Report at Selling Price"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSalesCostPrice 
            Caption         =   "Sales Report at Cost Price"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSalRetSum 
            Caption         =   "Sales Return Summary"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPurchRetSum 
            Caption         =   "Purchase Return Summary"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuInvAdjSum 
            Caption         =   "Inventory Adjustment Summary"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSalesDistrib 
            Caption         =   "Sales Distribution Report"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSalesDB 
            Caption         =   "Sales DB for Extract"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu MenuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuScripting 
         Caption         =   "Scripting"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintPref 
         Caption         =   "Printing Preferences"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInvAudit 
         Caption         =   "Inventory Audit"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuQuit 
      Caption         =   "&Quit"
   End
End
Attribute VB_Name = "Warehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Warehouse.Caption = SysBranch & " v17.02.13"
    '" v15.12.21" = Sorting by partnumber ang transfer og picking requested by mam jean
    'v16.01.29 = qty on hand vs. allocated qty sa BPO
    'v16.04.16 = detailed outgoing summary
    'v17.01.07 = online po for mcparts depot rizal & binos tagum
    'V17.02.13 = PO with Intransit, PO exclude 3S & include Cabantian, MCPart Depot Rizal, Back to Old Pricing.
   
    FrDate = DateValue("1/1/1977")
    ToDate = DateValue("1/1/1977")
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Ans As Integer, i As Integer
    Dim MyForms As String
    
    MyForms = ""
    
    For i = 0 To Forms.Count - 1
        MyForms = Forms(i).Caption
    Next i
    If MyForms <> SysBranch Then
        MsgBox "Please Close the following forms:" & vbCrLf & MyForms, vbExclamation + vbOKOnly, "Warning"
        Cancel = True
        Exit Sub
    End If
    
    Ans = MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Quit")
End Sub


Private Sub mnuStockTrans_Click()
    
    frmSPtrans.Show
      
End Sub





Private Sub mnuWhouse_Click()
    SPWarehouse.Show
End Sub

Private Sub MenuQuit_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Quit?", vbYesNo, "Close Program")
    If Ans = vbNo Then Exit Sub
    LYRcon.Close
    End

End Sub

Private Sub mnuAFL_Click()
    ExpAdvLiq.Show vbModal
End Sub

Private Sub mnuAging_Click()
    AgingNew.Show
End Sub

Private Sub mnuAllocBBO_Click()
    AllocBBO.Show
End Sub

Private Sub mnuBackOrder_Click()
    SPbckOrdSum.Show
End Sub

Private Sub mnuBackup_Click()
    Dim myFile As String, myFile2 As String
    Dim Qry As String, BranchDB As String

    myFile = Dir("C:\MIS\BACKUP\MYSQLDUMP.EXE")
    myFile2 = Dir("C:\MIS\BACKUP\RAR.EXE")
    If UCase$(myFile) <> "MYSQLDUMP.EXE" Or UCase$(myFile2) <> "RAR.EXE" Then
        MsgBox "Backup File Systems missing, Please consult IT Support Group.", vbExclamation + vbOKOnly, "Warning"
        Exit Sub
    End If
    Qry = "SELECT settings.Value From settings Where settings.DESCRIPTION = 'BACKUP' AND settings.REMARKS = 'FILENAME'"
    myFile = GetFieldValue(Qry)
    With cDlgBkUp
        .CancelError = True
        .Filter = "Backup Files (*.RAR)|(*.RAR)"
        .InitDir = App.Path
        myFile = myFile & Format(Now, "mmddyyhhmm")
        .filename = myFile & ".RAR"
        .ShowSave
        myFile2 = .filename
    End With
    If UCase$(SysBranch) = "MRC WAREHOUSE" Then
        BranchDB = "mrcwarehouse"
    Else
        MsgBox "Backup File Systems missing, Please consult IT Support Group.", vbExclamation + vbOKOnly, "Warning"
        Exit Sub
    End If
    
    Open "Backup.bat" For Output As #1
    Print #1, "@echo off"
    Print #1, "echo Processing your BACKUP for " & myFile & ", Please wait....."
    Print #1, "C:\MIS\BACKUP\mysqldump -hlocalhost -uroot -plaborday " & BranchDB & " > C:\MIS\BACKUP\" & myFile & ".SQL"
    Print #1, "echo This may take several minutes... Please wait...."
    Print #1, "C:\MIS\BACKUP\rar a -plaborday " & "C:\MIS\BACKUP\" & myFile & ".rar C:\MIS\BACKUP\" & myFile & ".SQL"
    Print #1, "COPY C:\MIS\BACKUP\" & myFile & ".RAR " & myFile2
    Print #1, "echo BACKUP is now finished"
    Print #1, "del C:\MIS\BACKUP\" & myFile & ".SQL"
    Print #1, "del Backup.bat"
    Close #1
    Shell "Backup.bat", vbNormalNoFocus

End Sub

Private Sub mnuBBOPO_Click()
    BrnBackOrderPO.Show
End Sub
Private Sub mnuBBOsumm_Click()
    frmHSummary.MySource = "BRANCH BACK ORDER"
    frmHSummary.Show
End Sub

Private Sub mnuBranchRecDis_Click()
    frmHSummary.MySource = "BRANCH RECEIVING DISCREPANCY"
    frmHSummary.Show
End Sub

Private Sub mnuCardex_Click()
    SPCardex.Show vbModal
End Sub

Private Sub mnuCardexRed_Click()
    SPCardexRed.Show vbModal
End Sub


Private Sub mnuCashflow_Click()
    CashFlowCashier.Show
End Sub

Private Sub mnuChecks_Click()
    arCheckSumm.Show
End Sub

Private Sub mnuCustomer_Click()
    AcctCustomer.Show vbModal
End Sub

Private Sub mnuCV_Click()
    CashVoucher.Show
End Sub

Private Sub mnuDailyCash_Click()
    DailyCash.Show
End Sub

Private Sub mnuDailyExpense_Click()
    DisburseSumm.Show
End Sub

Private Sub mnuDailySales_Click()
    SPSalesRpt.Show
End Sub

Private Sub mnuDMCM_Click()
    DMCMadj.Show
End Sub

Private Sub mnuEmp_Click()
    Employees.Show
End Sub

Private Sub mnuInvAdj_Click()
   InvntryADJ.Show
End Sub

Private Sub mnuInvAdjSum_Click()
    FrmSummary.mySrc = "Inventory Adjustment"
    FrmSummary.Show
End Sub

Private Sub mnuInvAudit_Click()
    AuditInvntry.Show
End Sub

Private Sub mnuMerchandise_Click()
    SPrptINV.Show
End Sub

Private Sub mnuMonColl_Click()
    MonthlyColl.Show
End Sub

Private Sub mnuMonSales_Click()
    MonthlySales.Show
End Sub

Private Sub mnuMonthlyInv_Click()
    frmHSummary.MySource = "MONTHLY MERCHANDISE INVENTORY"
    frmHSummary.Show
End Sub

Private Sub mnuPartsProp_Click()
    SPupdate.Show
End Sub

Private Sub mnuPCV_Click()
    Disbursement.Show
End Sub

Private Sub mnuPOBranch_Click()
    FrmSummary.mySrc = "Branch Purchase Order"
    FrmSummary.Show
End Sub

Private Sub mnuPOsupplier_Click()
    SPordering.Show
End Sub

Private Sub mnuPr_Click()
    arPR.Show
End Sub

Private Sub mnuPrice_Click()
    SPpricing.Show
End Sub

Private Sub mnuPurchaseReturn_Click()
    PurchaseReturn.Show
End Sub

Private Sub mnuPurchRetSum_Click()
    FrmSummary.mySrc = "Purchase Return"
    FrmSummary.Show
End Sub

Private Sub mnuReceipts_Click()
    Receipt.Show
End Sub

Private Sub mnuReceiving_Click()
    SPincoming.Show
End Sub

Private Sub mnuReceivingSum_Click()
    InvDtlRecRpt.Show
End Sub

Private Sub mnuTrasferSum_click()
    
    frmHSummary.MySource = "DETAILED STOCKTRANSFER"
    frmHSummary.Show
    
End Sub

Private Sub mnuSalesCostPrice_Click()
    InvDtlCostSalRpt.Show
End Sub

Private Sub mnuSalesDB_Click()
    SalesDB.Show
End Sub

Private Sub mnuSalesDistrib_Click()
    SalesDistribution.Show
End Sub

Private Sub mnuSaleSellPrice_Click()
    InvDtlSalRpt.Show
End Sub

Private Sub mnuSalesReturn_Click()
    SalesReturn.Show
End Sub

Private Sub mnuSalRetSum_Click()
    FrmSummary.mySrc = "Sales Return"
    FrmSummary.Show

End Sub

Private Sub mnuStatement_Click()
    BillingNew.Show
End Sub

Private Sub mnuStockTransfer_Click()
    frmSPtrans.Show
End Sub

Private Sub mnuSupplier_Click()
    EMPsupplier.Show vbModal
End Sub

Private Sub mnuUser_Click()
    frmUser.Show vbModal
    
End Sub



