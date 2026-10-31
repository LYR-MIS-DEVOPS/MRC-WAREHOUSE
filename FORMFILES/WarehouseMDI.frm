VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Warehouse 
   BackColor       =   &H00808080&
   Caption         =   "MRC Warehouse System"
   ClientHeight    =   11040
   ClientLeft      =   225
   ClientTop       =   -1620
   ClientWidth     =   20475
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
      Top             =   10665
      Width           =   20475
      _ExtentX        =   36116
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "WarehouseMDI.frx":1C2E6
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
            Picture         =   "WarehouseMDI.frx":1CED8
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
            Object.Width           =   22119
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "WarehouseMDI.frx":1D7B2
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
      End
      Begin VB.Menu mnuUser 
         Caption         =   "User Profile"
      End
      Begin VB.Menu mnuCustomer 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnuSupplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu mnuBranch 
         Caption         =   "Branch"
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
         Begin VB.Menu mnuWSPOsupplier 
            Caption         =   "Wholesale P.O to Supplier "
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPOBranch 
            Caption         =   "P.O. from Branch"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuBReplenish 
            Caption         =   "Branch Replenishment"
            Enabled         =   0   'False
            Begin VB.Menu mnuDataEntry 
               Caption         =   "Parts Matrix Data Entry"
               Enabled         =   0   'False
            End
            Begin VB.Menu mnuBrnchAssign 
               Caption         =   "PM Branch Assignment"
               Enabled         =   0   'False
            End
            Begin VB.Menu mnuBranchQty 
               Caption         =   "Branch Standard Quantity"
               Enabled         =   0   'False
            End
            Begin VB.Menu mnuPICKList 
               Caption         =   "Picking List Summary"
               Enabled         =   0   'False
            End
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
            Caption         =   "Retail Receiving Receipts"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuWSReceiving 
            Caption         =   "Wholesale Receiving Receipts"
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
            Caption         =   "PO Stock Transfer"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRStockTransfer 
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
            Caption         =   "Retail Merchandise Inventory"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuWSMerchandise 
            Caption         =   "Wholesale Merchandise Inventory"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCardex 
            Caption         =   "Retail Cardex"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuWSCardex 
            Caption         =   "Wholesale Cardex"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCardexRed 
            Caption         =   "Retail Cardex (RED)"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuWSCardexRed 
            Caption         =   "Wholesale Cardex (RED)"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuMonthlyInv 
            Caption         =   "Monthly Merchandise Inventory"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuBackOrder 
            Caption         =   "Back Order"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuBranchRecDis 
            Caption         =   "Branch Receiving Discrepancy"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSDisc 
            Caption         =   "Branch Discrepancy Process"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRInvInq 
            Caption         =   "Retail Inventory Inquiry"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MenuAdjustment 
         Caption         =   "Adjustments"
         Begin VB.Menu mnuPrice 
            Caption         =   "Pricing Discount"
         End
         Begin VB.Menu mnuPricingDiscount 
            Caption         =   "Pricing Discount"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSRPupdater 
            Caption         =   "SRP Updater"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSRPBReq 
            Caption         =   "SRP Branch Request"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPartsProp 
            Caption         =   "Parts Properties"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuInvAdj 
            Caption         =   "Retail Inventory Adjustment"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuWSInvAdj 
            Caption         =   "Wholesale Inventory Adjustment"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPOAdj 
            Caption         =   "Transaction Roll Back"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSDiscAdj 
            Caption         =   "Settlement of Discrepancy Adj."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPartsCatalog 
            Caption         =   "Parts Catalog"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuAdjReq 
            Caption         =   "Adjustment Request"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCancelTrans 
            Caption         =   "Cancellation of Transaction"
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
         Begin VB.Menu mnuInvoice 
            Caption         =   "Costing"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRFP 
            Caption         =   "Request For Payment"
            Begin VB.Menu mnuRFPTrade 
               Caption         =   "RFP Trade"
            End
            Begin VB.Menu mnuRFPNonTrade 
               Caption         =   "RFP Non-Trade"
            End
         End
         Begin VB.Menu mnuTreasuryRFP 
            Caption         =   "Request For Payment(Treasury)"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRFPRecon 
            Caption         =   "RFP Reconciliation"
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
         Begin VB.Menu mnuDetailedPO 
            Caption         =   "Detailed Purchase Order Report"
         End
         Begin VB.Menu mnuReceivingSum 
            Caption         =   "Detailed Receiving Report"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTrasferSum 
            Caption         =   "Detailed Outgoing Report"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDtldServRatio 
            Caption         =   "Detailed Supplier Service Ratio"
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
         Begin VB.Menu mnuServiceRatio 
            Caption         =   "Service Ratio"
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
      Begin VB.Menu mnuIntrnstSumm 
         Caption         =   "Intransit Summary"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInvAudit 
         Caption         =   "Inventory Audit"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep3 
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
    
    Warehouse.Caption = SysBranch & " v26.04.07"
    '" v15.12.21" = Sorting by partnumber ang transfer og picking requested by mam jean
    'v16.01.29 = qty on hand vs. allocated qty sa BPO
    'v16.04.16 = detailed outgoing summary
    'v17.01.07 = online po for mcparts depot rizal & binos tagum
    'v17.04.12 = parts matrix
    'v17.05.23 = suppliers invoice w/ Check
    'v17.05.30 = Reyl access to PO suggested qty.
    'v17.06.21 = Fix Suppliers Invoice computation
    'v17.07.04 = Add Branch Van74 in Branch PO and Update Accts Payable
    'v17.07.12 = Add Branch Van73 & MRC Matina in Branch PO and Update Accts Payable
    'v17.07.15 = fix ConnectOnline Van73 & fix Suppliers Invoice form layout (suggested by Sir Jo)
    'v17.09.15 = Add SRP Updater
    'v17.09.28 = Add Branch Discrepancy Process and Settlement of Discrepancy Adj. Forms
    'v17.10.25 = SRP PRICING
    'v17.11.29 = CARDEX AND STOCKTRANSFER (DELIVERED DATE)
    'v18.03.14 = UPDATED SRP PRICING
    'v18.04.17 = New Stock Transfer (Replenishment)
    'v18.04.18 = Update PO to Supplier (SPInsXst)
    'v18.04.20 = insert Last Date-IN in PriceFactor
    'v18.04.29 = Update in PO to Supplier, P.O from Branch (Print Picking Form) and Add Column in P.O from Branch OrderType
    'v18.08.02 = COSTING,PRICE AND PARTS PROPERTIES
    'v18.08.15 = Parts Catalog
    'v18.08.30 = SPinxst Update
    'v18.09.03 = SRP PRICING
    'v18.09.13 = PO TO SUPPLIER
    'v18.09.20 = PO TO SUPPLIER -can add qty
    'v18.09.25 = Dili na kaFINAL c JOHN sa SRP Updater only SIR JO
    'v18.10.05 = save to excel sa srp updater
    'v18.10.15 = update sa Inventory Adjustment, Cardex Red ug Receiving Stocks
    'v18.10.25 = update sa printing sa Supplier Product Cost Receiving Receipt (SPCRR)
    'v18.11.20 = update sa SPCRR, Inventory Adj (BACK ORDER), SRP UPDATER, PO TO SUPPLIER and RECEIVING
    'v18.11.26 = update printed form of PO TO SUPPLIER FORM
    'v19.01.10 = update sa incoming summary
    'v19.03.09 = update sa WHOLESALE WAREHOUSE na TRANSACTION
    'v19.03.12 = update sa WHOLESALE INVENTORY ADJUSTMENT & CARDEX
    'v19.03.15 = update BRANCH PO printing
    'v19.03.23 = Update PO TO SUPPLIER UG BRANCH PO
    'v19.04.13 = Update SPCRR
    'v19.04.25 = Update SUPPLIER ORDER, DETAILED SUPPLIER SERVICE RATIO
    'v19.04.26 = Update SP Insert form
    'v19.04.30 = Update Price Factor
    'v19.05.08 = Update PO Summary
    'v19.05.25 = Validation in INCOMING
    'v19.06.04 = discount in SPCRR
    'v19.06.06 = Removed the Finalized PO Validation in PO Summary when Confirming the Date and Shipment date
    'v19.06.17 = WS MAX DISCOUNT in PRICE ADJUSMENT & PRICE FACTOR
    'v19.06.21 = gbutngan ug Total ang rptSTFOrdPORepl ..
    'v19.08.03 = g-Update ang print preview sa retail merchandise inv
    'v19.08.23 = g-Update ang computation sa BACK ORDER sa RETAIL CARDEX RED & CONDITION sa QUERY ni BACK ORDER SUMMARY
    'v19.09.04 = nagAdd ug another Validation sa Update PO na button .. VALIDATION: DAPAT dili mkita ni TAGUM WAREHOUSE ang PO ni BRANCH na para kang DAVAO WAREHOUSE
    'v19.09.24 = g-Update ang Update PO ug Save PO details sa FormSummary
    'v19.09.26 = added MRCAGDAO in branch total
    'v19.10.26 = updated in SRPUpdater and PriceFactor
    'v19.11.04 = bugs Fix in SRP Updater and PriceFactor
    'v19.11.06 = bugs Fix in SRP Updater and PriceFactor
    'v19.12.07 = RFP WH module
    'v19.12.12 = RFP WH updates
    'v19.12.18 = RFP WH updates in SUMMARY
    'v19.01.11 = Back to SPinsxst
    'v19.01.23 = Updates in LOADSPAREPARTS Query
    'v19.02.03 = Updates Settlement of Discrepancy Adjustment in Damage Item
    'v19.02.03 = Updates in SPORDERING FROM CWOV TO CWV
    'v19.02.18 = Add Date Deposited in RFP Treasury
    'v19.02.24 = Updated SPCRR Printable Report
    'v20.03.13 = Updated LOADSPAREPARTS in SPOrdering
    'v20.06.06 = Fixing bugs in RFP form
    'v20.08.01 = Update sa SPOrdering dapat COST W VAT ang mo reflect and RFP WH
    'v20.08.07 = Update sa RFP WH
    'v20.08.18 = Update sa SPCRR in OTHER AMOUNT can input Negative Sign
    'v20.10.01 = Update in SPordering spareparts logs and srp validation in price factor
    'v20.10.06 = Update in SPordering enable the exportbutton in report if final na ang PO and disable if pending
    'v20.10.09 = Add New Branch in SPordering
    'v20.10.13 = Added Module Retail Inventory Inquiry
    'v20.10.14 = Update the Purchase Order Printable Form
    'v20.10.30 = Update the Receiving Receipt (Validation of Packing list Print Date)
    'v20.11.05 = New Module RFP Non-Trade
    'v20.11.06 = Fix bug RFP Non-Trade
    'v20.11.13 = Fix bug RFP Non-Trade (Print)
    'v20.11.18 = Fix RFP Non-Trade (Print)
    'v20.12.01 = Fix EXPORT button in Print Preview in Wholesale Dept P.O
    'v21.01.04 = Update Account Title in RFP Trade and Non-Trade
    'v21.01.26 = Fix account titles display in RFP Trade and Non-Trade
    'v21.01.28 = Fix SPCRR Other Amount
    'v21.02.05 = Added Document Date Received and Company in RFP Trade and RFP Non-Trade
    'v21.02.06 = Update for using the LYR MOTORPARTS branch name
    'v21.02.17 = Fix in SPFind MaxDiscount
    'v21.02.23 = Adjustment Request
    'v21.03.12 = Cancellation Module and add Company fields in RFP
    'v21.03.13 = Fix in RFP CBD
    'v21.03.24 = Add consignee in RFP
    'v21.03.26 = fix consignee in RFP
    'v21.04.22 = fix particulars in RFP Other and add bayugan new branch
    'v21.05.03 = fix search in PO
    'v21.06.16 = Change the Executive Assistant of CFO to Accounting Division Head in RFP
    'v21.07.19 = editable ang received by sa rfp treasury
    'v21.09.09 = change approved by in PartTrans from embedded approved name to requery to settings
    'v21.10.15 = updates on adjustment request and cancellation module
    'V23.08.30 = Updates on Price Adjustment - Can edit Wholesale Price
    'v23.10.07 = Change Stocktransfer form to Picking Form and Update Stock transfer duplicate entry validation
    'v24.02.05 = Update the exit button in STF-print form
    'v24.02.29 = Update Price Adjustment for 3 different discounts
    'V24.05.15 = Fixed : (stock transfer module) upon exit on print form
                'Fixed : Merge description and description 2 on SP Discount module and Spareparts advance search
                'Update: Change Supplier Brand to Product on SP Properties module and Display Sub categories
                'Adding heading on picking form from  Branch PO to Warehouse module
                'Remove Price and total Column on Branch PO to Warehouse
    'v24.05.15 = Adding key accounts discount feature
    'v24.05.24 = ADDING APRROVED BY IN P.O FROM BRANCH MODULE AND REFLECT ON REPORT
    'v24.05.27 = Changing RE ORDER POINT to SUPPLIER PRODUCT CODE on SP Properties
    'v24.05.27 = Fix Intransit data from PO from supplier module
    'v24.06.03 = Change 6mSOQ formula to SQ-TTLQH-BO-INTRN
    'v24.06.05 = Remove updating backorder from spareparts table upon creating backorder adjustment
                'modify spcrr crystal report
    'v24.06.11 = modify available QTY on STF by warehouse merchandise invty less pending STF/PO.
    '24.07.23 =  1.  SKUs with Zero RmQ will not be included upon printing of Picking Form.
            '    2.  Final Stock Transfer Form will reflect Unserved Quantity per SKU.
            '    3.  Printed Final Stock Transfer, exclude the following:
            '        a.  SKUs with Zero Qty Served
            '        b.  Column for Order Quantity, Unserved Qty and Percentage (%)
            '    4.  Served Qty for Final Stock Transfer automatically deducted on Warehouse Merchandise Inventory even STF is not yet confirmed as in-transit.
            '    5.  Report on Service Ratio can now be extracted for necessary analysis related to inventory.
    '24.07.23 =  backorder fixed.
    '24.08.02 = Stock Transfer processing fixed (1.0.2)
    '24.10.01 = CANCELLATION ON PO FROM BRANCH (1.1.3) MIS240906-02388
    '24.11.05 = ADDITIONAL FIELDS ON ENROLLING NEW PART NUMBER particularly SUB-CATEGORY and PRODUCT BRAND on supplier order form (1.2.3) MIS240914-01088
    '24.11.21 = add excel extract on rfp trade and non-trade (1.3.3) MIS241107-00542
    '24.11.23 (1.4.3) = remove new price and old price on spcrr print preview and add reviewed by : MIS241121-02788
    '24.12.23 (1.5.3) = set price diff. and price amount default value to zero when old cost from is zero  : MIS240923-04788
    '25.03.01 (1.5.3) = fixed backorder module matches to cardex red data  : MIS250220-02788
    '25.05.05 (1.6.3) = add Detailed Purchase Order Report on Inventory menu : MIS250217-02688
    'v25.05.27 (1.6.4) = sort by partnumber upon printing STF
    'v25.09.08 (1.7.4) = Adding remarks and waybill on RR packing lists print form and adjust Packing # length to 20+ characters : MIS250904-01388
    'v25.09.19 (1.8.4) = feature added by automatic loads all the data from the Receiving Receipt : MIS250906-01488
    'v25.12.01 (1.9.4) = fixed filtering on inventory adjustment and settlement of discrepancy summary : MIS251104-03188'
    'v25.12.03 (1.10.4) = set PACKING NO to editable if RFP TRADE term is CBD : MIS251120-00842
    'v25.12.13 (1.11.4) = update transaction list and account list to include all LYRMP unit code 4, 12, 21: MIS251120-00842 & MIS251209-03196
    'v26.01.31 (1.12.04) = POSITIONS on RFP trade printing form set to dynamic by getting all variables on the settings table
    'v26.04.07 (1.14.04) = add SRRNO and date received column in printing and extracting accounts payable invoicing summary
    'V26.04.07 (1.15.04) = update signatory of SPCRR: reiewed by & recommendation for approval :MIS260209-02588
    
    
    FrDate = DateValue("1/1/1977")
    ToDate = DateValue("1/1/1977")
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ans As Integer, i As Integer
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
    
    ans = MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Quit")
End Sub


Private Sub mnuStockTrans_Click()
    
    frmSPtrans.Show
      
End Sub


Private Sub mnuWhouse_Click()
    SPWarehouse.Show
End Sub


Private Sub MenuQuit_Click()
    Dim ans As Integer
    
    ans = MsgBox("Do you want to Quit?", vbYesNo, "Close Program")
    If ans = vbNo Then Exit Sub
    LYRcon.Close
    End

End Sub

Private Sub mnuAdjReq_Click()
    frmOtherAdjustment.Show
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
        .filter = "Backup Files (*.RAR)|(*.RAR)"
        .InitDir = App.Path
        myFile = myFile & Format(Now, "mmddyyhhmm")
        .filename = myFile & ".RAR"
        .ShowSave
        myFile2 = .filename
    End With
    If UCase$(SysBranch) = "LYR MOTORPARTS CORP. - MAIN WAREHOUSE" Then
        BranchDB = "mrcwarehouse"
    'If UCase$(SysBranch) = "LYR MOTORPARTS CORP. - SOBRECARY WAREHOUSE" Then
        'BranchDB = "lyrmpsobrecarywh"
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

'************** CODE UPDATED: 04/10/2017 **********************
Private Sub mnuBranchQty_Click()
    frmSummaryRepl.mySrc = "Branch Quantity"
    frm_OnProcess2.OwnerForm = "View Summary Repl"
    frmSummaryRepl.Show
  
End Sub

Private Sub mnuBrnchAssign_Click()
    BranchAssign.Show vbModal
End Sub

Private Sub mnuDataEntry_Click()
    DataEntry.Show
End Sub

Private Sub mnuBrnchIdentifying_Click()
    frmSummaryRepl.mySrc = "Branch Replenishment Order"
    frmSummaryRepl.Show
    'BranchIdentifying.Show
End Sub
'********** END CODE UPDATED: 04/10/2017 **********************

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

Private Sub mnuIntrnstSumm_Click()
 frmIntrnstSum.Show
End Sub

Private Sub mnuInvAdj_Click()
   InvntryADJ.Show
End Sub

Private Sub mnuInvAdjSum_Click()
    FrmSummary.mySrc = "Inventory Adjustment"
    FrmSummary.Show
End Sub

Private Sub mnuInvAudit_Click()
    'AuditInvntry.Show
    frmMonthlyPhysicalInvSumm.Show
End Sub

Private Sub mnuInvoice_Click()
    frmAcctPayInvcg.Show
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

Private Sub mnuPartsCatalog_Click()
    frmDataEntry.Show
End Sub

Private Sub mnuPartsProp_Click()
    frmSPproperties.Show
    'SPupdate.Show
End Sub

Private Sub mnuPCV_Click()
    Disbursement.Show
End Sub

'************** CODE UPDATED: 04/10/2017 **********************
Private Sub mnuPICKList_Click()
    frmSummaryRepl.mySrc = "Branch Picking"
    frmSummaryRepl.Show
End Sub
'********** END CODE UPDATED: 04/10/2017 **********************

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
'    frmPriceFactor.Show
    frmPriceDiscountv2.Show
End Sub

Private Sub mnuPricingDiscount_Click()
    frmPricingDiscount.Show
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

Private Sub mnuDetailedPO_Click()
     frmDetailedPurchaseOrderReport.MySource = "Detailed Purchase Order"
     frmDetailedPurchaseOrderReport.Show
End Sub

Private Sub mnuRFPNonTrade_Click()
 frmRFPOther.Show
End Sub

Private Sub mnuRFPRecon_Click()
 FrmSummary.mySrc = "RFP RECONCILIATION"
 FrmSummary.Show
End Sub

Private Sub mnuRFPTrade_Click()
 frmRFPWH.Show
End Sub

Private Sub mnuRInvInq_Click()
    frmInventoryInquiry.Show vbModal
End Sub

Private Sub mnuRStockTransfer_Click()
    ParTrans.Show
End Sub

'************** CODE UPDATED: 04/10/2017 **********************
'Private Sub mnuReplSTrans_Click()
'    ReplStockTransfer.MySource = "REPLENISH STOCK TRANSFER"
'    ReplStockTransfer.Show
'End Sub
'********** END CODE UPDATED: 04/10/2017 **********************

Private Sub mnuSDisc_Click()
    FrmSummary.mySrc = "SETTLEMENT DISCREPANCY"
    FrmSummary.Show
End Sub

Private Sub mnuSDiscAdj_Click()
    SDiscInvtryAdj.Show
End Sub

Private Sub mnuServiceRatio_Click()
   
    frmServiceRatioSummary.MySource = "SERVICE RATIO"
     frmServiceRatioSummary.Show
     
End Sub

Private Sub mnuSRPBReq_Click()
    frmSRPREQ.Show
End Sub

Private Sub mnuSRPupdater_Click()
    frmSRPUpdater.Show
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
    'frmSPtrans.Show
    frmPOReplStockTrans.Show
End Sub

Private Sub mnuSupplier_Click()
    EMPsupplier.Show vbModal
End Sub

Private Sub mnuUser_Click()
    frmUser.Show vbModal
End Sub

Private Sub mnuWSCardex_Click()
    SPCardexWS.Show
End Sub

Private Sub mnuWSCardexRed_Click()
    SPCardexRedWS.Show
End Sub

Private Sub mnuWSInvAdj_Click()
    WSInvntryADJ.Show
End Sub

Private Sub mnuWSMerchandise_Click()
    SPrptINVWS.Show
End Sub

Private Sub mnuWSPOsupplier_Click()
    SPorderingWS.Show
End Sub

Private Sub mnuWSReceiving_Click()
    SPincomingWS.Show
End Sub

Private Sub mnuDtldServRatio_Click()
    frmHSummary.MySource = "DETAILED SUPPLIER SERVICE RATIO"
    frmHSummary.Show
End Sub

Private Sub mnuPOAdj_Click()
    frmRollBack.Show vbModal
    'frmPOAdj.Show
End Sub

Private Sub mnuCancelTrans_Click()
    frmCancellation.Show vbModal
End Sub



