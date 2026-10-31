VERSION 5.00
Begin VB.MDIForm BranchOrder 
   BackColor       =   &H00C0FFC0&
   Caption         =   "BRANCH PURCHASE ORDER"
   ClientHeight    =   10170
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuPO 
      Caption         =   "P.O."
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "QUIT"
   End
End
Attribute VB_Name = "BranchOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuPO_Click()
    SPordering.Show
End Sub

Private Sub mnuQuit_Click()
    
    LYRcon.Close
    End
    
End Sub

