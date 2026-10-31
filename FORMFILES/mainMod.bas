Attribute VB_Name = "mainMod"
Global LYRcon As New ADODB.Connection

Declare Function CreateFieldDefFile Lib "p2smon.dll" (lpUnk As Object, ByVal filename As String, ByVal bOverWriteExistingFile As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Global Const CB_FINDSTRING = &H14C
Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String

'user variables
Global UserID As Integer

Sub Main()
                          
    On Error GoTo LocalErr
    LYRcon.CursorLocation = adUseClient
    LYRcon.Open "dsn=SQLMCPDEPOT"
    BranchOrder.Show
    LogIn.Show vbModal
    Exit Sub
LocalErr:
    'RESIZE EXTRA SERIAL#:  A9E8C7C6D9D0F2E4B1B5
    MsgBox "Error connecting to server..." & vbCrLf & "Make sure you are connected to the Server before running the program.", vbExclamation + vbOKOnly, "Connection Failed"

End Sub

Public Function NumToText(dblVal As Double) As String
    Static Ones(0 To 9) As String
    Static Teens(0 To 9) As String
    Static Tens(0 To 9) As String
    Static Thousands(0 To 4) As String
    Static bInit As Boolean
    Dim i As Integer, bAllZeros As Boolean, bShowThousands As Boolean
    Dim strVal As String, strBuff As String, strTemp As String
    Dim nCol As Integer, nChar As Integer
    Debug.Assert dblVal >= 0

    If bInit = False Then
       
        bInit = True
        Ones(0) = "zero"
        Ones(1) = "one"
        Ones(2) = "two"
        Ones(3) = "three"
        Ones(4) = "four"
        Ones(5) = "five"
        Ones(6) = "six"
        Ones(7) = "seven"
        Ones(8) = "eight"
        Ones(9) = "nine"
        Teens(0) = "ten"
        Teens(1) = "eleven"
        Teens(2) = "twelve"
        Teens(3) = "thirteen"
        Teens(4) = "fourteen"
        Teens(5) = "fifteen"
        Teens(6) = "sixteen"
        Teens(7) = "seventeen"
        Teens(8) = "eighteen"
        Teens(9) = "nineteen"
        Tens(0) = ""
        Tens(1) = "ten"
        Tens(2) = "twenty"
        Tens(3) = "thirty"
        Tens(4) = "forty"
        Tens(5) = "fifty"
        Tens(6) = "sixty"
        Tens(7) = "seventy"
        Tens(8) = "eighty"
        Tens(9) = "ninety"
        Thousands(0) = ""
        Thousands(1) = "thousand"   'US numbering
        Thousands(2) = "million"
        Thousands(3) = "billion"
        Thousands(4) = "trillion"
    End If
On Error GoTo NumToTextError
    If dblVal - Int(dblVal) > 0 Then
        strBuff = "PESOS and " & Format((dblVal - Int(dblVal)) * 100, "00") & "/100"
    Else
        strBuff = "PESOS ONLY"
    End If
     strVal = CStr(Int(dblVal))
    bAllZeros = True
    For i = Len(strVal) To 1 Step -1
        nChar = Val(Mid$(strVal, i, 1))
        nCol = (Len(strVal) - i) + 1
        Select Case (nCol Mod 3)
            Case 1
                bShowThousands = True
                If i = 1 Then
                    strTemp = Ones(nChar) & " "
                ElseIf Mid$(strVal, i - 1, 1) = "1" Then
                    
                    strTemp = Teens(nChar) & " "
                    i = i - 1
                ElseIf nChar > 0 Then
                    
                    strTemp = Ones(nChar) & " "
                Else
                     bShowThousands = False
                    If Mid$(strVal, i - 1, 1) <> "0" Then
                        bShowThousands = True
                    ElseIf i > 2 Then
                        If Mid$(strVal, i - 2, 1) <> "0" Then
                            bShowThousands = True
                        End If
                    End If
                    strTemp = ""
                End If
                 If bShowThousands Then
                    If nCol > 1 Then
                        strTemp = strTemp & Thousands(nCol \ 3)
                        If bAllZeros Then
                            strTemp = strTemp & " "
                        Else
                            strTemp = strTemp & " "
                         End If
                    End If
                     bAllZeros = False
                End If
                strBuff = strTemp & strBuff
            Case 2
                If nChar > 0 Then
                    If Mid$(strVal, i + 1, 1) <> "0" Then
                        strBuff = Tens(nChar) & " " & strBuff
                    Else
                        strBuff = Tens(nChar) & " " & strBuff
                    End If
                End If
            Case 0
                If nChar > 0 Then
                    strBuff = Ones(nChar) & " hundred " & strBuff
                End If
        End Select
    Next i
    'strBuff = UCase$(Left$(strBuff, 1)) & Mid$(strBuff, 2)
    strBuff = UCase$(strBuff) '& Mid$(strBuff, 2)
EndNumToText:
    NumToText = strBuff
    Exit Function
NumToTextError:
    strBuff = "#Error#"
    Resume EndNumToText
End Function


Public Sub CenterFrm(ByVal Parentfrm As MDIForm, ByVal Childfrm As Form)
    Childfrm.Left = (Parentfrm.Width \ 2) - (Childfrm.Width \ 2)
    Childfrm.Top = (Parentfrm.Height \ 2) - (Childfrm.Height \ 2) - 600
End Sub


Public Function MyPaperSize() As Integer
    Dim MyPrn As Printer
    Dim i As Integer
    'Dim doctoprint As New System.Drawing.Printing.PrintDocument
    
    'MyPaperSize = MyPrn.PaperSize
    
   '         doctoprint.PrinterSettings.PrinterName = "YourPrinterName" '(ex. "Epson SQ-1170 ESC/P 2")
            For i = 0 To doctoprint.PrinterSettings.PaperSizes.Count - 1
                Dim rawKind As Integer
                If doctoprint.PrinterSettings.PaperSizes(i).PaperName = "ADEL 8.5X5.5" Then
                    rawKind = CInt(doctoprint.PrinterSettings.PaperSizes(i).GetType().GetField("ADEL 8.5X5.5", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).GetValue(doctoprint.PrinterSettings.PaperSizes(i)))
                    MyPaperSize = rawKind
                    Exit Function
               End If
            Next
     MyPaperSize = 0
End Function

Public Sub PrintMyForm(ByRef MyVal() As String, FormName As String, ByVal MyFlx As MSFlexGrid)
    Dim rsPref As New ADODB.Recordset
    Dim mcLyr As CRAXDRT.Application
    Dim mcReport As CRAXDRT.Report

    Dim Qry As String, z As Integer
    Dim i As Integer, j As Integer, NxtRow As Long, InsertHere As Integer
    Dim MyDesc() As String, myTop() As String, myLeft() As String, myLen() As String
    Dim myHeight() As String, myFsize() As String, myFbold() As String, myAlign() As String
    Dim mySubLeft() As String, mySubLen() As String, mySubBold() As String, mySubAlign() As String
    Dim txtObj As TextObject
    
    Set mcLyr = New CRAXDRT.Application
    Set mcReport = mcLyr.NewReport

    On Error GoTo LocalErr


    Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
          "WHERE (((PRINTPREF.FORMNAME)='" & FormName & "'));"
    rsPref.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsPref.RecordCount > 0 Then
        rsPref.MoveFirst
        'Assign data to variables
        MyDesc = Split(rsPref!Description, ",")
        myTop = Split(rsPref!Top, ",")
        myLeft = Split(rsPref!Left, ",")
        myLen = Split(rsPref!Length, ",")
        myHeight = Split(rsPref!Height, ",")
        myFsize = Split(rsPref!FontSize, ",")
        myFbold = Split(rsPref!FontBold, ",")
        myAlign = Split(rsPref!Alignment, ",")
        InsertHere = IIf(rsPref!INSERTINDEX <> 0, rsPref!INSERTINDEX, -1)
        'Start Printing here
        If FormName = "INVOICE" Or FormName = "REQUISITION" Then
            mcReport.PaperSize = crPaperEnvelopeB6
        ElseIf FormName = "JOB ORDER" Then
            mcReport.PaperSize = crDefaultPaperSize
        Else
            mcReport.PaperSize = crPaper10x14
            'mcReport.PaperSize = crDefaultPaperSize
        End If
        mcReport.TopMargin = 0
        mcReport.BottomMargin = 0
        mcReport.Sections(1).Height = 0
        mcReport.Sections(2).Height = 0
        With mcReport.Sections(3)
            .Height = rsPref!FORMHEIGHT
            For i = LBound(MyDesc) To UBound(MyDesc)
                If i = InsertHere Then
                    NxtRow = rsPref!LOOPSTARTTOP
                    mySubLeft = Split(rsPref!SUBLEFT, ",")
                    mySubLen = Split(rsPref!SUBLEN, ",")
                    mySubBold = Split(rsPref!SUBFBOLD, ",")
                    mySubAlign = Split(rsPref!SUBALIGN, ",")
                    If FormName = "Promissory Note" Then z = 10
                    For j = 1 To MyFlx.Rows - 1
                        For k = LBound(mySubLeft) To UBound(mySubLeft)
                            If Len(MyFlx.TextMatrix(j, k + 1)) > 55 And InStr(UCase$(FormName), "INVOICE") > 1 Then
                                Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
                                txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial Narrow": txtObj.Font.Size = 8: txtObj.Height = 600: txtObj.CanGrow = True ' Val(rsPref!LOOPHEIGHT)
                            Else
                                Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
                                txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Tahoma": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT)
                            End If
                        Next k
                        If FormName = "Promissory Note" Then    'Control rows in Promi
                            NxtRow = NxtRow + rsPref!LOOPINTERVAL - z
                            z = IIf(z = 70, 10, z + 10)
                        Else
                            NxtRow = NxtRow + rsPref!LOOPINTERVAL
                        End If
                    Next j
                End If
                If Len(MyVal(i)) > 75 And UCase$(FormName) = "SALES DR" And InStr(MyDesc(i), "ItemRow") > 0 Then
                    Set txtObj = .AddTextObject(MyVal(i), Val(myLeft(i)), Val(myTop(i)))
                    txtObj.Width = Val(myLen(i)): txtObj.HorAlignment = Val(myAlign(i)): txtObj.Font.Name = "Arial Narrow": txtObj.Font.Size = 8: txtObj.Height = 600: txtObj.Font.Bold = Val(myFbold(i))
                Else
                    Set txtObj = .AddTextObject(MyVal(i), Val(myLeft(i)), Val(myTop(i)))
                    txtObj.Width = Val(myLen(i)): txtObj.HorAlignment = Val(myAlign(i)): txtObj.Font.Name = "Tahoma": txtObj.Font.Size = Val(myFsize(i)): txtObj.Height = Val(myHeight(i)): txtObj.Font.Bold = Val(myFbold(i))
                End If
            Next i
        End With
        mcReport.Sections(4).Height = 0
        mcReport.Sections(5).Height = 0
        
        mcReport.PrintOut False, 1, , 1, 1
    Else
        MsgBox "No Printing Preference found!", vbCritical + vbOKOnly, "Unknown Form Name"
    End If
    rsPref.Close
    Exit Sub
LocalErr:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf & "Please inform your System Administrator", vbCritical + vbOKOnly, "Error"


End Sub


Function DTCpy(MyQry As String)
    Dim rsPrn As New ADODB.Recordset
    Dim myCnt As Integer, Tmp As String
    
    With rsPrn
        .Open MyQry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            myCnt = IIf(IsNull(.Fields(0)), 0, .Fields(0)) + 1
            .Fields(0) = myCnt
            .Update
        End If
        .Close
    End With
    If myCnt = 1 Then
        DTCpy = Format(Now, "mm/dd/yy hh:mm ampm")
    Else
        If myCnt = 2 Then
            Tmp = "nd"
        ElseIf myCnt = 3 Then
            Tmp = "rd"
        Else
            Tmp = "th"
        End If
        DTCpy = Format(Now, "mm/dd/yy hh:mm ampm") & " " & myCnt & Tmp & " copy"
    End If
End Function

Public Function RepAlign(aStr As String) As Integer
    Select Case aStr
        Case "crLeftAlign"
            RepAlign = 3
        Case "crHorCenterAlign"
            RepAlign = 1
        Case "crRightAlign"
            RepAlign = 4
    End Select
End Function

Public Function CtrlNumExist(Qry As String) As Boolean
    Dim rsExist As New ADODB.Recordset
    
    With rsExist
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            CtrlNumExist = True
        Else
            CtrlNumExist = False
        End If
        .Close
    End With
End Function


Public Function GetFieldValue(MyQ As String) As Variant
    Dim rsAll As New ADODB.Recordset
    
    With rsAll
        .Open MyQ, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValue = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValue = ""
        End If
        .Close
    End With
        
End Function
