Attribute VB_Name = "FormatSquarespaceOrders"
Option Explicit

' ============================================================================
'  Squarespace order formatter  --  Excel for Mac (also works on Windows)
'
'  HOW YOUR WIFE USES IT (after the one-time setup):
'    1. Open the Squarespace orders CSV in Excel (double-click it).
'    2. Click the "Format Orders" button on the toolbar.
'    -> A brand-new workbook opens with a clean, grouped packing list.
'       The original CSV is never changed.
'
'  WHAT IT DOES, on the sheet that is currently open:
'    - Fills the Shipping Name down onto every line of a multi-item order
'      (falls back to Billing Name if an order has no shipping name).
'    - Sorts by Shipping Name so each customer's rows sit together,
'      including a repeat customer's separate orders.
'    - Keeps only: Order ID, Lineitem quantity, Lineitem name,
'      Lineitem variant, Shipping Name.
'    - Centers everything, styles + freezes the header, shades each
'      customer block, and auto-sizes the columns.
'
'  Order numbers are shown left-padded to 5 digits (e.g. 01437) to match
'  Squarespace, even though Excel drops the leading zero when it opens a CSV.
'  Change ORDER_ID_MIN_WIDTH below if your order numbers are a different size.
' ============================================================================

Private Const ORDER_ID_MIN_WIDTH As Long = 5

Public Sub FormatSquarespaceOrders()
    Dim wsIn As Worksheet
    On Error Resume Next
    Set wsIn = ActiveSheet
    On Error GoTo 0
    If wsIn Is Nothing Then
        MsgBox "Open your Squarespace orders CSV first, then click the button.", vbExclamation, "Format Orders"
        Exit Sub
    End If

    Dim data As Variant
    data = wsIn.UsedRange.Value
    If Not IsArray(data) Then
        MsgBox "This sheet has no data. Open the orders CSV, then try again.", vbExclamation, "Format Orders"
        Exit Sub
    End If

    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)
    If nRows < 2 Then
        MsgBox "This sheet doesn't look like an orders export (no rows).", vbExclamation, "Format Orders"
        Exit Sub
    End If

    ' ---- locate the columns we need, by header name (row 1) ----
    Dim cOrder As Long, cQty As Long, cName As Long, cVar As Long, cShip As Long, cBill As Long
    cOrder = FindCol(data, "Order ID")
    cQty = FindCol(data, "Lineitem quantity")
    cName = FindCol(data, "Lineitem name")
    cVar = FindCol(data, "Lineitem variant")
    cShip = FindCol(data, "Shipping Name")
    cBill = FindCol(data, "Billing Name")

    Dim missing As String
    If cOrder = 0 Then missing = missing & vbCr & "   - Order ID"
    If cQty = 0 Then missing = missing & vbCr & "   - Lineitem quantity"
    If cName = 0 Then missing = missing & vbCr & "   - Lineitem name"
    If cVar = 0 Then missing = missing & vbCr & "   - Lineitem variant"
    If cShip = 0 Then missing = missing & vbCr & "   - Shipping Name"
    If Len(missing) > 0 Then
        MsgBox "This doesn't look like a Squarespace orders export." & vbCr & _
               "Couldn't find these column(s):" & missing, vbExclamation, "Format Orders"
        Exit Sub
    End If

    ' ---- pass 1: learn each order's shipping name (first non-empty wins) ----
    Dim shipByOrder As Collection
    Set shipByOrder = New Collection
    Dim r As Long, oid As String, s As String
    For r = 2 To nRows
        oid = CleanStr(data(r, cOrder))
        If Len(oid) > 0 Then
            s = CleanStr(data(r, cShip))
            If Len(s) > 0 Then AddIfNew shipByOrder, oid, s
        End If
    Next r
    ' fall back to Billing Name for any order with no shipping name at all
    If cBill > 0 Then
        For r = 2 To nRows
            oid = CleanStr(data(r, cOrder))
            If Len(oid) > 0 Then
                If Not HasKey(shipByOrder, oid) Then
                    s = CleanStr(data(r, cBill))
                    If Len(s) > 0 Then AddIfNew shipByOrder, oid, s
                End If
            End If
        Next r
    End If

    ' ---- build the output rows (skip blank trailing rows) ----
    Dim cap As Long: cap = nRows - 1
    Dim oOrder() As String, oQty() As Variant, oName() As String
    Dim oVar() As String, oShip() As String, idx() As Long
    ReDim oOrder(1 To cap): ReDim oQty(1 To cap): ReDim oName(1 To cap)
    ReDim oVar(1 To cap): ReDim oShip(1 To cap): ReDim idx(1 To cap)

    Dim m As Long: m = 0
    Dim nm As String
    For r = 2 To nRows
        oid = CleanStr(data(r, cOrder))
        nm = CleanStr(data(r, cName))
        If Len(oid) = 0 And Len(nm) = 0 Then GoTo NextRow  ' truly blank row
        m = m + 1
        oOrder(m) = PadOrder(oid)
        oQty(m) = data(r, cQty)
        oName(m) = nm
        oVar(m) = CleanStr(data(r, cVar))
        s = CleanStr(data(r, cShip))
        If Len(s) = 0 Then s = LookupOrEmpty(shipByOrder, oid)
        oShip(m) = s
        idx(m) = m
NextRow:
    Next r

    If m = 0 Then
        MsgBox "No order rows found on this sheet.", vbExclamation, "Format Orders"
        Exit Sub
    End If

    ' ---- stable sort by (shipping name, order id, original order) ----
    Dim i As Long, j As Long, keyv As Long
    For i = 2 To m
        keyv = idx(i)
        j = i - 1
        Do While j >= 1
            If CompareRows(oShip(idx(j)), oOrder(idx(j)), idx(j), _
                           oShip(keyv), oOrder(keyv), keyv) > 0 Then
                idx(j + 1) = idx(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        idx(j + 1) = keyv
    Next i

    ' ---- write a clean new workbook ----
    Application.ScreenUpdating = False

    Dim wbOut As Workbook
    Set wbOut = Workbooks.Add
    Dim ws As Worksheet
    Set ws = wbOut.Sheets(1)
    On Error Resume Next
    ws.Name = "Orders"
    On Error GoTo 0

    ' keep Order ID as text so leading zeros survive
    ws.Columns(1).NumberFormat = "@"

    ws.Cells(1, 1).Value = "Order ID"
    ws.Cells(1, 2).Value = "Lineitem quantity"
    ws.Cells(1, 3).Value = "Lineitem name"
    ws.Cells(1, 4).Value = "Lineitem variant"
    ws.Cells(1, 5).Value = "Shipping Name"

    Dim outArr() As Variant
    ReDim outArr(1 To m, 1 To 5)
    Dim rr As Long
    For i = 1 To m
        rr = idx(i)
        outArr(i, 1) = oOrder(rr)
        outArr(i, 2) = oQty(rr)
        outArr(i, 3) = oName(rr)
        outArr(i, 4) = oVar(rr)
        outArr(i, 5) = oShip(rr)
    Next i
    ws.Cells(2, 1).Resize(m, 5).Value = outArr

    Dim full As Range
    Set full = ws.Cells(1, 1).Resize(m + 1, 5)

    full.Font.Name = "Arial"
    full.Font.Size = 11
    full.HorizontalAlignment = xlCenter
    full.VerticalAlignment = xlCenter
    full.Borders.LineStyle = xlContinuous
    full.Borders.Color = RGB(217, 217, 217)
    full.Borders.Weight = xlThin

    ' header styling
    With ws.Cells(1, 1).Resize(1, 5)
        .Interior.Color = RGB(68, 84, 106)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' shade every other customer block
    Dim prevKey As String, band As Boolean
    prevKey = Chr$(1)
    band = False
    For i = 1 To m
        rr = idx(i)
        If LCase$(oShip(rr)) <> prevKey Then
            band = Not band
            prevKey = LCase$(oShip(rr))
        End If
        If band Then
            ws.Cells(i + 1, 1).Resize(1, 5).Interior.Color = RGB(242, 242, 242)
        End If
    Next i

    ' sizing
    ws.Columns("A:E").AutoFit
    If ws.Columns("C").ColumnWidth > 60 Then ws.Columns("C").ColumnWidth = 60
    If ws.Columns("D").ColumnWidth > 40 Then ws.Columns("D").ColumnWidth = 40
    If ws.Columns("E").ColumnWidth > 34 Then ws.Columns("E").ColumnWidth = 34

    ' filter + frozen header
    full.AutoFilter
    wbOut.Activate
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    ws.Cells(2, 1).Select

    Application.ScreenUpdating = True

    Dim custCount As Long
    custCount = CountCustomers(oShip, idx, m)
    MsgBox "Done! " & m & " line items from " & custCount & " customers, grouped and formatted." & vbCr & vbCr & _
           "Press Cmd+S to save this as an Excel file.", vbInformation, "Format Orders"
End Sub


' ---------- helpers ----------

Private Function FindCol(ByRef data As Variant, ByVal header As String) As Long
    Dim c As Long
    For c = 1 To UBound(data, 2)
        If StrComp(CleanStr(data(1, c)), header, vbTextCompare) = 0 Then
            FindCol = c
            Exit Function
        End If
    Next c
    FindCol = 0
End Function

Private Function CleanStr(ByVal v As Variant) As String
    If IsError(v) Then CleanStr = "": Exit Function
    If IsNull(v) Then CleanStr = "": Exit Function
    If IsEmpty(v) Then CleanStr = "": Exit Function
    CleanStr = Trim$(CStr(v))
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As Integer
    If Len(s) = 0 Then IsAllDigits = False: Exit Function
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch < 48 Or ch > 57 Then IsAllDigits = False: Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function PadOrder(ByVal s As String) As String
    If Len(s) = 0 Then PadOrder = "": Exit Function
    If IsAllDigits(s) And Len(s) < ORDER_ID_MIN_WIDTH Then
        PadOrder = String$(ORDER_ID_MIN_WIDTH - Len(s), "0") & s
    Else
        PadOrder = s
    End If
End Function

Private Sub AddIfNew(ByRef col As Collection, ByVal k As String, ByVal v As String)
    On Error Resume Next
    col.Add v, k          ' fails silently if key already present -> first wins
    On Error GoTo 0
End Sub

Private Function HasKey(ByRef col As Collection, ByVal k As String) As Boolean
    Dim tmp As Variant
    On Error Resume Next
    tmp = col(k)
    HasKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function LookupOrEmpty(ByRef col As Collection, ByVal k As String) As String
    On Error Resume Next
    LookupOrEmpty = col(k)
    If Err.Number <> 0 Then LookupOrEmpty = ""
    Err.Clear
    On Error GoTo 0
End Function

' returns >0 if row A should sort AFTER row B, <0 before, 0 equal
Private Function CompareRows(ByVal shipA As String, ByVal ordA As String, ByVal ixA As Long, _
                             ByVal shipB As String, ByVal ordB As String, ByVal ixB As Long) As Long
    Dim c As Long
    c = StrComp(LCase$(shipA), LCase$(shipB), vbTextCompare)
    If c <> 0 Then CompareRows = c: Exit Function
    c = StrComp(ordA, ordB, vbTextCompare)
    If c <> 0 Then CompareRows = c: Exit Function
    If ixA > ixB Then
        CompareRows = 1
    ElseIf ixA < ixB Then
        CompareRows = -1
    Else
        CompareRows = 0
    End If
End Function

Private Function CountCustomers(ByRef oShip() As String, ByRef idx() As Long, ByVal m As Long) As Long
    Dim i As Long, prev As String, n As Long
    prev = Chr$(1)
    For i = 1 To m
        If LCase$(oShip(idx(i))) <> prev Then
            n = n + 1
            prev = LCase$(oShip(idx(i)))
        End If
    Next i
    CountCustomers = n
End Function
