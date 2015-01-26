VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Service OI Forecast"
   ClientHeight    =   5985
   ClientLeft      =   6915
   ClientTop       =   2715
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   17.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   6975
   Begin VB.Frame Frame1 
      Caption         =   "Output"
      Height          =   3255
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   4815
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   480
         Width           =   4575
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Look Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   " Service OI             Forecast"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   6495
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu ReadMe 
         Caption         =   "Read Me"
         Enabled         =   0   'False
      End
      Begin VB.Menu break 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About Me"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sales_path, fcst_path As String
Dim month_arr(11) As String
Dim fcstApp As Excel.Application
Dim fcstBook As Excel.Workbook
Dim fcstSheet As Excel.Worksheet
Dim salesApp As Excel.Application
Dim salesBook As Excel.Workbook
Dim salesSheet As Excel.Worksheet
Dim ErrorMessage As String

Public Function find_col(info As String, sheet As Excel.Worksheet)
Dim n As Integer
n = 1
While n < 50
    If InStr(LCase(Trim(sheet.Cells(3, n).Value)), info) <> 0 Then
        find_col = n
        n = 100
    Else
        n = n + 1
    End If
Wend
If n = 50 Then
find_col = 0
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & info & " column not found"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
End If
End Function

Function check_row(col As Integer, info As String, sheet As Excel.Worksheet, sales_row As Integer, opt As String)
Dim row, counter, index, n As Integer
Dim result() As Integer
Dim break, pass As Boolean
Dim error_msg As String
row = 4
break = False
pass = False
While break = False
    If sheet.Cells(row, col).Value = "" Then
        counter = counter + 1
        row = row + 1
    Else
        counter = 0
        If sheet.Cells(row, col).Value = info Then
            ReDim Preserve result(index)
            result(index) = row
            index = index + 1
            row = row + 1
        Else
            row = row + 1
        End If
    End If
    If counter = 5 Then
    break = True
    End If
Wend
If index < 2 Then
    pass = True
Else
    Select Case opt
    Case "self"
        error_msg = "sales_fcst"
    Case "all"
        error_msg = "fcst_all"
    End Select
    While n < index
        error_msg = error_msg & " row " & result(n)
        n = n + 1
    Wend
    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & info & " (sales_fcst row " & sales_row & ")" & " found in " & error_msg
    Form2.Text1.SelStart = Len(Form2.Text1.Text)
    ErrorMessage = "Rpeated " & info & " (sales_fcst row " & sales_row & ")" & " found in " & error_msg
End If
check_row = pass
End Function

Function check(ToAdd As Excel.Worksheet, database As Excel.Worksheet, opt As String)
Dim toadd_no, database_no As Integer
Dim a As Integer
Dim pass As Integer
pass = True
toadd_no = find_col("no.", ToAdd)
database_no = find_col("no.", database)
a = 4
While ToAdd.Cells(a, toadd_no).Value <> ""
    If check_row(database_no, ToAdd.Cells(a, toadd_no).Value, database, a, opt) Then
        Form2.Text1.Text = Form2.Text1.Text & vbCrLf & ToAdd.Cells(a, toadd_no).Value & " pass"
        Form2.Text1.SelStart = Len(Form2.Text1.Text)
    Else
        Form2.Text1.Text = Form2.Text1.Text & vbCrLf & ToAdd.Cells(a, toadd_no).Value & " fail"
        Form2.Text1.SelStart = Len(Form2.Text1.Text)
        pass = False
    End If
    a = a + 1
Wend
check = pass
End Function

Public Function find_row(col As Integer, info As String, sheet As Excel.Worksheet)
Dim n, counter, result As Integer
Dim break As Boolean
n = 4
break = False
While break = False
    If sheet.Cells(n, col).Value = "" Then
        counter = counter + 1
        n = n + 1
    Else
        counter = 0
        If LCase(sheet.Cells(n, col).Value) = LCase(info) Then
            result = n
            break = True
        Else
            n = n + 1
        End If
    End If
    If counter = 5 Then
    break = True
    End If
Wend
find_row = result
End Function

Function insert_row(row_a As Integer, row_d As Integer, ToAdd As Excel.Worksheet, database As Excel.Worksheet)
Dim n, i, j As Integer
Dim continue As Boolean
continue = True
database.Cells(row_d, find_col("no.", database)).Value = ToAdd.Cells(row_a, find_col("no.", ToAdd)).Value
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "No. added"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
database.Cells(row_d, find_col("region", database)).Value = ToAdd.Cells(row_a, find_col("region", ToAdd)).Value
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Region added"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
database.Cells(row_d, find_col("bu", database)).Value = ToAdd.Cells(row_a, find_col("bu", ToAdd)).Value
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "BU added"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
database.Cells(row_d, find_col("div", database)).Value = ToAdd.Cells(row_a, find_col("div", ToAdd)).Value
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Div added"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
database.Cells(row_d, find_col("business type", database)).Value = ToAdd.Cells(row_a, find_col("business type", ToAdd)).Value
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Budiness Type added"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
database.Cells(row_d, find_col("client", database)).Value = ToAdd.Cells(row_a, find_col("client", ToAdd)).Value
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Client added"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
database.Cells(row_d, find_col("ender user", database)).Value = ToAdd.Cells(row_a, find_col("ender user", ToAdd)).Value
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Ender User added"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
database.Cells(row_d, find_col("qty of ender user in different", database)).Value = ToAdd.Cells(row_a, find_col("qty of ender user in different", ToAdd)).Value
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Qty of Ender User in different added"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
database.Cells(row_d, find_col("sales name", database)).Value = ToAdd.Cells(row_a, find_col("sales name", ToAdd)).Value
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Sales Name added"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
database.Cells(row_d, find_col("project name", database)).Value = ToAdd.Cells(row_a, find_col("project name", ToAdd)).Value
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Project Name added"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
While continue
    If Val(ToAdd.Cells(row_a, find_col(month_arr(n), ToAdd)).Value) <> 0 Then
        i = find_col(month_arr(n), database)
        j = find_col(month_arr(n), ToAdd)
        database.Cells(row_d, i + 3).Value = ToAdd.Cells(row_a, j + 3).Value
        database.Cells(row_d, i + 2).Value = ToAdd.Cells(row_a, j + 2).Value
        database.Cells(row_d, i + 1).Value = ToAdd.Cells(row_a, j + 1).Value
        Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Hit rates added"
        Form2.Text1.SelStart = Len(Form2.Text1.Text)
        database.Cells(row_d, i).Value = ToAdd.Cells(row_a, j).Value
        Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Amount added"
        Form2.Text1.SelStart = Len(Form2.Text1.Text)
        continue = False
    Else
        n = n + 1
    End If
    If n = 12 Then
        continue = False
    End If
Wend
End Function

Function available_row()
Dim n As Integer
Dim continue As Boolean
n = 4
continue = True
While continue
    If fcstSheet.Cells(n, find_col("sales name", fcstSheet)).Value = "" Then
        available_row = n
        continue = False
    Else
        n = n + 1
    End If
Wend
End Function

Function update(ToAdd As Excel.Worksheet, database As Excel.Worksheet)
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Updating database"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
Form1.Text1.Text = "Updating Database"
Dim row_a, row_d, col_a, col_d, curr_row, n, m, i, j, k, q, p As Integer
Dim sales_name As String
Dim continue, keep, added As Boolean
Dim cont_amt As Long
row_a = 4
col_a = find_col("no.", ToAdd)
While ToAdd.Cells(row_a, col_a).Value <> ""
    row_d = find_row(find_col("no.", database), ToAdd.Cells(row_a, col_a).Value, database)
    If row_d <> 0 Then
        Form2.Text1.Text = Form2.Text1.Text & vbCrLf & ToAdd.Cells(row_a, col_a).Value & " (" & row_a & ") -> " & row_d
        Form2.Text1.SelStart = Len(Form2.Text1.Text)
        Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Progresse status"
        Form2.Text1.SelStart = Len(Form2.Text1.Text)
        continue = True
        While continue
            If m = 12 Then
                continue = False
            Else
                i = find_col(month_arr(m), database)
                j = find_col(month_arr(m), ToAdd)
                cont_amt = CLng(database.Cells(row_d, i).Value) - CLng(ToAdd.Cells(row_a, j).Value)
                If cont_amt > 0 Then 'case delay
                    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & ToAdd.Cells(row_a, col_a).Value & " " & UCase(month_arr(m)) & " Delay"
                    Form2.Text1.SelStart = Len(Form2.Text1.Text)
                    Form1.Text1.Text = Form1.Text1.Text & vbCrLf & ToAdd.Cells(row_a, col_a).Value & " " & UCase(month_arr(m)) & " Delay"
                    Form1.Text1.SelStart = Len(Form2.Text1.Text)
                    database.Cells(row_d, i + 3).Value = ""
                    database.Cells(row_d, i + 2).Value = ""
                    database.Cells(row_d, i + 1).Value = ""
                    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Hit rates removed"
                    Form2.Text1.SelStart = Len(Form2.Text1.Text)
                    database.Cells(row_d, i).Value = 0.1
                    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "0.1 marked"
                    Form2.Text1.SelStart = Len(Form2.Text1.Text)
                    k = m + 1
                    keep = True
                    While keep
                        If k = 12 Then
                        keep = False
                        Else
                            p = find_col(month_arr(k), database)
                            q = find_col(month_arr(k), ToAdd)
                            If Val(ToAdd.Cells(row_d, q).Value) > 1 Then
                                database.Cells(row_d, p + 3).Value = ToAdd.Cells(row_a, q + 3)
                                database.Cells(row_d, p + 2).Value = ToAdd.Cells(row_a, q + 2)
                                database.Cells(row_d, p + 1).Value = ToAdd.Cells(row_a, q + 1)
                                Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Hit rates added"
                                Form2.Text1.SelStart = Len(Form2.Text1.Text)
                                database.Cells(row_d, p).Value = ToAdd.Cells(row_a, q)
                                Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Amount added"
                                Form2.Text1.SelStart = Len(Form2.Text1.Text)
                                keep = False
                            Else
                            k = k + 1
                            End If
                        End If
                    Wend
                    continue = False
                ElseIf cont_amt < 0 Then 'case ahead
                    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & ToAdd.Cells(row_a, col_a).Value & " " & UCase(month_arr(m)) & " Ahead"
                    Form2.Text1.SelStart = Len(Form2.Text1.Text)
                    Form1.Text1.Text = Form1.Text1.Text & vbCrLf & ToAdd.Cells(row_a, col_a).Value & " " & UCase(month_arr(m)) & " Ahead"
                    Form1.Text1.SelStart = Len(Form2.Text1.Text)
                    database.Cells(row_d, i + 3).Value = ToAdd.Cells(row_a, j + 3).Value
                    database.Cells(row_d, i + 2).Value = ToAdd.Cells(row_a, j + 2).Value
                    database.Cells(row_d, i + 1).Value = ToAdd.Cells(row_a, j + 1).Value
                    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Hit rates added"
                    Form2.Text1.SelStart = Len(Form2.Text1.Text)
                    database.Cells(row_d, i).Value = ToAdd.Cells(row_a, j).Value
                    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Amount added"
                    Form2.Text1.SelStart = Len(Form2.Text1.Text)
                    k = m + 1
                    keep = True
                    While keep
                        If k = 12 Then
                        keep = False
                        Else
                            p = find_col(month_arr(k), database)
                            If Val(database.Cells(row_d, p).Value) > 1 Then
                                database.Cells(row_d, p + 3).Value = ""
                                database.Cells(row_d, p + 2).Value = ""
                                database.Cells(row_d, p + 1).Value = ""
                                Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Hit rates removed"
                                Form2.Text1.SelStart = Len(Form2.Text1.Text)
                                database.Cells(row_d, p).Value = "1.0"
                                Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "1.0 marked"
                                Form2.Text1.SelStart = Len(Form2.Text1.Text)
                                keep = False
                            Else
                            k = k + 1
                            End If
                        End If
                    Wend
                    continue = False
                Else
                    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & ToAdd.Cells(row_a, col_a).Value & " " & UCase(month_arr(m)) & " Normal"
                    Form2.Text1.SelStart = Len(Form2.Text1.Text)
                    m = m + 1
                End If
            End If
        Wend
        Else
            'Call insert_case(row_a, ToAdd, database)
            Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "New case: " & ToAdd.Cells(row_a, col_a).Value
            Form2.Text1.SelStart = Len(Form2.Text1.Text)
            Form1.Text1.Text = Form1.Text1.Text & vbCrLf & "New case: " & ToAdd.Cells(row_a, col_a).Value
            Form1.Text1.SelStart = Len(Form2.Text1.Text)
            added = False
            sales_name = ToAdd.Cells(row_a, find_col("sales name", ToAdd)).Value
            curr_row = find_row(find_col("sales name", database), sales_name, database)
            If curr_row = 0 Then
                curr_row = available_row()
                database.Cells(curr_row, find_col("sales name", database)).Value = sales_name
            End If
            While database.Cells(curr_row, find_col("sales name", database)).Value = sales_name
                If database.Cells(curr_row, find_col("no.", database)).Value = "" Then
                    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Empty row found"
                    Form2.Text1.SelStart = Len(Form2.Text1.Text)
                    Call insert_row(Int(row_a), Int(curr_row), ToAdd, database)
                    added = True
                End If
                curr_row = curr_row + 1
            Wend
            If added = False Then
                database.Rows(curr_row).Insert
                Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "No empty row exsists" & vbCrLf & "New row inserted"
                Form2.Text1.SelStart = Len(Form2.Text1.Text)
                Call insert_row(Int(row_a), Int(curr_row), ToAdd, database)
            End If
        End If
        row_a = row_a + 1
        m = 0
Wend
End Function

Public Function comd1()
Form2.Show
Form2.Text1.Text = "Open fcst_all excel file, hidden"
Set fcstApp = New Excel.Application
Set fcstBook = fcstApp.Workbooks.Open(fcst_path)
fcstApp.Visible = False
Set fcstSheet = fcstBook.Worksheets("Fcst_details")
'open sales excel
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Open sales_fcst excel file, hidden"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
Set salesApp = New Excel.Application
Set salesBook = salesApp.Workbooks.Open(sales_path)
salesApp.Visible = False
Set salesSheet = salesBook.Worksheets(1)
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Check for repeated Case No. in sales_fcst"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
Form1.Text1.Text = "Check for repeated case No."
If check(salesSheet, salesSheet, "self") Then
    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "sales_fcst pass"
    Form2.Text1.SelStart = Len(Form2.Text1.Text)
    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Check for repeated Case No. in fcst_all"
    Form2.Text1.SelStart = Len(Form2.Text1.Text)
    If check(salesSheet, fcstSheet, "all") Then
        Call update(salesSheet, fcstSheet)
        salesApp.Visible = True
        fcstApp.Visible = True
    Else
        If MsgBox(ErrorMessage, vbOKOnly, "Error") = vbOK Then
            fcstApp.Visible = True
            salesApp.DisplayAlerts = False
            salesBook.Close
            salesApp.DisplayAlerts = True
            salesApp.Quit
            Set salesApp = Nothing
            Form1.Text1.Text = Form1.Text1.Text & vbCrLf & ErrorMessage
            Form1.Text1.SelStart = Len(Form2.Text1.Text)
        End If
    End If
Else
    If MsgBox(ErrorMessage, vbOKOnly, "Error") = vbOK Then
        salesApp.Visible = True
        fcstApp.DisplayAlerts = False
        fcstBook.Close
        fcstApp.DisplayAlerts = True
        fcstApp.Quit
        Set fcstApp = Nothing
        Form1.Text1.Text = Form1.Text1.Text & vbCrLf & ErrorMessage
        Form1.Text1.SelStart = Len(Form2.Text1.Text)
    End If
End If
Form2.Cls
Unload Form2
End Function

Private Sub delay(sec As Single)
Dim tmrEnd
tmrEnd = Timer() + sec
Do While Timer() < tmrEnd
DoEvents
Loop
End Sub

Private Sub Form_Load()
Me.Hide
frmSplash.Show
month_arr(0) = "oct"
month_arr(1) = "nov"
month_arr(2) = "dec"
month_arr(3) = "jan"
month_arr(4) = "feb"
month_arr(5) = "mar"
month_arr(6) = "apr"
month_arr(7) = "may"
month_arr(8) = "jun"
month_arr(9) = "jul"
month_arr(10) = "aug"
month_arr(11) = "sep"
delay (0.5)
Unload frmSplash
Me.Show
End Sub

Private Sub Command1_Click()
Form3.Show
End Sub

Private Sub Command2_Click()
If fcst_path = "" Then
    Form4.Show
Else
    Form5.Show
End If
End Sub

Private Sub command3_click()
Unload Me
End Sub

Private Sub sales_loc_Click()
Form3.Show
End Sub

Private Sub all_loc_Click()
Form4.Show
End Sub

Private Sub exit_click()
Unload Me
End Sub
