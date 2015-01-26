VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Look Up"
   ClientHeight    =   5310
   ClientLeft      =   7425
   ClientTop       =   3000
   ClientWidth     =   5805
   LinkTopic       =   "Form5"
   ScaleHeight     =   5310
   ScaleWidth      =   5805
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   480
      TabIndex        =   9
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mark"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   4815
      Begin VB.TextBox Text2 
         Height          =   420
         Left            =   1800
         TabIndex        =   7
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Others:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "1.0 (Ahead)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2640
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "0.1 (Delay)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3120
      TabIndex        =   0
      Text            =   "All"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Both upper case and lower case are acceptable"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Sales Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fcstApp As Excel.Application
Dim fcstBook As Excel.Workbook
Dim fcstSheet As Excel.Worksheet
Dim newApp As Excel.Application
Dim newBook As Excel.Workbook
Dim newSheet As Excel.Worksheet
Dim SName As String
Dim opt As String
Dim adt As String
Dim month_arr(11) As String
Dim new_exist As Boolean

Private Sub Form_Load()
SName = "All"
Text1.Text = SName
opt = ""
adt = ""
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
End Sub

Private Sub Text1_Change()
SName = Trim(Text1.Text)
End Sub

Private Sub Check1_Click()
opt = opt + "d"
End Sub

Private Sub Check2_Click()
opt = opt + "a"
End Sub

Private Sub Text2_Change()
adt = Text2.Text
If adt = "" Then
    Check3.Value = 0
Else
    Check3.Value = 1
    opt = opt + "t"
End If
End Sub

Private Sub Command1_Click()
If opt = "" Then
    MsgBox "Please select at least one condition to continue", vbOKOnly, "Error"
Else
    Unload Me
    Form1.Text1.Text = "Looking Up Cases"
    Form2.Show
    Form2.Text1.Text = "Looking Up Cases"
    Set fcstApp = New Excel.Application
    Set fcstBook = fcstApp.Workbooks.Open(Form1.fcst_path)
    fcstApp.Visible = False
    Set fcstSheet = fcstBook.Worksheets("Fcst_details")
    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Open fcst_all excel file, hidden"
    Form2.Text1.SelStart = Len(Form2.Text1.Text)
    Set newApp = New Excel.Application
    Set newBook = newApp.Workbooks.Add
    newApp.Visible = False
    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Output excel created, hidden"
    Form2.Text1.SelStart = Len(Form2.Text1.Text)
    If InStr(opt, "d") <> 0 Then
        Set newSheet = newApp.Application.Worksheets.Add
        newSheet.Name = "Delay"
        Call find_case("0.1")
    End If
    If InStr(opt, "a") <> 0 Then
        Set newSheet = newApp.Application.Worksheets.Add
        newSheet.Name = "Ahead"
        Call find_case("1")
    End If
    If InStr(opt, "t") <> 0 Then
        Set newSheet = newApp.Application.Worksheets.Add
        newSheet.Name = "Others"
        Call find_case(adt)
    End If
    fcstBook.Close savechanges:=False
    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "fcst_all closed"
    Form2.Text1.SelStart = Len(Form2.Text1.Text)
    fcstApp.Quit
    Set fcstApp = Nothing
    Form2.Text1.Text = ""
    Unload Form2
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Function contains(row As Integer, condition As String)
Dim n, result As Integer
For n = 0 To 11
    If fcstSheet.Cells(row, Form1.find_col(month_arr(n), fcstSheet)).Value = condition Then
        result = result + 1
    End If
Next n
contains = result
End Function

Private Function copy_row(row_n As Integer, row_d As Integer, eof_col As Integer)
Dim n As Integer
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Inserting case"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
For n = 1 To eof_col
    newSheet.Cells(row_n, n).Value = fcstSheet.Cells(row_d, n).Value
Next n
End Function

Private Function find_case(condition As String)
Dim row_d, row_n, n, empt_row, i As Integer
Dim continue As Boolean
Dim status As Dictionary
Dim sales_name As String
Dim counter, total As Integer
Dim num_keys As Integer
Dim all_keys() As Variant
row_d = 4
row_n = 4
n = 1
continue = True
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Adding headings"
Form2.Text1.SelStart = Len(Form2.Text1.Text)
While fcstSheet.Cells(3, n) <> ""
    newSheet.Cells(3, n).Value = fcstSheet.Cells(3, n).Value
    n = n + 1
Wend
Set status = New Dictionary
If LCase(SName) = "all" Then
    While continue
        sales_name = LCase(fcstSheet.Cells(row_d, Form1.find_col("sales name", fcstSheet)).Value)
        Form2.Text1.Text = Form2.Text1.Text & vbCrLf & row_d & ": "
        Form2.Text1.SelStart = Len(Form2.Text1.Text)
        If sales_name <> "" Then
            If status.Exists(sales_name) = False Then
                Form2.Text1.Text = Form2.Text1.Text & "Adding new sales dictionary: " & sales_name
                Form2.Text1.SelStart = Len(Form2.Text1.Text)
                status.Add sales_name, Array(0, 0)
            Else
                Form2.Text1.Text = Form2.Text1.Text & "Dictionary key " & sales_name & " exists"
                Form2.Text1.SelStart = Len(Form2.Text1.Text)
            End If
            If contains(Int(row_d), condition) <> 0 Then
                counter = status.Item(sales_name)(0) + 1
                total = status.Item(sales_name)(1) + contains(Int(row_d), condition)
                status.Item(sales_name) = Array(counter, total)
                Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Condition satisfied, updating counters"
                Form2.Text1.SelStart = Len(Form2.Text1.Text)
                Call copy_row(Int(row_n), Int(row_d), n - 1)
                row_n = row_n + 1
            End If
        Else
            empt_row = empt_row + 1
            Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Empty row found"
            Form2.Text1.SelStart = Len(Form2.Text1.Text)
        End If
        If empt_row = 5 Then
            continue = False
        End If
        row_d = row_d + 1
    Wend
Else
    row_d = Form1.find_row(Form1.find_col("sales name", fcstSheet), SName, fcstSheet)
    If row_d <> 0 Then
        Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Adding new sales dictionary: " & SName
        Form2.Text1.SelStart = Len(Form2.Text1.Text)
        status.Add SName, Array(0, 0)
        While continue
            If LCase(fcstSheet.Cells(row_d, Form1.find_col("sales name", fcstSheet)).Value) <> SName Then
                continue = False
            ElseIf fcstSheet.Cells(row_d, Form1.find_col("sales name", fcstSheet)).Value = "" Then
                empt_row = empty_row = 1
                Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Empty row found"
                Form2.Text1.SelStart = Len(Form2.Text1.Text)
            Else
                If contains(Int(row_d), condition) <> 0 Then
                    counter = status.Item(SName)(0) + 1
                    total = status.Item(SName)(1) + contains(Int(row_d), condition)
                    status.Item(SName) = Array(counter, total)
                    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Condition satisfied, updating counters"
                    Form2.Text1.SelStart = Len(Form2.Text1.Text)
                    Call copy_row(Int(row_n), Int(row_d), n - 1)
                    row_n = row_n + 1
                End If
            End If
            If empt_row = 5 Then
                continue = False
            End If
            row_d = row_d + 1
        Wend
    End If
End If
num_keys = status.Count
If num_keys = 0 Then
    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & SName & " does not exist"
    Form2.Text1.SelStart = Len(Form2.Text1.Text)
    Form1.Text1.Text = Form1.Text1.Text & vbCrLf & SName & " does not exist"
    Form1.Text1.SelStart = Len(Form1.Text1.Text)
    newBook.Close savechanges:=False
    newApp.Quit
    Set newApp = Nothing
Else
    ReDim all_keys(num_keys - 1)
    all_keys = status.Keys
    Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "Printing results on the main window"
    Form2.Text1.SelStart = Len(Form2.Text1.Text)
    If condition = "0.1" Then
        Form1.Text1.Text = Form1.Text1.Text & vbCrLf & "Case Delay"
        Form1.Text1.SelStart = Len(Form1.Text1.Text)
    ElseIf condition = "1" Then
        Form1.Text1.Text = Form1.Text1.Text & vbCrLf & "Case Ahead"
        Form1.Text1.SelStart = Len(Form1.Text1.Text)
    Else
        Form1.Text1.Text = Form1.Text1.Text & vbCrLf & "Others: " & condition
        Form1.Text1.SelStart = Len(Form1.Text1.Text)
    End If
    Form1.Text1.Text = Form1.Text1.Text & vbCrLf & "Sales Name" & vbTab & "Cases" & vbTab & "Total Num."
    Form1.Text1.SelStart = Len(Form1.Text1.Text)
    For i = 0 To num_keys - 1
        Form1.Text1.Text = Form1.Text1.Text & vbCrLf & UCase(all_keys(i)) & vbTab & status.Item(all_keys(i))(0) & vbTab & status.Item(all_keys(i))(1)
        Form1.Text1.SelStart = Len(Form1.Text1.Text)
    Next i
    newApp.Visible = True
End If
End Function
