VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Query Results"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPage 
      Height          =   405
      Left            =   5160
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ListBox lbxRecords 
      Height          =   2010
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblPage 
      Alignment       =   1  'Right Justify
      Caption         =   "Page"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Source Code Dimulai dari sini

Option Explicit
Dim db As Connection
Dim lCurrentPage As Long

Private Sub cmdNext_Click()
    lCurrentPage = lCurrentPage + 1
    Call LoadListBox(lCurrentPage)
End Sub

Private Sub cmdPrevious_Click()
    If lCurrentPage > 1 Then
        lCurrentPage = lCurrentPage - 1
        Call LoadListBox(lCurrentPage)
    End If
End Sub

Private Sub Form_Load()
    
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & AppPath & "test.mdb;"

    lCurrentPage = 1
    Call LoadListBox(lCurrentPage)

End Sub
Private Sub LoadListBox(lPage As Long)
    Dim adoPrimaryRS As ADODB.Recordset
    Dim lPageCount As Long
    Dim nPageSize As Integer
    Dim lCount As Long

    nPageSize = 7
    Set adoPrimaryRS = New Recordset
    adoPrimaryRS.Open "select * from numbers", db, adOpenStatic, adLockOptimistic

    adoPrimaryRS.PageSize = nPageSize
    lPageCount = adoPrimaryRS.PageCount
    If lCurrentPage > lPageCount Then
        lCurrentPage = lPageCount
    End If
    
    txtPage.Text = lPage
    
    adoPrimaryRS.AbsolutePage = lCurrentPage
    
    With lbxRecords
        .Clear
        lCount = 0
        Do While Not adoPrimaryRS.EOF
            .AddItem adoPrimaryRS("aNumber")
            lCount = lCount + 1
            If lCount = nPageSize Then
                Exit Do
            End If
            adoPrimaryRS.MoveNext
        Loop
        
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not db Is Nothing Then
        db.Close
    End If
    Set db = Nothing
End Sub
Public Function AppPath() As String
    
    Dim sAns As String
    sAns = App.Path
    If Right(App.Path, 1) <> "\" Then sAns = sAns & "\"
    AppPath = sAns

End Function
