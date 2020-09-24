VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmConverter 
   Caption         =   "Euro Converter"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   Icon            =   "frmConverter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Convert"
      Height          =   3855
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   4215
      Begin VB.Frame Frame3 
         Caption         =   "Value"
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   3975
         Begin VB.Label lblCurr 
            Caption         =   "Euros"
            Height          =   255
            Left            =   1800
            TabIndex        =   13
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblValue 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "&Convert"
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "0"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "From Euros To Selected Currency"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "From Selected Currency To Euros"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Amount"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.ListBox lstRate 
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   6600
      Width           =   2175
   End
   Begin VB.ListBox lstCurrency 
      Height          =   2985
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox rtbMain 
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   6240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmConverter.frx":0442
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Latest Exchange Rates"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   2775
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5400
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Available Currencies"
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()
    
    If lstCurrency.ListIndex < 0 Then
        MsgBox "Must select a currency!", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    
    If Val(txtAmount.Text) <= 0 Then
        MsgBox "Amount must be greater than zero!", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    
    If Option1.Value = True Then
        'from selected currency to euro
        lblValue = Val(txtAmount) / lstRate.List(lstCurrency.ListIndex)
    Else
        'from euro to selected currency
        lblValue = Val(txtAmount) * lstRate.List(lstCurrency.ListIndex)
    End If
    
End Sub

Private Sub cmdGet_Click()
    Screen.MousePointer = 11
    'get new exchange rates
    Inet1.URL = "http://www.x-rates.com/cgi-bin/cgicalc.cgi?value=1&base=EUR"
    rtbMain.Text = Inet1.OpenURL
    rtbMain.Text = Replace(rtbMain.Text, Chr(34), "")
    'parse the HTML
    'get start point
    a = InStr(1, rtbMain.Text, "Value of 1 Euro") + 16
    'get currency names
    Dim iStart As Long
    iStart = 1
    
    Do
        b = InStr(iStart, rtbMain.Text, "nbsp;&nbsp;&nbsp;&nbsp;")
        If b = 0 Then Exit Do
        b = b + 28
        c = InStr(b, rtbMain.Text, "font") - 2
        cName = Mid(rtbMain.Text, b, c - b)
        lstCurrency.AddItem cName
        
        d = InStr(c, rtbMain.Text, "html>") + 5
        e = InStr(d, rtbMain.Text, "</a>")
        vName = Mid(rtbMain.Text, d, e - d)
        lstRate.AddItem vName
        
        iStart = e
        
    Loop
    
    Screen.MousePointer = 0
    
End Sub

Private Sub lstCurrency_Click()
    If Option2.Value = True Then
        lblCurr.Caption = lstCurrency.List(lstCurrency.ListIndex)
    End If
End Sub

Private Sub Option1_Click()
    lblCurr.Caption = "Euros"
    lblValue.Caption = "0.0"
End Sub

Private Sub Option2_Click()
    lblCurr.Caption = lstCurrency.List(lstCurrency.ListIndex)
    lblValue.Caption = "0.0"
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
End Sub
