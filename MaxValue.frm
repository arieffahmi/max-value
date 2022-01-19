VERSION 5.00
Begin VB.Form MaxValue 
   Caption         =   "Max Value In Array"
   ClientHeight    =   1290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMaxValue 
      Height          =   495
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblMaxValueInArray 
      Caption         =   "Maximum Value In Array:"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "MaxValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Main()
    Dim a(1 To 10) As Single
    
    Dim i As Integer
    
    For i = 1 To 10
        a(i) = InputBox("Enter a number for Array " & i)
    Next

    txtMaxValue.Text = Calculate(a)
End Sub


Public Function Calculate(a() As Single) As Single

    Dim i As Integer
    Dim MaxValue As Integer
    MaxValue = 0
      
    For i = 1 To 10
        If a(i) > MaxValue Then
        MaxValue = a(i)
        End If
    Next
    
    Calculate = MaxValue
    
End Function


Private Sub Form_Load()
    Call Main
End Sub
