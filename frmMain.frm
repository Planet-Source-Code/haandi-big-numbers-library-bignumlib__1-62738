VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "HAANDI's Integer BigLib - Calculator"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Caption         =   "To Base"
      Height          =   855
      Left            =   2160
      TabIndex        =   16
      Top             =   15
      Width           =   1935
      Begin VB.ComboBox cmbBase 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":0000
         Left            =   165
         List            =   "frmMain.frx":00C1
         Style           =   2  'Dropdown-Liste
         TabIndex        =   17
         Top             =   330
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   105
      TabIndex        =   14
      Top             =   6435
      Width           =   1695
   End
   Begin VB.TextBox txtArg1 
      Height          =   975
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   13
      Top             =   1635
      Width           =   6015
   End
   Begin VB.TextBox txtResult 
      Height          =   975
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   12
      Top             =   5235
      Width           =   6015
   End
   Begin VB.TextBox txtArg2 
      Height          =   975
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   11
      Top             =   2835
      Width           =   6015
   End
   Begin VB.TextBox txtArg3 
      Height          =   975
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   10
      Top             =   4035
      Width           =   6015
   End
   Begin VB.TextBox txtAbout 
      Alignment       =   2  'Zentriert
      BackColor       =   &H8000000F&
      Height          =   345
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   6015
   End
   Begin VB.Frame Frame3 
      Caption         =   "BigLib Instruction"
      Height          =   855
      Left            =   4200
      TabIndex        =   2
      Top             =   0
      Width           =   1935
      Begin VB.ComboBox cmbInstructions 
         Height          =   315
         ItemData        =   "frmMain.frx":01B9
         Left            =   150
         List            =   "frmMain.frx":01ED
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   345
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "From Base"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      Begin VB.ComboBox cmbBase 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":027B
         Left            =   165
         List            =   "frmMain.frx":033C
         Style           =   2  'Dropdown-Liste
         TabIndex        =   3
         Top             =   330
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Height          =   375
      Left            =   4425
      TabIndex        =   0
      Top             =   6435
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "By HAANDI - 2005"
      Height          =   255
      Left            =   2025
      TabIndex        =   15
      Top             =   6495
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Result"
      Height          =   255
      Index           =   3
      Left            =   105
      TabIndex        =   9
      Top             =   4995
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Argument 3"
      Height          =   255
      Index           =   2
      Left            =   105
      TabIndex        =   8
      Top             =   3795
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Argument 2"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   2595
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Argument 1"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   1395
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbInstructions_Click()
If cmbInstructions.Text = "BigAdd" Then
txtAbout.Text = "Result = Arg1 + Arg2 ."
End If
If cmbInstructions.Text = "BigAnd" Then
txtAbout.Text = "Result = Arg1 AND Arg2 ."
End If
If cmbInstructions.Text = "BigDec" Then
txtAbout.Text = "Result = Arg1 - 1 ."
End If
If cmbInstructions.Text = "BigDiv" Then
txtAbout.Text = "Result = ceiling( Arg1 / Arg2 ) ."
End If
If cmbInstructions.Text = "BigInc" Then
txtAbout.Text = "Result = Arg1 + 1 ."
End If
If cmbInstructions.Text = "BigMod" Then
txtAbout.Text = "Result = Arg1 MOD Arg2 ."
End If
If cmbInstructions.Text = "BigMul" Then
txtAbout.Text = "Result = Arg1 * Arg2 ."
End If
If cmbInstructions.Text = "BigNot" Then
txtAbout.Text = "Result = NOT(Arg1) ."
End If
If cmbInstructions.Text = "BigOr" Then
txtAbout.Text = "Result = Arg1 OR Arg2 ."
End If
If cmbInstructions.Text = "BigPow" Then
txtAbout.Text = "Result = Arg1^Arg2 . "
End If
If cmbInstructions.Text = "BigPowMod" Then
txtAbout.Text = "Result = (Arg1^Arg2) MOD Arg3 ."
End If
If cmbInstructions.Text = "BigShl" Then
txtAbout.Text = "Result = Arg1 << Arg2 ."
End If
If cmbInstructions.Text = "BigShr" Then
txtAbout.Text = "Result = Arg1 >> Arg2 ."
End If
If cmbInstructions.Text = "BigSub" Then
txtAbout.Text = "Result = Arg1 - Arg2 ."
End If
If cmbInstructions.Text = "BigXor" Then
txtAbout.Text = "Result = Arg1 XOR Arg2 ."
End If
If cmbInstructions.Text = "BigBaseConvert" Then
txtAbout.Text = "Result[ToBase] = Arg1[FromBase] ."
Frame1.Enabled = True: Frame2.Enabled = True
cmbBase(0).Enabled = True: cmbBase(1).Enabled = True
Else
Frame1.Enabled = False: Frame2.Enabled = False
cmbBase(0).Enabled = False: cmbBase(1).Enabled = False
End If
End Sub

Private Sub cmdExecute_Click()
Dim sA As String, sB As String, sC As String, sR As String
cmdExecute.Enabled = False
DoEvents
sA = txtArg1.Text
sB = txtArg2.Text
sC = txtArg3.Text

If cmbInstructions.Text = "BigAdd" Then
sR = StrAdd(sA, sB)
End If
If cmbInstructions.Text = "BigAnd" Then
sA = ConvFromBase10(sA, "16")
sB = ConvFromBase10(sB, "16")
sR = BooleanOperation(sA, sB, "AND")
sR = ConvToBase10(sR, "16")
End If
If cmbInstructions.Text = "BigDec" Then
sR = StrDec(sA)
End If
If cmbInstructions.Text = "BigDiv" Then
sR = StrDiv(sA, sB)
End If
If cmbInstructions.Text = "BigInc" Then
sR = StrInc(sA)
End If
If cmbInstructions.Text = "BigMod" Then
sR = StrMod(sA, sB)
End If
If cmbInstructions.Text = "BigMul" Then
sR = StrMult(sA, sB)
End If
If cmbInstructions.Text = "BigNot" Then
sA = ConvFromBase10(sA, "16")
sR = BooleanOperation(sA, vbNullString, "NOT")
sR = ConvToBase10(sR, "16")
End If
If cmbInstructions.Text = "BigOr" Then
sA = ConvFromBase10(sA, "16")
sB = ConvFromBase10(sB, "16")
sR = BooleanOperation(sA, sB, "OR")
sR = ConvToBase10(sR, "16")
End If
If cmbInstructions.Text = "BigPow" Then
sR = StrPow(sA, sB)
End If
If cmbInstructions.Text = "BigPowMod" Then
sR = StrPowMod(sA, sB, sC)
End If
If cmbInstructions.Text = "BigShl" Then
sR = StrShl(sA, sB)
End If
If cmbInstructions.Text = "BigShr" Then
sR = StrShr(sA, sB)
End If
If cmbInstructions.Text = "BigSub" Then
sR = StrSub(sA, sB)
End If
If cmbInstructions.Text = "BigXor" Then
sA = ConvFromBase10(sA, "16")
sB = ConvFromBase10(sB, "16")
sR = BooleanOperation(sA, sB, "XOR")
sR = ConvToBase10(sR, "16")
End If
If cmbInstructions.Text = "BigBaseConvert" Then
sR = ConvertBases(sA, cmbBase(0).Text, cmbBase(1).Text)
End If
txtResult.Text = sR
cmdExecute.Enabled = True
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
cmbBase(0).Text = "10"
cmbBase(1).Text = "16"
cmbInstructions.Text = "BigAdd"
End Sub
