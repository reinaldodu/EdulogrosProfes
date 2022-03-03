VERSION 5.00
Begin VB.Form PROCONS_ALUM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTAR ALUMNO"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "PROCONS_ALUM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CARNET:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   825
      End
   End
End
Attribute VB_Name = "PROCONS_ALUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim alumno As maestroalum
If Text1.Text = "" Then
MsgBox "ESCRIBA UN NUMERO DE CARNET", 64, "ADVERTENCIA"
Text1.SetFocus
GoTo FIN
End If
If Text1.Text > 32000 Then
MsgBox "No.CARNET INVALIDO", 64, "ADVERTENCIA"
Text1.SetFocus
GoTo FIN
End If
On Error Resume Next
Err.Clear
If Dir("a:\datos\inicial.edu") = "" Then
MsgBox "INSERTE EL DISKETTE DE DATOS EN LA UNIDAD A", 16, "ADVERTENCIA"
GoTo FIN
End If
NAR = FreeFile
Open "a:\datos\cont.edu" For Input As #NAR
Input #NAR, i
Close #NAR
h = Val(Text1.Text)
If ((h > i - 1) Or (h < 1)) Then
MsgBox "REGISTRO NO EXISTE", 32
Text1.SetFocus
GoTo FIN
End If
Open "a:\datos\prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, h, alumno
Close #NAR
If RTrim(alumno.n_carnet) = "" Then
MsgBox "ALUMNO ESTA RETIRADO", 32
Text1.SetFocus
GoTo FIN
End If
CONS_ALUM.Text1.Text = alumno.n_carnet
CONS_ALUM.Text13.Text = alumno.n_matricula
CONS_ALUM.Text2.Text = alumno.nombres
CONS_ALUM.Text3.Text = alumno.apellidos
CONS_ALUM.Text11.Text = alumno.documento
CONS_ALUM.Text4.Text = alumno.f_nacimiento
CONS_ALUM.Text5.Text = alumno.rh
CONS_ALUM.Text6.Text = alumno.acudiente
CONS_ALUM.Text8.Text = alumno.tel_acu
CONS_ALUM.Text16.Text = alumno.padre
CONS_ALUM.Text17.Text = alumno.tel_pa
CONS_ALUM.Text18.Text = alumno.madre
CONS_ALUM.Text19.Text = alumno.tel_ma
CONS_ALUM.Text7.Text = alumno.direccion
CONS_ALUM.Text9.Text = alumno.jornada
CONS_ALUM.Text10.Text = alumno.año_ingre
CONS_ALUM.Text12.Text = alumno.grado
CONS_ALUM.Text14.Text = alumno.sexo
dd = Val(Left(alumno.f_nacimiento, 2))
mm2 = Right(alumno.f_nacimiento, 7)
mm = Val(Left(mm2, 2))
aaaa = Val(Right(alumno.f_nacimiento, 4))
aaaa = Year(Date) - aaaa
If mm > Month(Date) Then
aaaa = aaaa - 1
End If
If mm = Month(Date) Then
   If dd > Day(Date) Then
   aaaa = aaaa - 1
   End If
End If
CONS_ALUM.Text15.Text = aaaa
CONS_ALUM.Show
FIN:
End Sub

Private Sub Form_Load()
Text1.MaxLength = 5
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
GoTo CONC445
End If
If C$ < "0" Or C$ > "9" Then
KeyAscii = 0
Beep
End If
CONC445:
End Sub
