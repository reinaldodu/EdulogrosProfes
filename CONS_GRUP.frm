VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CONS_GRUP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de grupo"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9255
   Icon            =   "CONS_GRUP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Copiar"
      Height          =   735
      Left            =   7920
      Picture         =   "CONS_GRUP.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Copia la lista de estudiantes"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text4 
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
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   4935
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   320
         Left            =   3240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid MATI9 
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         BackColorBkg    =   -2147483633
         GridColor       =   12582912
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL ESTUDIANTES:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6840
      TabIndex        =   7
      Top             =   240
      Width           =   1755
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "CONS_GRUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Combo1.Text <> Combo1.List(0) Then
    Combo1.Text = Combo1.List(0)
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
'dim alumno As maestroalum
'dim profe As maestropro
'dim icur As inforcur
'dim alugru As grupoalu
Screen.MousePointer = 11
MATI9.Rows = 1
Label7.Caption = ""
Frame1.Caption = ""
Text4.Text = ""
YO = 0
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    Screen.MousePointer = 0
    End
End If
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    MsgBox "ESTE GRUPO NO LE CORRESPONDE", 48, "CONSULTA DE GRUPO"
    Combo1.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
YO = 1
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Combo1.Text Then
        J2 = RTrim(icur.grado)
        J1 = RTrim(icur.jornada)
        J3 = icur.director
        GoTo ALTU86
    End If
Wend
ALTU86:
Close #NAR
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, J3, profe
Close #NAR
Label7.Caption = "DIRECTOR(A): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
Frame1.Caption = "JORNADA: " & J1 & "   GRADO: " & J2 & "   GRUPO: " & Combo1.Text
leo = 0
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    leo = leo + 1
    Get #NAR, leo, alugru
Wend
Close #NAR
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
NAR = FreeFile
Open Ruta & "\prinalu.edu" For Random As #NAR Len = Len(alumno)
For TN = 1 To leo - 1
    Get #(NAR - 1), TN, alugru
    Get #NAR, (Val(alugru.num_carnet)), alumno
    MATI9.Rows = YO + 1
    MATI9.TextMatrix(YO, 0) = TN
    'MATI9.TextMatrix(YO, 1) = alumno.n_carnet
    'MATI9.TextMatrix(YO, 2) = alumno.n_matricula
    MATI9.TextMatrix(YO, 1) = RTrim(alumno.apellidos)
    MATI9.TextMatrix(YO, 2) = RTrim(alumno.nombres)
    'MATI9.TextMatrix(YO, 5) = RTrim(alumno.f_nacimiento)
    'dd = Val(Left(alumno.f_nacimiento, 2))
    'mm2 = Right(alumno.f_nacimiento, 7)
    'mm = Val(Left(mm2, 2))
    'aaaa = Val(Right(alumno.f_nacimiento, 4))
    'aaaa = Year(Date) - aaaa
    'If mm > Month(Date) Then
    '    aaaa = aaaa - 1
    'End If
    'If mm = Month(Date) Then
    '    If dd > Day(Date) Then
    '        aaaa = aaaa - 1
    '    End If
    'End If
    'MATI9.TextMatrix(YO, 6) = aaaa
    'MATI9.TextMatrix(YO, 7) = RTrim(alumno.acudiente)
    'MATI9.TextMatrix(YO, 8) = RTrim(alumno.tel_acu)
    YO = YO + 1
Next TN
Close #NAR
Close #(NAR - 1)
Text4.Text = YO - 1
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
If YO = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA COPIAR", 16, "COPIAR"
    Exit Sub
End If
Clipboard.Clear
cop = ""
For X = 1 To Val(Text4.Text)
        ape = RTrim(MATI9.TextMatrix(X, 1))
        nom = RTrim(MATI9.TextMatrix(X, 2))
        cop = cop + LTrim(ape & " " & nom) & vbCrLf
'        If X < 10 Then
'           cop = cop + LTrim(Str(X) & "   - " & ape & " " & nom) & vbCrLf
'        Else
'           cop = cop + LTrim(Str(X) & " - " & ape & " " & nom) & vbCrLf
'        End If
Next X
Clipboard.SetText cop
End Sub

Private Sub Form_Load()
'dim clavv As CLAVEPRO
'dim argra As areagr
MATI9.Row = 0
MATI9.Col = 0
MATI9.ColWidth(0) = 450
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "CD"
'MATI9.Col = 1
'MATI9.ColWidth(1) = 800
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "CARNET"
'MATI9.Col = 2
'MATI9.ColWidth(2) = 1100
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "MATRICULA"
MATI9.Col = 1
MATI9.ColWidth(1) = 2200
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "APELLIDOS"
MATI9.Col = 2
MATI9.ColWidth(2) = 2200
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "NOMBRES"
'MATI9.Col = 5
'MATI9.ColWidth(5) = 1100
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "FECH_NACIM"
'MATI9.Col = 6
'MATI9.ColWidth(6) = 600
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "EDAD"
'MATI9.Col = 7
'MATI9.ColWidth(7) = 3300
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "ACUDIENTE"
'MATI9.Col = 8
'MATI9.ColWidth(8) = 1300
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "TELEFONO"
If (Dir(Ruta & "areagra.edu") <> "") Then
    Command1.Enabled = True
    NAR = FreeFile
    'Open "a:\datos\clase.edu" For Random As #NAR Len = Len(clavv)
    'Get #NAR, 1, clavv
    'Close #NAR
    cona = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If argra.num_pro = Val(MENUPROFE.LBLNumProfe.Caption) Then
            VALI2 = False
            For I = 0 To (Combo1.ListCount - 1)
                If Combo1.List(I) = RTrim(argra.nom_grup) Then
                    VALI2 = True
                    Exit For
                End If
            Next I
            If VALI2 = False Then
                Combo1.AddItem RTrim(argra.nom_grup)
            End If
        End If
    Wend
    Close #NAR
    Combo1.Text = Combo1.List(0)
    If RTrim(Combo1.Text) = "" Then
        Command1.Enabled = False
    End If
Else
    Command1.Enabled = False
End If
YO = 0
End Sub
