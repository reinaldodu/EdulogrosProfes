VERSION 5.00
Begin VB.Form LIST_LOGRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir listas para el control de logros"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5895
   Icon            =   "LIST_LOGRO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   3615
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "LIST_LOGRO.frx":0442
         Left            =   120
         List            =   "LIST_LOGRO.frx":0455
         TabIndex        =   1
         Text            =   "PRIMERO"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "AREA:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PERIODO:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   120
      Picture         =   "LIST_LOGRO.frx":0483
      ScaleHeight     =   3675
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "LISTAS DE CONTROL POR PERIODO"
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton Option2 
         Caption         =   "LISTA &FINAL"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LISTA DE &TRABAJO"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "JORNADA"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
End
Attribute VB_Name = "LIST_LOGRO"
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
If KeyAscii = 13 Then
    Combo2.SetFocus
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If Command2.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command2_Click
End If
End Sub

Private Sub Command2_Click()
'dim profe As maestropro
'dim clavv As CLAVEPRO
'dim mate As infomater
'dim alumno As maestroalum
'dim alugru As grupoalu
'dim argra As areagr
'dim ini As inicio
RESP = MsgBox("DESEA IMPRIMIR ESTE LISTADO?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
Screen.MousePointer = 11
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
Close #NAR
y = 0
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
y = y + 1
Get #NAR, y, mate
If RTrim(mate.nom) = Combo2.Text Then
que = mate.num
End If
Wend
Close #NAR
'Open "a:\datos\clase.edu" For Random As #NAR Len = Len(clavv)
'Get #NAR, 1, clavv
'Close #NAR
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, Val(MENUPROFE.LBLNumProfe.Caption), profe
Close #NAR
If Option1.Value = True Then
Printer.ScaleMode = 7
PAG = 1
Printer.CurrentY = 1
Printer.CurrentX = 5.5
Printer.Font.Size = 12
Printer.Print "CONTROL DE LOGROS  PERIODO: " & Combo1.Text
Printer.Font.Size = 10
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "P�g." & PAG
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print ini.nombre;
Printer.CurrentX = 16.5
Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
Printer.CurrentX = 0.5
Printer.Print "PROFESOR: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
Printer.CurrentX = 11
Printer.Print "AREA: " & Combo2.Text;
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
CLO = 0
While Not EOF(NAR)
CLO = CLO + 1
Get #NAR, CLO, argra
If (((argra.num_area) = que) And (RTrim((argra.nom_grup)) = Frame2.Caption)) Then
Close #NAR
GoTo inn44
End If
Wend
Close #NAR
inn44:
Printer.CurrentX = 19
Printer.Print "IH: " & argra.ih
Printer.CurrentX = 0.5
Printer.Print "JORNADA: " & Label5.Caption;
Printer.CurrentX = 11
Printer.Print "GRUPO: " & Frame2.Caption
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print "CD";
Printer.CurrentX = 1.3
Printer.Print "APELLIDOS Y NOMBRES";
Printer.Print ""
Open Ruta & Frame2.Caption & ".gru" For Random As #NAR Len = Len(alugru)
leo = 0
While Not EOF(NAR)
leo = leo + 1
Get #NAR, leo, alugru
Wend
Close #NAR
Open Ruta & Frame2.Caption & ".gru" For Random As #NAR Len = Len(alugru)
NAR = FreeFile
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
For rr = 1 To leo - 1
Get #(NAR - 1), rr, alugru
Get #NAR, (Val(alugru.num_carnet)), alumno
Printer.CurrentX = 0.5
Printer.Print rr;
Printer.CurrentX = 1.3
Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
Printer.Print ""
If (rr Mod 50) = 0 Then
Printer.NewPage
PAG = PAG + 1
Printer.CurrentY = 1
Printer.CurrentX = 5.5
Printer.Font.Size = 12
Printer.Print "CONTROL DE LOGROS  PERIODO: " & Combo1.Text
Printer.Font.Size = 10
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "P�g." & PAG
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print ini.nombre;
Printer.CurrentX = 16.5
Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
Printer.CurrentX = 0.5
Printer.Print "PROFESOR: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
Printer.CurrentX = 11
Printer.Print "AREA: " & Combo2.Text;
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
CLO = 0
While Not EOF(NAR)
CLO = CLO + 1
Get #NAR, CLO, argra
If (((argra.num_area) = que) And (RTrim((argra.nom_grup)) = Frame2.Caption)) Then
Close #NAR
GoTo inn
End If
Wend
Close #NAR
inn:
Printer.CurrentX = 19
Printer.Print "IH: " & argra.ih
Printer.CurrentX = 0.5
Printer.Print "JORNADA: " & Label5.Caption;
Printer.CurrentX = 11
Printer.Print "GRUPO: " & Frame2.Caption
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print "CD";
Printer.CurrentX = 1.3
Printer.Print "APELLIDOS Y NOMBRES";
Printer.Print ""
End If
Next rr
Close #(NAR - 1)
Close #NAR
End If
If Option2.Value = True Then
Printer.ScaleMode = 7
PAG = 1
Printer.CurrentY = 1
Printer.CurrentX = 5.5
Printer.Font.Size = 12
Printer.Print "CONTROL DE LOGROS  PERIODO: " & Combo1.Text
Printer.Font.Size = 10
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "P�g." & PAG
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print ini.nombre;
Printer.CurrentX = 16.5
Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
Printer.CurrentX = 0.5
Printer.Print "PROFESOR: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
Printer.CurrentX = 11
Printer.Print "AREA: " & Combo2.Text;
NAR = FreeFile
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
CLO = 0
While Not EOF(NAR)
CLO = CLO + 1
Get #NAR, CLO, argra
If (((argra.num_area) = que) And (RTrim((argra.nom_grup)) = Frame2.Caption)) Then
Close #NAR
GoTo inn77
End If
Wend
Close #NAR
inn77:
Printer.CurrentX = 19
Printer.Print "IH: " & argra.ih
Printer.CurrentX = 0.5
Printer.Print "JORNADA: " & Label5.Caption;
Printer.CurrentX = 11
Printer.Print "GRUPO: " & Frame2.Caption
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print "CD";
Printer.CurrentX = 1.3
Printer.Print "APELLIDOS Y NOMBRES";
Printer.CurrentX = 10.5
Printer.Print "LOGROS Y/O DIFICULTADES";
Printer.CurrentX = 18.4
Printer.Print "JV";
Printer.CurrentX = 19.4
Printer.Print "FA"
Printer.CurrentX = 10.5
Printer.Print "LG1 LG2 LG3 LG4 LG5 LG6 LG7 LG8 LG9 LG10"
Open Ruta & Frame2.Caption & ".gru" For Random As #NAR Len = Len(alugru)
leo = 0
While Not EOF(NAR)
leo = leo + 1
Get #NAR, leo, alugru
Wend
Close #NAR
Open Ruta & Frame2.Caption & ".gru" For Random As #NAR Len = Len(alugru)
NAR = FreeFile
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
For rr = 1 To leo - 1
Get #(NAR - 1), rr, alugru
Get #NAR, (Val(alugru.num_carnet)), alumno
Printer.CurrentX = 0.5
Printer.Print rr;
Printer.CurrentX = 1.3
Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
Printer.Print ""
If (rr Mod 50) = 0 Then
Printer.NewPage
PAG = PAG + 1
Printer.CurrentY = 1
Printer.CurrentX = 5.5
Printer.Font.Size = 12
Printer.Print "CONTROL DE LOGROS  PERIODO: " & Combo1.Text
Printer.Font.Size = 10
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "P�g." & PAG
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print ini.nombre;
Printer.CurrentX = 16.5
Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
Printer.CurrentX = 0.5
Printer.Print "PROFESOR: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
Printer.CurrentX = 11
Printer.Print "AREA: " & Combo2.Text;
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
CLO = 0
While Not EOF(NAR)
CLO = CLO + 1
Get #NAR, CLO, argra
If (((argra.num_area) = que) And (RTrim((argra.nom_grup)) = Frame2.Caption)) Then
Close #NAR
GoTo inn87
End If
Wend
Close #NAR
inn87:
Printer.CurrentX = 19
Printer.Print "IH: " & argra.ih
Printer.CurrentX = 0.5
Printer.Print "JORNADA: " & Label5.Caption;
Printer.CurrentX = 11
Printer.Print "GRUPO: " & Frame2.Caption
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print "CD";
Printer.CurrentX = 1.3
Printer.Print "APELLIDOS Y NOMBRES";
Printer.CurrentX = 10.5
Printer.Print "LOGROS Y/O DIFICULTADES";
Printer.CurrentX = 18.4
Printer.Print "JV";
Printer.CurrentX = 19.4
Printer.Print "FA"
Printer.CurrentX = 10.5
Printer.Print "LG1 LG2 LG3 LG4 LG5 LG6 LG7 LG8 LG9 LG10"
End If
Next rr
Close #(NAR - 1)
Close #NAR
End If
Printer.EndDoc
Printer.Font.Size = 8
End If
Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
'dim mate As infomater
'dim clavv As CLAVEPRO
'dim argra As areagr
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") Then
    Command2.Enabled = True
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
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            NAR = NAR - 1
            VALI2 = False
            For I = 0 To (Combo2.ListCount - 1)
                If Combo2.List(I) = RTrim(mate.nom) Then
                    VALI2 = True
                    Exit For
                End If
            Next I
            If VALI2 = False Then
                Combo2.AddItem RTrim(mate.nom)
            End If
        End If
    Wend
    Close #NAR
    Combo2.Text = Combo2.List(0)
    If (RTrim(Combo2.Text) = "") Then
        Command2.Enabled = False
    End If
Else
    Command2.Enabled = False
End If
Option1.Value = True
End Sub
