VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form LOGRO_PEN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de logros pendientes"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9375
   Icon            =   "LOGRO_PEN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   9135
      Begin VB.Frame Frame4 
         Caption         =   "OPCIONES"
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   5640
         TabIndex        =   13
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton Command4 
            Caption         =   "&IMPRIMIR"
            Height          =   735
            Left            =   1800
            Picture         =   "LOGRO_PEN.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir la información de la lista."
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&GUARDAR"
            Height          =   735
            Left            =   240
            Picture         =   "LOGRO_PEN.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Guardar la información de la lista."
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   5415
         Begin VB.CommandButton Command1 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   3720
            TabIndex        =   4
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   960
            TabIndex        =   3
            Top             =   720
            Width           =   2535
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   960
            TabIndex        =   2
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "AREA   :"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO:"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   630
         End
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "LOGRO_PEN.frx":0CC6
      Left            =   7680
      List            =   "LOGRO_PEN.frx":0CD9
      TabIndex        =   1
      Text            =   "PRIMERO"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9135
      Begin MSFlexGridLib.MSFlexGrid MATI50 
         Height          =   4215
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7435
         _Version        =   393216
         Rows            =   1
         Cols            =   13
         BackColorBkg    =   12632256
         GridColor       =   12582912
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6360
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5760
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO:"
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
      Left            =   6720
      TabIndex        =   8
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "LOGRO_PEN"
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
If KeyAscii = 13 Then
Combo3.SetFocus
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
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
'dim mate As infomater
'dim notas As notis
'dim icur As inforcur
'dim logru As logris
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
If VALI44 = False Then
    Call Command3_Click
    If Err.Number <> 0 Then
        Exit Sub
    End If
End If
If Dir(Ruta & Combo2.Text & ".gru") = "" Then
    MsgBox "GRUPO NO LE CORRESPONDE", 64
    Combo2.SetFocus
    Exit Sub
End If
Screen.MousePointer = 11
MATI50.Rows = 1
Frame1.Caption = ""
Frame3.Caption = ""
Label4.Caption = 0
Label5.Caption = ""
Label6.Caption = ""
If RTrim(Combo1.Text) = "PRIMERO" Then
    lw = 1
End If
If RTrim(Combo1.Text) = "SEGUNDO" Then
    lw = 2
End If
If RTrim(Combo1.Text) = "TERCERO" Then
    lw = 3
End If
If RTrim(Combo1.Text) = "CUARTO" Then
    lw = 4
End If
If RTrim(Combo1.Text) = "FINAL" Then
    lw = 5
End If
fl = Left(Combo2.Text, 1)
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Combo2.Text Then
        ser = Left(icur.grado, 3)
    End If
Wend
Close #NAR
y = 0
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
    y = y + 1
    Get #NAR, y, mate
    If RTrim(mate.nom) = Combo3.Text Then
        que = mate.num
    End If
Wend
Close #NAR
If Dir(Ruta & fl & ser & que & lw & ".lgr") = "" Then
    MsgBox "NO EXISTE INFORMACION DE LOGROS", 64
    Screen.MousePointer = 0
    Exit Sub
End If
If Dir(Ruta & Combo2.Text & que & lw & ".obp") = "" Then
    MsgBox "DEBE GENERAR PRIMERO EL ARCHIVO DE LOGROS PENDIENTES", 48
    Screen.MousePointer = 0
    Exit Sub
End If
Label5.Caption = que & lw
ret = 0
Open Ruta & Combo2.Text & que & lw & ".obp" For Random As #NAR Len = Len(notas)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, notas
Wend
Close #NAR
FERT = 0
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    FERT = FERT + 1
    Get #NAR, FERT, logru
Wend
Close #NAR
Open Ruta & Combo2.Text & que & lw & ".obp" For Random As #NAR Len = Len(notas)
For J = 1 To ret - 1
    Get #NAR, J, notas
    MATI50.Rows = J + 1
    MATI50.TextMatrix(J, 0) = J
    NAR = FreeFile
    Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
    For I = 1 To 10
        If notas.area(I) <> 0 Then
            Get #NAR, notas.area(I), logru
            If (logru.indicador = "D") Or (logru.indicador = "N") Then
                MATI50.TextMatrix(J, I) = notas.area(I)
            End If
        End If
    Next I
    Close #NAR
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Get #NAR, (Val(notas.num_carnet)), alumno
    Close #NAR
    MATI50.TextMatrix(J, 11) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
    MATI50.TextMatrix(J, 12) = alumno.n_carnet
    NAR = NAR - 1
Next J
Close #NAR
Label6.Caption = "AREA: " & Combo3.Text
Frame1.Caption = Combo2.Text
Frame3.Caption = "PERIODO: " & Combo1.Text
Label4.Caption = MATI50.Rows - 1
Screen.MousePointer = 0
End Sub

Private Sub Command3_Click()
'dim notas As notis
If Val(Label4.Caption) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA GUARDAR", 16
    Exit Sub
End If
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
RESP = MsgBox("DESEA GUARDAR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
If RESP = vbYes Then
    Screen.MousePointer = 11
    If Dir(Ruta & Frame1.Caption & Label5.Caption & ".obp") <> "" Then
        Kill Ruta & Frame1.Caption & Label5.Caption & ".obp"
    End If
    NAR = FreeFile
    Open Ruta & Frame1.Caption & Label5.Caption & ".obp" For Random As #NAR Len = Len(notas)
    For I = 1 To Val(Label4.Caption)
        For J = 1 To 10
            If MATI50.TextMatrix(I, J) = "" Then
                notas.area(J) = 0
            Else
                notas.area(J) = MATI50.TextMatrix(I, J)
            End If
        Next J
        notas.num_carnet = Right(MATI50.TextMatrix(I, 12), 5)
        Put #NAR, I, notas
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Close #NAR
            Exit Sub
        End If
    Next I
    Close #NAR
End If
VALI44 = True
Screen.MousePointer = 0
End Sub

Private Sub Command4_Click()
'Dim ini As inicio
'Dim logru As logris
If Val(Label4.Caption) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA IMPRIMIR", 16
    Exit Sub
End If
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
If Dir(Ruta & fl & ser & Label5.Caption & ".lgr") = "" Then
    MsgBox "NO EXISTE INFORMACION DE LOGROS", 64
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton1, "IMPRIMIR")
If RESP = vbYes Then
Screen.MousePointer = 11
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
Close #NAR
Printer.ScaleMode = 7
PAG = 1
Printer.CurrentY = 1
Printer.CurrentX = 6
Printer.Print "LISTADO DE LOGROS PENDIENTES PERIODO " & Combo1.Text
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "Pág." & PAG
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print ini.nombre;
Printer.CurrentX = 11
Printer.Print "AREA: " & Combo3.Text
Printer.CurrentX = 0.5
Printer.Print "GRUPO: " & Combo2.Text;
Printer.CurrentX = 17
Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print "APELLIDOS Y NOMBRES";
Printer.CurrentX = 7
Printer.Print "No.";
Printer.CurrentX = 7.5
Printer.Print "LOGROS PENDIENTES"
Printer.Print ""
z = 1
For I = 1 To Val(Label4.Caption)
    'MATI50.Row = I
    h = 0
    For J = 1 To 10
        'MATI50.Col = J
        If MATI50.TextMatrix(I, J) = "" Then
            h = h + 1
        End If
    Next J
    If h <> 10 Then
        Printer.CurrentX = 0.5
        'MATI50.Col = 11
        Printer.Print MATI50.TextMatrix(I, 11);
        For r = 1 To 10
            'MATI50.Col = r
            If MATI50.TextMatrix(I, r) <> "" Then
            Open Ruta & fl & ser & Label5.Caption & ".lgr" For Random As #NAR Len = Len(logru)
            Get #NAR, Val(MATI50.TextMatrix(I, r)), logru
            Close #NAR
            If (z Mod 66) = 0 Then
                Printer.NewPage
                PAG = PAG + 1
                Printer.CurrentY = 1
                Printer.CurrentX = 6
                Printer.Print "LISTADO DE LOGROS PENDIENTES PERIODO " & Combo1.Text
                Printer.CurrentY = 1.5
                Printer.CurrentX = 19
                Printer.Print "Pág." & PAG
                Printer.Print ""
                Printer.CurrentX = 0.5
                Printer.Print ini.nombre;
                Printer.CurrentX = 11
                Printer.Print "AREA: " & Combo3.Text
                Printer.CurrentX = 0.5
                Printer.Print "GRUPO: " & Combo2.Text;
                Printer.CurrentX = 17
                Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
                Printer.Print ""
                Printer.CurrentX = 0.5
                Printer.Print "APELLIDOS Y NOMBRES";
                Printer.CurrentX = 7
                Printer.Print "No.";
                Printer.CurrentX = 7.5
                Printer.Print "LOGROS PENDIENTES"
                Printer.Print ""
            End If
            Printer.CurrentX = 7
            Printer.Print MATI50.TextMatrix(I, r);
            Printer.CurrentX = 7.5
            Printer.Print RTrim(logru.observ)
            z = z + 1
            End If
        Next r
        Printer.Print ""
        z = z + 1
    End If
Next I
Printer.EndDoc
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
'dim mate As infomater
'dim clavv As CLAVEPRO
'dim argra As areagr
MATI50.Row = 0
MATI50.Col = 0
MATI50.ColWidth(0) = 500
MATI50.Text = "COD"
MATI50.Col = 1
MATI50.ColWidth(1) = 400
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "OB1"
MATI50.Col = 2
MATI50.ColWidth(2) = 400
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "OB2"
MATI50.Col = 3
MATI50.ColWidth(3) = 400
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "OB3"
MATI50.Col = 4
MATI50.ColWidth(4) = 400
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "OB4"
MATI50.Col = 5
MATI50.ColWidth(5) = 400
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "OB5"
MATI50.Col = 6
MATI50.ColWidth(6) = 400
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "OB6"
MATI50.Col = 7
MATI50.ColWidth(7) = 400
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "OB7"
MATI50.Col = 8
MATI50.ColWidth(8) = 400
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "OB8"
MATI50.Col = 9
MATI50.ColWidth(9) = 400
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "OB9"
MATI50.Col = 10
MATI50.ColWidth(10) = 500
MATI50.CellForeColor = RGB(255, 255, 255)
MATI50.CellBackColor = RGB(0, 0, 150)
MATI50.Text = "OB10"
MATI50.Col = 11
MATI50.ColWidth(11) = 4200
MATI50.Text = "APELLIDOS Y NOMBRES"
MATI50.Col = 12
MATI50.ColWidth(12) = 1200
MATI50.Text = "No.CARNET"
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") Then
    Command1.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
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
            For I = 0 To (Combo2.ListCount - 1)
                If Combo2.List(I) = RTrim(argra.nom_grup) Then
                    VALI2 = True
                    Exit For
                End If
            Next I
            If VALI2 = False Then
                Combo2.AddItem RTrim(argra.nom_grup)
            End If
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            NAR = NAR - 1
            VALI2 = False
            For I = 0 To (Combo3.ListCount - 1)
                If Combo3.List(I) = RTrim(mate.nom) Then
                    VALI2 = True
                    Exit For
                End If
            Next I
            If VALI2 = False Then
                Combo3.AddItem RTrim(mate.nom)
            End If
        End If
    Wend
    Close #NAR
    Combo2.Text = Combo2.List(0)
    Combo3.Text = Combo3.List(0)
    If (RTrim(Combo2.Text) = "") Or (RTrim(Combo3.Text) = "") Then
        Command1.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
    End If
Else
    Command1.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
End If
Label4.Caption = 0
VALI44 = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If VALI44 = False Then
   Call Command3_Click
   If Err.Number <> 0 Then
         Cancel = 1
         Exit Sub
   End If
   Unload Me
Else
   Unload Me
End If
End Sub

Private Sub MATI50_click()
'dim logru As logris
MATI50.ToolTipText = ""
If MATI50.Col > 0 And MATI50.Col < 11 And MATI50.Row > 0 Then
   If MATI50.Text <> "" Then
      On Error Resume Next
      Err.Clear
      If Dir(Ruta & "inicial.edu") = "" Then
        MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
        End
      End If
      If Dir(Ruta & fl & ser & Label5.Caption & ".lgr") <> "" Then
         NAR = FreeFile
         Open Ruta & fl & ser & Label5.Caption & ".lgr" For Random As #NAR Len = Len(logru)
         Get #NAR, MATI50.Text, logru
         Close #NAR
         MATI50.ToolTipText = "(" & RTrim(logru.indicador) & ") " & RTrim(logru.observ)
      End If
   End If
End If
End Sub

Private Sub MATI50_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   If MATI50.Col > 0 And MATI50.Col < 11 And MATI50.Row > 0 Then
      If MATI50.Text <> "" Then
         MATI50.Text = ""
         VALI44 = False
      End If
   End If
End If
End Sub
