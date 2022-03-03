VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form OBSER_ALUM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta observaciones de alumno"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   Icon            =   "OBSER_ALUM.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "OBSER_ALUM.frx":0442
      Left            =   8040
      List            =   "OBSER_ALUM.frx":0455
      TabIndex        =   1
      Text            =   "PRIMERO"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   9255
      Begin VB.CommandButton Command2 
         Caption         =   "Copiar"
         Height          =   315
         Left            =   8520
         TabIndex        =   6
         ToolTipText     =   "Copiar las observaciones del alumno"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7800
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   7320
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "OBSER_ALUM.frx":0483
         Left            =   4200
         List            =   "OBSER_ALUM.frx":0485
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "OBSER_ALUM.frx":0487
         Left            =   720
         List            =   "OBSER_ALUM.frx":0489
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código alumno:"
         Height          =   195
         Left            =   6120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   3480
         TabIndex        =   10
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AREA:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      Begin MSFlexGridLib.MSFlexGrid MATI24 
         Height          =   4815
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8493
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6720
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
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
      Left            =   1440
      TabIndex        =   14
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ALUMNO(A): "
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
      TabIndex        =   13
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label4 
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
      Left            =   7080
      TabIndex        =   12
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "OBSER_ALUM"
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
Text1.SetFocus
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo1.SetFocus
End If
End Sub

Private Sub Command1_Click()
'dim alumno As maestroalum
'dim alugru As grupoalu
'dim mate As infomater
'dim argra As areagr
'dim notas As notis
'dim icur As inforcur
'dim logru As logris
If Combo1.Text = "" Then
    MsgBox "ESCOJA EL AREA", 64
    Combo1.SetFocus
    Exit Sub
End If
If Combo2.Text = "" Then
    Combo2.SetFocus
    MsgBox "ESCOJA EL GRUPO", 64
    Exit Sub
End If
If Text1.Text = "" Then
    MsgBox "ESCRIBA EL CODIGO", 64
    Text1.SetFocus
    Exit Sub
End If
Screen.MousePointer = 11
MATI24.Rows = 1
Label6.Caption = ""
Frame1.Caption = ""
Label7.Caption = 0
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    Screen.MousePointer = 0
    End
End If
If Dir(Ruta & RTrim(Combo2.Text) & ".GRU") = "" Then
    MsgBox "NO LE CORRESPONDE ESTE GRUPO", 64
    Screen.MousePointer = 0
    Exit Sub
End If
NAR = FreeFile
ret = 0
Open Ruta & RTrim(Combo2.Text) & ".GRU" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
If (Val(Text1.Text) < 1) Or (Val(Text1.Text) > ret - 1) Then
    MsgBox "CODIGO DE ALUMNO NO EXISTE EN ESTE GRUPO", 64
    Screen.MousePointer = 0
    Exit Sub
End If
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(Combo2.Text) = RTrim(icur.nom) Then
        YUS = Left(icur.grado, 3)
    End If
Wend
Close #NAR
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
que = 0
While Not EOF(NAR)
    que = que + 1
    Get #NAR, que, mate
    If RTrim(mate.nom) = RTrim(Combo1.Text) Then
        uu = mate.num
    End If
Wend
Close #NAR
If RTrim(Combo3.Text) = "PRIMERO" Then
    lw = 1
End If
If RTrim(Combo3.Text) = "SEGUNDO" Then
    lw = 2
End If
If RTrim(Combo3.Text) = "TERCERO" Then
    lw = 3
End If
If RTrim(Combo3.Text) = "CUARTO" Then
    lw = 4
End If
If RTrim(Combo3.Text) = "FINAL" Then
    lw = 5
End If
If Dir(Ruta & Left(Combo2.Text, 1) & YUS & uu & lw & ".LGR") = "" Then
    MsgBox "NO EXISTE INFORMACION", 64
    Screen.MousePointer = 0
    Exit Sub
End If
If Dir(Ruta & RTrim(Combo2.Text) & uu & lw & ".obs") = "" Then
    MsgBox "NO EXISTE INFORMACION", 64
    Screen.MousePointer = 0
    Exit Sub
End If
Open Ruta & RTrim(Combo2.Text) & ".gru" For Random As #NAR Len = Len(alugru)
Get #NAR, Val(Text1.Text), alugru
Close #NAR
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, (Val(alugru.num_carnet)), alumno
Close #NAR
Label6.Caption = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " (" & alumno.n_carnet & ")"
Open Ruta & RTrim(Combo2.Text) & uu & lw & ".obs" For Random As #NAR Len = Len(notas)
que = 0
While Not EOF(NAR)
    que = que + 1
    Get #NAR, que, notas
    If notas.num_carnet = alugru.num_carnet Then
        Frame1.Caption = "JUCIO VALORATIVO: " & notas.JV & "  FALLAS: " & notas.FA
        NAR = FreeFile
        Open Ruta & Left(Combo2.Text, 1) & YUS & uu & lw & ".LGR" For Random As #NAR Len = Len(logru)
        For J = 1 To 10
            If notas.area(J) <> 0 Then
                Get #NAR, (notas.area(J)), logru
                MATI24.Rows = J + 1
                MATI24.TextMatrix(J, 0) = logru.indicador
                MATI24.TextMatrix(J, 1) = RTrim(logru.observ)
            Else
                Exit For
            End If
        Next J
        Label7.Caption = J - 1
        Close #NAR
        NAR = NAR - 1
    End If
Wend
Close #NAR
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
If Val(Label7.Caption) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA COPIAR", 16, "COPIAR"
    Exit Sub
End If
Screen.MousePointer = 11
Clipboard.Clear
cop = ""
cop = Label6.Caption & vbCrLf & Frame1.Caption & vbCrLf & vbCrLf
For X = 1 To Val(Label7.Caption)
    INDI = RTrim(MATI24.TextMatrix(X, 0))
    OB = RTrim(MATI24.TextMatrix(X, 1))
    cop = cop + INDI & " - " & OB & vbCrLf
Next X
Close #NAR
Clipboard.SetText cop
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
'dim mate As infomater
'dim clavv As CLAVEPRO
'dim argra As areagr
MATI24.Row = 0
MATI24.Col = 0
MATI24.ColWidth(0) = 400
MATI24.CellForeColor = RGB(255, 255, 255)
MATI24.CellBackColor = RGB(0, 0, 150)
MATI24.Text = "IND"
MATI24.Col = 1
MATI24.ColWidth(1) = 8200
MATI24.CellForeColor = RGB(255, 255, 255)
MATI24.CellBackColor = RGB(0, 0, 150)
MATI24.Text = "OBSERVACION"
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") Then
    Command1.Enabled = True
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
            For I = 0 To (Combo1.ListCount - 1)
                If Combo1.List(I) = RTrim(mate.nom) Then
                    VALI2 = True
                    Exit For
                End If
            Next I
            If VALI2 = False Then
                Combo1.AddItem RTrim(mate.nom)
            End If
        End If
    Wend
    Close #NAR
    Combo1.Text = Combo1.List(0)
    Combo2.Text = Combo2.List(0)
    If (RTrim(Combo1.Text) = "") Or (RTrim(Combo2.Text) = "") Then
        Command1.Enabled = False
        Command2.Enabled = False
    End If
Else
    Command1.Enabled = False
    Command2.Enabled = False
End If
Text1.MaxLength = 2
Label7.Caption = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub
