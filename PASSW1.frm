VERSION 5.00
Begin VB.Form PASSW1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contraseña de acceso"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "PASSW1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "PROFES - EDULOGROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.TextBox Text2 
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
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PROFESOR No."
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CONTRASEÑA:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1155
      End
   End
End
Attribute VB_Name = "PASSW1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'dim clavv As CLAVEPRO
Screen.MousePointer = 11
If Text2.Text = "" Then
    MsgBox "ESCRIBA EL No. DE PROFESOR", 48, "CONTRASEÑA"
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
If Text1.Text = "" Then
    MsgBox "ESCRIBA EL PASSWORD", 48, "CONTRASEÑA"
    Text1.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If

On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    Screen.MousePointer = 0
    End
End If

If acceso(Text2.Text, Text1.Text) = True Then
    NAR = FreeFile
    Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
    Get #NAR, Val(Text2.Text), profe
    Close #NAR
    MENUPROFE.LBLNumProfe.Caption = Text2.Text
    MENUPROFE.Label1.Caption = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
    Unload Me
    MENUPROFE.Show
Else
    MsgBox "CONTRASEÑA INCORRECTA", 64, "CONTRASEÑA"
    Screen.MousePointer = 0
    End
End If
' REGISTRO DE ACCESO EN LA BITACORA
proflogs.numprofe = Val(MENUPROFE.LBLNumProfe.Caption)
proflogs.fecha = Date
proflogs.hora = Time
Open Ruta & "logs.edu" For Append As #NAR
Write #NAR, proflogs.numprofe, proflogs.fecha, proflogs.hora
If Err.Number <> 0 Then
    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
    Screen.MousePointer = 0
    Close #NAR
    Exit Sub
End If
Close #NAR
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
Err.Clear
'SI LA CARPETA DATOS EXISTE, CREA AUTOMATICAMENTE EL ARCHIVO BD.TXT CON LA RUTA
If Dir(App.Path & "\datos", vbDirectory) <> "" Then
    If Dir(App.Path & "\BD.txt") <> "" Then
        Kill App.Path & "\BD.txt"
    End If
    NAR = FreeFile
    Open App.Path & "\BD.TXT" For Output As #NAR
    Write #NAR, App.Path & "\datos\"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Close #NAR
        Exit Sub
    End If
    Close #NAR
    Ruta = App.Path & "\datos\"
Else
'SI NO EXISTE LA CARPETA DATOS VERIFICA RUTA EN BD.TXT O SOLICITA LA RUTA DE ACCESO A LOS DATOS
    If Dir(App.Path & "\BD.txt") = "" Then
        MsgBox "NO EXISTE LA CARPETA DE DATOS, SELECCIONE LA CARPETA QUE CONTIENE LOS DATOS DEL SISTEMA", 48, "DATOS DEL SISTEMA"
        'End
        DataBusca.Show
        Unload Me
        Exit Sub
    Else
        NAR = FreeFile
        Open (App.Path & "\BD.txt") For Input As #NAR
        Input #NAR, Ruta
        Close #NAR
        
        If Dir(Ruta & "inicial.edu") = "" Then
            MsgBox "LA RUTA DE ACCESO A LOS DATOS A CAMBIADO, SELECCIONE LA CARPETA QUE CONTIENE LOS DATOS DEL SISTEMA", 48, "DATOS DEL SISTEMA"
            'Screen.MousePointer = 0
            'End
            DataBusca.Show
            Unload Me
            Exit Sub
        End If
    End If
End If
Text1.MaxLength = 15
Text2.MaxLength = 3
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
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
