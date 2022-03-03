VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CONS_MATER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar carga académica"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9030
   Icon            =   "CONS_MATER.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin MSFlexGridLib.MSFlexGrid MATI6 
         Height          =   3975
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         BackColorBkg    =   -2147483633
         GridColor       =   12582912
      End
   End
End
Attribute VB_Name = "CONS_MATER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim mate As infomater
'Dim ini As inicio
If MATI6.Rows = 1 Then
    MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 32, "IMPRIMIR"
    Exit Sub
End If
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
RESP = MsgBox("DESEA IMPRIMIR LA CARGA ACADÉMICA?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
Close #NAR
Printer.ScaleMode = 7
Printer.CurrentY = 1
Printer.Font.Size = 10
Printer.CurrentX = 1
Printer.Print ini.nombre
Printer.CurrentX = 1
Printer.Print "CARGA ACADÉMICA - " & MENUPROFE.Label1.Caption
Printer.Print ""
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "ASIGNATURA";
Printer.CurrentX = 15
Printer.Print "I.H.";
Printer.CurrentX = 16
Printer.Print "GRUPO"
Printer.Print ""
For maa = 1 To (MATI6.Rows - 1)
    Printer.CurrentX = 1
    Printer.Print MATI6.TextMatrix(maa, 0);
    Printer.CurrentX = 15
    Printer.Print MATI6.TextMatrix(maa, 1);
    Printer.CurrentX = 16
    Printer.Print MATI6.TextMatrix(maa, 2)
Next maa
Printer.EndDoc
Printer.Font.Size = 8
End If
End Sub

Private Sub Form_Load()
'dim mate As infomater
'dim clavv As CLAVEPRO
'dim argra As areagr
MATI6.Rows = 1
L = 1
MATI6.Row = 0
MATI6.Col = 0
MATI6.ColWidth(0) = 5500
MATI6.CellForeColor = RGB(255, 255, 255)
MATI6.CellBackColor = RGB(0, 0, 150)
MATI6.Text = "ASIGNATURA"
MATI6.Col = 1
MATI6.ColWidth(1) = 500
MATI6.CellForeColor = RGB(255, 255, 255)
MATI6.CellBackColor = RGB(0, 0, 150)
MATI6.Text = "I.H."
MATI6.Col = 2
MATI6.ColWidth(2) = 2000
MATI6.CellForeColor = RGB(255, 255, 255)
MATI6.CellBackColor = RGB(0, 0, 150)
MATI6.Text = "GRUPO"

If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") Then
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
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            NAR = NAR - 1
            MATI6.Rows = L + 1
            MATI6.TextMatrix(L, 0) = RTrim(mate.nom) & " - (" & mate.num & ")"
            MATI6.TextMatrix(L, 1) = argra.ih
            MATI6.TextMatrix(L, 2) = RTrim(argra.nom_grup)
            L = L + 1
            
'            VALI2 = False
'            For I = 1 To (MATI6.Rows - 1)
'                If MATI6.TextMatrix(I, 1) = RTrim(mate.nom) Then
'                    VALI2 = True
'                    Exit For
'                End If
'            Next I
'            If VALI2 = False Then
'                MATI6.Rows = L + 1
'                MATI6.TextMatrix(L, 0) = mate.num
'                MATI6.TextMatrix(L, 1) = RTrim(mate.nom)
'                L = L + 1
'            End If

        End If
    Wend
    Close #NAR
    If L = 1 Then
        Command1.Enabled = False
    End If
Else
    Command1.Enabled = False
End If
MATI6.Col = 0
MATI6.Sort = 3
End Sub
