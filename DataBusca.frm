VERSION 5.00
Begin VB.Form DataBusca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar la ruta de datos del sistema"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6165
   Icon            =   "DataBusca.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "DataBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Err.Clear
RutaDir = Dir1.Path
'SE VERIFICA QUE LA CARPETA TENGA DATOS DEL SISTEMA EDULOGROS
If Dir(RutaDir & "\inicial.edu") = "" Then
    MsgBox "LA RUTA SELECCIONADA NO CONTIENE DATOS DE NOTAS, VERIFIQUE NUEVAMENTE.", 32, "ADVERTENCIA"
    Exit Sub
Else
    'SE VERIFICA QUE EXISTA EL ARCHIVO BD.TXT
    If Dir(App.Path & "\BD.txt") <> "" Then
        Kill App.Path & "\BD.txt"
    End If
    NAR = FreeFile
    Open App.Path & "\BD.TXT" For Output As #NAR
    Write #NAR, RutaDir & "\"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Close #NAR
        Exit Sub
    End If
    Close #NAR
    PASSW1.Show
    Unload Me
End If
End Sub

Private Sub Dir1_Change()
Frame1.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
