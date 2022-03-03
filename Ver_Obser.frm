VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Ver_Obser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Consultar observaciones"
   ClientHeight    =   1470
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   9795
   Icon            =   "Ver_Obser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MConObs 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   2
   End
End
Attribute VB_Name = "Ver_Obser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MConObs.Row = 0
MConObs.Col = 0
MConObs.ColWidth(0) = 400
MConObs.CellForeColor = RGB(255, 255, 255)
MConObs.CellBackColor = RGB(0, 0, 150)
MConObs.Text = "No."
MConObs.Col = 1
MConObs.ColWidth(1) = 400
MConObs.CellForeColor = RGB(255, 255, 255)
MConObs.CellBackColor = RGB(0, 0, 150)
MConObs.Text = "IND"
MConObs.Col = 2
MConObs.ColWidth(2) = 20000
MConObs.CellForeColor = RGB(255, 255, 255)
MConObs.CellBackColor = RGB(0, 0, 150)
MConObs.Text = "OBSERVACION"
MConObs.ScrollBars = flexScrollBarVertical
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
Ver_Ini = 0
NAR = FreeFile
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    Ver_Ini = Ver_Ini + 1
    Get #NAR, Ver_Ini, logru
Wend
Close #NAR
Ver_Obser.Caption = "Consulta de indicadores"
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
t = 1
For J = 1 To (Ver_Ini - 1)
    Get #NAR, J, logru
    If SWobserv = True Then
        If Trim(logru.indicador) <> "L" Then
            MConObs.Rows = t + 1
            MConObs.TextMatrix(t, 0) = t
            MConObs.TextMatrix(t, 1) = logru.indicador
            MConObs.TextMatrix(t, 2) = logru.observ
            t = t + 1
        End If
    Else
        If Trim(logru.indicador) = "L" Then
            MConObs.Rows = t + 1
            MConObs.TextMatrix(t, 0) = t
            MConObs.TextMatrix(t, 1) = logru.indicador
            MConObs.TextMatrix(t, 2) = logru.observ
            t = t + 1
        End If
    End If
Next J
Close #NAR
End Sub

Private Sub Form_Resize()
If Ver_Obser.Width < 1100 Then Exit Sub
MConObs.ColWidth(2) = Ver_Obser.Width - 1100
MConObs.Height = Ver_Obser.Height - 500
MConObs.Width = Ver_Obser.Width - 200
End Sub

Private Sub MConObs_Click()
If MConObs.Row > 0 Then
    MConObs.ToolTipText = Left(RTrim(MConObs.Text), 200)
End If
End Sub
