VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MENUPROFE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edulogros - Profes"
   ClientHeight    =   4125
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3315
   Icon            =   "MENUPROFE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "glogros"
            Object.ToolTipText     =   "GRABAR LOGROS"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "gdesemp"
            Object.ToolTipText     =   "GRABAR DESEMPEÑOS"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "gobserv"
            Object.ToolTipText     =   "GRABAR OBSERVACIONES"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "gporcent"
            Object.ToolTipText     =   "CONSULTAR PORCENTAJES Y DESEMPEÑOS"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ggrupo"
            Object.ToolTipText     =   "CONSULTAR GRUPO"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3870
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Object.ToolTipText     =   "PROFES - EDULOGROS V.2.5"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "03:40 p.m."
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   480
      Picture         =   "MENUPROFE.frx":0442
      ScaleHeight     =   2715
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   840
      Width           =   2415
      Begin ComctlLib.ImageList ImageList1 
         Left            =   480
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   9
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENUPROFE.frx":33B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENUPROFE.frx":36D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENUPROFE.frx":39EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENUPROFE.frx":3D06
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENUPROFE.frx":4020
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENUPROFE.frx":433A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENUPROFE.frx":4654
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENUPROFE.frx":496E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENUPROFE.frx":4C88
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   3000
   End
   Begin VB.Label LBLNumProfe 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2760
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Menu prograb 
      Caption         =   "&Grabar"
      Begin VB.Menu prologros 
         Caption         =   "&Logros"
         Shortcut        =   {F2}
      End
      Begin VB.Menu graba_PtjLgr 
         Caption         =   "Porcentajes de logros"
      End
      Begin VB.Menu liporlog 
         Caption         =   "-"
      End
      Begin VB.Menu proobservv 
         Caption         =   "&Desempeños"
         Shortcut        =   {F3}
      End
      Begin VB.Menu probole 
         Caption         =   "&Observaciones"
         Shortcut        =   {F4}
      End
      Begin VB.Menu corte4 
         Caption         =   "-"
      End
      Begin VB.Menu plansem 
         Caption         =   "Planeación"
         Begin VB.Menu memejes 
            Caption         =   "Ejes temáticos y contenidos"
         End
         Begin VB.Menu mencompetencias 
            Caption         =   "Competencias"
         End
         Begin VB.Menu ln_plans 
            Caption         =   "-"
         End
         Begin VB.Menu plansem2 
            Caption         =   "Planeación semanal"
         End
      End
      Begin VB.Menu corte5 
         Caption         =   "-"
      End
      Begin VB.Menu poyecprofes 
         Caption         =   "Proyectos"
         Visible         =   0   'False
      End
      Begin VB.Menu lineproy 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu grainfresu 
         Caption         =   "&Comentarios informe final"
      End
      Begin VB.Menu prosalsals 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu procons 
      Caption         =   "&Consultar"
      Begin VB.Menu progru 
         Caption         =   "&Grupo"
      End
      Begin VB.Menu evapper 
         Caption         =   "&Reporte de desempeños"
      End
      Begin VB.Menu RepLogPerd 
         Caption         =   "Reporte Logros perdidos y reaprendizaje"
      End
      Begin VB.Menu mateprocon 
         Caption         =   "&Carga académica"
      End
   End
   Begin VB.Menu prodata 
      Caption         =   "Datos"
      Begin VB.Menu copiadata 
         Caption         =   "Copiar datos"
      End
      Begin VB.Menu actualizadata 
         Caption         =   "Actualizar datos"
      End
   End
   Begin VB.Menu ayuprof 
      Caption         =   "Ay&uda"
      Begin VB.Menu ayuyupro 
         Caption         =   "Edulogros en Internet"
         Shortcut        =   {F1}
      End
      Begin VB.Menu linayupr 
         Caption         =   "-"
      End
      Begin VB.Menu acerpro 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu varipro 
      Caption         =   "variospro"
      Visible         =   0   'False
      Begin VB.Menu vagrabole 
         Caption         =   "Grabar &Logros"
      End
      Begin VB.Menu vagrades 
         Caption         =   "Grabar &Desempeños"
      End
      Begin VB.Menu vagraobs 
         Caption         =   "Grabar &Observaciones"
      End
      Begin VB.Menu conpromat 
         Caption         =   "Consultar &Materias"
      End
      Begin VB.Menu vaconsgru 
         Caption         =   "Consultar &Grupo"
      End
   End
End
Attribute VB_Name = "MENUPROFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub acerpro_Click()
ACERCADE.Show 1
End Sub

Private Sub actualizadata_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    On Error Resume Next
    Err.Clear
    If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
        MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
        End
    End If
    'NO DEJA ACTUALIZAR SISTEMA SINO SE ENCUENTRA DENTRO DE LA RED DEL COLEGIO
    If Dir(Ruta & "leyenda.edu") = "" Then
        MsgBox "DEBE INGRESAR DESDE LA RED DEL COLEGIO PARA ACTUALIZAR LOS DATOS EN EL SISTEMA", 16, "Actualizar datos"
        Exit Sub
    End If
    DataActualiza.Show 1
End If
End Sub

Private Sub ayuyupro_Click()
If Dir("C:\Archivos de programa\Internet Explorer\IEXPLORE.EXE") = "" Then
    MsgBox "EL NAVEGADOR WEB NO PUEDE INICIARSE, VERIFIQUE SU CONFIGURACIÓN", 48
Else
    If Dir(Ruta & "webhelp.txt") = "" Then
        MsgBox "NO EXISTE EL ARCHIVO DE LA PAGINA WEB", 64
    Else
        NAR = FreeFile
        Open Ruta & "webhelp.txt" For Input As #NAR
        Input #NAR, TTT
        Close #NAR
        Shell "C:\Archivos de programa\Internet Explorer\IEXPLORE.EXE " & TTT, vbNormalFocus
    End If
End If
End Sub

Private Sub conpromat_Click()
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
CONS_MATER.Show
End Sub

Private Sub copiadata_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    On Error Resume Next
    Err.Clear
    If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
        MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
        End
    End If
    'NO DEJA COPIAR DATOS SINO SE ENCUENTRA DENTRO DE LA RED DEL COLEGIO
    If Dir(Ruta & "leyenda.edu") = "" Then
        MsgBox "DEBE INGRESAR DESDE LA RED DEL COLEGIO PARA COPIAR LOS DATOS DEL SISTEMA", 16, "Copiar datos"
        Exit Sub
    End If
    DataCopia.Show 1
End If
End Sub

Private Sub evapper_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
REPORTE_PORCENT.Show
End Sub

Private Sub Form_Activate()
StatusBar1.Panels.Item(1).Text = Format(Date, "mmm/dd/yyyy")
End Sub

Private Sub form_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 2 Then
MENUPROFE.PopupMenu varipro
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RESP = MsgBox("DESEA SALIR DEL PROGRAMA PROFES-EDULOGROS?", vbYesNo + vbQuestion + vbDefaultButton1, "SALIR")
If RESP = vbYes Then
End
Else
Cancel = True
End If
End Sub

Private Sub graba_PtjLgr_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
If Dir(Ruta & "conf_logro.edu") <> "" Then
    Open Ruta & "conf_logro.edu" For Input As #NAR
    Input #NAR, ConfLgr
    Close #NAR
    If ConfLgr = 0 Then
        MsgBox "No se pueden grabar porcentajes de logros.  El sistema está configurado para obtener los porcentajes de forma automática.", 64, "ADVERTENCIA"
        Exit Sub
    End If
Else
    MsgBox "No se pueden grabar porcentajes de logros.  El sistema está configurado para obtener los porcentajes de forma automática.", 64, "ADVERTENCIA"
    Exit Sub
End If
Porcentaje_Logros.Show
End Sub

Private Sub grainfresu_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
INFREFINAL.Show
End Sub

Private Sub mateprocon_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
CONS_MATER.Show
End Sub

Private Sub notasprof_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
OBSER_ALUM.Show
End Sub

Private Sub memejes_Click()
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
Ejes_Contenidos.Show

End Sub

Private Sub mencompetencias_Click()
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
Competencias.Show
End Sub

Private Sub plansem2_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
planeacion_semanal.Show
End Sub

Private Sub poyecprofes_Click()
proyectos.Show
End Sub

Private Sub probole_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
GRABAR_OBSER.Show
End Sub

Private Sub progru_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
CONS_GRUP.Show
End Sub

Private Sub proimpgru_Click()
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
IMP_GRUP.Show
End Sub

Private Sub prologra_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
LOGRO_PEN.Show
End Sub

Private Sub prologros_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
COPEGA.Show 1
End Sub

Private Sub proobservv_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
GRABA_DESEMP.Show
End Sub

Private Sub prosalsals_Click()
End
End Sub

Private Sub RepLogPerd_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
Reporte_Logros.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
Case "glogros"
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
COPEGA.Show 1
Case "gdesemp"
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
GRABA_DESEMP.Show
Case "gobserv"
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
GRABAR_OBSER.Show
Case "gporcent"
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
REPORTE_PORCENT.Show
Case "ggrupo"
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
CONS_GRUP.Show
End Select
End Sub

Private Sub vaconsgru_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
CONS_GRUP.Show
End Sub

Private Sub vagrabole_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
COPEGA.Show 1
End Sub

Private Sub vagrades_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
GRABA_DESEMP.Show
End Sub
Private Sub vagraobs_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
GRABAR_OBSER.Show
End Sub


Private Sub vaimprigrup_Click()
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
IMP_GRUP.Show
End Sub

Private Sub valogrpend_Click()
On Error Resume Next
Err.Clear
If (Dir(Ruta & "inicial.edu") = "") And (Dir(Ruta & "infcur.edu") = "") And (Dir(Ruta & "materia.edu") = "") Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
LOGRO_PEN.Show
End Sub
