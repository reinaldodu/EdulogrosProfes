VERSION 5.00
Begin VB.Form DataCopia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione el directorio donde va a copiar los datos"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5430
   Icon            =   "DataCopiar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5175
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "DataCopia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Err.Clear
RutaDir = Dir1.Path
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
If Dir(RutaDir & "\datos", vbDirectory) <> "" Then
    MsgBox "Ya existen datos creados en este directorio, seleccione otro directorio para hacer la copia", 48, "Copiar datos"
    Exit Sub
End If
If Dir("c:\datos", vbDirectory) <> "" Then
    MsgBox "Necesita borrar o mover primero el directorio C:\DATOS\ para realizar la copia", 48, "Copiar datos"
    Exit Sub
End If
RESP = MsgBox("DESEA COPIAR LA INFORMACIÓN EN EL DIRECTORIO " & RutaDir, vbYesNo + vbQuestion + vbDefaultButton1, "COPIAR INFORMACION")
If RESP = vbYes Then
    Screen.MousePointer = 11
    'CREA EL DIRECTORIO DATOS
    MkDir RutaDir & "\DATOS"
    '*******HACER COPIA TEMPORAL EN C: **********
    MkDir "C:\DATOS"
    FileCopy Ruta & "INICIAL.EDU", "C:\DATOS\INICIAL.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "INFCUR.EDU", "C:\DATOS\INFCUR.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "PRINPRO.EDU", "C:\DATOS\PRINPRO.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "PRINALU.EDU", "C:\DATOS\PRINALU.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "CONT.EDU", "C:\DATOS\CONT.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "MATERIA.EDU", "C:\DATOS\MATERIA.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "AREAGRA.EDU", "C:\DATOS\AREAGRA.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    FileCopy Ruta & "CONTPRO.EDU", "C:\DATOS\CONTPRO.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    'Copia archivo de configuración de desempeños
    FileCopy Ruta & "CONF_DESEMP.EDU", "C:\DATOS\CONF_DESEMP.EDU"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
'    If Dir(Ruta & "RETIALU.EDU") <> "" Then
'        FileCopy Ruta & "RETIALU.EDU", "C:\DATOS\RETIALU.EDU"
'        If Err.Number <> 0 Then
'            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'    End If

    'SI EXISTE ARCHIVO DE CONFIGURACIÓN DE PORCENTAJES DE LOGROS, LO COPIA
    If Dir(Ruta & "conf_logro.edu") <> "" Then
        FileCopy Ruta & "conf_logro.edu", "C:\DATOS\conf_logro.edu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If

    FileCopy Ruta & "webhelp.txt", "C:\DATOS\webhelp.txt"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    que = 0
    Open Ruta & "CLAPRO.EDU" For Random As #NAR Len = Len(clavv)
    While Not EOF(NAR)
        que = que + 1
        Get #NAR, que, clavv
        If clavv.NUMERO = Val(MENUPROFE.LBLNumProfe.Caption) Then
            GoTo SAIRS
        End If
    Wend
SAIRS:
    Close #NAR
    Open "C:\DATOS\CLAPRO.EDU" For Random As #NAR Len = Len(clavv)
    Put #NAR, 1, clavv
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Close #NAR
        Exit Sub
    End If
    Close #NAR
    
    'COPIA ARCHIVO DE BLOQUEO DE LOGROS
    'Open "C:\DATOS\periodosL.edu" For Output As #NAR
    'Write #NAR, "1", "1", "1", "1", "1"
    'Close #NAR
    FileCopy Ruta & "periodosL.edu", "C:\DATOS\periodosL.edu"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'COPIA ARCHIVO DE BLOQUEO DE DESEMPEÑOS
    'Open "C:\DATOS\periodosD.edu" For Output As #NAR
    'Write #NAR, "1", "1", "1", "1", "1"
    'Close #NAR
    FileCopy Ruta & "periodosD.edu", "C:\DATOS\periodosD.edu"
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    ' COPIAR LOS GRUPOS EN DONDE DA CLASE EL PROFESOR
    CERD = 0
    NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        CERD = CERD + 1
        Get #NAR, CERD, argra
        If argra.num_pro = Val(MENUPROFE.LBLNumProfe.Caption) Then
            FileCopy Ruta & RTrim(argra.nom_grup) & ".gru", "C:\DATOS\" & RTrim(argra.nom_grup) & ".gru"
            If Err.Number <> 0 Then
                MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                Screen.MousePointer = 0
                Close #NAR
                Exit Sub
            End If
        End If
    Wend
    Close #NAR
    ' SI ES DIRECTOR DE GRUPO COPIA EL ARCHIVO DEL GRUPO
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        If icur.director = Val(MENUPROFE.LBLNumProfe.Caption) Then
            If Dir(Ruta & RTrim(icur.nom) & ".gru") <> "" Then
                FileCopy Ruta & RTrim(icur.nom) & ".gru", "C:\DATOS\" & RTrim(icur.nom) & ".gru"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
        End If
    Wend
    Close #NAR
    'COPIAR ARCHIVOS DE LOGROS, DESEMPEÑOS Y NOTAS
    For lw = 1 To 4
        CERD = 0
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
            CERD = CERD + 1
            Get #NAR, CERD, argra
            If argra.num_pro = Val(MENUPROFE.LBLNumProfe.Caption) Then
                If Dir(Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".dsp") <> "" Then
                    FileCopy Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".dsp", "C:\DATOS\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".dsp"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                End If
                If Dir(Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs") <> "" Then
                    FileCopy Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs", "C:\DATOS\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                End If
                If Dir(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") <> "" Then
                    If FileLen(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") <> 0 Then
                        FileCopy Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr", "C:\DATOS\" & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr"
                        If Err.Number <> 0 Then
                            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                            Screen.MousePointer = 0
                            Close #NAR
                            Exit Sub
                        End If
                    End If
                End If
                
                If Dir(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ptj") <> "" Then
                    'If FileLen(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ptj") <> 0 Then
                        FileCopy Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ptj", "C:\DATOS\" & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".ptj"
                        If Err.Number <> 0 Then
                            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                            Screen.MousePointer = 0
                            Close #NAR
                            Exit Sub
                        End If
                    'End If
                End If
                
            End If
        Wend
        Close #NAR
    Next lw
    '***** COPIAR DATOS DE C:\DATOS A RUTADIR + \DATOS ******
    NGRA = Dir("C:\DATOS\*.*")
    Do While NGRA <> ""
        FileCopy "C:\DATOS\" & NGRA, RutaDir & "\DATOS\" & NGRA
        Kill "C:\DATOS\" & NGRA
        NGRA = Dir
    Loop
    RmDir ("C:\DATOS")
    MsgBox "LA INFORMACIÓN SE COPIÓ CON EXITO", 64, "Copiar Datos"
    Screen.MousePointer = 0
    Unload Me
End If
End Sub

Private Sub Dir1_Change()
Frame1.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
