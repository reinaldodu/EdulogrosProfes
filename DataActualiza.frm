VERSION 5.00
Begin VB.Form DataActualiza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione el directorio para actualizar los datos del sistema central"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7095
   Icon            =   "DataActualiza.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "DataActualiza.frx":0442
      Left            =   5040
      List            =   "DataActualiza.frx":0455
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   6615
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO:"
      Height          =   195
      Left            =   4080
      TabIndex        =   3
      Top             =   360
      Width           =   780
   End
End
Attribute VB_Name = "DataActualiza"
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
'SE VERIFICA QUE LA CARPETA TENGA DATOS DEL SISTEMA EDULOGROS
If Dir(RutaDir & "\inicial.edu") = "" Then
    MsgBox "LA RUTA SELECCIONADA NO CONTIENE DATOS DE NOTAS PARA DESCARGAR, VERIFIQUE NUEVAMENTE.", 32, "ADVERTENCIA"
    Exit Sub
End If
If Combo1.Text = "PRIMERO" Then
    lw = 1
End If
If Combo1.Text = "SEGUNDO" Then
    lw = 2
End If
If Combo1.Text = "TERCERO" Then
    lw = 3
End If
If Combo1.Text = "CUARTO" Then
    lw = 4
End If
If Combo1.Text = "FINAL" Then
    lw = 5
End If
'SI RUTA ORIGEN Y RUTA DESTINO SON IGUALES NO HACER LA ACTUALIZACION
If Ruta = RutaDir Then
    MsgBox "EL DIRECTORIO ORIGEN Y DESTINO NO PUEDEN SER LOS MISMOS", 16, "Actualizar datos"
    Exit Sub
End If
'SE VERIFICA SI ESTÁ ACTIVO EL PERIODO PARA AGREGAR LOGROS
If VeriPeriodo(lw, "periodosL") = False Then
    MsgBox "NO PUEDE ACTUALIZAR LOS DATOS YA QUE LOS LOGROS DE ESTE PERIODO SE ENCUENTRAN BLOQUEADOS, COMUNIQUESE CON EL ADMINISTRADOR DEL SISTEMA", 16, "Actualizar datos"
    Exit Sub
End If
'SE VERIFICA SI ESTÁ ACTIVO EL PERIODO PARA AGREGAR DESEMPEÑOS
If VeriPeriodo(lw, "periodosD") = False Then
    MsgBox "NO PUEDE ACTUALIZAR LOS DATOS YA QUE LOS DESEMPEÑOS DE ESTE PERIODO SE ENCUENTRAN BLOQUEADOS, COMUNIQUESE CON EL ADMINISTRADOR DEL SISTEMA", 16, "Actualizar datos"
    Exit Sub
End If

'SE VERIFICA SI LOS DATOS SELECCIONADOS CORRESPONDEN AL PROFESOR QUE INGRESO
Open RutaDir & "\CLAPRO.EDU" For Random As #NAR Len = Len(clavv)
Get #NAR, 1, clavv
Close #NAR
If clavv.NUMERO <> Val(MENUPROFE.LBLNumProfe.Caption) Then
    MsgBox "LOS DATOS SELECCIONADOS NO CORRESPONDEN A LA CUENTA DE PROFESOR INGRESADA", 48, "ADVERTENCIA"
    Exit Sub
End If
RESP = MsgBox("DESEA ACTUALIZAR LOS DATOS DEL SISTEMA DESDE " & RutaDir, vbYesNo + vbQuestion + vbDefaultButton1, "Actualizar datos")
If RESP = vbYes Then
    Screen.MousePointer = 11
    CERD = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
    CERD = CERD + 1
    Get #NAR, CERD, argra
    If (argra.num_pro) = Val(MENUPROFE.LBLNumProfe.Caption) Then
        'OBTENER EL NOMBRE DE LA MATERIA
        NAR = FreeFile
        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
        Get #NAR, argra.num_area, mate
        Close #NAR
        NAR = NAR - 1
        'COPIA OBSERVACIONES
        If Dir(RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS") = "" Then
            RESP = MsgBox("Grupo " & Format(RTrim(argra.nom_grup), "<") & " no tiene observaciones para el área " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "info incompleta")
            If RESP = vbYes Then
                GoTo vaconti
            Else
                Close #NAR
                Screen.MousePointer = 0
                Exit Sub
            End If
        Else
            'If FileLen("A:\DATOS\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS") = 0 Then
            If FileLen(RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS") = 0 Then
                RESP = MsgBox("Grupo " & Format(RTrim(argra.nom_grup), "<") & " no tiene observaciones para el área " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "info incompleta")
                If RESP = vbYes Then
                    GoTo vaconti
                Else
                    Close #NAR
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            'FileCopy "A:\DATOS\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS", Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS"
            FileCopy RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS", Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".OBS"
            If Err.Number <> 0 Then
                MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                Screen.MousePointer = 0
                Close #NAR
                Exit Sub
            End If
        End If
vaconti:
        'COPIA DESEMPEÑOS
        If Dir(RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".DSP") = "" Then
            RESP = MsgBox("Grupo " & Format(RTrim(argra.nom_grup), "<") & " no tiene desempeños para el área " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "info incompleta")
            If RESP = vbYes Then
                GoTo vaconti2
            Else
                Close #NAR
                Screen.MousePointer = 0
                Exit Sub
            End If
        Else
            If FileLen(RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".DSP") = 0 Then
                RESP = MsgBox("Grupo " & Format(RTrim(argra.nom_grup), "<") & " no tiene desempeños para el área " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "info incompleta")
                If RESP = vbYes Then
                    GoTo vaconti2
                Else
                    Close #NAR
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            FileCopy RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".DSP", Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".DSP"
            If Err.Number <> 0 Then
                MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                Screen.MousePointer = 0
                Close #NAR
                Exit Sub
            End If
        End If
vaconti2:
        'COPIAR LOGROS
        fl = Left(argra.nom_grup, 1)
        'If Check1.Value = 1 Then
            'If Dir("A:\DATOS\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR") = "" Then
            If Dir(RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR") = "" Then
                RESP = MsgBox("No existen logros u observaciones para el grado " & Format(RTrim(argra.grado), "<") & " área " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "info incompleta")
                If RESP = vbYes Then
                    GoTo vaconti3
                Else
                    Close #NAR
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            Else
                'If FileLen("A:\DATOS\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR") = 0 Then
                If FileLen(RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR") = 0 Then
                    RESP = MsgBox("No existen logros u observaciones para el grado " & Format(RTrim(argra.grado), "<") & " área " & Trim(mate.nom) & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "info incompleta")
                    If RESP = vbYes Then
                        GoTo vaconti3
                    Else
                        Close #NAR
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
                'FileCopy "A:\DATOS\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR", Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR"
                FileCopy RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR", Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".LGR"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
        'End If
vaconti3:
        'COMENTARIOS EN EL INFORME FINAL.
        If Dir(RutaDir & "\LRF" & RTrim(argra.nom_grup) & ".LRF") <> "" Then
            FileCopy RutaDir & "\LRF" & RTrim(argra.nom_grup) & ".LRF", Ruta & "LRF" & RTrim(argra.nom_grup) & ".LRF"
            If Err.Number <> 0 Then
                MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                Screen.MousePointer = 0
                Close #NAR
                Exit Sub
            End If
        End If
        If Dir(RutaDir & "\ORF" & RTrim(argra.nom_grup) & ".ORF") <> "" Then
            FileCopy RutaDir & "\ORF" & RTrim(argra.nom_grup) & ".ORF", Ruta & "ORF" & RTrim(argra.nom_grup) & ".ORF"
            If Err.Number <> 0 Then
                MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                Screen.MousePointer = 0
                Close #NAR
                Exit Sub
            End If
        End If
        
        'COPIA ARCHIVOS DE PORCENTAJES DE LOGROS SI EXISTEN
        If Dir(RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".PTJ") <> "" Then
            FileCopy RutaDir & "\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".PTJ", Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".PTJ"
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
    'GUARDAR REGISTRO (LOG) DE ACTUALIZACION DE NOTAS EN EL SISTEMA
    ifnt.numprofe = Val(MENUPROFE.LBLNumProfe.Caption)
    ifnt.periodo = Combo1.Text
    ifnt.fecha = Date
    ifnt.hora = Time
    Open Ruta & "infnota.edu" For Append As #NAR
    Write #NAR, ifnt.numprofe, ifnt.periodo, ifnt.fecha, ifnt.hora
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Close #NAR
        Exit Sub
    End If
    Close #NAR
    MsgBox "SISTEMA ACTUALIZADO EXITOSAMENTE", 48, "Actualizar Datos"
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

Private Sub Form_Load()
Combo1 = Combo1.List(0)
End Sub
