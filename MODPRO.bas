Attribute VB_Name = "MODPRO"
Option Explicit
Public I As Integer, r As Integer, h As Integer, TT As Integer, VALI As Boolean, VALI2 As Boolean, VALI4 As Boolean, leo As Integer, plo As Integer
Public J As Integer, z As Integer, Y As Integer, s As Integer, t As Integer, PAG As Integer, dire As Integer, J1 As String, J2 As String, YUS As String
Public que As Integer, maa As Integer, NA As Integer, VALI80 As Boolean, VALI180 As Boolean, VALI380 As Boolean, cona As Integer, Ver_Ini As Integer, CLO As Integer, CROA As Integer, L As Integer, PEGG As String
Public pio As Integer, ret As Integer, ape As String, nom As String, FERT As Integer, RE22 As String, lw As Integer, ser As String, YO As Integer
Public fl As String, uu As Integer, JOJI As String, NAR As Integer, J3 As Integer, VALI44 As Boolean, TN As Integer, TTT As String, rr As Integer, CP As Integer
Public alumno As maestroalum, profe As maestropro, icur As inforcur, mate As infomater, alugru As grupoalu, argra As areagr, logru As logris, notas As notis, ini As inicio, clavv As CLAVEPRO, leyfin As leyenfin, obsfin As String * 750
Public Ruta As String, confdesemp As ini_desemp, notas_desemp As porcentaje_desemp, SWobserv As Boolean, Cont_Lgr As Integer, ifnt As infornoti, proflogs As bitacora, porcent_manual As porcentaje_manual
Public proyectosaula As proyectos, semanal_planeacion As planeacion_semanal, semanal_ejetematico As eje_tematico_semanal, semanal_contenidos As contenidos_semanal, semanal_competencias As competencias_semanal
Public ValiModifica As Boolean, Mod_Vr As String

Type maestroalum
n_matricula As Integer
n_carnet As String * 8
nombres As String * 30
apellidos As String * 30
documento As String * 15
f_nacimiento As String * 10
rh As String * 4
sexo As String * 1
padre As String * 50
tel_pa As String * 20
madre As String * 50
tel_ma As String * 20
acudiente As String * 50
tel_acu As String * 20
direccion As String * 60
jornada As String * 10
año_ingre As String * 4
grado As String * 15
End Type

Type maestropro
nombres As String * 30
apellidos As String * 30
documento As String * 15
fech_nacim As String * 10
rh As String * 4
direccion As String * 60
Telefono As String * 20
año_ingre As String * 4
especiali As String * 40
escalafon As String * 2
End Type

Type inforcur
nom As String
jornada As String
grado As String
director As Integer
End Type

Type infomater
nom As String * 50
num As Integer
End Type

Type grupoalu
num_carnet As String * 5
End Type

Type areagr
grado As String * 15
num_area As Integer
ih As Integer
num_pro As Integer
nom_grup As String * 20
End Type

Type logris
indicador As String * 5
observ As String * 800
End Type

Type notis
num_carnet As String * 5
FA As Integer
area(1 To 10) As Integer
End Type

Type infornoti
numprofe As Integer
periodo As String
fecha As Date
hora As String
End Type

Type inicio
ciudad As String
nombre As String
modalidad As String
Telefono As String
Rector As String
End Type

Type CLAVEPRO
NUMERO As Integer
PASSWW As String * 15
End Type

Type leyenfin
num_carnet As String * 5
fnob(1 To 5) As Integer
End Type

Type bitacora
numprofe As Integer
fecha As Date
hora As String
End Type

Type ini_desemp
grado As String * 15
desemp(1 To 4) As String * 5
recupera(1 To 4) As String * 5
rango(1 To 3) As Byte
End Type

Type porcentaje_desemp
num_carnet As String * 5
porcentaje(1 To 10) As Byte
recuperado(1 To 10) As Boolean
logro(1 To 10) As Byte
End Type

'PORCENTAJE MANUAL DE LOGROS
Type porcentaje_manual
porcent_logro As Byte
End Type

'**********************************
'ESTRUCTURA DE DATOS -PLANEACIONES-
'**********************************

' PLANEACION SEMANAL

Type planeacion_semanal
fecha As String * 20
eje As String * 100
contenidos As String * 100
competencia As String * 100
logros As String * 100
End Type

Type eje_tematico_semanal
'num_eje As Integer
txt_eje As String * 200
End Type

Type contenidos_semanal
'num_cont As Integer
txt_cont As String * 200
num_eje As Integer
End Type

Type competencias_semanal
'num_comp As Integer
cod_comp As String * 10
txt_comp As String * 700
num_logro As String * 50
End Type

'Type logro_competencia_semanal
'num_comp As Integer
'num_logro As Integer
'End Type

' PROYECTOS

Type proyectos
nombre As String * 500
responsables As String * 1000
poblacion As String * 500
objetivos As String * 2000
Competencias As String * 5000
metas As String * 2000
ejes_tematicos As String * 2000
metodologia As String * 2000
cronograma As String * 2000
recursos As String * 2000
evaluacion As String * 2000
observaciones As String * 2000
End Type



'****Función que carga la ruta de la base de datos desde el archivo BD.txt****
'Public Function RutaBD As String
'If Dir(App.Path & "\BD.txt") = "" Then
'    MsgBox "NO EXISTE EL ARCHIVO BD.txt", 48
'Else
'    NAR = FreeFile
'    Open (App.Path & "\BD.txt") For Input As #NAR
'    Input #NAR, RutaBD
'    Close #NAR
'End If
'End Function

'función que verifica si está habilitado el periodo académico
Public Function VeriPeriodo(Num_Periodo As Integer, ArchivoBloqueo As String) As Boolean
Dim FlagPeriodo As Boolean, p1 As String, p2 As String, p3 As String, p4 As String, p5 As String
VeriPeriodo = False
NAR = FreeFile
Open Ruta & ArchivoBloqueo & ".edu" For Input As #NAR
Input #NAR, p1, p2, p3, p4, p5
Close #NAR
If p1 = "1" And Num_Periodo = 1 Then
    VeriPeriodo = True
End If
If p2 = "1" And Num_Periodo = 2 Then
    VeriPeriodo = True
End If
If p3 = "1" And Num_Periodo = 3 Then
    VeriPeriodo = True
End If
If p4 = "1" And Num_Periodo = 4 Then
    VeriPeriodo = True
End If
If p5 = "1" And Num_Periodo = 5 Then
    VeriPeriodo = True
End If
End Function

'Función que verifica el acceso de los profesores
Public Function acceso(num_profe As Integer, password_profe As String) As Boolean
Dim avanza As Integer
acceso = False
avanza = 0
NAR = FreeFile
    Open Ruta & "CLAPRO.EDU" For Random As #NAR Len = Len(clavv)
    While Not EOF(NAR)
        avanza = avanza + 1
        Get #NAR, avanza, clavv
        If (clavv.NUMERO = Val(num_profe)) And (Trim(clavv.PASSWW) = Trim(password_profe)) Then
            acceso = True
            Close #NAR
            Exit Function
        End If
    Wend
    Close #NAR
End Function
