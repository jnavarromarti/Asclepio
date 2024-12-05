Option Explicit
Public MesActual As Date
Public DiaSeleccionado As Date


Sub Save()
    Dim Fila As Long
    Dim Especialidad As String
    Dim HojaDestino As Worksheet
    Dim NumeroPaciente As String
    Dim CamposFaltantes As String
    Dim DNI As String
    Dim Celda As Range
    Dim ActualizarDatos As Boolean
    Dim UltimoNumeroPaciente As Long
    Dim respuesta As VbMsgBoxResult
    
    CamposFaltantes = ""
    Especialidad = Hoja1.Range("S23").Value
    NumeroPaciente = Hoja1.Range("S3").Value
    DNI = Hoja1.Range("G17").Value

    If Hoja1.Range("H2").Value = "" Then CamposFaltantes = CamposFaltantes & "Nombre (H2)" & vbCrLf
    If Hoja1.Range("H5").Value = "" Then CamposFaltantes = CamposFaltantes & "Apellidos (H5)" & vbCrLf
    If NumeroPaciente = "" Then CamposFaltantes = CamposFaltantes & "Nº paciente (S3)" & vbCrLf
    If DNI = "" Then CamposFaltantes = CamposFaltantes & "DNI (G17)" & vbCrLf
    If Not ValidateDNI(DNI) Then
        MsgBox "El DNI proporcionado no es válido. Introduce un DNI válido."
        Exit Sub
    End If
    
    If CamposFaltantes <> "" Then
        MsgBox "Faltan los siguientes campos:" & vbCrLf & CamposFaltantes
        Exit Sub
    End If
    

    Select Case Especialidad
        Case "Radiología"
            Set HojaDestino = Hoja2
        Case "Logopedia"
            Set HojaDestino = Hoja3
        Case "Psicopedagogía"
            Set HojaDestino = Hoja4
        Case Else
            MsgBox "No hay especialistas de esto aún."
            Exit Sub
    End Select


    Call AutomaticallyAssign


    Set Celda = HojaDestino.Range("A:A").Find(What:=NumeroPaciente, LookIn:=xlValues, LookAt:=xlWhole)

    If Celda Is Nothing Then

        Fila = HojaDestino.Cells(Rows.Count, 1).End(xlUp).Row + 1
    Else
            
        ActualizarDatos = False
        If Hoja1.Range("H2").Value <> Celda.Offset(0, 1).Value Then ActualizarDatos = True
        If Hoja1.Range("H5").Value <> Celda.Offset(0, 2).Value Then ActualizarDatos = True
        If Hoja1.Range("G17").Value <> Celda.Offset(0, 3).Value Then ActualizarDatos = True
        If Hoja1.Range("G11").Value <> Celda.Offset(0, 4).Value Then ActualizarDatos = True
        If Hoja1.Range("G14").Value <> Celda.Offset(0, 5).Value Then ActualizarDatos = True
        If Hoja1.Range("G20").Value <> Celda.Offset(0, 6).Value Then ActualizarDatos = True
        If Hoja1.Range("G23").Value <> Celda.Offset(0, 7).Value Then ActualizarDatos = True
        If Hoja1.Range("G26").Value <> Celda.Offset(0, 8).Value Then ActualizarDatos = True
        If Hoja1.Range("G29").Value <> Celda.Offset(0, 9).Value Then ActualizarDatos = True
        If Hoja1.Range("G32").Value <> Celda.Offset(0, 10).Value Then ActualizarDatos = True
        If Hoja1.Range("G35").Value <> Celda.Offset(0, 11).Value Then ActualizarDatos = True
        If Hoja1.Range("N14").Value <> Celda.Offset(0, 12).Value Then ActualizarDatos = True
        If Especialidad <> Celda.Offset(0, 13).Value Then ActualizarDatos = True
        If Hoja1.Range("S29").Value <> Celda.Offset(0, 14).Value Then ActualizarDatos = True
        If Hoja1.Range("S26").Value <> Celda.Offset(0, 15).Value Then ActualizarDatos = True

        If ActualizarDatos Then
            Fila = Celda.Row
            MsgBox "Se ha actualizado el paciente."
        Else
            MsgBox "El paciente no ha sido modificado."
            Exit Sub
        End If
    End If

    HojaDestino.Cells(Fila, 1).Value = Hoja1.Range("S3").Value
    HojaDestino.Cells(Fila, 2).Value = Hoja1.Range("H2").Value
    HojaDestino.Cells(Fila, 3).Value = Hoja1.Range("H5").Value
    HojaDestino.Cells(Fila, 4).Value = Hoja1.Range("G17").Value
    HojaDestino.Cells(Fila, 5).Value = Hoja1.Range("G11").Value
    HojaDestino.Cells(Fila, 6).Value = Hoja1.Range("G14").Value
    HojaDestino.Cells(Fila, 7).Value = Hoja1.Range("G20").Value
    HojaDestino.Cells(Fila, 8).Value = Hoja1.Range("G23").Value
    HojaDestino.Cells(Fila, 9).Value = Hoja1.Range("G26").Value
    HojaDestino.Cells(Fila, 10).Value = Hoja1.Range("G29").Value
    HojaDestino.Cells(Fila, 11).Value = Hoja1.Range("G32").Value
    HojaDestino.Cells(Fila, 12).Value = Hoja1.Range("G35").Value
    HojaDestino.Cells(Fila, 13).Value = Hoja1.Range("N14").Value
    HojaDestino.Cells(Fila, 14).Value = Hoja1.Range("S23").Value
    HojaDestino.Cells(Fila, 15).Value = Hoja1.Range("S29").Value
    HojaDestino.Cells(Fila, 16).Value = Hoja1.Range("S26").Value

    Hoja1.Range("H2:H5").Value = ""
    Hoja1.Range("S3").Value = ""
    Hoja1.Range("G11:G35").Value = ""
    Hoja1.Range("N14").Value = ""
    Hoja1.Range("S23:S29").Value = ""
End Sub

Sub Search()
    Dim Celda As Range
    Dim NumeroPaciente As String
    Dim Nombre As String
    Dim Apellidos As String
    Dim Encontrado As Boolean
    Dim HojaActual As Worksheet
    Dim HojasBusqueda As Variant
    Dim i As Integer

    NumeroPaciente = Hoja1.Range("S3").Value
    Nombre = Hoja1.Range("H2").Value
    Apellidos = Hoja1.Range("H5").Value
  
    If NumeroPaciente = "" And (Nombre = "" Or Apellidos = "") Then
        MsgBox "Por favor, ingrese el número de paciente o el nombre y apellidos para realizar la búsqueda."
        Exit Sub
    End If
    
    Encontrado = False
    
    HojasBusqueda = Array(Hoja2, Hoja3, Hoja4)
    
    For i = LBound(HojasBusqueda) To UBound(HojasBusqueda)
        Set HojaActual = HojasBusqueda(i)
        
        If NumeroPaciente <> "" Then
            Set Celda = HojaActual.Range("A:A").Find(What:=NumeroPaciente, LookIn:=xlValues, LookAt:=xlWhole)
            If Not Celda Is Nothing Then
                Encontrado = True

                Hoja1.Range("S3").Value = Celda.Value
                Hoja1.Range("H2").Value = Celda.Offset(0, 1).Value
                Hoja1.Range("H5").Value = Celda.Offset(0, 2).Value
                Hoja1.Range("G17").Value = Celda.Offset(0, 3).Value
                Hoja1.Range("G11").Value = Celda.Offset(0, 4).Value
                Hoja1.Range("G14").Value = Celda.Offset(0, 5).Value
                Hoja1.Range("G20").Value = Celda.Offset(0, 6).Value
                Hoja1.Range("G23").Value = Celda.Offset(0, 7).Value
                Hoja1.Range("G26").Value = Celda.Offset(0, 8).Value
                Hoja1.Range("G29").Value = Celda.Offset(0, 9).Value
                Hoja1.Range("G32").Value = Celda.Offset(0, 10).Value
                Hoja1.Range("G35").Value = Celda.Offset(0, 11).Value
                Hoja1.Range("N14").Value = Celda.Offset(0, 12).Value
                Hoja1.Range("S23").Value = Celda.Offset(0, 13).Value
                Hoja1.Range("S29").Value = Celda.Offset(0, 14).Value
                Hoja1.Range("S26").Value = Celda.Offset(0, 15).Value
                Exit For
            End If
        End If
        
        
        If Not Encontrado And Nombre <> "" And Apellidos <> "" Then
            Set Celda = HojaActual.Range("B:B").Find(What:=Nombre, LookIn:=xlValues, LookAt:=xlWhole)
            If Not Celda Is Nothing And Celda.Offset(0, 1).Value = Apellidos Then
                Encontrado = True
                
                Hoja1.Range("S3").Value = Celda.Offset(0, -1).Value
                Hoja1.Range("H2").Value = Celda.Value
                Hoja1.Range("H5").Value = Celda.Offset(0, 1).Value
                Hoja1.Range("G17").Value = Celda.Offset(0, 2).Value
                Hoja1.Range("G11").Value = Celda.Offset(0, 3).Value
                Hoja1.Range("G14").Value = Celda.Offset(0, 4).Value
                Hoja1.Range("G20").Value = Celda.Offset(0, 5).Value
                Hoja1.Range("G23").Value = Celda.Offset(0, 6).Value
                Hoja1.Range("G26").Value = Celda.Offset(0, 7).Value
                Hoja1.Range("G29").Value = Celda.Offset(0, 8).Value
                Hoja1.Range("G32").Value = Celda.Offset(0, 9).Value
                Hoja1.Range("G35").Value = Celda.Offset(0, 10).Value
                Hoja1.Range("N14").Value = Celda.Offset(0, 11).Value
                Hoja1.Range("S23").Value = Celda.Offset(0, 12).Value
                Hoja1.Range("S29").Value = Celda.Offset(0, 13).Value
                Hoja1.Range("S26").Value = Celda.Offset(0, 14).Value
                Exit For
            End If
        End If
    Next i
    
    If Not Encontrado Then
        MsgBox "El paciente no se encuentra en el registro."
    End If
End Sub

Sub Delete()
    Dim Celda As Range
    Dim NumeroPaciente As String
    Dim HojaActual As Worksheet
    Dim HojasBusqueda As Variant
    Dim i As Integer
    Dim Encontrado As Boolean
     
    NumeroPaciente = Hoja1.Range("S3").Value
    
    If NumeroPaciente = "" Then
        MsgBox "Escribe el número del paciente que deseas eliminar."
        Exit Sub
    End If
    
    Encontrado = False
    
    HojasBusqueda = Array(Hoja2, Hoja3, Hoja4)
    
    For i = LBound(HojasBusqueda) To UBound(HojasBusqueda)
        Set HojaActual = HojasBusqueda(i)
        
  
        Set Celda = HojaActual.Range("A:A").Find(What:=NumeroPaciente, LookIn:=xlValues, LookAt:=xlWhole)
        If Not Celda Is Nothing Then
            Encontrado = True
            Celda.EntireRow.Delete
            MsgBox "El paciente ha sido eliminado correctamente."
            Exit For
        End If
    Next i
    

    If Not Encontrado Then
        MsgBox "El paciente no existe."
    End If
    
  
    Hoja1.Range("H2:H5").Value = ""
    Hoja1.Range("S3").Value = ""
    Hoja1.Range("G11:G35").Value = ""
    Hoja1.Range("N14").Value = ""
    Hoja1.Range("S23:S29").Value = ""
End Sub

Sub Clear()
    
    Hoja1.Range("H2:H5").Value = ""
    Hoja1.Range("S3").Value = ""
    Hoja1.Range("G11:G35").Value = ""
    Hoja1.Range("N14").Value = ""
    Hoja1.Range("S23:S29").Value = ""
End Sub

Sub Bill()
    Dim Especialidad As String
    Dim HojaDestino As Worksheet

    Especialidad = Hoja1.Range("S23").Value

    If Especialidad = "Radiología" Then
        Set HojaDestino = Hoja5

        Call ReferralClinic
    Else
        Set HojaDestino = Hoja6
        Hoja6.Range("A7").Value = Especialidad
    End If

    HojaDestino.Range("K3").Value = Hoja1.Range("S3").Value
    HojaDestino.Range("K4").Value = Hoja1.Range("H2").Value
    HojaDestino.Range("K5").Value = Hoja1.Range("H5").Value
    HojaDestino.Range("K2").Value = Hoja1.Range("G17").Value

    If HojaDestino.Range("C13").MergeCells Then
        HojaDestino.Range("C13").MergeArea.Cells(1, 1).Value = Date
    Else
        HojaDestino.Range("C13").Value = Date
    End If

    HojaDestino.Range("L13").Value = ObtenerNumeroFactura()

    If Especialidad = "Radiología" Then
        Sheets("Hoja4").Activate
    Else
        Sheets("Hoja5").Activate
    End If
End Sub

Sub GiveAppointment()
    Dim FechaCita As String
    Dim HoraCita As String
    Dim FechaHoraCita As Date
    Dim Fila As Long
    Dim HojaDestino As Worksheet
    Dim NumeroPaciente As String
    Dim CamposFaltantes As String
    
    Set HojaDestino = Hoja8
    
    NumeroPaciente = Hoja1.Range("S3").Value
    If NumeroPaciente = "" Then
        MsgBox "Por favor, ingrese el número de paciente antes de asignar una cita."
        Exit Sub
    End If
    
    CamposFaltantes = ""
    If Hoja1.Range("H2").Value = "" Then CamposFaltantes = CamposFaltantes & "Nombre (H2)" & vbCrLf
    If Hoja1.Range("H5").Value = "" Then CamposFaltantes = CamposFaltantes & "Apellidos (H5)" & vbCrLf
    If Hoja1.Range("G17").Value = "" Then CamposFaltantes = CamposFaltantes & "DNI (G17)" & vbCrLf
    If Hoja1.Range("G20").Value = "" Then CamposFaltantes = CamposFaltantes & "Teléfono (G20)" & vbCrLf

    If CamposFaltantes <> "" Then
        MsgBox "Faltan los siguientes campos antes de asignar una cita:" & vbCrLf & CamposFaltantes
        Exit Sub
    End If

    FechaCita = InputBox("¿Qué día quiere la cita? (dd/mm/aaaa)", "Asignar Cita")
    If IsDate(FechaCita) = False Then
        MsgBox "Fecha no válida. Intente de nuevo."
        Exit Sub
    End If

    HoraCita = InputBox("¿A qué hora quiere la cita? (hh:mm)", "Asignar Cita")
    If Len(HoraCita) <> 5 Or Mid(HoraCita, 3, 1) <> ":" Or Not IsNumeric(Left(HoraCita, 2)) Or Not IsNumeric(Right(HoraCita, 2)) Then
        MsgBox "Hora no válida. Intente de nuevo en formato hh:mm."
        Exit Sub
    End If

    On Error Resume Next
    FechaHoraCita = CDate(FechaCita & " " & HoraCita)
    On Error GoTo 0
    If Err.Number <> 0 Then
        MsgBox "Fecha y hora no válidas. Por favor, intente de nuevo."
        Exit Sub
    End If

    Fila = HojaDestino.Cells(Rows.Count, 1).End(xlUp).Row + 1

    HojaDestino.Cells(Fila, 1).Value = Hoja1.Range("S3").Value
    HojaDestino.Cells(Fila, 2).Value = Hoja1.Range("H2").Value
    HojaDestino.Cells(Fila, 3).Value = Hoja1.Range("H5").Value
    HojaDestino.Cells(Fila, 4).Value = Hoja1.Range("S23").Value
    HojaDestino.Cells(Fila, 5).Value = HoraCita
    HojaDestino.Cells(Fila, 6).Value = FechaCita
    HojaDestino.Cells(Fila, 7).Value = Hoja1.Range("G20").Value
    HojaDestino.Cells(Fila, 8).Value = Hoja1.Range("G17").Value

    MsgBox "Cita asignada para " & Hoja1.Range("H2").Value & " el " & FechaCita & " a las " & HoraCita
End Sub


Sub PrintIt()
    Dim Especialidad As String
    Dim HojaDestino As Worksheet
    Dim NumCopias As Integer
    Dim respuesta As VbMsgBoxResult
    Dim FechaHoy As String
    Dim FacturaGuardada As Boolean

    Especialidad = Hoja1.Range("S23").Value

    If Especialidad = "Radiología" Then
        Set HojaDestino = Hoja5
    Else
        Set HojaDestino = Hoja6
    End If

    NumCopias = InputBox("¿Cuántas copias desea imprimir?", "Impresión de Factura", 1)
    If Not IsNumeric(NumCopias) Or NumCopias < 0 Then
        MsgBox "Número de copias inválido. Se cancelará la impresión.", vbExclamation
        Exit Sub
    End If

    FechaHoy = Format(Date, "dd/mm/yyyy")
    If HojaDestino.Range("C13").Value <> FechaHoy Then
        HojaDestino.Range("C13").Value = FechaHoy
    End If

    FacturaGuardada = SaveBill(HojaDestino)

    If FacturaGuardada Then
        HojaDestino.PrintOut Copies:=NumCopias

        With HojaDestino
            .Range("K2:K5").ClearContents
            .Range("A19:A27").ClearContents
            .Range("E19:E27").ClearContents
            .Range("I19:I27").ClearContents
            .Range("K19:K27").ClearContents
        End With

        MsgBox "Factura impresa y datos del paciente eliminados.", vbInformation
    Else
        MsgBox "Error al guardar la factura. La impresión fue cancelada.", vbExclamation
    End If
End Sub

Function ValidateDNI(DNI As String) As Boolean
    Dim Numeros As String
    Dim Letra As String
    Dim Letras As String
    Dim PosicionLetra As Integer
    Dim PrimerCaracter As String
    
    Letras = "TRWAGMYFPDXBNJZSQVHLCKET"
    
    If Len(DNI) <> 9 Then
        ValidateDNI = False
        Exit Function
    End If

    PrimerCaracter = Left(DNI, 1)
    Numeros = Mid(DNI, 2, 7)
    Letra = Right(DNI, 1)

    If PrimerCaracter = "X" Then
        Numeros = "0" & Numeros
    ElseIf PrimerCaracter = "Y" Then
        Numeros = "1" & Numeros
    ElseIf PrimerCaracter = "Z" Then
        Numeros = "2" & Numeros
    ElseIf IsNumeric(PrimerCaracter) Then

        Numeros = PrimerCaracter & Numeros
    Else

        ValidateDNI = False
        Exit Function
    End If

    If Not IsNumeric(Numeros) Then
        ValidateDNI = False
        Exit Function
    End If

    PosicionLetra = CLng(Numeros) Mod 23

    If Mid(Letras, PosicionLetra + 1, 1) = UCase(Letra) Then
        ValidateDNI = True
    Else
        ValidateDNI = False
    End If
End Function

Sub AutomaticallyAssign()
    Dim NumeroPaciente As String
    Dim respuesta As VbMsgBoxResult
    Dim UltimoNumeroPaciente As Long
    
    NumeroPaciente = Hoja1.Range("S3").Value
    
    If NumeroPaciente = "" Then
        respuesta = MsgBox("¿Desea asignar un número de paciente automático?", vbYesNo)
        If respuesta = vbYes Then
            UltimoNumeroPaciente = GetLastPatientNumber()
            NumeroPaciente = UltimoNumeroPaciente + 1
            Hoja1.Range("S3").Value = NumeroPaciente
            MsgBox "Se ha asignado el número de paciente: " & NumeroPaciente
        Else
            Exit Sub
        End If
    End If
End Sub

Function GetLastPatientNumber() As Long
    Dim UltimoNumeroHoja2 As Long
    Dim UltimoNumeroHoja3 As Long
    Dim UltimoNumeroHoja4 As Long
    
    UltimoNumeroHoja2 = Hoja2.Cells(Hoja2.Rows.Count, 1).End(xlUp).Value
    UltimoNumeroHoja3 = Hoja3.Cells(Hoja3.Rows.Count, 1).End(xlUp).Value
    UltimoNumeroHoja4 = Hoja4.Cells(Hoja4.Rows.Count, 1).End(xlUp).Value
    
    GetLastPatientNumber = Application.WorksheetFunction.Max(UltimoNumeroHoja2, UltimoNumeroHoja3, UltimoNumeroHoja4)
End Function

Sub ReferralClinic()
    Dim Derivacion As String

    Derivacion = InputBox("¿Desde qué clínica se le deriva?")

    Hoja5.Range("A31").Value = Derivacion
End Sub
Sub SearchBill()
    Dim FechaFactura As String
    Dim Celda As Range
    Dim HojaDestino As Worksheet
    Dim HojaRegistro As Worksheet
    Dim Especialidad As String
    Dim Encontrado As Boolean
    Dim i As Long

    Set HojaRegistro = Hoja9

    FechaFactura = InputBox("¿Qué día desea recuperar la factura? (dd/mm/aaaa)", "Buscar Factura")
    
    If IsDate(FechaFactura) = False Then
        MsgBox "Fecha no válida. Intente de nuevo."
        Exit Sub
    End If

    Encontrado = False

    For i = 2 To HojaRegistro.Cells(HojaRegistro.Rows.Count, 1).End(xlUp).Row
        If HojaRegistro.Cells(i, 2).Value = CDate(FechaFactura) Then

            Especialidad = HojaRegistro.Cells(i, 7).Value

            If Especialidad = "Radiología" Then
                Set HojaDestino = Hoja5
            Else
                Set HojaDestino = Hoja6
            End If

            HojaDestino.Range("L13").Value = HojaRegistro.Cells(i, 1).Value
            HojaDestino.Range("C13").Value = HojaRegistro.Cells(i, 2).Value
            HojaDestino.Range("K2").Value = HojaRegistro.Cells(i, 3).Value
            HojaDestino.Range("K3").Value = HojaRegistro.Cells(i, 4).Value
            HojaDestino.Range("K4").Value = HojaRegistro.Cells(i, 5).Value
            HojaDestino.Range("K5").Value = HojaRegistro.Cells(i, 6).Value
            HojaDestino.Range("A7").Value = HojaRegistro.Cells(i, 7).Value

            HojaDestino.Range("A19").Value = HojaRegistro.Cells(i, 8).Value
            HojaDestino.Range("A21").Value = HojaRegistro.Cells(i, 9).Value
            HojaDestino.Range("A23").Value = HojaRegistro.Cells(i, 10).Value
            HojaDestino.Range("A25").Value = HojaRegistro.Cells(i, 11).Value
            HojaDestino.Range("A27").Value = HojaRegistro.Cells(i, 12).Value

            HojaDestino.Range("E19").Value = HojaRegistro.Cells(i, 13).Value
            HojaDestino.Range("E21").Value = HojaRegistro.Cells(i, 14).Value
            HojaDestino.Range("E23").Value = HojaRegistro.Cells(i, 15).Value
            HojaDestino.Range("E25").Value = HojaRegistro.Cells(i, 16).Value
            HojaDestino.Range("E27").Value = HojaRegistro.Cells(i, 17).Value

            HojaDestino.Range("I32").Value = HojaRegistro.Cells(i, 18).Value

            If Especialidad = "Radiología" Or HojaDestino.Name = "Hoja5" Then
                HojaDestino.Range("A31").Value = HojaRegistro.Cells(i, 19).Value
            End If

            Encontrado = True
            Exit For
        End If
    Next i

    If Not Encontrado Then
        MsgBox "No se encontró ninguna factura con la fecha especificada.", vbExclamation
    Else
        MsgBox "Factura cargada correctamente.", vbInformation
    End If
End Sub

Function ObtenerNumeroFactura() As String
    Dim UltimoNumeroFactura As Long
    Dim NuevaFactura As String
    Dim AnioActual As String
    Dim HojaRegistro As Worksheet
    Set HojaRegistro = Hoja9

    UltimoNumeroFactura = HojaRegistro.Cells(HojaRegistro.Rows.Count, 1).End(xlUp).Value

    UltimoNumeroFactura = UltimoNumeroFactura + 1

    AnioActual = Right(Year(Date), 2)
    NuevaFactura = AnioActual & " " & Format(UltimoNumeroFactura, "00000")
    
    ObtenerNumeroFactura = NuevaFactura
End Function

Function SaveBill(HojaFactura As Worksheet) As Boolean
    On Error GoTo ErrorHandler

    Dim Fila As Long
    Dim HojaRegistro As Worksheet
    Dim UltimoNumeroFactura As Long
    Dim Especialidad As String
    Dim ReferenciaClinica As String
    
    Set HojaRegistro = Hoja9
    Especialidad = HojaFactura.Range("A7").Value

    Fila = HojaRegistro.Cells(HojaRegistro.Rows.Count, 1).End(xlUp).Row + 1

    HojaRegistro.Cells(Fila, 1).Value = HojaFactura.Range("L13").Value
    HojaRegistro.Cells(Fila, 2).Value = HojaFactura.Range("C13").Value
    HojaRegistro.Cells(Fila, 3).Value = HojaFactura.Range("K2").Value
    HojaRegistro.Cells(Fila, 4).Value = HojaFactura.Range("K3").Value
    HojaRegistro.Cells(Fila, 5).Value = HojaFactura.Range("K4").Value
    HojaRegistro.Cells(Fila, 6).Value = HojaFactura.Range("K5").Value
    HojaRegistro.Cells(Fila, 7).Value = HojaFactura.Range("A7").Value

    HojaRegistro.Cells(Fila, 8).Value = HojaFactura.Range("A19").Value
    HojaRegistro.Cells(Fila, 9).Value = HojaFactura.Range("A21").Value
    HojaRegistro.Cells(Fila, 10).Value = HojaFactura.Range("A23").Value
    HojaRegistro.Cells(Fila, 11).Value = HojaFactura.Range("A25").Value
    HojaRegistro.Cells(Fila, 12).Value = HojaFactura.Range("A27").Value

    HojaRegistro.Cells(Fila, 13).Value = HojaFactura.Range("E19").Value
    HojaRegistro.Cells(Fila, 14).Value = HojaFactura.Range("E21").Value
    HojaRegistro.Cells(Fila, 15).Value = HojaFactura.Range("E23").Value
    HojaRegistro.Cells(Fila, 16).Value = HojaFactura.Range("E25").Value
    HojaRegistro.Cells(Fila, 17).Value = HojaFactura.Range("E27").Value

    HojaRegistro.Cells(Fila, 18).Value = HojaFactura.Range("I32").Value

    If Especialidad = "Radiología" Then
        ReferenciaClinica = HojaFactura.Range("A31").Value
        HojaRegistro.Cells(Fila, 19).Value = ReferenciaClinica
    End If

    SaveBill = True
     MsgBox "Factura guardada."
    Exit Function

ErrorHandler:
    SaveBill = False
     MsgBox "Error al guardar factura."
End Function

Public Sub FinderDayAppointments(FechaSeleccionada As Date)
    Dim FechaString As String
    FechaString = Format(FechaSeleccionada, "dd/mm/yyyy")

    AppointmentFinder FechaString
End Sub

Function AppointmentFinder(FechaSeleccionada As String)
    Dim HojaCitas As Worksheet
    Dim HojaDestino As Worksheet
    Dim Celda As Range
    Dim NombrePaciente As String
    Dim ApellidosPaciente As String
    Dim DNIPaciente As String
    Dim Especialidad As String
    Dim NumeroPaciente As String
    Dim HoraCita As String
    Dim CeldaDestino As Long
    Dim Encontrado As Boolean
    Dim TextoCita As String
    Dim PosicionActual As Long
    Dim TextoFinal As String

    Set HojaDestino = Hoja7
    Set HojaCitas = Hoja8

    Encontrado = False

    HojaDestino.Range("Q7:W16, Q19:W29").ClearContents

    Set Celda = HojaCitas.Range("F:F").Find(What:=FechaSeleccionada, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not Celda Is Nothing Then
        Do
            NumeroPaciente = Celda.Offset(0, -5).Value
            NombrePaciente = Celda.Offset(0, -4).Value
            ApellidosPaciente = Celda.Offset(0, -3).Value
            DNIPaciente = Celda.Offset(0, 1).Value
            Especialidad = Celda.Offset(0, -2).Value
            HoraCita = Format(Celda.Offset(0, -1).Value, "hh:mm")

            TextoCita = "•" & NumeroPaciente & " " & NombrePaciente & " " & ApellidosPaciente & " " & DNIPaciente & " " & Especialidad

            Select Case HoraCita
                Case "10:00": CeldaDestino = 7
                Case "10:30": CeldaDestino = 8
                Case "11:00": CeldaDestino = 9
                Case "11:30": CeldaDestino = 10
                Case "12:00": CeldaDestino = 11
                Case "12:30": CeldaDestino = 12
                Case "13:00": CeldaDestino = 13
                Case "13:30": CeldaDestino = 14
                Case "14:00": CeldaDestino = 15
                Case "14:30": CeldaDestino = 16
                Case "15:30": CeldaDestino = 19
                Case "16:00": CeldaDestino = 20
                Case "16:30": CeldaDestino = 21
                Case "17:00": CeldaDestino = 22
                Case "17:30": CeldaDestino = 23
                Case "18:00": CeldaDestino = 24
                Case "18:30": CeldaDestino = 25
                Case "19:00": CeldaDestino = 26
                Case "19:30": CeldaDestino = 27
                Case "20:00": CeldaDestino = 28
                Case "20:30": CeldaDestino = 29
                Case Else: CeldaDestino = 0
            End Select

            TextoFinal = HojaDestino.Range("Q" & CeldaDestino & ":W" & CeldaDestino).Value
            If TextoFinal <> "" Then
                TextoFinal = TextoFinal & vbCrLf & TextoCita
            Else
                TextoFinal = TextoCita
            End If

            If CeldaDestino > 0 Then
                HojaDestino.Range("Q" & CeldaDestino & ":W" & CeldaDestino).Value = TextoFinal

                With HojaDestino.Range("Q" & CeldaDestino & ":W" & CeldaDestino)
                    PosicionActual = 1
                    .Characters(PosicionActual, Len(NumeroPaciente)).Font.Bold = True
                    PosicionActual = PosicionActual + Len(NumeroPaciente) + 1
                    PosicionActual = PosicionActual + Len(NombrePaciente & " " & ApellidosPaciente) + 1
                    .Characters(PosicionActual, Len(DNIPaciente)).Font.Bold = True
                    PosicionActual = PosicionActual + Len(DNIPaciente) + 1
                    .Characters(PosicionActual, Len(Especialidad)).Font.Bold = True
                    .Characters(PosicionActual, Len(Especialidad)).Font.Underline = xlUnderlineStyleSingle
                End With
                
                Encontrado = True
            End If

            Set Celda = HojaCitas.Range("F:F").FindNext(Celda)
            
        Loop While Not Celda Is Nothing And Celda.Address <> HojaCitas.Range("F:F").Find(What:=FechaSeleccionada, LookIn:=xlValues, LookAt:=xlWhole).Address
    End If
End Function
