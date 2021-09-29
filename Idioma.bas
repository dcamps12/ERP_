Attribute VB_Name = "Idioma"
Public vlAceptar As String
Public vlApertura As String
Public vlSeleccionar As String
Public vlCerrar As String
Public vlCerrado As String
Public vlCancelar As String
Public vlConcepto As String
Public vlCierre As String
Public vlExplotacion As String
Public vlRegistro As String
Public vlIra As String
Public vlFecha As String
Public vlPagina As String
Public vlCuenta As String
Public vlDesde As String
Public vlHasta As String
Public vlAltas As String
Public vlBajas As String
Public vlModificaciones As String
Public vlConsultas As String
Public vlModificar As String
Public vlNuevo As String
Public vlRespetar As String
Public vlAntes As String
Public vlDesesperado As String
Public vlSeleccionFechas As String
Public vlSeleccionCuentas As String
Public vlSeleccionCentroCoste As String
Public vlDe As String
Public vlImpresionCorrecta As String
Public vlPendiente As String
Public vlCompletado As String
Public vlSi As String
Public vlNo As String

Public vlSeleccion, vlSalidaPor, vlDiario, vlDiarioGeneral, vlAsiento As String
Public vlTotalizarPor, vlImprimir As String
Public vlSoloResumen, vlImpresion As String

Public vlPedirBorrar As String
Public vlBloqueado As String
Public vlSinInformacion As String
Public vlListadoCancelado As String
Public vlListadoDe As String

Public vlDatosSeleccion, vlSeleccionPor, vlValor, vlResultado As String
Public vlTotal  As String
Public vlSumaySigue As String
Public vlSaldoAnterior As String
Public vlTitulo As String
Public vlDebe As String, vlHaber As String
Public vlSaldo As String
Public vlSaldoInicial As String
Public vlSaldoFinal As String
Public vlAcumulado As String
Public vlPeriodo As String

Public vlEnero As String, vlFebrero As String, vlMarzo As String, vlAbril  As String
Public vlMayo  As String, vlJunio As String, vlJulio As String, vlAgosto As String
Public vlSeptiembre  As String, vlOctubre As String, vlNoviembre As String, vlDiciembre As String

Public vlAsiCaption As String
Public vlAsiCuenta As String
Public vlAsiNomCuenta As String
Public vlAsiCentro As String
Public vlAsiDocumento As String
Public vlAsiCodcon As String
Public vlAsiDesccon As String
Public vlAsiTipo As String
Public vlAsiImporte As String

Public Const ID1043_CODIGO = 4301
Public Const ID1043_CODIGODESC = 4302
Public Const ID1043_TITULO = 4303
Public Const ID1043_TITULODESC = 4304
Public Const ID1043_CAPTIONWINDOW = 4305

Public vlD As String, vlH As String

Public vlArticulo As String
Public Sub LeerLiterales(Optional Idioma As String)



   If Idioma = "" Then
      Idioma = cCfg.Idioma
   End If
   
   vlArticulo = GetTitulo("articulo", "public")
   vlDatosSeleccion = GetTitulo("datosseleccion")
   vlSinInformacion = GetDescripcion("sininformacion")
   vlListadoCancelado = GetDescripcion("listadocancel")
   vlListadoDe = GetDescripcion("listadode")
   
   'Para que la carga al iniciar el programa sea m�s r�pida
   If Idioma = "esp" Then
      vlPendiente = "Pendiente"
      vlCompletado = "Completado"
      vlSi = "S�"
      vlNo = "No"
      vlD = "D"
      vlH = "H"
      vlDe = "De"
      vlAceptar = "Aceptar"
      vlSeleccionar = "Seleccionar"
      vlSeleccionFechas = "Selecci�n fechas"
      vlSeleccionCuentas = "Selecci�n cuentas"
      vlSeleccionCentroCoste = "Selecci�n C.Coste"
      vlCerrar = GetTitulo("cerrar", "public")
      vlCerrado = "Cerrado"
      vlCancelar = "Cancelar"
      vlConcepto = "Concepto"
      vlRegistro = "Registro"
      vlIra = "Ir a"
      vlFecha = "Fecha"
      vlPagina = "P�gina"
      vlCuenta = "Cuenta"
      vlDebe = "Debe"
      vlHaber = "Haber"
      vlSaldo = "Saldo"
      vlSaldoInicial = "Saldo Inicial"
      vlSaldoFinal = "Saldo Final"
      vlTitulo = "T�tulo"
      vlAcumulado = "Acumulado"
      vlPeriodo = "Periodo"
      vlDesde = "Desde"
      vlHasta = "Hasta"
      vlAltas = "Altas"
      vlBajas = "Bajas"
      vlModificaciones = "Modificaciones"
      vlConsultas = "Consultas"
      vlModificar = "Modificar"
      vlNuevo = "Nuevo"
      vlRespetar = "Respetar fecha"
      vlAntes = "Antes si es posible"
      vlDesesperado = "Desesperado"
      vlSeleccion = "Selecci�n"
      vlSalidaPor = "Salida por"
      vlDiario = "Diario"
      vlDiarioGeneral = "Diario General"
      vlAsiento = "Asiento"
      vlTotalizarPor = "Totalizar por"
      vlTotal = "Total"
      vlSumaySigue = "Suma y sigue"
      vlSaldoAnterior = "Saldo anterior"
      vlImprimir = "Imprimir"
      vlSoloResumen = "S�lo resumen"
      vlImpresion = "Impresi�n"
      vlPedirBorrar = "Confirma la eliminaci�n?"
      vlBloqueado = "Registro en tratamiento por otro usuario"
      vlSeleccionPor = "Selecci�n por"
      vlValor = "Valor"
      vlResultado = "Resultado"
      vlApertura = "Apertura"
      vlExplotacion = "Explotaci�n"
      vlCierre = "Cierre"
      vlEnero = "Enero"
      vlFebrero = "Febrero"
      vlMarzo = "Marzo"
      vlAbril = "Abril"
      vlMayo = "Mayo"
      vlJunio = "Junio"
      vlJulio = "Julio"
      vlAgosto = "Agosto"
      vlSeptiembre = "Septiembre"
      vlOctubre = "Octubre"
      vlNoviembre = "Noviembre"
      vlDiciembre = "Diciembre"
   
      vlAsiCaption = "Datos Asientos"
      vlAsiCuenta = "Cuenta#Cuenta contable"
      vlAsiNomCuenta = "Nombre#Nombre de la cuenta contable"
      vlAsiCentro = "C.Coste#Centro de coste"
      vlAsiDocumento = "Docum.#Documento"
      vlAsiCodcon = "C�d. concepto#C�digo concepto"
      vlAsiDesccon = "Concepto#Concepto del apunte"
      vlAsiTipo = "Tipo#Tipo del apunte"
      vlAsiImporte = "Importe#Importe del apunte"
      
      vlImpresionCorrecta = "�Ha finalizado correctamente la impresion?"
   ElseIf Idioma = "por" Then
      vlPendiente = "Pendente"
      vlCompletado = "Completado"
      vlSi = "Si"
      vlNo = "No"
      vlD = "D"
      vlH = "H"
      vlDe = "De"
      vlAceptar = "Aceitar"
      vlSeleccionar = "Selecionar"
      vlSeleccionFechas = "Sele��o Datas"
      vlSeleccionCuentas = "Sele��o contas"
      vlSeleccionCentroCoste = "Sele��o C. Custos"
      vlCerrar = "Fechar"
      vlCerrado = "Fechado"
      vlCancelar = "Cancelar"
      vlConcepto = "Conceito"
      vlRegistro = "Registo"
      vlIra = "Ir para"
      vlFecha = "Data"
      vlPagina = "P�gina"
      vlCuenta = "Conta"
      vlDebe = "Deve"
      vlHaber = "Haver"
      vlSaldo = "Saldo"
      vlSaldoInicial = "Saldo Inicial"
      vlSaldoFinal = "Saldo Final"
      vlTitulo = "T�tulo"
      vlAcumulado = "Acumulado"
      vlPeriodo = "Per�odo"
      vlDesde = "Desde"
      vlHasta = "At�"
      vlAltas = "Altas"
      vlBajas = "Baixas"
      vlModificaciones = "Modifica��es"
      vlConsultas = "Consultas"
      vlModificar = "Modificar"
      vlNuevo = "Novo"
      vlRespetar = "Data de respeito"
      vlAntes = "Antes se possivel"
      vlDesesperado = "Desesperado"
      vlSeleccion = "Sele��o"
      vlSalidaPor = "Sa�da por"
      vlDiario = "Di�rio"
      vlDiarioGeneral = "Di�rio geral"
      vlAsiento = "Registro"
      vlTotalizarPor = "Totalizar por"
      vlTotal = "Total"
      vlSumaySigue = "Soma e segue"
      vlSaldoAnterior = "Saldo anterior"
      vlImprimir = "Imprimir"
      vlSoloResumen = "S� resumo"
      vlImpresion = "Impress�o"
      vlPedirBorrar = "Confirma a elimina��o?"
      vlBloqueado = "Registo a ser tratado por outro utilizador"
      vlSeleccionPor = "Sele��o por"
      vlValor = "Valor"
      vlResultado = "Resultado"
      vlApertura = "Abertura"
      vlExplotacion = "Explora��o"
      vlCierre = "Fecho"
      vlEnero = "Janeiro"
      vlFebrero = "Fevereiro"
      vlMarzo = "Mar�o"
      vlAbril = "Abril"
      vlMayo = "Maio"
      vlJunio = "Junho"
      vlJulio = "Julho"
      vlAgosto = "Agosto"
      vlSeptiembre = "Setembro"
      vlOctubre = "Outubro"
      vlNoviembre = "Novembro"
      vlDiciembre = "Dezembro"
   
      vlAsiCaption = "Dados Registros"
      vlAsiCuenta = "Conta#Conta contabil�stica"
      vlAsiNomCuenta = "Nome#Nome da conta contabil�stica"
      vlAsiCentro = "C.Custo#Centro de custo"
      vlAsiDocumento = "Docum.#Documenta��o"
      vlAsiCodcon = "Conceito#Conceito do c�digo"
      vlAsiDesccon = "Conceito#Conceito do nota"
      vlAsiTipo = "Tipo#Tipo de nota"
      vlAsiImporte = "Valor#Valor da nota"
      
      vlImpresionCorrecta = "Finalizou corretamente a impress�o?"
   Else 'cat
      vlPendiente = "Pendent"
      vlCompletado = "Completat"
      vlSi = "S�"
      vlNo = "No"
      vlD = "D"
      vlH = "H"
      vlDe = "De"
      vlAceptar = "Acceptar"
      vlSeleccionar = "Seleccionar"
      vlSeleccionFechas = "Selecci� dates"
      vlSeleccionCuentas = "Selecci� comptes"
      vlSeleccionCentroCoste = "Selecci� C.Costos"
      vlCerrar = "Tancar"
      vlCerrado = "Tancat"
      vlCancelar = "Cancel�lar"
      vlConcepto = "Concepte"
      vlRegistro = "Registre"
      vlIra = "Anar a"
      vlFecha = "Data"
      vlPagina = "P�gina"
      vlCuenta = "Compte"
      vlDebe = "Deure"
      vlHaber = "Haver"
      vlSaldo = "Saldo"
      vlSaldoInicial = "Saldo Inicial"
      vlSaldoFinal = "Saldo Final"
      vlTitulo = "T�tol"
      vlAcumulado = "Acumulat"
      vlPeriodo = "Per�ode"
      vlDesde = "Des de"
      vlHasta = "Fins"
      vlAltas = "Altes"
      vlBajas = "Baixes"
      vlModificaciones = "Modificacions"
      vlConsultas = "Consultes"
      vlModificar = "Modificar"
      vlNuevo = "Nou"
      vlRespetar = "Respectar data"
      vlAntes = "Abans si es possible"
      vlDesesperado = "Desesperat"
      vlSeleccion = "Selecci�"
      vlSalidaPor = "Sortida per"
      vlDiario = "Diari"
      vlDiarioGeneral = "Diari General"
      vlAsiento = "Assentament"
      vlTotalizarPor = "Totalitzar per"
      vlTotal = "Total"
      vlSumaySigue = "Suma i continua"
      vlSaldoAnterior = "Saldo anterior"
      vlImprimir = "Imprimir"
      vlSoloResumen = "Sols resum"
      vlImpresion = "Impressi�"
   
      vlPedirBorrar = "Confirma l'eliminaci�?"
      vlBloqueado = "Registre en tractament per un altre usuari"
      
      vlSeleccionPor = "Selecci� per"
      vlValor = "Valor"
      vlResultado = "Resultat"
      
      vlApertura = "Apertura"
      vlExplotacion = "Explotaci�"
      vlCierre = "Tancament"
      vlEnero = "Gener"
      vlFebrero = "Febrer"
      vlMarzo = "Mar�"
      vlAbril = "Abril"
      vlMayo = "Maig"
      vlJunio = "Juny"
      vlJulio = "Juliol"
      vlAgosto = "Agost"
      vlSeptiembre = "Setembre"
      vlOctubre = "Octubre"
      vlNoviembre = "Novembre"
      vlDiciembre = "Desembre"
   
      vlAsiCaption = "Dades Assentaments"
      vlAsiCuenta = "Compte#Compte comptable"
      vlAsiNomCuenta = "Nombre#Nombre del compte comptable"
      vlAsiCentro = "C.Costos#Centre de costos"
      vlAsiDocumento = "Docum.#Document"
      vlAsiCodcon = "Codi concepte#Codi concepte"
      vlAsiDesccon = "Concepte#Concepte de l'apunt"
      vlAsiTipo = "Tipus#Tipus d'apunt"
      vlAsiImporte = "Import#Import de l'apunt"
      
      vlImpresionCorrecta = "Ha finalitzat correctament la impressi�?"
   End If
End Sub

Public Function GetTitulo(Campo As String, Optional Tabla As String) As String
   If IsMissing(Tabla) Then
      Tabla = "public"
   End If
   If Tabla = "" Then
      Tabla = "public"
   End If
   
   On Error GoTo ergo
   GetTitulo = SysCols(Tabla & "|" & Campo).Titulo
   Exit Function
   
ergo:
   GetTitulo = ""
End Function

Public Function GetDescripcion(Campo As String, Optional Tabla As String) As String
   If IsMissing(Tabla) Then
      Tabla = "public"
   End If
   If Tabla = "" Then
      Tabla = "public"
   End If
   
   On Error GoTo ergo
   GetDescripcion = SysCols(Tabla & "|" & Campo).Descripcion
   Exit Function
ergo:
   GetDescripcion = ""
End Function

