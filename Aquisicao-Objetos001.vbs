On Error Resume Next

Class retornoSolve

	Dim solve
	Dim ra
	Dim dec
	Dim x
	Dim y
	
End Class

Class registroPlano 

	Dim objeto
	Dim ra
	Dim dec
	Dim exposicao
	Dim filtro
	Dim binning
	Dim frames
	Dim stack
	Dim pixelsX
	Dim pixelsY
	Dim maximo
	Dim minimo
	Dim periodo
	Dim tipo
	Dim raDecimal
	Dim decDecimal
	Dim sequencia
	Dim vetorRa
	Dim vetorDec
	
End Class

Class alvo

	Dim alvo
	Dim x
	Dim y
	
End Class

Dim localDec 'Declinacao local do Zenite em graus
Dim localRa 'RA local em horas decimais
Dim contador
Dim quantidadeCap
Dim ano
Dim mes
Dim dia
Dim hora
Dim minuto
Dim segundo
Dim agora
Dim pasta
Dim complemento
Dim margemHorizonte
Dim tentativaAlvo
Dim expApontamento
Dim toleranciaSolve
Dim setPointCamera
Dim exposicaoGuider
Dim flagGuider
Dim cameraApontamento
Dim complementoSolve
Dim binningCaptura
Dim binningProcura
Dim flagDebug
Dim flagDogsHeaven
Dim flagCalibra
Dim flagCalibraAlvo
Dim flagConfirmacao
Dim flagCool
Dim flagWarmup
Dim flagSolve
Dim delimitador
Dim escolheTelescopio
Dim capturado

Dim arquivoTexto
Dim arquivoLog
Dim arquivoRetorno
Dim arquivoTemp
Dim arquivoConfig
Dim arquivoCapturados
Dim linhaDados
Dim regDados
Dim subframeX
Dim subframeY

Dim diretorioLocal
Dim diretorioFlat
Dim diretorioLight
Dim diretorioImagem
Dim diretorioTemp
Dim diretorioDark
Dim diretorioBias
Dim diretorioCalibracao
Dim diretorioLog
Dim diretorioRet
Dim diretorioGuider
Dim diretorioDogsHeaven
Dim diretorioReport
Dim diretorioProntas
Dim diretorioStack

Dim telescopio
Dim fsoPasta
Dim arquivo
Dim reg
Dim Log
Dim ret
Dim conf
Dim statusReport
Dim arrayPlano ()
Dim arrayCapturados ()
Dim indicePlano
Dim IndiceCapturados
Dim erroSolve

Dim nomesFiltros
Dim flagVariavel
Dim indiceConf
Dim arrayConf ()
Dim confDados
Dim prefixoArquivo
Dim otimizaGoto
Dim contadorCap
Dim flagSequencia
Dim dither
Dim flagGrandeCampo
Dim isoGrandeCampo
Dim binningGrandeCampo
Dim flagCalibraGrandeCampo
Dim exposicaoGrandeCampo
Dim prefixoGrandeCampo
Dim timeoutGrandeCampo
Dim trocaMeridiano


exibeTela "Copyright (C) 2016-2O21 - Observatório Adhara"
exibeTela "AQObjetos V1.00"
exibeTela "Aquisição automatizada de objetos celestes"
exibeTela " "

arquivoConfig = "C:\Observatorio Adhara\Config.txt"

' Inicializacao das objetos
	
Set telescopioMaxim = CreateObject("Maxim.Application") 'Objeto telescopio do Maxim
Set camera = CreateObject("MaxIm.CCDCamera")
Set fso = WScript.CreateObject("Scripting.FileSystemObject") 'Objeto File System
Set doc = WScript.CreateObject("MaxIm.Document") 'Objeto de documentos do Maxim
Set app = WScript.CreateObject("MaxIm.Application")'Objeto de aplicacoes do Maxim
Set conv = CreateObject("ASCOM.Utilities.Util")'Objeto de conversoes do ASCOM
Set novas = CreateObject("ASCOM.Astrometry.NOVAS.NOVAS31")'Objeto de conversoes do ASCOM

Set plano = New registroPlano
Set resultadoAlvo = New alvo

' Verifica a existencia do arquivo config

If Not fso.FileExists (arquivoConfig) Then

	exibeTela "Arquivo " & arquivoConfig & " inexistente, tecle <ENTER> para finalizar."
	WScript.StdIn.ReadLine
	mataObjetos ()
	WScript.Quit
	
End If

' Le arquivo Config

Set objArquivoTexto = fso.GetFile (arquivoConfig)
Set conf = objArquivoTexto.OpenAsTextStream (1, -2)

indiceConf = 0
	
Do Until conf.AtEndOfStream
	
	linhaDados = conf.ReadLine
	confDados = Split(linhaDados, ",")
	ReDim Preserve arrayConf (indiceConf + 1)
	arrayConf (indiceConf) = confDados (0)
	indiceConf = indiceConf + 1

Loop

conf.Close

' Inicializacao das variaveis

arquivoTexto = arrayConf (0)
flagVariavel = CInt(arrayConf (1))
flagDebug = CInt(arrayConf (2))
delimitador = arrayConf (3)
setPointCamera = CDbl(arrayConf (4))
quantidadeCap = CInt(arrayConf (5))
flagSequencia = CInt(arrayConf (6))
margemHorizonte = CDbl(arrayConf (7))
tentativaAlvo = CInt(arrayConf (8))
expApontamento = CDbl(arrayConf (9))
toleranciaSolve = CDbl(arrayConf (10)) / 60
cameraApontamento = CInt(arrayConf (11))
binningProcura = CInt(arrayConf (12))
exposicaoGuider = CDbl(arrayConf (13))
flagCalibra = CInt(arrayConf (14))
flagCalibraAlvo = CInt(arrayConf (15))
diretorioLocal = arrayConf (16) + "\"
flagDogsHeaven = CInt(arrayConf (17))
diretorioDogsHeaven = arrayConf (18) + "\"
prefixoArquivo = arrayConf (19)
subframeX = CInt(arrayConf (20))
subframeY = CInt(arrayConf (21))
otimizaGoto = CInt(arrayConf (25))
escolheTelescopio  = CInt(arrayConf (26))
flagGuider = CInt(arrayConf (27))
flagCool = CInt(arrayConf (28)) 
flagWarmup = CInt(arrayConf (29))
dither = CDbl(arrayConf (30))
flagSolve = CInt(arrayConf (31))
erroSolve = CInt(arrayConf (32))
flagGrandeCampo = CInt(arrayConf (33))
isoGrandeCampo = CInt(arrayConf (34))
binningGrandeCampo = CInt(arrayConf (35))
flagCalibraGrandeCampo = CInt(arrayConf (36))
exposicaoGrandeCampo = CInt(arrayConf (37))
prefixoGrandeCampo = arrayConf (38)
timeoutGrandeCampo = CInt(arrayConf (39))
trocaMeridiano = CInt(arrayConf (43))
ano = CStr(Year(Now))
mes = zeroEsquerda(Month(Now))
dia = zeroEsquerda(Day(Now))
hora = zeroEsquerda(Hour(Now))
minuto = zeroEsquerda(Minute(Now))
segundo = zeroEsquerda(Second(Now))
agora = Now
complemento = Ano + Mes + dia +"\"
diretorioCalibracao = diretorioLocal + "Calibracao\"
diretorioFlat = diretorioLocal + "Calibracao\Flats\"
diretorioDark = diretorioLocal + "Calibracao\Dark\"
diretorioBias = diretorioLocal + "Calibracao\Bias\"
diretorioGuider = diretorioLocal + "Calibracao\Guider\"
diretorioLight = diretorioLocal + "Lights\"
diretorioTemp = diretorioLocal + "Temp\"
diretorioLog = diretorioLocal + "Log\"
diretorioRet = diretorioLocal + "Ret\"
diretorioImagem = diretorioLight + complemento
diretorioStack = diretorioImagem + "Stack\"
diretorioReport = diretorioDogsHeaven + "Reports\"
diretorioProntas = diretorioDogsHeaven + "Prontas\"
arquivoLog = diretorioLog + "Log-" + Ano + Mes + dia + "-" + hora + minuto + segundo + ".txt"
arquivoRetorno = diretorioRet + "Ret-" + Ano + Mes + dia + "-" + hora + minuto + segundo + ".txt"
arquivoTemp = diretorioRet + "Temp.txt"
arquivoCapturados = diretorioLog + "Capturados.txt"
complementoSolve = "-PlateSolve0"

' Inicializa objeto Telescopio

If (escolheTelescopio = 1) then

	Set chsr =  CreateObject("ASCOM.Utilities.Chooser")
	chsr.DeviceType = "Telescope"
	scopeProgID = chsr.Choose(scopeProgID)
	Set telescopio = CreateObject(scopeProgID)

Else

	Set telescopio = CreateObject("Maxpoint.Telescope") 'Objeto telescopio Maxpoint
		
End If

' Cria pastas 

criaDiretorio (diretorioCalibracao)
criaDiretorio (diretorioFlat)
criaDiretorio (diretorioDark)
criaDiretorio (diretorioBias)
criaDiretorio (diretorioGuider)
criaDiretorio (diretorioLight)
criaDiretorio (diretorioImagem)
criaDiretorio (diretorioTemp)
criaDiretorio (diretorioLog)
criaDiretorio (diretorioRet)

' Deleta os arquivos no diretorio Temp

If (flagDebug = 0) Then
'lixo
	fso.DeleteFile(diretorioTemp + "*.*")
	
End If

' Cria arquivo de log

Set log = fso.CreateTextFile(arquivoLog, 8, True)

' Verifica a existencia do arquivo texto

If Not fso.FileExists (diretorioLocal + arquivoTexto) Then

	exibeTela "Arquivo " & arquivoTexto & " inexistente, tecle <ENTER> para finalizar."
	WScript.StdIn.ReadLine
	mataObjetos ()
	WScript.Quit

End If

' Verifica a existencia do arquivo temporario

If fso.FileExists (arquivoTemp) Then

	fso.DeleteFile(arquivoTemp)
	
End If

' Cria arquivo de temporario

Set ret = fso.CreateTextFile(arquivoTemp, 8, True)

' Verifica arquivo de capturados

If (flagSequencia = 1) Then

	If fso.FileExists (arquivoCapturados) Then

		trataCapturados " ", 0

	End If
	
End If

' Conecta as cameras 

camera.LinkEnabled = True
camera.DisableAutoShutdown = True

If Not camera.LinkEnabled Then

	exibeTela "Falha ao conectar as câmeras, tecle <ENTER> para finalizar."
   	WScript.StdIn.ReadLine
   	mataObjetos ()
   	WScript.Quit
	
End If

' Configura camera

If (flagCool = 1) Then

	camera.TemperatureSetpoint = setPointCamera
	camera.CoolerOn = True

	If (camera.Temperature > setPointCamera) Then

		exibeTela "Aguardando o resfriamento da câmera..."
		exibeTela " "
	
	End If

	Do While camera.Temperature > setPointCamera

		WScript.Sleep 10
		
	Loop
	
End If

nomesFiltros = camera.FilterNames

' Conecta o telescopio

telescopioMaxim.TelescopeConnected = True 
telescopio.Unpark
telescopio.Tracking = False
exibeTela Telescopio.Description

' Cria grupos de calibração

exibeTela "Carregando frames de calibração"
exibeTela " "
diretorioCalibracao = diretorioBias + ";" + diretorioDark + ";" + diretorioFlat  + ";" + diretorioGuider
app.CreateCalibrationGroups diretoriocalibracao, 1, 1, False
	
Do While telescopio.Slewing

	WScript.Sleep 1

Loop

' Calibracao da guiagem

telescopio.Tracking = True

If (flagDebug = 0 And flagGuider = 1) Then

' Aponta para o Zenite
'lixo
	localDec = telescopio.SiteLatitude 
	localRa = telescopio.SiderealTime - 0.5
	telescopio.SlewToCoordinates localRa, localDec  

	Do While telescopio.Slewing

		WScript.Sleep 1

	Loop
	
	exibeTela "Azimute: " & CStr(FormatNumber (telescopio.Azimuth, 2, , vbTrue))
	exibeTela "Altitude: " & CStr(FormatNumber (telescopio.Altitude, 2, , vbTrue))
	exibeTela "RA: " & conv.HoursToHMS (telescopio.SiderealTime, "h", "m", "s", 2)
	exibeTela "DEC: " & conv.DegreesToDMS (telescopio.SiteLatitude, Chr(176), "´", Chr(34), 2)
	exibeTela " "
	exibeTela "Calibrando a guiagem... "
	exibeTela " "
'lixo
	If (controlaGuiagem (1, exposicaoGuider, telescopio.SiteLatitude) = 0) Then  ' - 1

		exibeTela "Calibragem Ok... "
		exibeTela " "
		flagGuider = 1
	
	Else

		exibeTela "Calibragem falhou... "
		exibeTela " "
		flagGuider = 0
	
	End If
	
Else 

	exibeTela " "	
	
End If
	
' Carrega arquivo de planos

Set objArquivoTexto = fso.GetFile (diretorioLocal + arquivoTexto)
Set reg = objArquivoTexto.OpenAsTextStream (1, -2)

linhaDados = reg.ReadLine
indicePlano = 0

Do Until reg.AtEndOfStream 

	ReDim Preserve arrayPlano (indicePlano + 1)
	Set arrayPlano (indicePlano) = New registroPlano
	Set arrayPlano (indicePlano) = leDadosTexto (delimitador, flagVariavel)
	indicePlano = indicePlano + 1

Loop

reg.Close

If (otimizaGoto = 1) Then

	exibeTela "Otimizando plano de captura..... "
	exibeTela " "

	Set arrayPlano = classificaPlano (arrayPlano)

End If

' Inicio da aquisicao

exibeTela "Iniciando a aquisição..... "
exibeTela " "

indicePlano = -1
flagConfirmacao = 1
contadorCap = 1
statusReport = " "

Do Until (indicePlano >= UBound (arrayPlano) And flagConfirmacao = 0) 

	If (flagDogsHeaven = 0) Then

		flagConfirmacao = 0
	
	Else

		flagConfirmacao = statusConfirmacao ()

	End If
	
	If (flagConfirmacao = 1) Then
	
		Do 

			statusReport = confereReport ()
			flagConfirmacao = statusConfirmacao ()

		Loop While statusReport <> " "  And flagConfirmacao = 1
		
		
	End If

	indicePlano = indicePlano + 1

	If (indicePlano < UBound (arrayPlano)) Then
	
		If (flagSequencia = 1) Then 

			capturado = procuraCapturados (arrayPlano (indicePlano).objeto, arrayPlano (indicePlano).filtro)
			
		Else
		
			capturado = 0
			
		End If

		If (capturado = 0) Then 
		
			If (confereCeu (margemHorizonte, arrayPlano (indicePlano).raDecimal, arrayPlano (indicePlano).decDecimal) = 1) Then
	
				exibeDadosTexto arrayPlano (indicePlano), flagVariavel

				Set resultadoAlvo = procuraAlvo (arrayPlano (indicePlano).objeto, arrayPlano (indicePlano).raDecimal, arrayPlano (indicePlano).decDecimal, arrayPlano (indicePlano).ra, arrayPlano (indicePlano).dec, cameraApontamento, binningProcura, expApontamento, arrayPlano (indicePlano).filtro, 0, arrayPlano (indicePlano).exposicao)

				exibeTela " "
	
				If (resultadoAlvo.alvo = 1) Then	
				
					exibeTela "Alvo encontrado, iniciando a aquisição... "
					exibeTela " "
					capturaAlvo arrayPlano (indicePlano).objeto, arrayPlano (indicePlano).raDecimal, arrayPlano (indicePlano).decDecimal, arrayPlano (indicePlano).exposicao, arrayPlano (indicePlano).filtro, arrayPlano (indicePlano).binning, arrayPlano (indicePlano).frames, arrayPlano (indicePlano).stack, flagGuider, exposicaoGuider, arrayPlano(indicePlano).pixelsX, arrayPlano(indicePlano).pixelsY, arrayPlano (indicePlano).maximo, arrayPlano (indicePlano).minimo, arrayPlano (indicePlano).periodo, arrayPlano (indicePlano).tipo, flagVariavel, prefixoArquivo, resultadoAlvo.x, resultadoAlvo.y

				Else
		
					exibeTela "Alvo fora de alcance... "
					exibeTela " "
			
				End If
			
				If (flagSequencia = 1) Then
			
					trataCapturados  arrayPlano (indicePlano).objeto, arrayPlano (indicePlano).filtro
				
				End if

				contadorCap = contadorCap + 1
			
				If (contadorCap > quantidadeCap) Then
			
					indicePlano = UBound (arrayPlano) + 1
				
				End If 
		
			End If
			
		End If
		
	Else
	
		WScript.Sleep 15000		

	End If
	
Loop

If (flagGrandeCampo = 1) Then
	
	geraGrandeCampo " ", 0

End If

app.CloseAll
ret.Close
fso.CopyFile arquivoTemp, arquivoRetorno, True

If fso.FileExists (arquivoCapturados) Then
		
	fso.DeleteFile (arquivoCapturados)
		
End If

' Retorna o telescipio para posicao Home

telescopio.Park

' Aquece a camera

If (flagWarmup = 1) Then

	warmUp ()
	
End If

exibeTela "Tempo total de execução: " & CStr(TimeSerial (0, 0, DateDiff("s", agora, Now)))
exibeTela ""
exibeTela "Tecle <ENTER> para finalizar."
Log.Close

' Desconecta as cameras 

camera.LinkEnabled = False
camera.DisableAutoShutdown = False

' Desconecta o telescopio

telescopioMaxim.TelescopeConnected = False

WScript.StdIn.ReadLine
mataObjetos ()


' Funcao para inserir o zero a esquerda

Function zeroEsquerda(valor)

   if  Valor < 10 Then
   
      zeroEsquerda = "0" + CStr(valor)

   else

      zeroEsquerda =  CStr(valor)
   
   End If

End Function

' Funcao para criar as pastas

Function criaDiretorio(novoDiretorio)

   If Not fso.FolderExists (novoDiretorio)  Then
   
      Set fsoPasta = fso.CreateFolder(novoDiretorio)
   
   End If

End Function

' Funcao selecionar a ultima pasta

Function selecionaDiretorio(pastaOrigem)

	Dim ultimaPasta

   	For Each subfolder in fso.GetFolder (pastaOrigem).SubFolders
   	
            ultimaPasta =  subfolder.Name

   	Next
   	
   	selecionaDiretorio = ultimaPasta

End Function

' Funcao para ler os dados do arquivo texto

Function leDadosTexto (leDelimitador, leVariavel)

	Dim dadosLidos
	Dim regDadosLidos

	Set rDados = New registroPlano

	dadosLidos = reg.ReadLine
	regDadosLidos = Split(dadosLidos, leDelimitador)
	rDados.objeto = regDadosLidos (0)
	rDados.ra = regDadosLidos (1)	
	rDados.dec = regDadosLidos (2)
	rDados.exposicao = CInt(regDadosLidos (3))
	rDados.filtro = CInt(regDadosLidos (4))
	rDados.binning = CInt(regDadosLidos (5))
	rDados.frames = CInt(regDadosLidos (6))
	rDados.stack = CInt(regDadosLidos (7))
	rDados.pixelsX = CInt(regDadosLidos (8))
	rDados.pixelsY = CInt(regDadosLidos (9))
	
	If (leVariavel = 1) Then
		
		rDados.maximo = regDadosLidos (10)
		rDados.minimo = regDadosLidos (11)
		rDados.periodo = regDadosLidos (12)
		rDados.tipo = regDadosLidos (13)
		
	End If
	
	rDados.raDecimal = conv.HMSToHours (rDados.ra)
	rDados.decDecimal = conv.DMSToDegrees (rDados.dec)

	Set leDadosTexto = rDados
	
End Function

' Funcao para exibir os dados do arquivo texto

Function exibeDadosTexto (eDados, eVariavel)

	exibeTela "Objeto: " & eDados.objeto
	exibeTela "RA: " & eDados.ra
	exibeTela "DEC: " & eDados.dec
	exibeTela "Exposição: " & CStr(eDados.exposicao) & " segundos"
	
	If (UBound (nomesFiltros) > 0) Then 
	
		exibeTela "Filtro: " & nomesFiltros (eDados.filtro)
	
	End If
	
	exibeTela "Frames: " & CStr(eDados.frames) & " Empilhamento: " & CStr(eDados.stack)
	
	If (eDados.pixelsX = 0) Then
	
		exibeTela "Pixels X: Full"

	Else  

		exibeTela "Pixels X: " & CStr(eDados.pixelsX)
		
	End If
	
	If (eDados.pixelsY = 0) Then
	
		exibeTela "Pixels Y: Full"

	Else  

		exibeTela "Pixels Y: " & CStr(eDados.pixelsY)
		
	End If
	
	If (eVariavel = 1) Then
	
		exibeTela "Magnitudes - Max: " & eDados.maximo & "  Min: " & eDados.minimo
		exibeTela "Período: " & eDados.periodo & " dias"
		exibeTela "Tipo: " & eDados.tipo
		
	End If
				
	exibeTela " "
	
End Function

' Funcao para conferir se o alvo está aparente no ceu

Function confereCeu (horizonte, raCeu, decCeu)

	Dim flagCeu

	Set teleCalc = CreateObject("ScopeSim.Telescope") 
	
	teleCalc.Connected = True
	teleCalc.Tracking = True
	teleCalc.SlewToCoordinates  raCeu, decCeu
	
	Do While teleCalc.Slewing

		WScript.sleep  1

	Loop
	
	If (teleCalc.Altitude >= horizonte) Then
	
		flagCeu = 1
		
	Else
	
		flagCeu = 0
		
	End If
	
	teleCalc.Connected = False
	Set teleCalc = Nothing
	confereCeu = flagCeu

End Function

' Funcao para apontar o telescopio

Function procuraAlvo (procObjeto, procRa, procDec, procRaStr, procDecStr, procCamera, procBinning, procExposicao, procFiltro, procOpcaoSolve, procExposicaoCaptura)

	Dim contAlvo
	Dim contErro
		
	Set prSolve = New retornoSolve
	Set prAlvo = New alvo
	
	exibeTela "Procurando o alvo... "
	exibeTela " "
	camera.SetFullFrame ()
	telescopio.TargetRightAscension = procRa
	telescopio.TargetDeclination = procDec
	contErro = 0
	prAlvo.alvo = 0

	For contAlvo = 1 to tentativaAlvo
	
		If (trocaMeridiano > 0) Then
exibeTela "ENTRANDO"
			confereMeridiano procRa, contAlvo, tentativaAlvo, procExposicao, procExposicaoCaptura
exibeTela "SAINDO"
		End If

		prSolve.solve = 0
		prSolve.ra = 0
 		prSolve.dec = 0
		prSolve.x = 0
		prSolve.y = 0
		exibeTela "Tentativa " & CStr(contAlvo) & "..."
		telescopio.SlewToTarget
		
		Do While telescopio.Slewing

			WScript.Sleep 1

		Loop

		WScript.sleep  3000
		
		If (procCamera = 1) Then
		
			camera.BinX procBinning
			camera.BinY procBinning
			camera.Expose procExposicao, 1, procFiltro

			Do While Not camera.ImageReady

				WScript.Sleep 1

			Loop

			camera.SetFITSKey "VOBJETO", procObjeto
			camera.SetFITSKey "VRA", conv.HoursToHMS (procRa, " ", " ", " ", 3)		
			camera.SetFITSKey "VDEC", conv.DegreesToDMS (procDec, " ", " ", " ", 3)	
			arquivo = diretorioTemp + procObjeto + complementoSolve + CStr(zeroEsquerda(contAlvo)) + ".fit"

			If (flagDebug = 1) Then
		
				camera.Document.Close
				
				If fso.FileExists (arquivo) Then
				
					doc.OpenFile arquivo
					
				Else 
				
					prAlvo.alvo = 0
					prAlvo.x = 0
					prAlvo.y = 0
					Exit For
						
				End If
			
			Else
'lixo			
		
'camera.Document.Close
'doc.OpenFile arquivo

				If (flagCalibraAlvo = 1) Then
				
					camera.Calibrate
					
				End If
'lixo				
				camera.SaveImage arquivo
										
			End if
			
		Else

			camera.GuiderExpose procExposicao

			Do While camera.GuiderRunning

				WScript.Sleep 1

			Loop
			
			app.CurrentDocument.SetFITSKey "VOBJETO", procObjeto
			app.CurrentDocument.SetFITSKey "VRA", conv.HoursToHMS (procRa, " ", " ", " ", 3)	
			app.CurrentDocument.SetFITSKey "VDEC", conv.DegreesToDMS (procDec, " ", " ", " ", 3)
			arquivo = diretorioTemp + procObjeto + complementoSolve + CStr(zeroEsquerda(contAlvo)) + ".fit"
			
			If (flagDebug = 1) Then
			
				app.CurrentDocument.Close
				doc.OpenFile arquivo
				
			Else
		
				If (flagCalibraAlvo = 1) Then
		
					app.CurrentDocument.Calibrate
					
				End If
			
				app.CurrentDocument.SaveFile arquivo , 3, False, 3, 1
				
			End If

		End If

		Set prSolve = confereCoordenadas (arquivo, procOpcaoSolve, 0)
		
'exibeTela CStr (prSolve.solve)
'exibeTela CStr (prSolve.ra)
'exibeTela CStr (prSolve.dec)
'exibeTela CStr (prSolve.x)
'exibeTela CStr (prSolve.y)
	
		If (prSolve.solve = 1) Then
	
			telescopio.SyncToCoordinates prSolve.ra, prSolve.dec

			WScript.sleep  1000	
			contErro = 0
					
			If (Abs (prSolve.ra  - procRa) <= toleranciaSolve And Abs (Abs (prSolve.dec) - Abs (procDec)) <=  toleranciaSolve) Then
			
				prAlvo.alvo = 1
				prAlvo.x = prSolve.x
				prAlvo.y = prSolve.y
				Exit For
				
			Else
			
				exibeTela "RA: " & conv.HoursToHMS (prSolve.ra, "h", "m", "s",2) & "    " & "DEC: " & conv.DegreesToDMS (prSolve.dec, Chr(176), "´", Chr(34), 2)
				exibeTela " "
				prAlvo.alvo = 0
				prAlvo.x = 0
				prAlvo.y = 0
				
			End If
			
		Else 
		
			contErro = contErro + 1			
					
		End If
		
		If (contErro = erroSolve) Then
		
			prAlvo.alvo = 0
			prAlvo.x = 0
			prAlvo.y = 0
			Exit For
			
		End If
			
	Next
	
	app.CurrentDocument.Close
	Set procuraAlvo = prAlvo
	
End Function

' Funcao para capturar as imagens


Function capturaAlvo (capObjeto, capRa, capDec, capExposicao, capFiltro, capBinning, capFrames, capStack, capGuiagem, capExpoGuiagem, capPixelsX, capPixelsY, capMaximo, capMinimo, capPeriodo, capTipo, capVariavel, capPrefixo, capX, CapY)

	Dim contExposicao
	Dim nomeArquivo
	Dim contadorStack
	Dim sequenciaStack
	Dim resultadoSolve
	Dim contador
	Dim arquivoDogsHeaven
	Dim posicaoX
	Dim posicaoY
	Dim dataImagem
	Dim nomeArquivoGrandeCampo
	Dim flagTimeout
	Dim contadorTimeout
	Dim flagMeridiano
	
	
	Set rSolve = New retornoSolve
	Set meridianoAlvo = New alvo
	
	contadorStack = 0
	sequenciaStack = 1
	dataImagem = Left (complemento, 8)
	
	camera.BinX capBinning
	camera.BinY capBinning
	
	If (capPixelsX = 0 And capPixelsY = 0) Then
	
		camera.SetFullFrame ()
		
	Else 
	
'WScript.Echo capPixelsX
'WScript.Echo capPixelsY
'WScript.Echo capX
'WScript.Echo capY	
'WScript.Echo capBinning
'WScript.Echo " "
	
		If (capBinning = 1) Then
		
			posicaoX = capX * binningProcura
			posicaoY = capY * binningProcura

			
		Else

		 	posicaoX = capX * (binningProcura + 1) / capBinning
			posicaoY = capY * (binningProcura + 1) / capBinning
				
		End If	
		
		camera.StartX = Int(posicaoX - (capPixelsX / 2))
		camera.StartY = Int(posicaoY - (capPixelsY / 2))
		camera.NumX = capPixelsX
		camera.NumY = capPixelsY
		
'WScript.Echo camera.StartX
'WScript.Echo camera.StartY
'WScript.Echo camera.NumX
'WScript.Echo camera.NumY
'WScript.Echo " "		
	
	End if
	
	If (capGuiagem = 1) Then
	
		controlaGuiagem 2, capExpoGuiagem, capDec
		
	End If
	
	For contExposicao = 1 to capFrames
	
		If (trocaMeridiano > 0) Then
		
			If (flagGrandeCampo = 1 And contExposicao = 1) Then
			
				flagMeridiano = confereMeridiano (capRa, 1, 1, exposicaoGrandeCampo, 0)
				
			Else

				flagMeridiano = confereMeridiano (capRa, 1, 1, capExposicao, 0)
				
			End If
			
			If (flagMeridiano = 1) Then

				If (capGuiagem = 1) Then

					controlaGuiagem 0, 0, 0
		
				End If
				
				Set meridianoAlvo = procuraAlvo (capObjeto, capRa, capDec, conv.HoursToHMS (capRa, " ", " ", " ", 3), conv.DegreesToDMS (capDec, " ", " ", " ", 3), cameraApontamento, binningProcura, expApontamento, capFiltro, 0, 0)

				If (meridianoAlvo.alvo = 1) Then	
				
					exibeTela "Troca de medidiano bem sucedida, retomando a aquisição... "
					exibeTela " "

				Else
		
					exibeTela "Troca de medidiano mau sucedida, cancelando a aquisição... "
					exibeTela " "
					Exit For
			
				End If
				
				If (capGuiagem = 1) Then
	
					controlaGuiagem 2, capExpoGuiagem, capDec
		
				End If
			
			End If

		End If
		
		If (flagGrandeCampo = 1 And contExposicao = 1) Then
	
			nomeArquivoGrandeCampo = diretorioImagem + prefixoGrandeCampo + "_" + capObjeto + Ano + Mes + dia + "-" + hora + minuto + segundo + ".fit"
			geraGrandeCampo nomeArquivoGrandeCampo, exposicaoGrandeCampo

		End If
	
		camera.Expose capExposicao, 1, capFiltro

		Do While Not camera.ImageReady

			WScript.Sleep 1

		Loop

		If (capVariavel = 1) Then
	
			camera.SetFITSKey "VOBJETO", capObjeto
			camera.SetFITSKey "VRA", conv.HoursToHMS (capRa, " ", " ", " ", 3)
			camera.SetFITSKey "VDEC", conv.DegreesToDMS (capDec, " ", " ", " ", 3)
			camera.SetFITSKey "VEXPOSICAO", capExposicao
			camera.SetFITSKey "VFILTRO", capFiltro
			camera.SetFITSKey "VBINNING", capBinning
			camera.SetFITSKey "VFRAMES", capFrames	
			camera.SetFITSKey "VSTACK", capStack	
			camera.SetFITSKey "VMAXIMO", capMaximo						
			camera.SetFITSKey "VMINIMO", capMinimo
			camera.SetFITSKey "VPERIODO", capPeriodo
			camera.SetFITSKey "VTIPO", capTipo
			
		End If

		nomeArquivo = capObjeto + "0" + zeroEsquerda(capFiltro) + "-0" + zeroEsquerda(contExposicao) + ".fit"
		arquivo = diretorioImagem + nomeArquivo

		If (flagDebug = 0) Then
		
			If (flagCalibra = 1) Then	
			
				camera.Calibrate
				
			End If
'lixo		
			camera.SaveImage arquivo
					
		End If



		exibeTela arquivo
		ret.WriteLine nomeArquivo
		camera.Document.Close
		
		If (capStack > 1) Then
		
			contadorStack = contadorStack + 1
			
			If (contadorStack = 1) Then
			
				criaDiretorioStack ()
			
			End If
		
			fso.CopyFile arquivo, diretorioStack
'			doc.OpenFile arquivo

			If (contadorStack = capStack) Then
			
				exibeTela " "
				exibeTela "Empilhando imagens"
				exibeTela " "
				arquivo = diretorioStack + "*.fit"
				doc.CombineFiles arquivo, 0, False, 2, True
				nomeArquivo = capObjeto + "-S" + "0" + zeroEsquerda(capFiltro) + "-0" + zeroEsquerda(sequenciaStack) + ".fit"
				arquivo = diretorioImagem + nomeArquivo
				doc.SaveFile arquivo , 3, False, 3, 1
				ret.WriteLine nomeArquivo
'				doc.Close
				contadorStack = 0
				sequenciaStack = sequenciaStack + 1
				doc.Close
								
				If (flagSolve = 1 Or flagDogsHeaven = 1) Then
				
					Set rSolve  = confereCoordenadas (arquivo, 1, 0)
			
				End If

				If (flagDogsHeaven = 1 And rSolve.solve = 1) Then

					exibeTela " "				
					exibeTela "Movendo arquivo para: " & diretorioDogsHeaven & "AProcessar\"
					exibeTela " "	
					nomeArquivo = Replace (nomeArquivo, capObjeto, (capObjeto + "-" + dataImagem))
					arquivoDogsHeaven = diretorioDogsHeaven + "AProcessar\" + capPrefixo + nomeArquivo
				
					If fso.FileExists (arquivoDogsHeaven) Then 
					
						fso.DeleteFile arquivoDogsHeaven, True
						
					End if
					
					fso.MoveFile arquivo, arquivoDogsHeaven
					
				End If
				
			End If
			
		Else

			exibeTela " "	
			
			If (flagSolve = 1 Or flagDogsHeaven = 1) Then
				
				Set rSolve  = confereCoordenadas (arquivo, 1, 0)
			
			End If 
			
			If (flagDogsHeaven = 1 And capStack = 1 And rSolve.solve = 1) Then
				
				exibeTela " "
				exibeTela "Movendo arquivo para: " & diretorioDogsHeaven & "AProcessar\"	
				exibeTela " "	
				nomeArquivo = Replace (nomeArquivo, capObjeto, (capObjeto + "-" + dataImagem))
				arquivoDogsHeaven = diretorioDogsHeaven + "AProcessar\" + capPrefixo + nomeArquivo
				fso.MoveFile arquivo, arquivoDogsHeaven
					
			End If
			
		End If
		
'		If (capGuiagem = 1 And dither > 0) Then
		
'			geraDither ()
			
'		End if
		
	Next

	If (flagGrandeCampo = 1) Then
	
		flagTimeout = 0
		contadorTimeout = 0
		

		exibeTela "Aguardando imagem de grande campo - Timeout: " & timeoutGrandeCampo & "s"
		exibeTela " "
	
		Do Until flagTimeout = 1 Or fso.FileExists (nomeArquivoGrandeCampo)
		
			If (contadorTimeout < 10) Then 
			
				WScript.Sleep timeoutGrandeCampo * 100
				contadorTimeout = contadorTimeout + 1
				
			Else
			
				flagTimeout = 1
				
			End If
		
		Loop 
	
		If (fso.FileExists (nomeArquivoGrandeCampo)) Then
		
			exibeTela nomeArquivoGrandeCampo & " - Encontrado"
			exibeTela " "
			ret.WriteLine nomeArquivoGrandeCampo
		
			If (flagDogsHeaven = 1) Then
		
				Set rSolve = confereCoordenadas (nomeArquivoGrandeCampo, 1, 1)

				If (rSolve.solve = 1) Then

					exibeTela "Copiando arquivo para: " & diretorioDogsHeaven & "AProcessar\"
					exibeTela " "
					fso.CopyFile nomeArquivoGrandeCampo, diretorioDogsHeaven + "AProcessar\", True
					
				End If
					
			End If
			
		Else
		
			exibeTela nomeArquivoGrandeCampo & " - Não encontrado"
			exibeTela " "
				
		End If
		
	End If
	
	If (capGuiagem = 1) Then

		controlaGuiagem 0, 0, 0
		
	End If

	ret.WriteLine "*"
	app.CloseAll
	
End Function	
	
' Funcao para aquecer a camera lentamente

Function warmUp ()

	exibeTela "Câmera Warm-up..."
	exibeTela " "
	 
	Do While camera.TemperatureSetpoint  < 0
	
		camera.TemperatureSetpoint = 0
		WScript.Sleep tempoZero * 1000
		
	Loop
	
	camera.TemperatureSetpoint = 20
	
End Function

' Funcao para exibir o conteudo na tela e gravar no arquivo de log

Function exibeTela (exibir)

	Dim linha
	
	WScript.Echo exibir
	linha = Now () & "	" & exibir
	Log.WriteLine linha

	
End Function

' Funcao para tranformar graus em radianos

Function graus2Rad (valor)

	Dim pi
	
	pi = 4 * Atn(1)

	graus2Rad = valor * (pi / 180)
	
End Function	

' Funcao para tranformar radianos em graus

Function rad2Graus (valor)

	Dim pi
	
	pi = 4 * Atn(1)

	rad2Graus = valor * (180 / pi)
	
End Function

' Funcao para tranformar horas em radianos

Function hora2Rad (valor)

	hora2Rad = valor * 0.2618
	
End Function

' Funcao para tranformar radianos em horas

Function rad2Hora (valor)

	rad2Hora = valor / 0.2618
	
End Function

Function hora2Graus (valor)

	hora2Graus = valor * 15
	
End Function

Function graus2Hora (valor)

	graus2Hora = valor / 15
	
End Function

' Funcao para controlar a guiagem

' 0 - Encerra  1 - Calibra  2 - Guia

Function controlaGuiagem (opcao, tempo, declinacao)

	If (opcao = 0) Then
	
		camera.GuiderStop
		controlaGuiagem = 0
		
	End If
	
	If (opcao = 1) Then
	
		Err.Clear
	
		camera.GuiderExpose tempo
		
		Do While camera.GuiderRunning

			WScript.Sleep 1

		Loop
		
		camera.GuiderAutoSelectStar = True
		camera.GuiderDeclination = declinacao
		
		camera.GuiderCalibrate (tempo)
		
		Do While camera.GuiderCalState = 1 

			WScript.Sleep 1

		Loop
		
		If (camera.GuiderCalState = 2) Then  'Err.number <> 0 And 
		
			controlaGuiagem = 0	
			
		Else
		
			controlaGuiagem = 1
			
		End If
				
	End If
	
	If (opcao = 2) Then

		camera.GuiderExpose tempo
		
		Do While camera.GuiderRunning

			WScript.sleep  1

		Loop
		
		camera.GuiderAutoSelectStar = True
		camera.GuiderDeclination = declinacao
		camera.GuiderTrack (tempo)
		controlaGuiagem = 0
		WScript.Sleep 20000
		
	End If

End Function

' Plate Solve

Function confereCoordenadas (imagem, opcao, catalogo) ' Opcao 0 - Apontamento 1 - Captura  / Catalogo 0 - Geral 1 - Grande Campo

	Dim arquivoSolve
	Dim arquivoApm
	Dim apm
	Dim linhaApm
	Dim regApm
	Dim ra
	Dim dec
	Dim placaX
	Dim PlacaY
	Dim dimX
	Dim dimY
	Dim linhaComando
	Dim flagLicenca
	Dim centroX
	Dim centroY 
	
	Set rSolve = New retornoSolve
	Set sys = CreateObject("Shell.Application")
	
	If fso.FileExists (imagem) Then

		arquivoSolve = Replace (imagem, ".fit", ".apm")
		doc.OpenFile imagem
		centroX = CInt(doc.GetFITSKey ("NAXIS1")) / 2
		centroY = CInt(doc.GetFITSKey ("NAXIS2")) / 2
		ra = hora2Rad (conv.HMSToHours (doc.GetFITSKey ("OBJCTRA")))
		dec = graus2Rad (conv.DMSToDegrees (doc.GetFITSKey ("OBJCTDEC")))
		placaX = CStr (doc.GetFITSKey ("XPIXSZ")) / CStr (doc.GetFITSKey ("FOCALLEN")) * 206.265
		PlacaY = CStr (doc.GetFITSKey ("YPIXSZ")) / CStr (doc.GetFITSKey ("FOCALLEN")) * 206.265
		dimX = graus2Rad (CInt ((CStr (doc.GetFITSKey ("NAXIS1")) * placaX) / 60)) / 60
		dimY = graus2Rad (CInt ((CStr (doc.GetFITSKey ("NAXIS2")) * PlacaY) / 60)) / 60
		linhaComando = ra & "," & dec & "," & dimX & "," & dimY & "," & "999" & "," & arquivo & "," & "5"
		doc.close

		If (flagSequencia = 1) Then 

			If fso.FileExists (arquivoSolve) Then
		
				fso.DeleteFile (arquivoSolve)
		
			End If
			
		End If

		sys.ShellExecute "C:\Observatorio Adhara\PlateSolve2.exe", linhaComando
	
		Do While Not fso.FileExists (arquivoSolve)

			WScript.Sleep 10
			WScript.Sleep 1
	
		Loop
	
		Set arquivoApm = fso.GetFile (arquivoSolve)

		Do While arquivoApm.Size = 0

			WScript.Sleep 10
			WScript.Sleep 1
	
		Loop

		Set apm =  arquivoApm.OpenAsTextStream (1, -2)
		linhaApm = apm.ReadLine
		regApm = Split(linhaApm, ",")
		ra = regApm (0)
		dec = regApm (1)
		apm.Close
		fso.DeleteFile (arquivoApm)
		
	Else
		
		ra = 999
		dec = 999
		
	End If		

	If (ra <> 999 and dec <> 999)  Then
		
		If (opcao = 1) Then 
		
			Set plate = CreateObject("PinPoint.Plate")

			plate.SigmaAboveMean = 4         	' Amount above noise : default = 4.0 - 2.0 - 
			plate.MinimumBrightness = 200     	' Minimum star brightness: default = 200
			plate.CatalogMaximumMagnitude = 20  ' Maximum catalog magnitude: default = 20.0
			plate.CatalogExpansion = 0.3    	' Area expansion of catalog to account for misalignment: default = 0.3 - 0.8
			plate.MinimumStarSize = 2         	' Minimum star size in pixels: default = 2
			plate.CatalogMinimumMagnitude = -2
			plate.MaxSolveTime = 5000
			plate.UseFaintStars = False
			plate.AttachFITS imagem
			plate.RightAscension = rad2Hora (ra)
			plate.Declination = rad2Graus (dec)
			plate.ArcsecPerPixelHoriz = placaX
			plate.ArcsecPerPixelVert = PlacaY
			flagLicenca = CInt(arrayConf (24))
			
' GSC-ACT; accurate and easy to use 3 - GSC  10 - UCAC3 11 - UCAC4  - 5 - USNO A2.0	
		
			If (catalogo = 0) Then
			
				plate.Catalog =  CInt(arrayConf (22))
				plate.CatalogPath = arrayConf (23)


			Else

				plate.Catalog =  CInt(arrayConf (41))
				plate.CatalogPath = arrayConf (42)
				

			End if
		
'WScript.Echo plate.CacheImageStars
'WScript.Echo plate.ReadFITSValue ("OBJCTRA")
'WScript.Echo plate.RightAscension 
'WScript.Echo plate.ReadFITSValue ("OBJCTDEC")
'WScript.Echo plate.Declination
'WScript.Echo plate.ArcsecPerPixelHoriz
'WScript.Echo plate.ArcsecPerPixelVert 
		
			plate.FindImageStars
			plate.FindCatalogStars
	
			Set stars = plate.ImageStars
	
			exibeTela CStr(stars.Count) & " Estrelas encontradas..."
					
			Err.Clear
			plate.Solve
	
'exibeTela "Resultado"
'exibeTela CStr(plate.Solved)
'exibeTela CStr(Err.number)
	
			If (Err.number <> 0) Then
	
'WScript.Echo "entrou Erro"
'WScript.Echo Err.number

				exibeTela "Falha ao tentar o platesolve..."
 				rSolve.solve = 0
 				rSolve.ra = 0
				rSolve.dec = 0
				rSolve.x = 0
				rSolve.y = 0
 		
 			Else
 	
				exibeTela "Platesolve Ok..."
				exibeTela "Residual medio " & CStr(FormatNumber (plate.MatchAvgResidual, 2, , vbTrue)) & "  Ordem " &  CStr(plate.MatchFitOrder)
				
				If (flagDebug = 0) Then

					plate.UpdateFITS()
				
				End If
			
 				rSolve.solve = 1			
				rSolve.ra = plate.RightAscension
				rSolve.dec = plate.Declination
	
				If (plate.SkyToXy (conv.HMSToHours (plate.ReadFITSValue ("VRA")), conv.DMSToDegrees (doc.GetFITSKey ("VDEC")))) Then
			
					rSolve.x = plate.ScratchX
					rSolve.y = plate.ScratchY
				
				Else
			
					rSolve.x = 0
					rSolve.y = 0
				
				End If				
				
'WScript.Echo conv.HoursToHMS (plate.RightAscension, "h", "m", "s", 2)
'WScript.Echo plate.RightAscension 
'WScript.Echo conv.DegreesToDMS (plate.Declination, chr(176), "´", chr(34), 2)
'WScript.Echo plate.Declination
'WScript.Echo FormatNumber (Abs (plate.RightAscension  - ra), 10, , vbTrue)
'WScript.Echo FormatNumber (Abs (Abs (plate.Declination) - Abs (dec)), 10, , vbTrue)

				plate.DetachFITS

			End If
	 	
	 	Else
			 	
	 	 	rSolve.solve = 1
 			rSolve.ra = rad2Hora (ra)
			rSolve.dec = rad2Graus (dec)
			rSolve.x = centroX
			rSolve.y = centroY
			exibeTela "Platesolve Ok..."
			
 		End If
 			
  		Set plate = Nothing
 		Set stars = Nothing
 		
	Else
	
		rSolve.solve = 0
 		rSolve.ra = 0
		rSolve.dec = 0
		rSolve.x = 0
		rSolve.y = 0		
		
	End If

 	Set sys = Nothing
 	Set apm = Nothing
 	Set arquivoApm = Nothing
 	
  	Set confereCoordenadas = rSolve
 		
End Function

' Verifica o status do arquivo de confirmacao

Function statusConfirmacao

Dim arquivoConfirmacao
Dim dadosConfirmacao

arquivoConfirmacao = diretorioReport + "\SCRIPTSTATUS.TXT" 

	If fso.FileExists (arquivoConfirmacao) Then

		Set objArquivoTexto = fso.GetFile (arquivoConfirmacao)
		Set confirmacao = objArquivoTexto.OpenAsTextStream (1, -2)

		dadosConfirmacao = CInt(confirmacao.ReadLine)
		confirmacao.Close
		
	Else
	
		dadosConfirmacao = 0

	End If
	
	Set objArquivoTexto = Nothing
	Set confirmacao = Nothing
	
	statusConfirmacao = dadosConfirmacao
	
End Function

' Verifica a existencia de arquivos de report

Function confereReport ()

	Dim arquivoReport
	Dim dadosReport (6)
	Dim indReport
	Dim rArquivo
	Dim rSequencia
	
	Set planoReport = New registroPlano
	Set repAlvo = New alvo
	
	indReport = 0
	retornoAlvo = 0
	
	arquivoReport = procuraArquivo (diretorioReport, "OBJETOSUSPEITO_")

	If (arquivoReport <> " ") Then 
	
		Set objArquivoReport = fso.GetFile (diretorioReport + arquivoReport)
		Set report = objArquivoReport.OpenAsTextStream (1, -2)
	
		Do Until report.AtEndOfStream
	
			dadosReport (indReport) = report.ReadLine
			indReport = indReport + 1

		Loop
	
		report.Close
	
		rArquivo = dadosReport (0)
		rSequencia = Mid (arquivoReport, (InStr (arquivoReport, "_") + 1), InStrRev (arquivoReport, ".") - InStr (arquivoReport, "_") - 1)
		planoReport.objeto = "Conf_" + zeroEsquerda(rSequencia) + "_" + Right (dadosReport (0), Len (dadosReport (0)) - Len (prefixoArquivo))
		planoReport.objeto = Left (planoReport.objeto, InStrRev (planoReport.objeto, ".") - 1)
		planoReport.raDecimal = graus2Hora (CDbl (Mid (dadosReport (2), 4, 11)))
		planoReport.decDecimal = CDbl (Mid (dadosReport (2), 23, 11))
		planoReport.ra = Mid (dadosReport (3), 5, 15)
		planoReport.dec = Mid (dadosReport (3), 25, 15)
		planoReport.maximo = CDbl(Mid (dadosReport (4), 5, 5))
'planoReport.raDecimal = 14.492183
'planoReport.decDecimal = -62.67525000	
'WScript.Echo rArquivo
'WScript.Echo planoReport.objeto
'WScript.Echo planoReport.raDecimal
'WScript.Echo planoReport.decDecimal
'WScript.Echo planoReport.ra
'WScript.Echo planoReport.dec
'WScript.Echo planoReport.maximo
'WScript.Echo " "

		If (confereCeu (margemHorizonte, planoReport.raDecimal, planoReport.decDecimal) = 1) Then

			doc.OpenFile (diretorioProntas + rArquivo)
			planoReport.exposicao = CDbl(doc.GetFITSKey("VEXPOSIC"))
			planoReport.filtro = CInt(doc.GetFITSKey("VFILTRO"))
			planoReport.frames = CInt(doc.GetFITSKey("VFRAMES"))
			planoReport.stack = CInt(doc.GetFITSKey("VSTACK"))
			planoReport.Binning = CInt(doc.GetFITSKey("VBINNING"))
			exibeDadosTexto planoReport, 0
			doc.Close

			Set repAlvo = procuraAlvo (planoReport.objeto, planoReport.raDecimal, planoReport.decDecimal, planoReport.ra, planoReport.dec, cameraApontamento, binningProcura, planoReport.exposicao, planoReport.filtro, 0, planoReport.exposicao)
			exibeTela " "
			
			If (repAlvo.alvo = 1) Then	
				
				exibeTela "Alvo encontrado, iniciando a aquisição... "
				exibeTela " "
				capturaAlvo planoReport.objeto, planoReport.raDecimal, planoReport.decDecimal, planoReport.exposicao, planoReport.filtro, planoReport.binning, planoReport.frames, planoReport.stack, flagGuider, exposicaoGuider, subframeX, subframeY, 0, 0, 0, " ", 0, prefixoArquivo, repAlvo.x, repAlvo.y

			Else
		
				exibeTela "Alvo fora de alcance... "
				exibeTela " "
			
			End If
		
		End If

		If (repAlvo.alvo = 1) Then
		
			fso.DeleteFile (diretorioReport + arquivoReport)
			
		Else
	
			fso.MoveFile diretorioReport + arquivoReport, diretorioReport + "Falha_" + arquivoReport
			
		End If
			
		
	End If
	
	confereReport = arquivoReport 

End Function

' Verifica a existencia de determinado arquivo

Function procuraArquivo(procDiretorio, procArquivo)

	Dim arquivoTemp

	Set pastaLeitura = fso.GetFolder(procDiretorio)
	Set arquivoLeitura = PastaLeitura.Files

	arquivoTemp = " "

	For each fileIdx In arquivoLeitura
   
 '  		WScript.Echo UCase(Left(fso.GetFileName (fileIdx.name),  Len(procArquivo)))
      
		If UCase(Left(fso.GetFileName (fileIdx.name),  Len(procArquivo))) = procArquivo Then

			arquivoTemp = fileIdx.name
			Exit For
            
 		End If
      
	Next
	
	Set pastaLeitura = Nothing
	Set arquivoLeitura = Nothing

	procuraArquivo = ArquivoTemp

End Function

Function classificaPlano ( listaObjetos ())

	Dim claIndice
	Dim raHorizonteOeste
	Dim periodo
	Dim totalArray
	Dim i
	Dim j
	Dim k
	Dim colunas 
		
    colunas = Array (13776, 4592, 1968, 861, 336, 112, 48, 21, 7, 3, 1)
	
	Set regAux = New registroPlano
	
	raHorizonteOeste = Telescopio.SiderealTime - 6
	
	If (raHorizonteOeste < 0) Then
	
		raHorizonteOeste = 23.999999 + raHorizonteOeste
	
	End if
	
	totalArray = UBound (listaObjetos)
	
'WScript.Echo raHorizonteOeste
'WScript.Echo " "	
	
	For claIndice = 0 To totalArray
	
		If (listaObjetos(claIndice).raDecimal  >= raHorizonteOeste) then
		
			listaObjetos (claIndice).vetorRa = listaObjetos (claIndice).raDecimal  - raHorizonteOeste
			
		Else
		 
		 	listaObjetos (claIndice).vetorRa = 23.999999 -  raHorizonteOeste + listaObjetos (claIndice).raDecimal  
		 	
		End If
		
		listaObjetos (claIndice).vetorDec = distanciaObjeto (raHorizonteOeste, 0, listaObjetos (claIndice).raDecimal, listaObjetos (claIndice).decDecimal, listaObjetos (claIndice).vetorRa)
		listaObjetos (claIndice).sequencia = (listaObjetos (claIndice).vetorRa + listaObjetos (claIndice).vetorDec) / 2
		
'WScript.Echo listaObjetos (claIndice).objeto + "	"  + CStr(FormatNumber (listaObjetos (claIndice).raDecimal, 6, , vbTrue))  + "	"  + CStr(FormatNumber (CStr(listaObjetos (claIndice).vetorRa), 6, , vbTrue)) + "	"  + CStr(FormatNumber (CStr(listaObjetos (claIndice).vetorDec), 6, , vbTrue)) + "	"  + CStr(FormatNumber (CStr(listaObjetos (claIndice).sequencia), 6, , vbTrue))
		
	Next
	
'WScript.Echo " "

	For k = 0 To UBound (colunas)
	
		periodo = colunas (k)

		For i = periodo To totalArray
		
			j = i
					
			Set regAux = listaObjetos (i)
			
			Do While j >= periodo
			
				If (listaObjetos (j - periodo).sequencia <= regAux.sequencia) Then
				
					Exit Do
					
				End If
				
				Set listaObjetos (j) = listaObjetos (j - periodo)
				
				j = j - periodo
				
			Loop
			
			Set listaObjetos (j) = regAux	
		
		Next
	
	next
	
'WScript.Echo " "	
	
'For claIndice = 0 To UBound (listaObjetos)
	
'	WScript.Echo listaObjetos (claIndice).objeto + "	"  + CStr(FormatNumber (listaObjetos (claIndice).raDecimal, 6, , vbTrue))  + "	"  + CStr(FormatNumber (CStr(listaObjetos (claIndice).vetorRa), 6, , vbTrue)) + "	"  + CStr(FormatNumber (CStr(listaObjetos (claIndice).vetorDec), 6, , vbTrue)) + "	"  + CStr(FormatNumber (CStr(listaObjetos (claIndice).sequencia), 6, , vbTrue))
		
'Next

'WScript.Echo " "
	
End Function 

Function distanciaObjeto (raInicial, decInicial, raFinal, decFinal, vetor)

	Dim distancia
	Dim circumpolar
	
	distancia = Sin (graus2Rad (decInicial)) * Sin (graus2Rad (decFinal)) + Cos (graus2Rad (decInicial)) * Cos (graus2Rad (decFinal)) * Cos (hora2Rad (raInicial) - hora2Rad (raFinal))
	distancia = rad2Graus (Atn(-distancia / Sqr(-distancia * distancia + 1)) + 2 * Atn(1))
	circumpolar = -1 * (90 + telescopio.SiteLatitude)

'WScript.Echo decFinal	
'WScript.Echo circumpolar	

	If (vetor > 12 And decFinal > circumpolar) Then
	
		distancia = 360 - distancia
		
	End if
	
	distanciaObjeto = distancia

End Function

' Le e grava objetos capturados

Function trataCapturados (objeto, filtro)

	Dim cap
	Dim objArquivoCapturados
	Dim i

	If (objeto = " ") Then 
		
		Set objArquivoCapturados = fso.GetFile (arquivoCapturados)
		Set cap = objArquivoCapturados.OpenAsTextStream (1, -2)
		
		indiceCapturados = 0

		Do Until cap.AtEndOfStream

			ReDim Preserve arrayCapturados (indiceCapturados + 1)
			arrayCapturados (IndiceCapturados) = cap.ReadLine
			indiceCapturados = indiceCapturados + 1

		Loop
		
	Else
	
		If (Not fso.FileExists (arquivoCapturados)) Then
		
			Set cap = fso.CreateTextFile(arquivoCapturados, 8, True)		
			Set cap = Nothing

		End If
				
		Set objArquivoCapturados = fso.GetFile (arquivoCapturados)
		Set cap = objArquivoCapturados.OpenAsTextStream (8, -2)
			
		cap.WriteLine objeto + CStr (filtro)
		ReDim Preserve arrayCapturados (UBound(arrayCapturados) + 1)
		arrayCapturados (UBound(arrayCapturados) - 1) = objeto + CStr (filtro)
		
	End If
	
	cap.Close
	
	Set cap = Nothing
	Set objArquivoCapturados = Nothing
	
End Function

' Verifica que objeto já foi procurado / capturado

Function procuraCapturados (objeto, filtro)

	Dim indice
	Dim resutado
	
	resutado = 0

	For Indice = 0 To UBound (arrayCapturados)

		If (arrayCapturados (indice) =  objeto + CStr (filtro)) Then
		
			resultado = 1
			Exit for
			
		End If
	
	Next 

	procuraCapturados = resultado

End function	

' Escolhe a direcao do Dither

Function geraDither

	Dim direcao
	
	direcao = Rnd 
	
	If (direcao < 0.25) Then
	
		moveDither (0)
		
	Else
	
		If (direcao >= 0.25 And direcao < 0.50) Then
		
			moveDither (1)
			
		Else
		
			If (direcao >= 0.50 And direcao < 0.75) Then
			
				moveDither (2)
				
			Else
			
				moveDither (3)
				
			End If
		
		End If
	   
	   
	End If
		
End Function

' Executa o movimento do dither

Function moveDither (direcao)

	camera.GuiderMove direcao, dither

	Do While camera.GuiderMoving

			WScript.Sleep 1

	Loop
	
	WScript.Sleep 2000

End Function

' Abora programa caso seja pressionado ESC

Function abortaProcesso ()

	Dim tecla
	
	tecla = WScript.StdIn.Read
	
	If (tecla = Chr (8)) Then
	
		mataObjetos ()
		WScript.Quit
		
	End If
	
	WScript.Sleep 1

End Function

' Cria pasta para o stack

Function criaDiretorioStack ()

	If fso.FolderExists (diretorioStack) Then

		fso.DeleteFolder Left (diretorioStack, Len (diretorioStack) - 1), True
		fso.CreateFolder (diretorioStack)
		
	Else
		
		fso.CreateFolder (diretorioStack)	
	
	End If

End Function

' Gera arquivo de grande campo

Function geraGrandeCampo (arquivo, exposicao)

	Dim arqGrandeCampo
	
	arqGrandeCampo = diretorioTemp + "GrandeCampo.txt"

	Set agc = fso.CreateTextFile(arqGrandeCampo, 8, True)
	agc.WriteLine exposicao
	agc.WriteLine arquivo
	agc.Close
	
	

End Function

' Confere o meridiano
' Valores convertidos para horas

Function confereMeridiano(ra, sequencia, tentativas, exposicao, exposicaoCaptura)

	Dim exposicoesRestantes
	Dim margemSeguranca
	Dim exposicaoTotal
	Dim meridiano
	Dim meridianoVirtualLeste
	Dim meridianoVirtualOeste
	Dim intervalo
	Dim raFinal
		
	margemSeguranca = 1.1 ' 10%
	intervalo = trocaMeridiano / 7200
	
	If (sequencia = 1) Then
	
		exposicoesRestantes = tentativas

	Else

		exposicoesRestantes = tentativas - sequencia

	End If
	
	exposicaoTotal = (((exposicoesRestantes * exposicao) + exposicaoCaptura) * margemSeguranca) / 3600
	raFinal = ra - exposicaoTotal
	
	If (raFinal < 0) Then
	
		raFinal = raFinal + 23.999999
	
	End If	
	
	meridiano = telescopio.SiderealTime
	meridianoVirtualLeste = calcMedirianoVirtual (meridiano, intervalo)
	meridianoVirtualOeste = calcMedirianoVirtual (ra, intervalo)
exibeTela "ra -> " & FormatNumber (ra, 4, , vbTrue)
exibeTela "intervalo -> " & FormatNumber (intervalo, 4, , vbTrue)
exibeTela "meridiano -> " & FormatNumber (meridiano, 4, , vbTrue)
exibeTela "raFinal -> " & FormatNumber (raFinal, 4, , vbTrue)
exibeTela "meridianoVirtualLeste -> " & FormatNumber (meridianoVirtualLeste, 4, , vbTrue)
exibeTela " "
	If (raFinal < meridianoVirtualLeste And  (ra - meridiano) > 0) Then

		exibeTela "Aguadando a passagem do meridiano..."
		exibeTela " "
		meridianoVirtualOeste = calcMedirianoVirtual (ra, intervalo)
		
		Do 

			WScript.Sleep 10000
			meridiano = telescopio.SiderealTime
exibeTela "meridiano -> " & FormatNumber (meridiano, 4, , vbTrue)
exibeTela "meridianoVirtualOeste -> " & FormatNumber (meridianoVirtualOeste, 4, , vbTrue)
' exibeTela "meridianoVirtual - meridiano -> " & FormatNumber (meridianoVirtual - meridiano, 4, , vbTrue)
exibeTela " "		
		Loop Until (meridianoVirtualOeste < meridiano)
		
		confereMeridiano = 1

	Else
	
		confereMeridiano = 0
	
	End If
	

End Function

' Calcula o meridiano virtual

Function calcMedirianoVirtual (calcMeridiano, calcIntervalo)

	Dim meridianoTemp

	meridianoTemp = calcMeridiano + calcIntervalo
	
	If (meridianoTemp > 23.999999) Then
	
		meridianoTemp = 23.999999 - meridianoTemp 
	
	End If
	
	calcMedirianoVirtual = meridianoTemp
	
End Function

' Limpa a memoria

Function mataObjetos ()

	Set telescopioMaxim = Nothing
	Set camera = Nothing
	Set fso = Nothing
	Set doc = Nothing
	Set app = Nothing
	Set conv = Nothing
	Set novas = Nothing

End Function

