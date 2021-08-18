
		telescopio.SlewToCoordinates ra, dec
		
		Do While telescopio.Slewing

			WScript.sleep  1

		Loop

		WScript.sleep  1500
		
		If (cameraAlvo = 1) Then
		
			camera.Expose expApontamento, 1, regFiltro

			Do While Not camera.ImageReady

				WScript.sleep  1

			Loop
			
			camera.SetFITSKey "VARIAVEL", regVariavel
			camera.SetFITSKey "VRA", regRa
			camera.SetFITSKey "VDEC", regDec
			arquivo = diretorioTemp + regVariavel + complementoSolve + CStr(zeroEsquerda(contAlvo)) + ".fits"
			'camera.SaveImage arquivo
			
		Else
		
			camera.GuiderExpose expApontamento
	
			Do While camera.GuiderRunning

				WScript.sleep  1

			Loop
		
			app.CurrentDocument.SetFITSKey "VARIAVEL", regVariavel
			app.CurrentDocument.SetFITSKey "VRA", regRa
			app.CurrentDocument.SetFITSKey "VDEC", regDec		
			arquivo = diretorioTemp + regVariavel + complementoSolve + CStr(zeroEsquerda(contAlvo)) + ".fits"
			'app.CurrentDocument.SaveFile arquivo , 3, True, 0, 1
	
		End If
		
		doc.OpenFile arquivo


















		
			If (Abs (plate.RightAscension  - ra) < toleranciaSolve And Abs (plate.Declination - dec) <  toleranciaSolve) Then
			
				plate.DetachFITS
				resultadoSolve = 2
				exibeTela "Aqui 002"
				Exit For
				
			Else
			
				exibeTela "Aqui 003"
				linha = "RA: " + conv.HoursToHMS (plate.RightAscension, ":", ":", " ",2) + "    " + "DEC: " + conv.DegreesToDMS (plate.Declination, "º", "´", "´´",2)
				exibeTela linha
				plate.DetachFITS
				resultadoSolve = 1
				
			End If
			
			
			
