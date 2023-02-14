Dim WsShell
Set WsShell = CreateObject("WScript.shell")
Dim SessionNumber  
Dim dtewait
Dim emailObj

strComputer = "."

Set objNetwork = CreateObject("Wscript.Network")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 ''' Processo que será verificado '''''''
Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'wscript.exe'")
cdr = 0
 ''' elimina o processo definido '''
 For each Processo in ColProcesses
	cdr = cdr + 1
	If cdr > 1 then
		result = MsgBox ("O programa ja esta rodando, deseja fechar?", vbYesNo, "Yes No Example")
		Select Case result
		Case vbYes
			For each ProcessoWS in ColProcesses
				ProcessoWs.Terminate()
			Next
		Case vbNo
			WScript.Quit
		End Select
	end if
Next

result = MsgBox ("Desligar Computador Apos o Termino?", vbYesNo, "Yes No Example")

Select Case result
Case vbYes
    DesligaPC = 1
Case vbNo
    DesligaPC = 0
End Select

dtewait = DateAdd("M",2,""&Day(Date)&"/"&Month(Date)&"/"&Year(Date)+1&" 17:00:00")
Data = Date + 1
minum = minute(Now) 
do Until (Now() > dtewait)
	if Date = Data then 
	'if minute(Now) = minum then
		If Not IsObject(application) Then
		   
			on error resume next
			Set SapGuiAuto  = GetObject("SAPGUI")
				if Err.Number <> 0 then
					WsShell.Run "https://www.myweg.net/irj/portal?NavigationTarget=pcd:portal_content/net.weg.folder.weg/net.weg.folder.core/net.weg.folder.roles/net.weg.role.ecc/net.weg.iview.ecc"
					WScript.Sleep 10000
					FreakoutAndFixTheError
					err.clear
					WScript.Sleep 10000
					If Not IsObject(application) Then
						Set SapGuiAuto  = GetObject("SAPGUI")
						Set application = SapGuiAuto.GetScriptingEngine
					End If
					
					If Not IsObject(connection) Then
					
						on error resume next
						Set connection = application.Children(0)
						
						if Err.Number <> 0 then
						msgbox("Favor iniciar o SAP antes de execultar.")
						Wscript.quit
						end if
					
					end if
				end if
			Set application = SapGuiAuto.GetScriptingEngine
		End If
		If Not IsObject(connection) Then
			
			on error resume next
			Set connection = application.Children(0)

			if Err.Number <> 0 then
				WsShell.Run "https://www.myweg.net/irj/portal?NavigationTarget=pcd:portal_content/net.weg.folder.weg/net.weg.folder.core/net.weg.folder.roles/net.weg.role.ecc/net.weg.iview.ecc"
				WScript.Sleep 10000
				FreakoutAndFixTheError
				err.clear
				WScript.Sleep 10000
				If Not IsObject(application) Then
					Set SapGuiAuto  = GetObject("SAPGUI")
					Set application = SapGuiAuto.GetScriptingEngine
				End If
				
				If Not IsObject(connection) Then
				
					on error resume next
					Set connection = application.Children(0)
					
					if Err.Number <> 0 then
					msgbox("Favor iniciar o SAP antes de execultar.")
					Wscript.quit
					end if
				
				end if
			end if
			
		End If

		If Not IsObject(session) Then
		   Set session    = connection.Children(0)
		End If
		If IsObject(WScript) Then
		   WScript.ConnectObject session,     "on"
		   WScript.ConnectObject application, "on"
		End If

		do While connection.Children.Count > 5
			Response = MsgBox("Favor fechar uma Janela do SAP para Continuar!", vbRetryCancel, "Aviso")
			If Response = vbCancel then
				WScript.Quit
			end if
		loop

			on error resume next
			session.findById("wnd[0]").sendVKey 0
			FreakoutAndFixTheError
			err.clear
				
				
		Dim FileSys
		Dim FileShell
		Dim objFolder
		Dim colFiles
		Dim objFile
		Dim nulin
		Dim OPJG
		Dim Ordem
		Dim codpro
		Dim Forne
		Dim numet
		Dim DataO
		Dim locall
		Dim line
		Dim linea
		Dim str
		Dim status
		Dim Nesce
		Dim agru
		Dim numf
		Dim DocID
		Dim Doc
		Dim LineOV
		Dim ItemID
		Dim Linetest
		Dim OPnum

		OPJG = ""
		nulin = ""
		Ordem = ""
		codpro = ""
		Forne = ""
		numet = ""
		DataO = ""
		locall = ""
		line = ""
		linea = ""
		str = ""
		status = ""
		Status2 = ""
		Nesce = ""
		agru = ""
		numf = ""
		OPexit = 0
		OPnum = 0
		DocID = 0
		Doc = 0
		LineOV = 0
		ItemID = 0
		Linetest = 0
		QtdOrd = 0
		QtdOrd2 = 0

		Set WsShell = CreateObject("WScript.Shell")
		UserName = WsShell.ExpandEnvironmentStrings("%UserName%")
		Set FileSys = CreateObject("Scripting.FileSystemObject")
		Set FileShell = CreateObject("Shell.Application")

		Data = Day(Date) & "." & Month(Date) & "." & Year(Date)
		Path = "C:\Users\" + UserName + "\Desktop\Faturamento\"


		'nVar = InputBox(" Bem Vindo Sr(a). " & UserName  _
		'& vbNewLine & " __________________________________________ "_
		'& vbNewLine & " ( 01 ) - Fazer Extração e Gerar Relatorio " _
		'& vbNewLine & " ( 02 ) - Apenas Gerar Relatorio" _
		'& vbNewLine & " ( 03 ) - Continuar Relatorio" _
		'& vbNewLine & " __________________________________________ "_
		'& vbNewLine & " ", "Execultaveis", "01")

		ender = Path + "Faturamento "& Data &".xls"
		If FileSys.FileExists(ender) Then
		 nVar = "02"
		else
		 nVar = "01"
		End If

		 'session.findbyid("wnd[0]").close
		 'session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
		 
		If nVar <> "01" and nVar <> "1" and nVar <> "02" and nVar <> "2" Then
		else
				If nVar = "1" or nVar = "01" Then
					session.createSession
					Wscript.Sleep 1000
					SessionNumber = connection.Children.Count-1
					Wscript.Sleep 1000
					Set session = connection.sessions.Item(Cint(SessionNumber))
					Wscript.Sleep 100

					session.findById("wnd[0]").maximize
					session.findById("wnd[0]/tbar[0]/okcd").text = "ZTSD024"
					session.findById("wnd[0]").sendVKey 0
					session.findById("wnd[0]/tbar[1]/btn[17]").press
					session.findById("wnd[1]/usr/txtENAME-LOW").text = "claudioj"
					session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
					session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
					session.findById("wnd[1]").sendVKey 0
					session.findById("wnd[1]/tbar[0]/btn[8]").press
					session.findById("wnd[0]/usr/btn%_S_VSTEL_%_APP_%-VALU_PUSH").press
					session.findById("wnd[1]/tbar[0]/btn[16]").press
					session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "BA28"
					session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "BB41"
					session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
					session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
					session.findById("wnd[1]/tbar[0]/btn[8]").press
					session.findById("wnd[0]/tbar[1]/btn[8]").press
					session.findById("wnd[0]/tbar[1]/btn[45]").press
					session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
					session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
					session.findById("wnd[1]/tbar[0]/btn[0]").press
					session.findById("wnd[1]/usr/ctxtDY_PATH").text = Path
					session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Faturamento "& Data &".xls"
					session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 26
					session.findById("wnd[1]/tbar[0]/btn[0]").press
					session.findById("wnd[0]/tbar[0]/btn[15]").press
					session.findById("wnd[0]/tbar[0]/btn[15]").press
				end if

				endereco = Path + "Faturamento "& Data &".xls"

				'if nVar = "3" or nVar = "03" then
				'	endereco = Path + "Faturamento "& Data &"old.xls"
				'end if

				Set obj = CreateObject("Excel.Application")
				Set objExcel = obj.Workbooks.Open(endereco)
				objExcel.Application.Visible = True
				objExcel.Sheets("Faturamento "& Data).Activate


				if nVar = "1" or nVar = "01" or nVar = "2" or nVar = "02" then
					objExcel.Sheets("Faturamento "& Data).Range("1:7").Delete
					objExcel.Sheets("Faturamento "& Data).Range("2:2").Delete
					objExcel.Sheets("Faturamento "& Data).Columns("A:B").Delete
					objExcel.Sheets("Faturamento "& Data).Columns("K:P").Delete
					objExcel.Sheets("Faturamento "& Data).Columns("F:I").NumberFormat = "0,"

					objExcel.Sheets("Faturamento "& Data).Range("K1") = "Ordems"
					objExcel.Sheets("Faturamento "& Data).Range("L1") = "Status Local"
					objExcel.Sheets("Faturamento "& Data).Range("M1") = "Tratativas"
					objExcel.Sheets("Faturamento "& Data).Range("C:C").copy
					objExcel.Sheets("Faturamento "& Data).Range("N:N").PasteSpecial -4163

					ulin = objExcel.Sheets("Faturamento "& Data).Range("A999999").End(-4162).Row

					objExcel.Sheets("Faturamento "& Data).Columns("A:A").ColumnWidth = 20
					objExcel.Sheets("Faturamento "& Data).Columns("B:B").ColumnWidth = 11
					objExcel.Sheets("Faturamento "& Data).Columns("E:E").ColumnWidth = 20

					objExcel.Sheets("Faturamento "& Data).Range("P2").FormulaR1C1 = "=VALUE(RC[-10])"
					objExcel.Sheets("Faturamento "& Data).Range("P2").Copy
					objExcel.Sheets("Faturamento "& Data).Range("P2:P"& ulin).PasteSpecial -4123
					objExcel.Sheets("Faturamento "& Data).Range("P2:P"& ulin).Copy
					objExcel.Sheets("Faturamento "& Data).Range("F2:F"& ulin).PasteSpecial -4163
					objExcel.Sheets("Faturamento "& Data).Range("Q2").FormulaR1C1 = "=VALUE(RC[-10])"
					objExcel.Sheets("Faturamento "& Data).Range("Q2").Copy
					objExcel.Sheets("Faturamento "& Data).Range("Q2:Q"& ulin).PasteSpecial -4123
					objExcel.Sheets("Faturamento "& Data).Range("Q2:Q"& ulin).Copy
					objExcel.Sheets("Faturamento "& Data).Range("G2:G"& ulin).PasteSpecial -4163
					objExcel.Sheets("Faturamento "& Data).Range("R2").FormulaR1C1 = "=VALUE(RC[-10])"
					objExcel.Sheets("Faturamento "& Data).Range("R2").Copy
					objExcel.Sheets("Faturamento "& Data).Range("R2:R"& ulin).PasteSpecial -4123
					objExcel.Sheets("Faturamento "& Data).Range("R2:R"& ulin).Copy
					objExcel.Sheets("Faturamento "& Data).Range("H2:H"& ulin).PasteSpecial -4163
					objExcel.Sheets("Faturamento "& Data).Range("S2").FormulaR1C1 = "=VALUE(RC[-10])"
					objExcel.Sheets("Faturamento "& Data).Range("S2").Copy
					objExcel.Sheets("Faturamento "& Data).Range("S2:S"& ulin).PasteSpecial -4123
					objExcel.Sheets("Faturamento "& Data).Range("S2:S"& ulin).Copy
					objExcel.Sheets("Faturamento "& Data).Range("I2:I"& ulin).PasteSpecial -4163
					objExcel.Sheets("Faturamento "& Data).Columns("I:I").Style = "Currency"
					objExcel.Sheets("Faturamento "& Data).Columns("P:S").ClearContents

					objExcel.Sheets("Faturamento "& Data).Columns("F:H").HorizontalAlignment = 3
					objExcel.Sheets("Faturamento "& Data).Columns("F:H").VerticalAlignment = 3

					objExcel.Sheets("Faturamento "& Data).Range("P1") = "CNPA"
					objExcel.Sheets("Faturamento "& Data).Columns("P:P").HorizontalAlignment = 3

					objExcel.Sheets("Faturamento "& Data).Columns("F:H").ColumnWidth = 6
					objExcel.Sheets("Faturamento "& Data).Columns("K:K").ColumnWidth = 11
					objExcel.Sheets("Faturamento "& Data).Columns("L:L").ColumnWidth = 26
					objExcel.Sheets("Faturamento "& Data).Columns("M:M").ColumnWidth = 20
					objExcel.Sheets("Faturamento "& Data).Columns("N:N").ColumnWidth = 6
					objExcel.Sheets("Faturamento "& Data).Range("P1").ColumnWidth = 5
					objExcel.Sheets("Faturamento "& Data).Columns("J").Replace ".", "/"
					objExcel.Sheets("Faturamento "& Data).Columns("J").NumberFormat = "dd/mm/yyyy"
					objExcel.Sheets("Faturamento "& Data).Range("J:J").copy
					objExcel.Sheets("Faturamento "& Data).Range("O:O").PasteSpecial -4163
					objExcel.Sheets("Faturamento "& Data).Columns("O").NumberFormat = "dd/mm/yyyy"
				end if

					'Mid(Trim(Forne), 1, 20) = "61379-THOR MAQUINAS E MONTAGENS LTDA"
					'Mid(Trim(Forne), 1, 20) = "509167-GRATT INDUSTRIA DE MAQUINAS LTDA"
					'Mid(Trim(Forne), 1, 20) = "498700-LINDSAY AMERICA DO SUL LTDA"
					'Mid(Trim(Forne), 1, 20) = "16543-MEBRAFE INST.EQUIP.FRIG.LTDA"
					'Mid(Trim(Forne), 1, 20) = "4359-A CARNEVALLI & CIA LTDA"
					'Mid(Trim(Forne), 1, 20) = "785316-HEDEL MAQS. E EQUIPS. LTDA"
					'Mid(Trim(Forne), 1, 20) = "842543-EXTRUSION SYSTEM INDUSTRIA E COMERC"
					'Mid(Trim(Forne), 1, 20) = "17833-RULLI STANDARD IND.COM.DE MAQS.LTDA"
					'Mid(Trim(Forne), 1, 20) = "344907-WORTEX MAQUINAS E EQUIPAMENTOS LTDA"


				'if nVar = "3" or nVar = "03" then
				'nulin = objExcel.Sheets("Faturamento "& Data).Range("K999999").End(-4162).Row - 1
				'else
				nulin = 2	
				'end if

				If nVar <> "1" or nVar <> "01" Then
					session.createSession
					Wscript.Sleep 1000
					SessionNumber = connection.Children.Count-1
					Wscript.Sleep 1000
					Set session = connection.sessions.Item(Cint(SessionNumber))
				end if

				Do Until objExcel.Sheets("Faturamento "& Data).Range("A"& nulin) = ""
					'if nulin = 200 or nulin = 400 or nulin = 600 or nulin = 800 or nulin = 1000 or nulin = 1200 or nulin = 1400 or nulin = 1600 or nulin = 1800 or nulin = 2000 or nulin = 2200 or nulin = 2400 or nulin = 2600 or nulin = 2800 or nulin = 3000 then
					if nulin = 1000 then
					
						strComputer = "."
						Set objNetwork = CreateObject("Wscript.Network")
						Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
						Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'saplgpad.exe'")
						 ''' elimina o processo definido '''
						For each Processo in ColProcesses
							Processo.Terminate()
						Next
						
						Set objNetwork = CreateObject("Wscript.Network")
						Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
						Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'iexplore.exe'")
						 ''' elimina o processo definido '''
						For each Processo in ColProcesses
							Processo.Terminate()
						Next
						
						WsShell.Run "https://www.myweg.net/irj/portal?NavigationTarget=pcd:portal_content/net.weg.folder.weg/net.weg.folder.core/net.weg.folder.roles/net.weg.role.ecc/net.weg.iview.ecc"
						WScript.Sleep 5000

						Set SapGuiAuto  = GetObject("SAPGUI")
						Set application = SapGuiAuto.GetScriptingEngine
						Set connection = application.Children(0)
						Set session    = connection.Children(0)
							
						If IsObject(WScript) Then
						   WScript.ConnectObject session,     "on"
						   WScript.ConnectObject application, "on"
						End If
						
						session.createSession
						Wscript.Sleep 1000
						SessionNumber = connection.Children.Count-1
						Wscript.Sleep 1000
						Set session = connection.sessions.Item(Cint(SessionNumber))
					end if


					SessionNumber = connection.Children.Count-1
					Set session = Nothing
					Set session    = connection.sessions.Item(Cint(SessionNumber))
					
					OPJG = 0
					
					Forne = objExcel.Sheets("Faturamento "& Data).Range("A"& nulin)
					Doc = objExcel.Sheets("Faturamento "& Data).Range("B"& nulin)
					QtdeDis = objExcel.Sheets("Faturamento "& Data).Range("F"& nulin)
					QtdeSol = objExcel.Sheets("Faturamento "& Data).Range("G"& nulin)
					EstoqDisp = objExcel.Sheets("Faturamento "& Data).Range("H"& nulin)
					LineOV = objExcel.Sheets("Faturamento "& Data).Range("N"& nulin)
					ItemID = objExcel.Sheets("Faturamento "& Data).Range("D"& nulin)
					
					If Mid(Trim(Forne), 1, 10) = "61379-THOR" or Mid(Trim(Forne), 1, 12) = "509167-GRATT" or _
					Mid(Trim(Forne), 1, 14) = "498700-LINDSAY" or Mid(Trim(Forne), 1, 13) = "16543-MEBRAFE" or _
					Mid(Trim(Forne), 1, 17) = "4359-A CARNEVALLI" or Mid(Trim(Forne), 1, 12) = "785316-HEDEL" or _
					Mid(Trim(Forne), 1, 16) = "842543-EXTRUSION" or Mid(Trim(Forne), 1, 11) = "17833-RULLI" or _
					Mid(Trim(Forne), 1, 13) = "344907-WORTEX" or Mid(Trim(Forne), 1, 14) = "803687-LINDNER" or _
					Mid(Trim(Forne), 1, 10) = "796935-WEG" then
						objExcel.Sheets("Faturamento "& Data).Range("A"& nulin&":Q"& nulin).Interior.Color = 10092492 '65535 RGB(6,55,35) .ColorIndex = 36
					end if	
					
						if QtdeDis = QtdeSol and QtdeSol <= EstoqDisp then
						objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
						objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
						objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "-----------"
						else
							'msgBOX("iNICIO")
							'session.findById("wnd[0]").maximize
							session.findById("wnd[0]/tbar[0]/okcd").text = "MD04"
							session.findById("wnd[0]").sendVKey 0

							'Entrar seção
							session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").text = ItemID
							session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = "1321"
							session.findById("wnd[0]").sendVKey 0
							
							'Fim Entrar seção
					'=====================================  VERIFICAÇÃO 1320 ORDEMS =======================================================================
							If session.findById("wnd[0]/sbar").messagetype = "E" then
							
								session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = "1320"
								session.findById("wnd[0]").sendVKey 0
								If session.findById("wnd[0]/sbar").messagetype = "E" then
									session.findById("wnd[0]").sendVKey 0
									objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Nao tem1"
								end if
								
								linea = 1
								tt = ""
								Do until session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&linea&"]").text = ""
								'msgbox("Tipo "&session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&linea&"]").text &" 1320")
									session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& linea &"]").setFocus
									
									tt = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&linea&"]").text
									if Mid(Trim(tt), 1, 1) = "_" then
										exit do
									end if
									
									if linea > 20 then
										session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ").verticalScrollbar.position = linea - 15
									end if
									DataOR = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& linea &"]").text
									Nesce = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG01[4,"& linea &"]").text
									Ordem =	session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[3,"& linea &"]").text
									if session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&linea&"]").text <> "EstDep" then
										if Mid(Trim(Ordem), 11, 1) = "/" then
											objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = Mid(Trim(Ordem), 1, 10)
										else
											objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = Mid(Trim(Ordem), 1, 12)
										end if
									end if
									objExcel.Sheets("Faturamento "& Data).Range("N1") = DataOR
									objExcel.Sheets("Faturamento "& Data).Range("N1").Replace ".", "/"
									objExcel.Sheets("Faturamento "& Data).Range("N1").NumberFormat = "dd/mm/yyyy"
									
									'msgbox("DataOR "& DataOR &" 1321 | "& objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) &" < "& date + 5 )
									If Mid(Trim(DataOR), 1, 1) <> "_" and session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&linea&"]").text <> "EstDep" then ' and objExcel.Sheets("Faturamento "& Data).Range("N1") < date + temp then

										if tt = "OrdPro" or tt = "OrdPla" then
											'msgbox("Nesc "& Nesce &" 1320")
											if Nesce > 0 then
												if tt = "OrdPla" then
													objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
													objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321a - Ordem Planejada"
													OPJG = 1
												else
													session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[3,"&linea&"]").setFocus
													session.findById("wnd[0]").sendVKey 2
													session.findById("wnd[1]/tbar[0]/btn[7]").press
													'"S" Sucess "W" Warning "E" Error "A" Abort "I" Information
													If session.findById("wnd[0]/sbar").messagetype = "W" then
														session.findById("wnd[0]").sendVKey 0
													end if

													on error resume next
														session.findById("wnd[0]/tbar[1]/btn[6]").press
														
														if Err.Number <> 0 then
															session.findById("wnd[0]/tbar[0]/btn[3]").press
															session.findById("wnd[0]").sendVKey 2
															session.findById("wnd[1]/tbar[0]/btn[7]").press
															session.findById("wnd[0]/tbar[1]/btn[6]").press
															
															On error resume next
															session.findById("wnd[0]").sendVKey 0
															FreakoutAndFixTheError
															err.clear
														end if
													
													session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1,0]").setFocus
													ttp = 0
													do until ttp = 10
														agru = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MATXT[2,"& ttp &"]").text
														if Mid(Trim(agru), 1, 9) = "AGRUPADOR" then
															codpro = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1,"& ttp &"]").text
															exit do
														end if
														ttp = ttp + 1
													loop
													'descitemagru = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MATXT[2,0]").text

													session.findById("wnd[0]/tbar[0]/btn[3]").press
													session.findById("wnd[0]/tbar[0]/btn[3]").press

													if codpro = "" then 
														OPJG = 1
														objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
														objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
														objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321a - Ordem Planejada"
													else
														session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-MATNR").text = codpro
														objExcel.Sheets("Faturamento "& Data).Range("C"& nulin) = codpro
														session.findById("wnd[0]").sendVKey 0

														session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-BERID").text = "1321"
														session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-WERKS").text = "1321"
														session.findById("wnd[0]").sendVKey 0
														OPJG = 0
													end if
													exit do
												end if
											else
											
												if Mid(Trim(Nesce), 1, 1) = "_" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Concluido"
													OPJG = 1
													exit do
												end if

												session.findById("wnd[0]").sendVKey 37
												session.findById("wnd[0]/usr/cntlGRID_1000/shellcont/shell/shellcont[1]/shell")
												set GRID = session.findById("wnd[0]/usr/cntlGRID_1000/shellcont/shell/shellcont[1]/shell")
												GRID.setCurrentCell 0,""
												
												
												ttov = 0
												do until ttov = -1
													
													DocTest = GRID.getcellvalue (ttov, "EXTRA")
													
													if Mid(Trim(DocTest), 15, 1) = "0" then
														Linetest = Mid(Trim(DocTest), 16, 2)
													else
														Linetest = Mid(Trim(DocTest), 15, 3)
													end if
													
													if Mid(Trim(DocTest), 2, 1) = "0" then
														DocID = Mid(Trim(DocTest), 3, 8)
													else
														DocID = Mid(Trim(DocTest), 2, 9)
													end if
													'msgbox(Doc &"="& DocID &"|"& LineOV &"="& Linetest)
													if CStr(Doc) = DocID and CStr(LineOV) = Linetest then
														if OPnum > 0 then
															OPnum = OPnum + 1
														else
															OPnum = 1
														end if
														OPexit = 1
														exit do
													end if
													
													On error resume next
													DatteTest = GRID.getcellvalue (ttov + 1, "DAT00")
													
													if Err.Number <> 0 then
														FreakoutAndFixTheError
														err.clear
														exit do
													end if
													
													ttov = ttov + 1	
												loop
												'msgbox(OPnum)
												session.findById("wnd[0]/tbar[0]/btn[3]").press
												
												if tt = "OrdPla" then
													objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
													objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Ordem Planejada"
													OPJG = 1
												'else
												'	objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Concluido"
												'	OPJG = 1
												end if
											
												
											end if

										else
											objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
											if session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "Fornec" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Fornecido"
												OPJG = 1	
											elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "OrdPla" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Liberar Ordem"
												OPJG = 1
											elseif Nesce < 0 then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Fazer Picking"
												OPJG = 1
											elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "EstDep" then
												OPJG = 1
											elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "SolCnt" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Transferir"
												OPJG = 1
											elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "ReqCmp" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Ordem Compra"
												OPJG = 1
											end if
										end if
										
										
									else
										objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
										if session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "Fornec" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Fornecido.1"
											OPJG = 1	
										elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "OrdPla" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Liberar Ordem.1"
											OPJG = 1
										elseif Nesce < 0 then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Fazer Picking.1"
											OPJG = 1
										elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "EstDep" then
											OPJG = 1
										elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "SolCnt" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - SOL. CNT.1"
											OPJG = 1
										elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "ReqCmp" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Ordem Compra.1"
											OPJG = 1
										end if
									end if	
									
									linea = linea + 1
									
								loop
								'msgBOX("fim")
							end if
					'===================================== FIM VERIFICAÇÃO 1320 ORDEMS =======================================================================
							if OPJG <> 1 then
								line = 1
								
									'status = 0
									'Do until session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text = "______"
									'msgbox("Tipo "&session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text &" 1321")
									

								session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& line &"]").setFocus
								DataOR = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& line &"]").text

								'msgbox("Data "& DataOR &" 1321")
								If Mid(Trim(DataOR), 1, 1) <> "_" and session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text <> "EstDep" then

									status = 0
									tpo = ""
									Do until session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text = ""
									'msgbox("Tipo "&session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text &" 1321")
										session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& line &"]").setFocus
										tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text
										if Mid(Trim(tpo), 1, 1) = "_" or line > 100 then
											if line > 100 then
												'objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "Verificar"
											end if
											exit do
										end if
										
										if line > 20 then
											session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ").verticalScrollbar.position = line - 15
										end if
										DataOR = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& line &"]").text
										Disp = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG02[5,"& line &"]").text
										Ordem =	session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[3,"& line &"]").text
										
										if session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text <> "EstDep" then
											if Mid(Trim(Ordem), 11, 1) = "/" then
												Ordem2 = Mid(Trim(Ordem), 1, 10)
											else
												Ordem2 = Mid(Trim(Ordem), 1, 12)
											end if
										end if
										
										Nesce = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG01[4,"& line &"]").text
										
										objExcel.Sheets("Faturamento "& Data).Range("N1") = DataOR
										objExcel.Sheets("Faturamento "& Data).Range("N1").Replace ".", "/"
										objExcel.Sheets("Faturamento "& Data).Range("N1").NumberFormat = "dd/mm/yyyy"
										
										'msgbox("DataOR "& DataOR &" 1321 | "& objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) &" < "& date + 5 )
										if objExcel.Sheets("Faturamento "& Data).Range("N1") > objExcel.Sheets("Faturamento "& Data).Range("O"& nulin) + 20 then'date + temp then 'objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) then
											'objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Data OP e Maior que data OV"
											exit do
										end if
										
										if tpo = "OrdPro" or tpo = "OrdPla" then
											'msgbox("Nesc "&Nesce &" 1321")
											if Nesce > 0 then

				'=========================================Definição de quantidade por Ordem =============================================
												if Mid(Trim(Nesce), 1, 1) = "_" then
													exit do
												end if

												session.findById("wnd[0]").sendVKey 37
												session.findById("wnd[0]/usr/cntlGRID_1000/shellcont/shell/shellcont[1]/shell")
												set GRID = session.findById("wnd[0]/usr/cntlGRID_1000/shellcont/shell/shellcont[1]/shell")
												GRID.setCurrentCell 0,""
												
												
												ttov = 0
												do until ttov = -1
													
													DocTest = GRID.getcellvalue (ttov, "EXTRA")
													
													if Mid(Trim(DocTest), 15, 1) = "0" then
														Linetest = Mid(Trim(DocTest), 16, 2)
													else
														Linetest = Mid(Trim(DocTest), 15, 3)
													end if
													
													if Mid(Trim(DocTest), 2, 1) = "0" then
														DocID = Mid(Trim(DocTest), 3, 8)
													else
														DocID = Mid(Trim(DocTest), 2, 9)
													end if
													'msgbox(Doc &"="& DocID &"|"& LineOV &"="& Linetest)
													if CStr(Doc) = DocID and CStr(LineOV) = Linetest then
														if OPnum > 0 then
															OPnum = OPnum + 1
														else
															OPnum = 1
														end if
														OPexit = 1
														exit do
													end if
													
													On error resume next
													DatteTest = GRID.getcellvalue (ttov + 1, "DAT00")
													
													if Err.Number <> 0 then
														FreakoutAndFixTheError
														err.clear
														exit do
													end if
													
													ttov = ttov + 1	
												loop
												'msgbox(OPnum)
												session.findById("wnd[0]/tbar[0]/btn[3]").press
												
												if tpo = "OrdPla" and OPexit > 0 then
													objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
													objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Ordem Planejada"
													OPexit = 0
												'else
													'objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
													'objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
													'objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Liberar Ordem"
												end if
												
												if OPexit > 0 then
													if tpo = "OrdPla" then
													else
													OPexit = 0
													ttov = 0
													'msgbox(OPnum)
													
													if OPnum > 1 then
														objExcel.Sheets("Faturamento "& Data).Range("A"& nulin &":K"& nulin).copy
														'msgbox("NEW LINE")
														nulin = nulin + 1
														objExcel.Sheets("Faturamento "& Data).Range("A"& nulin &":K"& nulin).Insert -4121
														objExcel.Sheets("Faturamento "& Data).Range("N"& nulin &":O"& nulin).Insert -4121
														objExcel.Sheets("Faturamento "& Data).Range("A"& nulin &":K"& nulin).PasteSpecial -4163
														objExcel.Sheets("Faturamento "& Data).Range("N"& nulin - 1 &":O"& nulin - 1 ).copy
														objExcel.Sheets("Faturamento "& Data).Range("N"& nulin &":O"& nulin).PasteSpecial -4163
														objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("O"& nulin)
														
													end if
													
				'=========================================Definição de quantidade por Ordem =============================================

														session.findById("wnd[0]").sendVKey 2
														session.findById("wnd[1]/tbar[0]/btn[8]").press
														
														if session.Info.Program = "SAPLCMFE" then
															session.findById("wnd[0]/tbar[0]/btn[3]").press
															session.findById("wnd[0]").sendVKey 2
															session.findById("wnd[1]/tbar[0]/btn[7]").press
															
														end if
																
														'"S" Sucess "W" Warning "E" Error "A" Abort "I" Information
														If session.findById("wnd[0]/sbar").messagetype = "W" then
															session.findById("wnd[0]").sendVKey 0
														end if

														On error resume next
														session.findById("wnd[1]/tbar[0]/btn[0]").press
														FreakoutAndFixTheError
														err.clear

															
														On error resume next
														session.findById("wnd[0]/tbar[1]/btn[28]").press
														FreakoutAndFixTheError
														err.clear
														desc = ""
														
														'"S" Sucess "W" Warning "E" Error "A" Abort "I" Information
														If session.findById("wnd[0]/sbar").messagetype = "S" then
															session.findById("wnd[0]/tbar[1]/btn[5]").press
														else
															session.findById("wnd[1]/usr/btnDY_VAROPTION3").press
															session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/ctxtRESBD-MATNR[1,0]").setFocus
															codfalt = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/ctxtRESBD-MATNR[1,0]").text
															numf = 0
															
															do until session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/ctxtRESBD-MATNR[1,"&numf&"]").text = ""
															codfalt = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/ctxtRESBD-MATNR[1,"&numf&"]").text
															descfalt = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/txtRESBD-MATXT[8,"&numf&"]").text
															locall = codfalt &" - "& descfalt
															if Mid(Trim(locall), 12, 9) <> "EMBALAGEM" then
																desc = desc & vbNewLine & Mid(Trim(locall), 1, 20)
															end if
															if numf >= 5 then
																exit do
															end if
																numf = numf + 1
															loop
															
															session.findById("wnd[0]/tbar[0]/btn[12]").press
															session.findById("wnd[0]/tbar[1]/btn[5]").press
															
														end if
														'msgbox(desc)
														if desc = "" then
															objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "----"
														else
															objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = desc
														end if

														If session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[8,1]").text = "BLOQ LIB" _
														or session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[8,1]").text = "CNPA BLOQ" or session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[8,1]").text = "CONF BLOQ LIB" then

															objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
															objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "Jaragua Block"


															Forne = objExcel.Sheets("Faturamento "& Data).Range("A"& nulin)
															
															If Mid(Trim(Forne), 1, 10) = "61379-THOR" or Mid(Trim(Forne), 1, 12) = "509167-GRATT" or _
															Mid(Trim(Forne), 1, 14) = "498700-LINDSAY" or Mid(Trim(Forne), 1, 13) = "16543-MEBRAFE" or _
															Mid(Trim(Forne), 1, 17) = "4359-A CARNEVALLI" or Mid(Trim(Forne), 1, 12) = "785316-HEDEL" or _
															Mid(Trim(Forne), 1, 16) = "842543-EXTRUSION" or Mid(Trim(Forne), 1, 11) = "17833-RULLI" or _
															Mid(Trim(Forne), 1, 13) = "344907-WORTEX" or Mid(Trim(Forne), 1, 14) = "803687-LINDNER"  or _
															Mid(Trim(Forne), 1, 10) = "796935-WEG" then
																linea = 0
																do until session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").text = ""
																	session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").setFocus
																	locall = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[7,"& linea &"]").text
																	DataO = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").text
																	If Mid(Trim(locall), 1, 7) = "EMBALAR" or Mid(Trim(locall), 1, 10) = "FECHAMENTO" then
																		objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = DataO
																		objExcel.Sheets("Faturamento "& Data).Range("J"& nulin).Replace ".", "/"
																		exit do
																	end if
																	linea = linea + 1
																loop
															else 
																if Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 1 then
																	objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 3
																elseif Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 7 then
																	objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 2
																else 
																	objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1
																end if
															end if


																session.findById("wnd[0]/tbar[0]/btn[15]").press
																On error resume next
																session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
																FreakoutAndFixTheError
																err.clear

																ext = 0
																status = 1
															exit do

														else

															linea = 0
															session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,1]").setFocus
															str = 0

															do until session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").text = ""
																session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").setFocus
																Status = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[8,"& linea &"]").text
																Status2 = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[8,"& linea + 1 &"]").text
																Ordem =	session.findById("wnd[0]/usr/subORD_HEADER:SAPLCOVG:0801/txtPSFC_DISP-AUFNR").text
																DataO = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").text
																locall = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[7,"& linea &"]").text
																
																
																if Mid(Trim(DataO), 1, 1) = "_" then
																	exit do
																end if

																If Mid(Trim(locall), 1, 7) = "EMBALAR" or Mid(Trim(locall), 1, 10) = "FECHAMENTO" then
																	Forne = objExcel.Sheets("Faturamento "& Data).Range("A"& nulin)
																	If Mid(Trim(Forne), 1, 10) = "61379-THOR" or Mid(Trim(Forne), 1, 12) = "509167-GRATT" or _
																	Mid(Trim(Forne), 1, 14) = "498700-LINDSAY" or Mid(Trim(Forne), 1, 13) = "16543-MEBRAFE" or _
																	Mid(Trim(Forne), 1, 17) = "4359-A CARNEVALLI" or Mid(Trim(Forne), 1, 12) = "785316-HEDEL" or _
																	Mid(Trim(Forne), 1, 16) = "842543-EXTRUSION" or Mid(Trim(Forne), 1, 11) = "17833-RULLI" or _
																	Mid(Trim(Forne), 1, 13) = "344907-WORTEX" or Mid(Trim(Forne), 1, 14) = "803687-LINDNER" or _
																	Mid(Trim(Forne), 1, 10) = "796935-WEG" then
																		objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = DataO
																		objExcel.Sheets("Faturamento "& Data).Range("J"& nulin).Replace ".", "/"
																	else
																		if Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 1 then
																			objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 3
																		elseif Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 7 then
																			objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 2
																		else 
																			objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1
																		end if
																	end if
																	'exit do
																end if


																if str = 0 then 'or Status = "CNPA LIB" Or Status = "CONF LIB" then
																
																			if linea > 0 and Status = "CNPA CFMA LIB" and Status2 = "LIB" or linea > 0 and Status = "CNPA LIB" and Status2 = "LIB" or linea > 0 and Status = "CNPA CFMA LIB" and Status2 = "CNPA LIB" or linea > 0 and  Status = "CNPA CFMA LIB" and Status2 = "CNPA CFMA LIB" or linea > 0 and Status = "CNPA LIB" and Status2 = "CNPA LIB" or linea > 0 and Status = "CNPA CFMA LIB  PLAN" and Status2 = "LIB" or linea > 0 and Status = "CNPA CFMA LIB  PLAN" and Status2 = "CNPA CFMA LIB  PLAN" or linea > 0 and Status = "CNPA CFMA LIB  PLAN" and Status2 = "CNPA CFMA LIB" or linea > 0 and Status = "CNPA CFMA LIB  PLAN" and Status2 = "CNPA LIB" then
																				objExcel.Sheets("Faturamento "& Data).Range("P"& nulin) = "X"
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				QtdOrd = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LMNGA[10,"& linea &"]").text
																				QtdOrd2 = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LMNGA[10,"& linea - 1 &"]").text
																				if QtdOrd = QtdOrd2 then
																					objExcel.Sheets("Faturamento "& Data).Range("Q"& nulin) = "X"
																				end if
																				str = 1
																			elseif Status = "LIB" and Status2 = "LIB" then
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				str = 1
																			elseif Status = "LIB  PLAN" and Status2 = "LIB" or Status = "LIB  PLAN" and Status2 = "CNPA LIB" or Status = "LIB  PLAN" and Status2 = "CONF CFMA LIB" or Status = "LIB  PLAN" and Status2 = "CNPA CFMA LIB  PLAN" then
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				str = 1
																			elseif Status = "LIB" and Status2 = "ELIM LIB" or Status = "CNPA CFMA LIB" and Status2 = "ELIM LIB" or Status = "CNPA LIB" and Status2 = "ELIM LIB" or Status = "CNPA CFMA LIB  PLAN" and Status2 = "ELIM LIB" then
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				str = 1
																			elseif Status = "CNPA CFMA LIB" and Mid(Trim(Status2), 1, 1) = "_" or Status = "CNPA LIB" and Mid(Trim(Status2), 1, 1) = "_" or Status = "CNPA CFMA LIB" and Status2 = "" or Status = "CNPA LIB" and Status2 = "" or Status = "CNPA CFMA LIB  PLAN" and Status2 = "" then
																				objExcel.Sheets("Faturamento "& Data).Range("P"& nulin) = "X"
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				str = 1	
																			elseif Status = "CONF LIB" and Mid(Trim(Status2), 1, 1) = "_" or Status = "CONF CFMA LIB" and Mid(Trim(Status2), 1, 1) = "_" or Status = "LIB" and Mid(Trim(Status2), 1, 1) = "_" or Status = "CONF LIB" and Status2 = "" or Status = "CONF CFMA LIB" and Status2 = "" or Status = "LIB" and Status2 = "" then
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				str = 1
																			elseif Status = "ABER" and Mid(Trim(Status2), 1, 1) = "_" then
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "ABER"
																			end if
																			
																	'if Status = "CNPA CFMA LIB" or Status = "CNPA LIB" then
																	'	objExcel.Sheets("Faturamento "& Data).Range("P"& nulin) = "X"
																	'	locallb = locall
																	'else
																	'	locallc = locall
																	'end if
																	'if locallb = "" then
																	'	objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locallc
																	'	str = 1
																	'else
																	'	objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locallb
																	'end if
																	
																	objExcel.Sheets("Faturamento "& Data).Range("C"& nulin) = codpro
																	objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = Ordem2
																	
																	if desc = "" then
																		'objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "____"
																	end if
																	
																end if

																linea = linea + 1
															loop
																if objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "" then
																	objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "1321 - Verificar"
																end if
																session.findById("wnd[0]/tbar[0]/btn[15]").press
																On error resume next
																session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
																FreakoutAndFixTheError
																err.clear
																
																'status = 1
															'exit do
															
														end if
													end if
												end if
											else
												tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text
												if tpo = "Fornec" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Fornecido"
												elseif tpo = "OrdPla" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Liberar Ordem"
												elseif tpo = "EstDep" then
												
												elseif tpo = "SolCnt" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Transferir"
												elseif tpo = "ReqCmp" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Ordem Compra"
												end if
											end if
											'if status = 1 then
											'	exit do
											'end if
										end if
										If objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "" then
											tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text
											if tpo = "Fornec" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Fornecido.1"
											elseif tpo = "OrdPla" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Liberar Ordem.1"
											elseif tpo = "EstDep" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Estoque Depo.1"
											elseif tpo = "SolCnt" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - SOL. CNT.1"
											elseif tpo = "ReqCmp" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Ordem Compra.1"
											end if
										end if
										
										line = line + 1
									loop
									if OPnum = 0 then
										objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
										objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
										if Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 1 then
											objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 3
										elseif Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 7 then
											objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 2
										else 
											objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1
										end if
										tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text
										if tpo = "Fornec" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Fornecido.2"
										elseif tpo = "OrdPla" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Liberar Ordem.2"
										elseif tpo = "EstDep" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Estoque Depo.2"
										elseif tpo = "SolCnt" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Transferir.2"
										elseif tpo = "ReqCmp" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Ordem Compra.2"
										elseif tpo = "OrdCli" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "Separar"	
										elseif tpo = "OrdPro" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "Verificar"
										end if
										if objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "" then
											if objExcel.Sheets("Faturamento "& Data).Range("N1") > objExcel.Sheets("Faturamento "& Data).Range("O"& nulin) + 20 then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Data OP e Maior que data OV"
											end if
										end if
									end if
									'session.findById("wnd[0]/tbar[0]/btn[15]").press
								else
						'===================================== VERIFICAÇÃO 1320 ORDEMS =======================================================================
									session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-BERID").text = "1320"
									session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-WERKS").text = "1320"
									session.findById("wnd[0]").sendVKey 0
									If session.findById("wnd[0]/sbar").messagetype = "I" then
										session.findById("wnd[0]").sendVKey 0
										objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Nao tem3"
									end if
									
									line = 1
									tt = ""
									Do until session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text = ""
									'msgbox("Tipo "&session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text &" 1320")
										tt = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text
										
										if Mid(Trim(tt), 1, 1) = "_" then
											exit do
										end if
										
										if line > 20 then
											session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ").verticalScrollbar.position = line - 15
										end if
										DataOR = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& line &"]").text
										Nesce = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG01[4,"& line &"]").text
										Ordem =	session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[3,"& line &"]").text
										if session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text <> "EstDep" then
											if Mid(Trim(Ordem), 11, 1) = "/" then
												objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = Mid(Trim(Ordem), 1, 10)
											else
												objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = Mid(Trim(Ordem), 1, 12)
											end if
										end if
										
										objExcel.Sheets("Faturamento "& Data).Range("N1") = DataOR
										objExcel.Sheets("Faturamento "& Data).Range("N1").Replace ".", "/"
										objExcel.Sheets("Faturamento "& Data).Range("N1").NumberFormat = "dd/mm/yyyy"
										
										'msgbox("DataOR "& DataOR &" 1320 | "& objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) &" < "& date + 5 )
										If Mid(Trim(DataOR), 1, 1) <> "_" and session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&linea&"]").text <> "EstDep" then 'and objExcel.Sheets("Faturamento "& Data).Range("N1") < date + temp then
										
											if tt = "OrdPro" or tt = "OrdPla" then
												'msgbox("Nesc "& Nesce &" 1320")
												if Nesce > 0 then
													if tt = "OrdPla" then
														objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
														objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
														objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321b - Ordem Planejada"
														OPJG = 1
													else
														session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[3,"&line&"]").setFocus
														session.findById("wnd[0]").sendVKey 2
														session.findById("wnd[1]/tbar[0]/btn[7]").press
														'"S" Sucess "W" Warning "E" Error "A" Abort "I" Information
														If session.findById("wnd[0]/sbar").messagetype = "W" then
															session.findById("wnd[0]").sendVKey 0
														end if
														
														
														on error resume next
														session.findById("wnd[0]/tbar[1]/btn[6]").press
														
														if Err.Number <> 0 then
															session.findById("wnd[0]/tbar[0]/btn[3]").press
															session.findById("wnd[0]").sendVKey 2
															session.findById("wnd[1]/tbar[0]/btn[7]").press
															session.findById("wnd[0]/tbar[1]/btn[6]").press
															
															On error resume next
															session.findById("wnd[0]").sendVKey 0
															FreakoutAndFixTheError
															err.clear
														end if
														
														session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1,0]").setFocus
														
														ttp = 0
														do until ttp = 10
															agru = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MATXT[2,"& ttp &"]").text
															if Mid(Trim(agru), 1, 9) = "AGRUPADOR" then
																codpro = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1,"& ttp &"]").text
																exit do
															end if
															ttp = ttp + 1
														loop

														session.findById("wnd[0]/tbar[0]/btn[3]").press
														session.findById("wnd[0]/tbar[0]/btn[3]").press

														if codpro = "" then
															OPJG = 1
															objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
															objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
															objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321a - Ordem Planejada"
														else
															session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-MATNR").text = codpro
															objExcel.Sheets("Faturamento "& Data).Range("C"& nulin) = codpro
															session.findById("wnd[0]").sendVKey 0

															session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-BERID").text = "1321"
															session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/ctxtRM61R-WERKS").text = "1321"
															session.findById("wnd[0]").sendVKey 0
															If session.findById("wnd[0]/sbar").messagetype = "I" then
																session.findById("wnd[0]").sendVKey 0
																objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Não tem"
															end if
															OPJG = 0
														end if
														exit do
													end if
												else
													
													if Mid(Trim(Nesce), 1, 1) = "_" then
														objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Concluido"
														OPJG = 1
														exit do
													end if

													session.findById("wnd[0]").sendVKey 37
													session.findById("wnd[0]/usr/cntlGRID_1000/shellcont/shell/shellcont[1]/shell")
													set GRID = session.findById("wnd[0]/usr/cntlGRID_1000/shellcont/shell/shellcont[1]/shell")
													GRID.setCurrentCell 0,""
													
													
													ttov = 0
													do until ttov = -1
														
														DocTest = GRID.getcellvalue (ttov, "EXTRA")
														
														if Mid(Trim(DocTest), 15, 1) = "0" then
															Linetest = Mid(Trim(DocTest), 16, 2)
														else
															Linetest = Mid(Trim(DocTest), 15, 3)
														end if
														
														if Mid(Trim(DocTest), 2, 1) = "0" then
															DocID = Mid(Trim(DocTest), 3, 8)
														else
															DocID = Mid(Trim(DocTest), 2, 9)
														end if
														'msgbox(Doc &"="& DocID &"|"& LineOV &"="& Linetest)
														if CStr(Doc) = DocID and CStr(LineOV) = Linetest then
															if OPnum > 0 then
																OPnum = OPnum + 1
															else
																OPnum = 1
															end if
															OPexit = 1
															exit do
														end if
														
														On error resume next
														DatteTest = GRID.getcellvalue (ttov + 1, "DAT00")
														
														if Err.Number <> 0 then
															FreakoutAndFixTheError
															err.clear
															exit do
														end if
														
														ttov = ttov + 1	
													loop
													'msgbox(OPnum)
													session.findById("wnd[0]/tbar[0]/btn[3]").press
													
													if tt = "OrdPla" then
														objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
														objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
														objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Ordem Planejada"
														OPJG = 1
													'else
														'objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Concluido"
														'OPJG = 1
													end if
													
												end if

											
											else
												objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
												if session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "Fornec" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Fornecido.2"
													OPJG = 1	
												elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "OrdPla" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Liberar Ordem.2"
													OPJG = 1
												elseif Nesce < 0 then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Fazer Picking.2"
													OPJG = 1
												elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "EstDep" then
													OPJG = 1
												elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "SolCnt" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Transferir.2"
													OPJG = 1
												elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "ReqCmp" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Ordem Compra.2"
													OPJG = 1
												end if
											end if

										else
											objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
											if session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "Fornec" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Fornecido.3"
													OPJG = 1	
												elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "OrdPla" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Liberar Ordem.3"
													OPJG = 1
												elseif Nesce < 0 then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Fazer Picking.3"
													OPJG = 1
												elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "EstDep" then
													OPJG = 1
												elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "SolCnt" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Transferir.3"
													OPJG = 1
												elseif session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text = "ReqCmp" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Ordem Compra.3"
													OPJG = 1
												end if
											exit do
										end if	
										
										line = line + 1
									loop
								
						'===================================== FIM VERIFICAÇÃO 1320 ORDEMS =======================================================================


						'===================================== VERIFICAÇÃO 1321 ORDEMS =======================================================================
									If OPJG <> 1 then
											
										line = 1
								
									'status = 0
									'Do until session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text = "______"
									'msgbox("Tipo "&session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text &" 1321")
									

								session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& line &"]").setFocus
								DataOR = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& line &"]").text

								'msgbox("Data "& DataOR &" 1321")
								If Mid(Trim(DataOR), 1, 1) <> "_" and session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text <> "EstDep" then

									status = 0
									tpo = ""
									Do until session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text = ""
									'msgbox("Tipo "&session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text &" 1321")
										session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& line &"]").setFocus
										tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text
										if Mid(Trim(tpo), 1, 1) = "_" or line > 100 then
											if line > 100 then
												'objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "Verificar"
											end if
											exit do
										end if
										
										if line > 20 then
											session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ").verticalScrollbar.position = line - 15
										end if
										DataOR = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1,"& line &"]").text
										Disp = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG02[5,"& line &"]").text
										Ordem =	session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[3,"& line &"]").text
										
										if session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,"&line&"]").text <> "EstDep" then
											if Mid(Trim(Ordem), 11, 1) = "/" then
												Ordem2 = Mid(Trim(Ordem), 1, 10)
											else
												Ordem2 = Mid(Trim(Ordem), 1, 12)
											end if
										end if
										
										Nesce = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG01[4,"& line &"]").text
										
										objExcel.Sheets("Faturamento "& Data).Range("N1") = DataOR
										objExcel.Sheets("Faturamento "& Data).Range("N1").Replace ".", "/"
										objExcel.Sheets("Faturamento "& Data).Range("N1").NumberFormat = "dd/mm/yyyy"
										
										'msgbox("DataOR "& DataOR &" 1321 | "& objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) &" < "& date + 5 )
										if objExcel.Sheets("Faturamento "& Data).Range("N1") > objExcel.Sheets("Faturamento "& Data).Range("O"& nulin) + 20 then'date + temp then 'objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) then
											'objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Data OP e Maior que data OV"
											exit do
										end if
										
										if tpo = "OrdPro" or tpo = "OrdPla" then
											'msgbox("Nesc "&Nesce &" 1321")
											if Nesce > 0 then

				'=========================================Definição de quantidade por Ordem =============================================
												if Mid(Trim(Nesce), 1, 1) = "_" then
													exit do
												end if

												session.findById("wnd[0]").sendVKey 37
												session.findById("wnd[0]/usr/cntlGRID_1000/shellcont/shell/shellcont[1]/shell")
												set GRID = session.findById("wnd[0]/usr/cntlGRID_1000/shellcont/shell/shellcont[1]/shell")
												GRID.setCurrentCell 0,""
												
												
												ttov = 0
												do until ttov = -1
													
													DocTest = GRID.getcellvalue (ttov, "EXTRA")
													
													if Mid(Trim(DocTest), 15, 1) = "0" then
														Linetest = Mid(Trim(DocTest), 16, 2)
													else
														Linetest = Mid(Trim(DocTest), 15, 3)
													end if
													
													if Mid(Trim(DocTest), 2, 1) = "0" then
														DocID = Mid(Trim(DocTest), 3, 8)
													else
														DocID = Mid(Trim(DocTest), 2, 9)
													end if
													'msgbox(Doc &"="& DocID &"|"& LineOV &"="& Linetest)
													if CStr(Doc) = DocID and CStr(LineOV) = Linetest then
														if OPnum > 0 then
															OPnum = OPnum + 1
														else
															OPnum = 1
														end if
														OPexit = 1
														exit do
													end if
													
													On error resume next
													DatteTest = GRID.getcellvalue (ttov + 1, "DAT00")
													
													if Err.Number <> 0 then
														FreakoutAndFixTheError
														err.clear
														exit do
													end if
													
													ttov = ttov + 1	
												loop
												'msgbox(OPnum)
												session.findById("wnd[0]/tbar[0]/btn[3]").press
												
												if tpo = "OrdPla" and OPexit > 0 then
													objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
													objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Ordem Planejada"
													OPexit = 0
												'else
												'	objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
												'	objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
												'	objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Liberar Ordem"
												end if
												
												if OPexit > 0 then
													if tpo = "OrdPla" then
													else
													OPexit = 0
													ttov = 0
													'msgbox(OPnum)
													'objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = Ordem2
													
													if OPnum > 1 then
														objExcel.Sheets("Faturamento "& Data).Range("A"& nulin &":K"& nulin).copy
														'msgbox("NEW LINE")
														nulin = nulin + 1
														objExcel.Sheets("Faturamento "& Data).Range("A"& nulin &":K"& nulin).Insert -4121
														objExcel.Sheets("Faturamento "& Data).Range("N"& nulin &":O"& nulin).Insert -4121
														objExcel.Sheets("Faturamento "& Data).Range("A"& nulin &":K"& nulin).PasteSpecial -4163
														objExcel.Sheets("Faturamento "& Data).Range("N"& nulin - 1 &":O"& nulin - 1 ).copy
														objExcel.Sheets("Faturamento "& Data).Range("N"& nulin &":O"& nulin).PasteSpecial -4163
														objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("O"& nulin)
														
													end if
													
				'=========================================Definição de quantidade por Ordem =============================================

														session.findById("wnd[0]").sendVKey 2
														session.findById("wnd[1]/tbar[0]/btn[8]").press
														
														if session.Info.Program = "SAPLCMFE" then
															session.findById("wnd[0]/tbar[0]/btn[3]").press
															session.findById("wnd[0]").sendVKey 2
															session.findById("wnd[1]/tbar[0]/btn[7]").press
															
														end if
																
														'"S" Sucess "W" Warning "E" Error "A" Abort "I" Information
														If session.findById("wnd[0]/sbar").messagetype = "W" then
															session.findById("wnd[0]").sendVKey 0
														end if

														On error resume next
														session.findById("wnd[1]/tbar[0]/btn[0]").press
														FreakoutAndFixTheError
														err.clear

															
														On error resume next
														session.findById("wnd[0]/tbar[1]/btn[28]").press
														FreakoutAndFixTheError
														err.clear
														desc = ""
														
														'"S" Sucess "W" Warning "E" Error "A" Abort "I" Information
														If session.findById("wnd[0]/sbar").messagetype = "S" then
															session.findById("wnd[0]/tbar[1]/btn[5]").press
														else
															session.findById("wnd[1]/usr/btnDY_VAROPTION3").press
															session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/ctxtRESBD-MATNR[1,0]").setFocus
															codfalt = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/ctxtRESBD-MATNR[1,0]").text
															numf = 0
															
															do until session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/ctxtRESBD-MATNR[1,"&numf&"]").text = ""
															codfalt = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/ctxtRESBD-MATNR[1,"&numf&"]").text
															descfalt = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/txtRESBD-MATXT[8,"&numf&"]").text
															locall = codfalt &" - "& descfalt
															if Mid(Trim(locall), 12, 9) <> "EMBALAGEM" then
																desc = desc & vbNewLine & Mid(Trim(locall), 1, 20)
															end if
															if numf >= 5 then
																exit do
															end if
																numf = numf + 1
															loop
															
															session.findById("wnd[0]/tbar[0]/btn[12]").press
															session.findById("wnd[0]/tbar[1]/btn[5]").press
															
														end if
														'msgbox(desc)
														if desc = "" then
															objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "----"
														else
															objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = desc
														end if

														If session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[8,1]").text = "BLOQ LIB" _
														or session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[8,1]").text = "CNPA BLOQ" or session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[8,1]").text = "CONF BLOQ LIB" then

															objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
															objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "Jaragua Block"


															Forne = objExcel.Sheets("Faturamento "& Data).Range("A"& nulin)
															
															If Mid(Trim(Forne), 1, 10) = "61379-THOR" or Mid(Trim(Forne), 1, 12) = "509167-GRATT" or _
															Mid(Trim(Forne), 1, 14) = "498700-LINDSAY" or Mid(Trim(Forne), 1, 13) = "16543-MEBRAFE" or _
															Mid(Trim(Forne), 1, 17) = "4359-A CARNEVALLI" or Mid(Trim(Forne), 1, 12) = "785316-HEDEL" or _
															Mid(Trim(Forne), 1, 16) = "842543-EXTRUSION" or Mid(Trim(Forne), 1, 11) = "17833-RULLI" or _
															Mid(Trim(Forne), 1, 13) = "344907-WORTEX" or Mid(Trim(Forne), 1, 14) = "803687-LINDNER" or _
															Mid(Trim(Forne), 1, 10) = "796935-WEG" then
																linea = 0
																do until session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").text = ""
																	session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").setFocus
																	locall = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[7,"& linea &"]").text
																	DataO = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").text
																	If Mid(Trim(locall), 1, 7) = "EMBALAR" or Mid(Trim(locall), 1, 10) = "FECHAMENTO" then
																		objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = DataO
																		objExcel.Sheets("Faturamento "& Data).Range("J"& nulin).Replace ".", "/"
																		exit do
																	end if
																	linea = linea + 1
																loop
															else 
																if Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 1 then
																	objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 3
																elseif Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 7 then
																	objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 2
																else 
																	objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1
																end if
															end if


																session.findById("wnd[0]/tbar[0]/btn[15]").press
																On error resume next
																session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
																FreakoutAndFixTheError
																err.clear

																ext = 0
																status = 1
															exit do

														else

															linea = 0
															session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,1]").setFocus
															str = 0

															do until session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").text = ""
																session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").setFocus
																Status = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[8,"& linea &"]").text
																Status2 = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-VSTTXT[8,"& linea + 1 &"]").text
																Ordem =	session.findById("wnd[0]/usr/subORD_HEADER:SAPLCOVG:0801/txtPSFC_DISP-AUFNR").text
																DataO = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,"& linea &"]").text
																locall = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LTXA1[7,"& linea &"]").text
																
																
																if Mid(Trim(DataO), 1, 1) = "_" then
																	exit do
																end if

																If Mid(Trim(locall), 1, 7) = "EMBALAR" or Mid(Trim(locall), 1, 10) = "FECHAMENTO" then
																	Forne = objExcel.Sheets("Faturamento "& Data).Range("A"& nulin)
																	If Mid(Trim(Forne), 1, 10) = "61379-THOR" or Mid(Trim(Forne), 1, 12) = "509167-GRATT" or _
																	Mid(Trim(Forne), 1, 14) = "498700-LINDSAY" or Mid(Trim(Forne), 1, 13) = "16543-MEBRAFE" or _
																	Mid(Trim(Forne), 1, 17) = "4359-A CARNEVALLI" or Mid(Trim(Forne), 1, 12) = "785316-HEDEL" or _
																	Mid(Trim(Forne), 1, 16) = "842543-EXTRUSION" or Mid(Trim(Forne), 1, 11) = "17833-RULLI" or _
																	Mid(Trim(Forne), 1, 13) = "344907-WORTEX" or Mid(Trim(Forne), 1, 14) = "803687-LINDNER" or _
																	Mid(Trim(Forne), 1, 10) = "796935-WEG" then
																		objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = DataO
																		objExcel.Sheets("Faturamento "& Data).Range("J"& nulin).Replace ".", "/"
																	else
																		if Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 1 then
																			objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 3
																		elseif Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 7 then
																			objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 2
																		else 
																			objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1
																		end if
																	end if
																	'exit do
																end if


																if str = 0 then 'or Status = "CNPA LIB" Or Status = "CONF LIB" then
																
																			if linea > 0 and Status = "CNPA CFMA LIB" and Status2 = "LIB" or linea > 0 and Status = "CNPA LIB" and Status2 = "LIB" or linea > 0 and Status = "CNPA CFMA LIB" and Status2 = "CNPA LIB" or linea > 0 and  Status = "CNPA CFMA LIB" and Status2 = "CNPA CFMA LIB" or linea > 0 and Status = "CNPA LIB" and Status2 = "CNPA LIB" or linea > 0 and Status = "CNPA CFMA LIB  PLAN" and Status2 = "LIB" or linea > 0 and Status = "CNPA CFMA LIB  PLAN" and Status2 = "CNPA CFMA LIB  PLAN" or linea > 0 and Status = "CNPA CFMA LIB  PLAN" and Status2 = "CNPA CFMA LIB" or linea > 0 and Status = "CNPA CFMA LIB  PLAN" and Status2 = "CNPA LIB" then
																				objExcel.Sheets("Faturamento "& Data).Range("P"& nulin) = "X"
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				QtdOrd = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LMNGA[10,"& linea &"]").text
																				QtdOrd2 = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/txtAFVGD-LMNGA[10,"& linea - 1 &"]").text
																				if QtdOrd = QtdOrd2 then
																					objExcel.Sheets("Faturamento "& Data).Range("Q"& nulin) = "X"
																				end if
																				str = 1
																			elseif Status = "LIB" and Status2 = "LIB" then
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				str = 1
																			elseif Status = "LIB  PLAN" and Status2 = "LIB" or Status = "LIB  PLAN" and Status2 = "CNPA LIB" or Status = "LIB  PLAN" and Status2 = "CONF CFMA LIB" or Status = "LIB  PLAN" and Status2 = "CNPA CFMA LIB  PLAN" then
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				str = 1
																			elseif Status = "LIB" and Status2 = "ELIM LIB" or Status = "CNPA CFMA LIB" and Status2 = "ELIM LIB" or Status = "CNPA LIB" and Status2 = "ELIM LIB" or Status = "CNPA CFMA LIB  PLAN" and Status2 = "ELIM LIB" then
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				str = 1
																			elseif Status = "CNPA CFMA LIB" and Mid(Trim(Status2), 1, 1) = "_" or Status = "CNPA LIB" and Mid(Trim(Status2), 1, 1) = "_" or Status = "CNPA CFMA LIB" and Status2 = "" or Status = "CNPA LIB" and Status2 = "" or Status = "CNPA CFMA LIB  PLAN" and Status2 = "" then
																				objExcel.Sheets("Faturamento "& Data).Range("P"& nulin) = "X"
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				str = 1	
																			elseif Status = "CONF LIB" and Mid(Trim(Status2), 1, 1) = "_" or Status = "CONF CFMA LIB" and Mid(Trim(Status2), 1, 1) = "_" or Status = "LIB" and Mid(Trim(Status2), 1, 1) = "_" or Status = "CONF LIB" and Status2 = "" or Status = "CONF CFMA LIB" and Status2 = "" or Status = "LIB" and Status2 = "" then
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locall
																				str = 1
																			elseif Status = "ABER" and Mid(Trim(Status2), 1, 1) = "_" then
																				objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "ABER"
																			end if
																	'if Status = "CNPA CFMA LIB" or Status = "CNPA LIB" then
																	'	objExcel.Sheets("Faturamento "& Data).Range("P"& nulin) = "X"
																	'	locallb = locall
																	'else
																	'	locallc = locall
																	'end if
																	'if locallb = "" then
																	'	objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locallc
																	'	str = 1
																	'else
																	'	objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = locallb
																	'end if
																	
																	objExcel.Sheets("Faturamento "& Data).Range("C"& nulin) = codpro
																	objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = Ordem2
																	
																	
																	if desc = "" then
																		'objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "____"
																	end if
																	
																end if

																linea = linea + 1
															loop
																if objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "" then
																	objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "1321 - Verificar"
																end if
																session.findById("wnd[0]/tbar[0]/btn[15]").press
																On error resume next
																session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
																FreakoutAndFixTheError
																err.clear
																
																'status = 1
															'exit do
															
														end if
													end if
												end if
											else
												tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text
												if tpo = "Fornec" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Fornecido"
												elseif tpo = "OrdPla" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Liberar Ordem"
												elseif tpo = "EstDep" then
												
												elseif tpo = "SolCnt" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Transferir"
												elseif tpo = "ReqCmp" then
													objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Ordem Compra"
												end if
											end if
											'if status = 1 then
											'	exit do
											'end if
										end if
										If objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "" then
											tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text
											if tpo = "Fornec" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Fornecido.1"
											elseif tpo = "OrdPla" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Liberar Ordem.1"
											elseif tpo = "EstDep" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Estoque Depo.1"
											elseif tpo = "SolCnt" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - SOL. CNT.1"
											elseif tpo = "ReqCmp" then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Ordem Compra.1"
											end if
										end if
										
										line = line + 1
									loop
									if OPnum = 0 then
										objExcel.Sheets("Faturamento "& Data).Range("K"& nulin) = "-----------"
										objExcel.Sheets("Faturamento "& Data).Range("L"& nulin) = "-----------"
										if Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 1 then
											objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 3
										elseif Weekday(objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1) = 7 then
											objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 2
										else 
											objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) = objExcel.Sheets("Faturamento "& Data).Range("J"& nulin) - 1
										end if
										tpo = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2,1]").text
										if tpo = "Fornec" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Fornecido.2"
										elseif tpo = "OrdPla" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Liberar Ordem.2"
										elseif tpo = "EstDep" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Estoque Depo.2"
										elseif tpo = "SolCnt" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Transferir.2"
										elseif tpo = "ReqCmp" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Ordem Compra.2"
										elseif tpo = "OrdCli" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "Separar.2"
										elseif tpo = "OrdPro" then
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "Verificar.2"
										end if
										if objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "" then
											if objExcel.Sheets("Faturamento "& Data).Range("N1") > objExcel.Sheets("Faturamento "& Data).Range("O"& nulin) + 20 then
												objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1321 - Data OP e Maior que data OV"
											end if
										end if
										
									end if
									'session.findById("wnd[0]/tbar[0]/btn[15]").press
										
										else
											objExcel.Sheets("Faturamento "& Data).Range("M"& nulin) = "1320 - Nao tem6"
											
										end if
											'session.findById("wnd[0]/tbar[0]/btn[15]").press
										OPJG = 0
									end if
								
								end if
						'===================================== FIM VERIFICAÇÃO 1321 ORDEMS =======================================================================
								
							'loop
								
								OPJG = 0
							end if '---- SE OPJG = 1
						end if '---- SE FOR 1 - 1 - 1 
					do until session.Info.Transaction = "SESSION_MANAGER"
						session.findById("wnd[0]/tbar[0]/btn[3]").press
					loop
					If objExcel.Sheets("Faturamento "& Data).Range("C"& nulin) < 10000 then
						objExcel.Sheets("Faturamento "& Data).Range("C"& nulin) = ""
					end if
					OPnum = 0
					nulin = nulin + 1
					Ordem = ""
					codpro = ""
					Forne = ""
					numet = ""
					DataO = ""
					locall = ""
					line = ""
					linea = ""
					str = ""
					status = ""
					Nesce = ""
					agru = ""
					numf = ""
					
				loop

				do until session.Info.Transaction = "SESSION_MANAGER"
					session.findById("wnd[0]/tbar[0]/btn[3]").press
				loop

				session.findById("wnd[0]/tbar[0]/btn[15]").press
		end if

		objExcel.Sheets("Faturamento "& Data).Range("N1") = "Line"
		objExcel.Sheets("Faturamento "& Data).Range("P1") = "CNPA"
		objExcel.Sheets("Faturamento "& Data).Range("Q1") = "Err.AP"
		objExcel.Sheets("Faturamento "& Data).Range("R1") = "Prior."
		objExcel.Sheets("Faturamento "& Data).Range("R2").FormulaR1C1 = _
				"=IF(RC[-15]=""61379-THOR MAQUINAS E MONTAGENS LTDA"",1,IF(RC[-15]=""509167-GRATT INDUSTRIA DE MAQUINAS LTDA"",2,IF(RC[-15]=""498700-LINDSAY AMERICA DO SUL LTDA"",3,IF(RC[-15]=""16543-MEBRAFE INST.EQUIP.FRIG.LTDA"",4,IF(RC[-15]=""4359-A CARNEVALLI & CIA LTDA"",5,IF(RC[-15]=""785316-HEDEL MAQS. E EQUIPS. LTDA"",6,IF(RC[-15]=""842543-EXTRUSION SYSTEM INDUSTRIA E COMERC" & _
				""",7,IF(RC[-15]=""17833-RULLI STANDARD IND.COM.DE MAQS.LTDA"",8,IF(RC[-15]=""796935-WEG EQUIP. ELETRICOS S/A - EOLICA"",9,0)))))))))" & _
				""
		objExcel.Sheets("Faturamento "& Data).Range("R2").copy
		objExcel.Sheets("Faturamento "& Data).Range("R2:R"& nulin-1).PasteSpecial -4123
		objExcel.Sheets("Faturamento "& Data).Range("R2:R"& nulin-1).copy
		objExcel.Sheets("Faturamento "& Data).Range("R2:R"& nulin-1).PasteSpecial -4163

		objExcel.Sheets("Faturamento "& Data).Range("A1:Q1").Insert -4121
		objExcel.Sheets("Faturamento "& Data).Range("A1:Q1").Interior.Color = 65535'10092492 '65535 RGB(6,55,35) .ColorIndex = 36
		objExcel.Sheets("Faturamento "& Data).Range("A1").FormulaR1C1 = "Quantidade Ovs:"
		objExcel.Sheets("Faturamento "& Data).Range("B1").FormulaR1C1 = "=COUNT(R[2]C:R["& nulin-1 &"]C)"
		objExcel.Sheets("Faturamento "& Data).Range("C1").FormulaR1C1 = "Valor Medio OVs:"
		objExcel.Sheets("Faturamento "& Data).Range("E1").FormulaR1C1 = "=AVERAGE(R[2]C[4]:R["& nulin-1 &"]C[4])"
		objExcel.Sheets("Faturamento "& Data).Range("G1").FormulaR1C1 = "Valor Total OVs:"
		'objExcel.Sheets("Faturamento "& Data).Range("F1").FormulaR1C1 = "=SUM(R[2]C:R["& nulin &"]C)"
		'objExcel.Sheets("Faturamento "& Data).Range("G1").FormulaR1C1 = "=SUM(R[2]C:R["& nulin &"]C)"
		'objExcel.Sheets("Faturamento "& Data).Range("H1").FormulaR1C1 = "=SUM(RC[-3]:RC[-2])"
		objExcel.Sheets("Faturamento "& Data).Range("I1").FormulaR1C1 = "=SUM(R[2]C:R["& nulin-1 &"]C)"
		objExcel.Sheets("Faturamento "& Data).Range("J1").FormulaR1C1 = "=R["& nulin-1 &"]C-R[2]C&"" Dias"""
		objExcel.Sheets("Faturamento "& Data).Range("L1").FormulaR1C1 = "Quantidade OVs Concluidas:"
		objExcel.Sheets("Faturamento "& Data).Range("M1").FormulaR1C1 = "=COUNTIF(R[2]C:R["& nulin-1 &"]C,""-----------"")"
		objExcel.Sheets("Faturamento "& Data).Range("O1").FormulaR1C1 = "Qtd.Parcial:"
		objExcel.Sheets("Faturamento "& Data).Range("P1").FormulaR1C1 = "=COUNTIF(R[2]C:R["& nulin-1 &"]C,""X"")"
		'objExcel.Sheets("Faturamento "& Data).Range("A1:Q1").copy
		'objExcel.Sheets("Faturamento "& Data).Range("A1:Q1").PasteSpecial -4163


		'XEdge = 8 'Top
		'XEdge = 9 'Bottom
		'XEdge = 7 'Left
		'XEdge = 10 'Right
		'objExcel.Sheets("Faturamento "& Data).Range("A1:Q1").Borders(XEdge).LineStyle = 1
		'objExcel.Sheets("Faturamento "& Data).Range("A1:Q1").Borders(XEdge).Weight = 4

		objExcel.Sheets("Faturamento "& Data).Range("A1:R1").Borders.LineStyle = 1
		objExcel.Sheets("Faturamento "& Data).Range("A2:R2").Borders.LineStyle = 1

		objExcel.Sheets("Faturamento "& Data).Range("A2:R2").Interior.Color = 15102720
		objExcel.Sheets("Faturamento "& Data).Range("A2:R2").Font.Color = RGB(255, 255, 255)
		objExcel.Sheets("Faturamento "& Data).Rows("1:2").Font.Bold = True

		obj.ActiveWorkbook.SaveAs Path + "Faturamento "& Data & "old.xls", 51
		obj.ActiveWorkbook.Close
		obj.Quit
		Set objWorkbook = Nothing
		Set objExcel = Nothing
		Set session = Nothing		
			'Data = Date + 1		
		exit do
		end if

Loop

xOutMsg = 	"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""><span id=""j_id0:emailTemplate:j_id3:j_id4:j_id6"">" & _
			"<html>" & _
            "<body>" & _
			"Bom dia<br/><br/>" & _
			"Segue em anexo relatório de Faturamento "& Data &".<br/><br/>" & _
			"</body>" & _
			"</html></span>" & _
			"<span style=""color:#104E8B""><b>Paulo Antonio</b></span style=""color:#104E8B""><br/>" & _
			"<span style=""color:#104E8B; font-size:12px; font-family:'Arial'"">CRM Automação (Itajai)</span><br/>" & _
			"<span style=""color:#104E8B; font-size:12px; font-family:'Arial'"">Fone:+55 (47) 3276 7424</span><br/>" & _
			"<span style=""color:#104E8B; font-size:12px; font-family:'Arial'"">WEG Drives & Conrols</span><br/>" & _
			"<span style=""color:#0000CD; font-size:12px; font-family:'Arial'"">www.weg.net</span><br/>"
            
			'"<u>New line with underline</u><br />
			'<p style='font-family:calibri;font-size:25'>Font size</p>"

Set emailObj = CreateObject("CDO.Message")

Set emailConfig = emailObj.Configuration
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.weg.net"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")    = 2 
emailConfig.Fields.Update

emailObj.From = "Paulo Antonio Muller dos Santos <pauloantonio@weg.net>" '"no-reply@weg.net"
emailObj.To = "Silene Krause Clementino <silenek@weg.net>"
emailObj.Cc = "William Luis Caminada <wilianc@weg.net>;Paulo Antonio Muller dos Santos <pauloantonio@weg.net>;Luana Americo Fagundes da Silva <luanaamerico@weg.net>"
emailObj.Subject = "Relatorio Faturamento" & Data
emailObj.HTMLBody = xOutMsg
emailObj.AddAttachment Path + "Faturamento "& Data & "old.xls"

emailObj.Send
Set emailObj = Nothing

if DesligaPC = 1 then
WsShell.Run "cmd"
WScript.Sleep 100
WsShell.SendKeys "shutdown /f" 
WsShell.SendKeys "~"
end if
