<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: FacilidadesPedido.asp
'	- Responsável		: Vital
'	- Descrição			: Lista as facilidade do Pedido

Function FacilidadesPedido(objRSFac)

	Dim objDic
	Dim strChave
	Dim strRet
	Dim intIndex
	Dim strTronco		
	Dim strBastidor		
	Dim strRegua		
	Dim strposicao		
	Dim strFila			
	Dim strLateral		
	Dim strTipoCabo		
	Dim strCxEmenda		
	Dim strDominio		
	Dim strNo			
	Dim strObs			
	Dim strRepresentacao

	Set objDic = Server.CreateObject("Scripting.Dictionary") 
	' alterado por PRSS - A LINHA ESTAVA ERRONEAMENTE COMENTADA E SEM INFORMAÇÃO DE QUEM A COMENTOU
	Set objRSFac = db.execute("CLA_sp_sel_facilidade null,null,null," & dblAcfId)
	' /PRSS
	
	intIndex = 1
	If Not objRSFac.eof then

		strRet = strRet & "<table  border=0 cellspacing=1 cellpadding=0  width=760>"

		While not objRSFac.Eof

				Select Case objRSFac("Sis_Id")
					Case 1
						strChave = objRSFac("Fac_TimeSlot")
						strRepresentacao = objRSFac("Fac_TimeSlot")
					Case Else	
						strChave = objRSFac("Fac_Par")
						strRepresentacao = objRSFac("Fac_Par")

				End Select		

				if not isNull(objRSFac("Fac_Representacao")) then
					strRepresentacao = objRSFac("Fac_Representacao")
				End if

				strNroAcessoPtaEbt	= objRSFac("Acf_NroAcessoPtaEBT")
				strNroAcessoPtaCli	= objRSFac("Acf_NroAcessoPtaCli")
				strNroAcessoCCTOPRov= objRSFac("Acf_NroAcessoCCTOProvedor")

				strTronco		= objRSFac("Fac_Tronco")
				strBastidor		= objRSFac("Fac_Bastidor")
				strRegua		= objRSFac("Fac_Regua")
				strposicao		= objRSFac("Fac_posicao")
				strFila			= objRSFac("Fac_Fila")

				strLateral		= objRSFac("Fac_Lateral")
				strTipoCabo		= objRSFac("Fac_TipoCabo")
				strCxEmenda		= objRSFac("Fac_CxEmenda")
				strDominio		= objRSFac("Fac_Dominio")
				strNo			= objRSFac("Fac_No")
				strSlot			= objRSFac("Fac_Slot")
				strPorta		= objRSFac("Fac_Porta")
				strObs			= objRSFac("Acf_Obs")
				dblSisId		= objRSFac("Sis_ID")

				Select Case dblSisId
					Case 1
					    ' PRSS/LPEREZ - 11/11/2005
						if  Not objDic.Exists(strRepresentacao&strBastidor&strRegua&strposicao&strDominio&strNo&strSlot&strPorta) then ' PRSS/LPEREZ - 11/11/2005
							Call objDic.Add(strRepresentacao&strBastidor&strRegua&strposicao&strDominio&strNo&strSlot&strPorta,objRSFac("Fac_Id")) ' PRSS/LPEREZ - 11/11/2005
							if intIndex = 1 then
								strRet = strRet & "<tr class=clsSilver2>"
								strRet = strRet & "<td nowrap>&nbsp;Nº Acso Pta EBT</td>"
								strRet = strRet & "<td nowrap>&nbsp;Nº Acso Pta CLI</td>"
								strRet = strRet & "<td nowrap>&nbsp;Nº Acso CCTO Prov</td>"
								strRet = strRet & "<td>&nbsp;Fila</td>"
								strRet = strRet & "<td>&nbsp;Bastidor</td>"
								strRet = strRet & "<td>&nbsp;Regua</td>"
								strRet = strRet & "<td>&nbsp;Posição</td>"
								strRet = strRet & "<td>&nbsp;Domínio</td>"
								strRet = strRet & "<td>&nbsp;Nó</td>"
								strRet = strRet & "<td>&nbsp;Slot</td>"
								strRet = strRet & "<td>&nbsp;Porta</td>"
								strRet = strRet & "<td nowrap>&nbsp;Timeslot</td>"
								strRet = strRet & "<td width=200>&nbsp;Obs</td>"
								strRet = strRet & "</tr>"
							End if	
							strRet = strRet & "<tr class=clsSilver>"
							strRet = strRet & "<td >&nbsp;" & strNroAcessoPtaEbt & "</td>"
							strRet = strRet & "<td >&nbsp;" & strNroAcessoPtaCli & "</td>"
							strRet = strRet & "<td >&nbsp;" & strNroAcessoCCTOPRov & "</td>"
							strRet = strRet & "<td >&nbsp;" & strFila & "</td>"
							strRet = strRet & "<td >&nbsp;" & strBastidor & "</td>"
							strRet = strRet & "<td >&nbsp;" & strRegua & "</td>"
							strRet = strRet & "<td >&nbsp;" & strposicao & "</td>"
							strRet = strRet & "<td >&nbsp;" & strDominio & "</td>"
							strRet = strRet & "<td >&nbsp;" & strNo & "</td>"
							strRet = strRet & "<td >&nbsp;" & strSlot & "</td>"
							strRet = strRet & "<td >&nbsp;" & strPorta & "</td>"
							strRet = strRet & "<td nowrap>&nbsp;" & strRepresentacao & "</td>"
							strRet = strRet & "<td >&nbsp;" & strObs & "</td>"
							strRet = strRet & "</tr>"
						End if
					Case 2

						If objRSFac("Int_ID") <> "" and Not isnull(objRSFac("Int_ID")) then

							if intIndex = 1 then
								strRet = strRet & "<tr class=clsSilver2>"
								strRet = strRet & "<td nowrap>&nbsp;Nº Acso Pta EBT</td>"
								strRet = strRet & "<td nowrap>&nbsp;Nº Acso Pta CLI</td>"
								strRet = strRet & "<td nowrap>&nbsp;Nº Acso CCTO Prov</td>"
								strRet = strRet & "<td width=70>&nbsp;Tronco</td>"
								strRet = strRet & "<td width=50>&nbsp;Par</td>"
								strRet = strRet & "<td width=50>&nbsp;Coordenada</td>"
								strRet = strRet & "<td width=320>&nbsp;Observação</td>"
								strRet = strRet & "</td>"
								strRet = strRet & "</tr>"
							End if
							strRet = strRet & "<tr class=clsSilver>"
							strRet = strRet & "<td >&nbsp;" & strNroAcessoPtaEbt & "</td>"
							strRet = strRet & "<td >&nbsp;" & strNroAcessoPtaCli & "</td>"
							strRet = strRet & "<td >&nbsp;" & strNroAcessoCCTOPRov & "</td>"
							strRet = strRet & "<td >&nbsp;" & strTronco & "</td>"
							strRet = strRet & "<td nowrap>&nbsp;" & strChave & "</td>"
							strRet = strRet & "<td nowrap>&nbsp;"
							If objRSFac("Int_ID") <> "" and Not isnull(objRSFac("Int_ID")) then
								Set objRSInt = db.execute("CLA_sp_sel_interligacao " & objRSFac("Int_ID"))
								While not objRSInt.eof 
									strRet = strRet & objRSInt("Int_CorOrigem") & "&nbsp;&nbsp;>&nbsp;&nbsp;" & objRSInt("Int_CorDestino")
									objRSInt.movenext
								Wend
							End if
							strRet = strRet & "<td>&nbsp;" & strObs & "</td>"
							strRet = strRet & "</tr>"
						Else

							if  Not objDic.Exists(strRepresentacao) then
								Call objDic.Add(strRepresentacao,objRSFac("Fac_Id")) 
								if intIndex = 1 then
									strRet = strRet & "<tr class=clsSilver2>"
									strRet = strRet & "<td nowrap>&nbsp;Nº Acso Pta EBT</td>"
									strRet = strRet & "<td nowrap>&nbsp;Nº Acso Pta CLI</td>"
									strRet = strRet & "<td nowrap>&nbsp;Nº Acso CCTO Prov</td>"
									strRet = strRet & "<td width=70>&nbsp;Tronco</td>"
									strRet = strRet & "<td width=50>&nbsp;Par</td>"
									strRet = strRet & "<td width=50>&nbsp;PADE/PAC</td>"
									strRet = strRet & "<td width=320>&nbsp;Observação</td>"
									strRet = strRet & "</td>"
									strRet = strRet & "</tr>"
								End if
								strRet = strRet & "<tr class=clsSilver>"
								strRet = strRet & "<td >&nbsp;" & strNroAcessoPtaEbt & "</td>"
								strRet = strRet & "<td >&nbsp;" & strNroAcessoPtaCli & "</td>"
								strRet = strRet & "<td >&nbsp;" & strNroAcessoCCTOPRov & "</td>"
								strRet = strRet & "<td >&nbsp;" & strTronco & "</td>"
								strRet = strRet & "<td nowrap>&nbsp;" & strRepresentacao & "</td>"
								strRet = strRet & "<td nowrap>&nbsp;"
								If objRSFac("Int_ID") <> "" and Not isnull(objRSFac("Int_ID")) then
									Set objRSInt = db.execute("CLA_sp_sel_interligacao " & objRSFac("Int_ID"))
									While not objRSInt.eof 
										strRet = strRet & objRSInt("Int_CorOrigem") & "&nbsp;&nbsp;>&nbsp;&nbsp;" & objRSInt("Int_CorDestino")
										objRSInt.movenext
									Wend
								End if
								strRet = strRet & "<td>&nbsp;" & strObs & "</td>"
								strRet = strRet & "</tr>"

							End if	
						End if
					Case 3

						If objRSFac("Int_ID") <> "" and Not isnull(objRSFac("Int_ID")) then
							if intIndex = 1 then
								strRet = strRet & "<tr class=clsSilver2>"
								strRet = strRet & "<td nowrap>&nbsp;Nº do Acesso</td>"
								strRet = strRet & "<td width=70>Cabo</td>"
								strRet = strRet & "<td width=50>Par</td>"
								strRet = strRet & "<td nowrap>PADE</td>"
								strRet = strRet & "<td nowrap>PADE/PAC</td>"
								strRet = strRet & "<td>Derivação</td>"
								strRet = strRet & "<td nowrap>Tipo Cabo</td>"
								strRet = strRet & "<td width=200>Obs</td>"
								strRet = strRet & "</tr>"
							End if
							strRet = strRet & "<tr class=clsSilver>"
							strRet = strRet & "<td >&nbsp;" & strNroAcessoPtaEbt & "</td>"
							strRet = strRet & "<td >&nbsp;" & strTronco & "</td>"
							strRet = strRet & "<td nowrap>&nbsp;" & strChave & "</td>"
							strRet = strRet & "<td >&nbsp;" & strCxEmenda & "</td>"
							strRet = strRet & "<td nowrap>&nbsp;"
							If objRSFac("Int_ID") <> "" and Not isnull(objRSFac("Int_ID")) then
								Set objRSInt = db.execute("CLA_sp_sel_interligacao " & objRSFac("Int_ID"))
								If not objRSInt.eof then
									strRet = strRet & objRSInt("Int_CorOrigem") & "&nbsp;&nbsp;>&nbsp;&nbsp;" & objRSInt("Int_CorDestino")
									objRSInt.movenext
								End if
							End if
							strRet = strRet & "<td >&nbsp;" & strLateral & "</td>"
							strRet = strRet & "<td >&nbsp;" & strTipoCabo & "</td>"
							strRet = strRet & "<td>&nbsp;" & strObs & "</td>"
							strRet = strRet & "</tr>"
						Else
							if  Not objDic.Exists(strRepresentacao) then
								Call objDic.Add(strRepresentacao,objRSFac("Fac_Id")) 
									if intIndex = 1 then
										strRet = strRet & "<tr class=clsSilver2>"
										strRet = strRet & "<td nowrap>&nbsp;Nº do Acesso</td>"
										strRet = strRet & "<td width=70>Cabo</td>"
										strRet = strRet & "<td width=50>Par</td>"
										strRet = strRet & "<td nowrap>PADE</td>"
										strRet = strRet & "<td nowrap>PADE/PAC</td>"
										strRet = strRet & "<td>Derivação</td>"
										strRet = strRet & "<td nowrap>Tipo Cabo</td>"
										strRet = strRet & "<td width=200>Obs</td>"
										strRet = strRet & "</tr>"
									End if
									strRet = strRet & "<tr class=clsSilver>"
									strRet = strRet & "<td >&nbsp;" & strNroAcessoPtaEbt & "</td>"
									strRet = strRet & "<td >&nbsp;" & strTronco & "</td>"
									strRet = strRet & "<td nowrap>&nbsp;" & strRepresentacao & "</td>"
									strRet = strRet & "<td >&nbsp;" & strCxEmenda & "</td>"
									strRet = strRet & "<td nowrap>&nbsp;"
									If objRSFac("Int_ID") <> "" and Not isnull(objRSFac("Int_ID")) then
										Set objRSInt = db.execute("CLA_sp_sel_interligacao " & objRSFac("Int_ID"))
										If not objRSInt.eof then
											strRet = strRet & objRSInt("Int_CorOrigem") & "&nbsp;&nbsp;>&nbsp;&nbsp;" & objRSInt("Int_CorDestino")
											objRSInt.movenext
										End if
									End if
									strRet = strRet & "<td >&nbsp;" & strLateral & "</td>"
									strRet = strRet & "<td >&nbsp;" & strTipoCabo & "</td>"
									strRet = strRet & "<td>&nbsp;" & strObs & "</td>"
									strRet = strRet & "</tr>"
								End if	
							End if	
				End Select		
			'End if
			intIndex = intIndex + 1
			objRSFac.MoveNext
		Wend
	End if
	if trim(strRet) <> "" then
		strRet = strRet & "</table>"
	End if	
	FacilidadesPedido = strRet
End Function
%>