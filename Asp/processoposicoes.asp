<!--#include file="../inc/data.asp"-->
<%
Function ResgatarPosicoesLivres(dblRecId,intRede,intWidthTable,blnLink)

	Set rs = db.execute("CLA_SP_CONS_POSICOES_LIVRES " & dblRecId & ",'" & request("txtTronco") & "','"& request("txtPar") & "','"& request("txtBastidor") & "','"& request("txtRegua") & "','"& request("txtPorta") & "','"& request("txtTimeslot") & "','"& request("txtDominio") & "','"& request("txtNO") & "'," & request("txtNroReg") & ",'" & Request("cboTipoCabo") & "','" & Request("txtLateral") & "','" & Request("txtCaixaEmenda") & "'")
	if not rs.Eof and Not rs.Bof then

		Select Case intRede

			Case 1	'DET

				strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=""" & intWidthTable & "px"">"
				strRet = strRet  & "<tr>"
				strRet = strRet  & "<th>&nbsp;Bastidor</th>"
				strRet = strRet  & "<th>&nbsp;Régua</th>"
				strRet = strRet  & "<th>&nbsp;Porta</th>"
				strRet = strRet  & "<th>&nbsp;TimeSlot</th>"
				strRet = strRet  & "<th>&nbsp;Domínio</th>"
				strRet = strRet  & "<th>&nbsp;Nó</th>"
				strRet = strRet  & "</tr>"

				cores = "#dddddd"
				blnDet = false	
				While Not rs.eof
						
					if not isNull(rs("Fac_Bastidor")) then
						if cores = "#dddddd" then cores = "#eeeeee"	else cores = "#dddddd"	end if
						valores = "\'" & Trim(rs("Fac_Bastidor")) & "\'"
						valores = valores & ",\'" & Trim(rs("Fac_Regua")) & "\'"
						valores = valores & ",\'" & Trim(rs("Fac_Porta")) & "\'"
						valores = valores & ",\'" & Trim(rs("Fac_Timeslot")) & "\'"
						valores = valores & ",\'" & Trim(rs("Fac_Dominio")) & "\'"
						valores = valores & ",\'" & Trim(rs("Fac_No")) & "\'"

						strRet = strRet  & "<tr>"
						if blnLink then
							strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;<a href=\'#\' onClick=""javascript:copyPaste(" & valores & ")"">" & rs("Fac_Bastidor") & "</a></td>"
						Else
							strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Bastidor") & "</td>"
						End if	
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Regua") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Porta") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Timeslot") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Dominio") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_No") & "</td>"
						strRet = strRet  & "</tr>"
						blnDet = true	
					End if	

					rs.movenext

				Wend

				strRet = strRet  & "</table>"
				if blnDet then
					Response.Write "<script language=javascript>parent.spnPosicoes.innerHTML = '" & strRet &"'</script>"
				Else
					Response.Write "<script language=javascript>alert('Facilidade não encontrada.')</script>"
				End if	
				
			Case 2	'NDET

				strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=""" & intWidthTable & "px"">"
				strRet = strRet  & "<tr>"
				strRet = strRet  & "<th>&nbsp;Tronco</th>"
				strRet = strRet  & "<th>&nbsp;Par</th>"
				strRet = strRet  & "</tr>"

				cores = "#dddddd"
				blnNDet = false
				While Not rs.eof

					if not isNull(rs("Fac_Tronco")) then
						if cores = "#dddddd" then cores = "#eeeeee"	else cores = "#dddddd"	end if
						valores = "\'" & Trim(rs("Fac_Tronco")) & "\'"
						valores = valores & ",\'" & Trim(rs("Fac_Par")) & "\'"
						strRet = strRet  & "<tr>"
						if blnLink then
							strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;<a href=\'#\' onClick=""javascript:copyPaste(" & valores & ");"">" & rs("Fac_Tronco") & "</a></td>"
						Else	
							strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Tronco") & "</td>"
						End if	
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Par") & "</td>"
						strRet = strRet  & "</tr>"
						blnNDet = true
					End if	

					rs.MoveNext
				Wend	

				strRet = strRet  & "</table>"
				if blnNDet then
					Response.Write "<script language=javascript>parent.spnPosicoes.innerHTML = '" & strRet &"'</script>"
				Else
					Response.Write "<script language=javascript>alert('Facilidade não encontrada.')</script>"
				End if	

			Case 3 'ADE

				strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=""" & intWidthTable & "px"">"
				strRet = strRet  & "<tr>"
				strRet = strRet  & "<th>&nbsp;Cabo</th>"
				strRet = strRet  & "<th>&nbsp;Par</th>"
				strRet = strRet  & "<th>&nbsp;Derivação</th>"
				strRet = strRet  & "<th>&nbsp;Tipo Cabo</th>"
				strRet = strRet  & "<th>&nbsp;PADE</th>"
				strRet = strRet  & "</tr>"

				cores = "#dddddd"
				blnAde = false						
				While Not rs.eof
						
					if not isNull(rs("Fac_Tronco"))then
						if cores = "#dddddd" then cores = "#eeeeee"	else cores = "#dddddd"	end if
						valores = ""
						valores = valores & "\'" & Trim(rs("Fac_Tronco")) & "\'"
						valores = valores & ",\'" & Trim(rs("Fac_Par")) & "\'"
						valores = valores & ",\'" & Trim(rs("Fac_Lateral")) & "\'"
						valores = valores & ",\'" & Trim(rs("Fac_TipoCabo")) & "\'"
						valores = valores & ",\'" & Trim(rs("Fac_CxEmenda")) & "\'"

						strRet = strRet  & "<tr>"
						if blnLink then
							strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;<a href=\'#\' onClick=""javascript:copyPaste(" & valores & ")"">" & rs("Fac_Tronco") & "</a></td>"
						Else	
							strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Tronco") & "</td>"
						End if	
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Par") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Lateral") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_TipoCabo") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_CxEmenda") & "</td>"
						strRet = strRet  & "</tr>"
						blnAde = true
					End if	
					rs.movenext

				Wend
					
				strRet = strRet  & "</table>"
				if blnAde then
					Response.Write "<script language=javascript>parent.spnPosicoes.innerHTML = '" & strRet &"'</script>"
				Else	
					Response.Write "<script language=javascript>alert('Facilidade não encontrada.')</script>"
				End if	

		End Select
	Else
		Response.Write "<script language=javascript>alert('Facilidade não encontrada.')</script>"
	End if

End Function

Function ResgatarPosicoesOcupadas(dblRecId,intRede,intWidthTable,blnLink)
	
	'Response.Write ("CLA_SP_CONS_POSICOES_OCUPADAS " & dblRecId & ",'" & request("txtTronco") & "','"& request("txtPar") & "','"& request("txtBastidor") & "','"& request("txtRegua") & "','"& request("txtPorta") & "','"& request("txtTimeslot") & "','"& request("txtDominio") & "','"& request("txtNO") & "','" & Request("cboTipoCabo") & "','" & Request("txtLateral") & "','" & Request("txtCaixaEmenda") & "'") & "<br>"
	Set rs = db.execute("CLA_SP_CONS_POSICOES_OCUPADAS " & dblRecId & ",'" & request("txtTronco") & "','"& request("txtPar") & "','"& request("txtBastidor") & "','"& request("txtRegua") & "','"& request("txtPorta") & "','"& request("txtTimeslot") & "','"& request("txtDominio") & "','"& request("txtNO") & "','" & Request("cboTipoCabo") & "','" & Request("txtLateral") & "','" & Request("txtCaixaEmenda") & "'")
	if not rs.Eof and Not rs.Bof then

		Select Case intRede

			Case 1	'DET

				strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=""" & intWidthTable & "px"">"
				strRet = strRet  & "<tr>"
				strRet = strRet  & "<th>&nbsp;Bastidor</th>"
				strRet = strRet  & "<th>&nbsp;Régua</th>"
				strRet = strRet  & "<th>&nbsp;Porta</th>"
				strRet = strRet  & "<th>&nbsp;TimeSlot</th>"
				strRet = strRet  & "<th>&nbsp;Domínio</th>"
				strRet = strRet  & "<th>&nbsp;Nó</th>"
				strRet = strRet  & "<th>&nbsp;Nº Acesso</th>"
				strRet = strRet  & "<th>&nbsp;Pedido</th>"
				strRet = strRet  & "<th>&nbsp;Cliente</th>"

				strRet = strRet  & "</tr>"

				cores = "#dddddd"
				blnDet = false	
				While Not rs.eof
						
					if not isNull(rs("Fac_Bastidor")) then
						if cores = "#dddddd" then cores = "#eeeeee"	else cores = "#dddddd"	end if

						strRet = strRet  & "<tr>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Bastidor") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Regua") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Porta") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Timeslot") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Dominio") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_No") & "</td>"

						strRet = strRet  & "<td bgcolor="  & cores & ">&nbsp;" & rs("Fac_NroAcesso") & "</td>"
						if blnLink then
							strRet = strRet  & "<td bgcolor="  & cores & " nowrap>&nbsp;<a href=\'#\' onClick=\'javascript:DetalharSolicitacao(" & RS("Sol_ID") & ")\' >"  & ucase(rs("Acp_Prefixo")) & "-" & right("00000" & rs("Acp_Numero"), 5) & "/" & rs("Acp_Ano") & "</a></td>"
						Else
							strRet = strRet  & "<td bgcolor="  & cores & " nowrap>&nbsp;"  & ucase(rs("Acp_Prefixo")) & "-" & right("00000" & rs("Acp_Numero"), 5) & "/" & rs("Acp_Ano") & "</td>"
						End if	

						strRet = strRet  & "<td bgcolor="  & cores & ">&nbsp;" & trim(rs("Cli_Nome")) & "</td>"

						strRet = strRet  & "</tr>"
						blnDet = true	
					End if	

					rs.movenext

				Wend

				strRet = strRet  & "</table>"
				if blnDet then
					Response.Write "<script language=javascript>parent.spnPosicoes.innerHTML = '" & strRet &"'</script>"
				Else
					Response.Write "<script language=javascript>alert('Facilidade não encontrada.')</script>"
				End if	
				
			Case 2	'NDET

				strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=""" & intWidthTable & "px"">"
				strRet = strRet  & "<tr>"
				strRet = strRet  & "<th>&nbsp;Tronco</th>"
				strRet = strRet  & "<th>&nbsp;Par</th>"
				strRet = strRet  & "<th>&nbsp;Coordenada</th>"
				strRet = strRet  & "<th>&nbsp;Nº Acesso</th>"
				strRet = strRet  & "<th>&nbsp;Pedido</th>"
				strRet = strRet  & "<th>&nbsp;Cliente</th>"
				strRet = strRet  & "</tr>"

				cores = "#dddddd"
				blnNDet = false
				While Not rs.eof

					if not isNull(rs("Fac_Tronco")) then
						if cores = "#dddddd" then cores = "#eeeeee"	else cores = "#dddddd"	end if
						strRet = strRet  & "<tr>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Tronco") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Par") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & " nowrap>&nbsp;"
						if rs("Int_ID") <> "" and not isnull(rs("Int_ID")) then
							set inter = db.execute("CLA_sp_sel_interligacao " & rs("Int_ID"))
							if not inter.eof then
								strRet = strRet  & inter("Int_CorOrigem") & "&nbsp;><br>&nbsp;" & inter("Int_CorDestino")
								inter.movenext
							end if
						End if
						strRet = strRet  & "</td>"

						strRet = strRet  & "<td bgcolor="  & cores & ">&nbsp;" & rs("Fac_NroAcesso") & "</td>"
						if blnLink then
							strRet = strRet  & "<td bgcolor="  & cores & " nowrap>&nbsp;<a href=\'#\' onClick=\'javascript:DetalharSolicitacao(" & RS("Sol_ID") & ")\' >"  & ucase(rs("Acp_Prefixo")) & "-" & right("00000" & rs("Acp_Numero"), 5) & "/" & rs("Acp_Ano") & "</a></td>"
						Else
							strRet = strRet  & "<td bgcolor="  & cores & " nowrap>&nbsp;"  & ucase(rs("Acp_Prefixo")) & "-" & right("00000" & rs("Acp_Numero"), 5) & "/" & rs("Acp_Ano") & "</td>"
						End if	
						strRet = strRet  & "<td bgcolor="  & cores & ">&nbsp;" & trim(rs("Cli_Nome")) & "</td>"

						strRet = strRet  & "</tr>"

						blnNDet = true
					End if	

					rs.MoveNext
				Wend	

				strRet = strRet  & "</table>"
				if blnNDet then
					Response.Write "<script language=javascript>parent.spnPosicoes.innerHTML = '" & strRet &"'</script>"
				Else
					Response.Write "<script language=javascript>alert('Facilidade não encontrada.')</script>"
				End if	

			Case 3 'ADE

				strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=""" & intWidthTable & "px"">"
				strRet = strRet  & "<tr>"
				strRet = strRet  & "<th>&nbsp;Cabo</th>"
				strRet = strRet  & "<th>&nbsp;Par</th>"
				strRet = strRet  & "<th>&nbsp;PADE/PAC</th>"
				strRet = strRet  & "<th>&nbsp;Derivação</th>"
				strRet = strRet  & "<th>&nbsp;Tipo Cabo</th>"
				strRet = strRet  & "<th>&nbsp;PADE</th>"
				strRet = strRet  & "<th>&nbsp;Nº Acesso</th>"
				strRet = strRet  & "<th>&nbsp;Pedido</th>"
				strRet = strRet  & "<th>&nbsp;Cliente</th>"

				strRet = strRet  & "</tr>"

				cores = "#dddddd"
				blnAde = false						
				While Not rs.eof
						
					if not isNull(rs("Fac_Tronco"))then
						if cores = "#dddddd" then cores = "#eeeeee"	else cores = "#dddddd"	end if

						strRet = strRet  & "<tr>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Tronco") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Par") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & " nowrap>&nbsp;"
						if rs("Int_ID") <> "" and not isnull(rs("Int_ID")) then
							set inter = db.execute("CLA_sp_sel_interligacao " & rs("Int_ID"))
							if not inter.eof then
								strRet = strRet  & inter("Int_CorOrigem") & "&nbsp;><br>&nbsp;" & inter("Int_CorDestino")
								inter.movenext
							end if
						End if
						strRet = strRet  & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_Lateral") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_TipoCabo") & "</td>"
						strRet = strRet  & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_CxEmenda") & "</td>"

						strRet = strRet  & "<td bgcolor="  & cores & ">&nbsp;" & rs("Fac_NroAcesso") & "</td>"
						if blnLink then
							strRet = strRet  & "<td bgcolor="  & cores & " nowrap>&nbsp;<a href=\'#\' onClick=\'javascript:DetalharSolicitacao(" & RS("Sol_ID") & ")\' >"  & ucase(rs("Acp_Prefixo")) & "-" & right("00000" & rs("Acp_Numero"), 5) & "/" & rs("Acp_Ano") & "</a></td>"
						Else
							strRet = strRet  & "<td bgcolor="  & cores & " nowrap>&nbsp;"  & ucase(rs("Acp_Prefixo")) & "-" & right("00000" & rs("Acp_Numero"), 5) & "/" & rs("Acp_Ano") & "</td>"
						End if							
						strRet = strRet  & "<td bgcolor="  & cores & ">&nbsp;" & trim(rs("Cli_Nome")) & "</td>"

						strRet = strRet  & "</tr>"

						blnAde = true
					End if	
					rs.movenext

				Wend
					
				strRet = strRet  & "</table>"
				if blnAde then
					Response.Write "<script language=javascript>parent.spnPosicoes.innerHTML = '" & strRet &"'</script>"
				Else	
					Response.Write "<script language=javascript>alert('Facilidade não encontrada.')</script>"
				End if	

		End Select
	Else
		Response.Write "<script language=javascript>alert('Facilidade não encontrada.')</script>"
	End if


End Function

Function ConsultarCoordenadasOcupadas(strCoordenada,dblRecId)

	Set rs = db.execute("CLA_sp_sel_coordenadaocupada '" & strCoordenada & "'," & dblRecId)

	If not rs.eof then

		strHtml = ""
		strHtml = strHtml & "<table border=0 cellspacing=1 cellpadding=0 cellspacing=0 width=100% >"
		strHtml = strHtml & "<tr>"
		strHtml = strHtml & "<th>&nbsp;PADE/PAC</th>"
		strHtml = strHtml & "<th>&nbsp;Nro. Acesso</th>"
		strHtml = strHtml & "<th>&nbsp;Nro. Pedido</th>"
		strHtml = strHtml & "<th>&nbsp;Cliente</th>"
		strHtml = strHtml & "</tr>"

		strHtml = strHtml & "<tr>"
		cores = "#dddddd"
		do while not rs.eof
			if cores = "#dddddd" then
				cores = "#eeeeee"
			else
				cores = "#dddddd"
			end if
			strHtml = strHtml & "<tr>"
			strHtml = strHtml & "<td bgcolor=" & cores & ">&nbsp;" & rs("Int_CorOrigem") & " > " & rs("Int_CorDestino") & "</td>"
			strHtml = strHtml & "<td bgcolor=" & cores & ">&nbsp;" & rs("Fac_NroAcesso") & "</td>"
			strHtml = strHtml & "<td bgcolor=" & cores & ">&nbsp;" & ucase(rs("acp_prefixo")) & "-" & right("00000" & rs("acp_numero"),5) & "/" & rs("acp_ano") & "</td>"
			strHtml = strHtml & "<td bgcolor=" & cores & ">&nbsp;" & rs("cli_nome") & "</td>"
			strHtml = strHtml & "</tr>"
			rs.movenext
		loop
		strHtml = strHtml & "</tr>"
		strHtml = strHtml & "</table>"

		Response.Write "<script language=javascript>parent.spnConsulta.innerHTML = '" & strHtml & "'</script>"
		
	Else
		Response.Write "<script language=javascript>alert('Registro não encontrado')</script>"
	End if

End Function

Dim dblRecId
Dim intRede

Select Case Trim(Request.Form("hdnAcao"))

	Case "ResgatarPosicoesLivres"

		dblRecId = Request.Form("hdnRecId")
		intRede = Request.Form("hdnRede") 

		Call ResgatarPosicoesLivres(dblRecId,intRede,510,true)

	Case "ResgatarPosicoesOcupadas"

		dblRecId = Request.Form("hdnRecId")
		intRede = Request.Form("hdnRede") 

		Call ResgatarPosicoesOcupadas(dblRecId,intRede,510,true)

	Case "ConsultarPosicoesOcupadas"
		Call ConsultarCoordenadasOcupadas(Trim(Request.Form("coordenada")),Trim(Request.Form("rec_id")))

End Select 
%>
