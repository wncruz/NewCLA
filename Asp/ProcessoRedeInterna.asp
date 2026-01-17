<!--#include file="../inc/data.asp"-->
<%
Function GravarRedeInterna()

	Dim Int_ID(4)

	Set Rec = db.execute("CLA_sp_check_recurso2 " & request("cboLocalInstala") & ", " & request("cboDistLocalInstala") & ", " & request("cboProvedor") & ", " & request("cboSistema") & ",null,0")
	DBAction = rec("ret")

	if DBAction = 0 then

		rec_id = rec("Rec_ID")
		Set rec = db.execute("CLA_sp_view_recurso " & rec_id)
		
		dblEscId = rec("Esc_ID") 
		dblDstId = rec("Dst_ID") 

		strDtInicioTemp = inverte_data(request("txtDtInicioTemp"))
		strDtFimTemp = inverte_data(request("txtDtFimTemp"))
		strDtPedido = inverte_data(request("hdnDtPedido"))
			
		For i = 1 to 4
			If request("txtCoordenada"&i) <> "" then
				Set rs = db.execute("CLA_sp_sel_interorigem '" & request("txtCoordenada"&i) & "'," & dblEscId & "," & dblDstId & ",null")
				DBAction = rs("ret")
				If DBAction = 100 and request("hdnRdiId") <> "" then
					if cdbl(rs("rdi_id")) = cdbl(request("hdnRdiId")) then
						DBAction = 0
					end if
				End if
				If DBAction > 0 then
					set rs = nothing
					Select Case DBAction
						Case 80
							Response.Write "<script language=javascript>alert('A " & i & "º Coordenada já esta sendo utilizada.');</script>"	
						Case 100
							Response.Write "<script language=javascript>alert('A " & i & "º Coordenada já está sendo utilizada em um pedido de rede interna.');</script>"	
						Case Else
							Response.Write "<script language=javascript>parent.resposta('"&DBAction&"','');</script>"	
					End Select
					Response.End 
				End if
				Int_ID(i) = rs("Int_ID")
			Else
				set rs = nothing
			End if
		Next

		if DBAction = 0 then

			Vetor_Campos(1)="adInteger,4,adParamInput," & request("hdnRdiId")
			Vetor_Campos(2)="adInteger,4,adParamInput," & Int_ID(1)
			Vetor_Campos(3)="adInteger,4,adParamInput," & Int_ID(2)
			Vetor_Campos(4)="adInteger,4,adParamInput," & Int_ID(3)
			Vetor_Campos(5)="adInteger,4,adParamInput," & Int_ID(4)
			Vetor_Campos(6)="adWChar,25,adParamInput," & strDtPedido
			Vetor_Campos(7)="adWChar,50,adParamInput," & request("txtNomeCli")
			Vetor_Campos(8)="adWChar,60,adParamInput," & request("txtEndereco")
			Vetor_Campos(9)="adWChar,30,adParamInput," & request("txtDesignacao")
			Vetor_Campos(10)="adWChar,50,adParamInput," & request("txtNroContrato")
			Vetor_Campos(11)="adWChar,25,adParamInput," & strDtInicioTemp
			Vetor_Campos(12)="adWChar,25,adParamInput," & strDtFimTemp
			Vetor_Campos(13)="adWChar,300,adParamInput," & request("txtObs")
			Vetor_Campos(14)="adInteger,4,adParamInput," & rec_id
			Call APENDA_PARAM("CLA_sp_ins_redeinterna",14,Vetor_Campos)

			Set objRS = ObjCmd.Execute
			DBAction = objRS("Ret")
			if DBAction = 1 then
				Response.Write "<script language=javascript>parent.resposta('" & DBAction& "','');parent.document.forms[0].hdnRdiId.value="&objRS("Rdi_Id")&";</script>"
			Else
				Response.Write "<script language=javascript>parent.resposta('" & DBAction& "','');</script>"
			End if	
		End if
	Else	
		Response.Write "<script language=javascript>parent.resposta('" & DBAction& "','');</script>"
	End if
End Function

Function EditarRdi(intRdiId)

	Vetor_Campos(1)="adInteger,4,adParamInput," & intRdiId
	Call APENDA_PARAM("CLA_sp_sel_redeinterna",1,Vetor_Campos)

	Set objRS = ObjCmd.Execute
	If not objRS.Eof and Not objRS.Bof then

		if not isnull(objRS("Esc_ID")) then
			set objRSDst = db.execute("CLA_sp_view_recursodistribuicao " & objRS("Esc_ID"))
			dblDstId = objRS("Dst_ID")

			strCboRet = "<Select name=cboDistLocalInstala style=""width:200px"">"
			strCboRet = strCboRet & "<Option value=""""></Option>"
			While Not objRSDst.Eof
				if Trim(objRSDst("Dst_ID")) = Trim(dblDstId) then strSel = " selected "
				strCboRet = strCboRet & "<Option value=""" & Trim(objRSDst("Dst_ID")) & """" & strSel & ">" & objRSDst("Dst_Desc") & "</Option>"
				strSel = ""
				objRSDst.MoveNext
			Wend
			strCboRet = strCboRet & "</Select>"
			'Distribuidor da estação de entrega setado
			Response.Write "<script language=javascript>parent.spnDistLocalInstala.innerHTML = '" & TratarAspasJS(strCboRet) & "'</script>"
		Else
			strCboRet = "<Select name=cboDistLocalInstala style=""width:200px"">"
			strCboRet = strCboRet & "<Option value=""""></Option>"
			strCboRet = strCboRet & "</Select>"
			'Distribuidor da estação de entrega setado
			Response.Write "<script language=javascript>parent.spnDistLocalInstala.innerHTML = '" & TratarAspasJS(strCboRet) & "'</script>"
		End if


		strHtmlRet= "<script language=javascript>with(parent.document.forms[0]){" 
		strHtmlRet = strHtmlRet & "cboLocalInstala.value = """ & objRS("Esc_ID") & """;"
		strHtmlRet = strHtmlRet & "cboProvedor.value = """ & objRS("Pro_ID") & """;" 
		strHtmlRet = strHtmlRet & "hdnDstId.value = """ & objRS("Dst_ID") & """;" 
		'strHtmlRet = strHtmlRet & "parent.ResgatarDistribuicao(document.forms[0].cboLocalInstala,0);" 
		strHtmlRet = strHtmlRet & "txtNomeCli.value = """ & objRS("Rdi_CliNome") & """;" 
		strHtmlRet = strHtmlRet & "txtEndereco.value = """ & objRS("Rdi_Endereco") & """;" 
		strHtmlRet = strHtmlRet & "txtDtInicioTemp.value = """ & objRS("Rdi_DtInicio") & """;" 
		strHtmlRet = strHtmlRet & "txtDtFimTemp.value = """ & objRS("Rdi_DtTermino") & """;" 
		strHtmlRet = strHtmlRet & "txtDesignacao.value = """ & objRS("Rdi_Designacao") & """;" 
		strHtmlRet = strHtmlRet & "txtNroContrato.value = """ & objRS("Rdi_NroContrato") & """;" 
		strHtmlRet = strHtmlRet & "txtObs.value = """ & objRS("Rdi_Obs") & """;" 
		if not isNull(objRS("Rdi_DataPed")) then
			strHtmlRet = strHtmlRet & "hdnDtPedido.value = """ & split(objRS("Rdi_DataPed")," ")(0) & """;" 
		End if	
		strHtmlRet = strHtmlRet & "txtCoordenada1.value = """ & objRS("Int_CorOrigem1") & """;" 
		strHtmlRet = strHtmlRet & "txtCoordenada2.value = """ & objRS("Int_CorOrigem2") & """;" 
		strHtmlRet = strHtmlRet & "txtCoordenada3.value = """ & objRS("Int_CorOrigem3") & """;" 
		strHtmlRet = strHtmlRet & "txtCoordenada4.value = """ & objRS("Int_CorOrigem4") & """;" 
		if not isNull(objRS("Int_CorDestino1")) then strHtmlRet = strHtmlRet & "parent.spnCoordenada1.innerHTML = """ & objRS("Int_CorOrigem1") & " > " & objRS("Int_CorDestino1") & """;" Else strHtmlRet = strHtmlRet & "parent.spnCoordenada1.innerHTML = """";" End if
		if not isNull(objRS("Int_CorDestino2")) then strHtmlRet = strHtmlRet & "parent.spnCoordenada2.innerHTML = """ & objRS("Int_CorOrigem2") & " > " & objRS("Int_CorDestino2") & """;" Else strHtmlRet = strHtmlRet & "parent.spnCoordenada2.innerHTML = """";" End if
		if not isNull(objRS("Int_CorDestino3")) then strHtmlRet = strHtmlRet & "parent.spnCoordenada3.innerHTML = """ & objRS("Int_CorOrigem3") & " > " & objRS("Int_CorDestino3") & """;" Else strHtmlRet = strHtmlRet & "parent.spnCoordenada3.innerHTML = """";" End if
		if not isNull(objRS("Int_CorDestino4")) then strHtmlRet = strHtmlRet & "parent.spnCoordenada4.innerHTML = """ & objRS("Int_CorOrigem4") & " > " & objRS("Int_CorDestino4") & """;" Else strHtmlRet = strHtmlRet & "parent.spnCoordenada4.innerHTML = """";" End if
		strHtmlRet = strHtmlRet & "parent.spnDtPedido.innerHTML = """ & Formatar_Data(objRS("Rdi_DataPed")) & """;" 

		strHtmlRet = strHtmlRet & "}</script>"
		Response.Write strHtmlRet
	Else
		Response.Write "<script laguage=javascript>alert('Rede Interna não encontrada.')</script>"	
	End if
End Function

Function RemoverRdi()

	Vetor_Campos(1)="adInteger,4,adParamInput," & request("hdnRdiId")
	Call APENDA_PARAM("CLA_sp_del_redeinterna",1,Vetor_Campos)
	Set objRS = ObjCmd.Execute
	if objRS("Ret") > 0 then
		Response.Write "<script laguage=javascript>parent.resposta(3,'');</script>"	
	Else
		Response.Write "<script laguage=javascript>alert('Rede Interna não encontrada.')</script>"	
	End if

End Function

Select Case Trim(Request.Form("hdnAcao"))

	Case "ListaInterligacaoLivre"
		Set objRSRec = db.execute("CLA_sp_check_recurso2 " & request("cboLocalInstala") & ", " & request("cboDistLocalInstala") & ", " & request("cboProvedor") & ", " & request("cboSistema") & ",null,0")
		Set DBAction = objRSRec("ret")

		If DBAction <> 0 then
			Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
		Else
			intRecId = objRSRec("Rec_ID")
			Response.Write "<script language=javascript>parent.PosicoesLivre('L'," & intRecId & ");</script>"
		End if
	Case "GravarRedeInterna"
		Call GravarRedeInterna()
	Case "EditarRdi"
		EditarRdi(Request.Form("hdnRdiId"))
	Case "RemoverRdi"
		RemoverRdi()

End Select  
%>