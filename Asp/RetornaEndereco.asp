<!--#include file="../inc/data.asp"-->
<%

'Response.End
	dim objXmlDoc 
	dim objRSCid,ndUF,ndCidsigla,sRetorno
	
	Dim strCboRet
	Dim intCount
  Dim strCep 
  Dim sTipoCEP
  
  strCep   = Request("hdnCEP")
  sTipoCEP = Request("hdnTipoCEP")

 	If sTipoCEP = "1" Then 	 	
		Set objRS = db.execute("CLA_SP_VIEW_CEP '" & strCep & "'")
	Else
		Set objRS = db.execute("CLA_SP_VIEW_CEP null," & strCep)		
	End If

	If Not objRS.eof and  Not objRS.bof then
		intCount = 0
		strCboRet = "<Select name=cboCEPS onChange=""ProcurarCEPX(2)"">"
		strCboRet = strCboRet & "<Option value="""">Selecione um CEP</Option>"
		While Not objRS.Eof
			strCboRet = strCboRet & "<Option value=" & Trim(objRS("Cep_ID")) & ">" & TratarAspasJS(Trim(objRS("RuaCompleta"))) & " - " & TratarAspasJS(Trim(objRS("Cep"))) & "</Option>"			
			objRS.MoveNext
			intCount = intCount + 1
		Wend
		strCboRet = strCboRet & "</Select>"
		if intCount > 1 then 'Retorna um combo com os CEPS encontrados
					Response.Write "<script language=javascript>parent.spnCEPS.innerHTML = '" & strCboRet & "'</script>"
		Else
			objRS.MoveFirst
						Response.Write "<script language=javascript>with (parent.document.forms[0]){" & _
						"cboLogrEnd.value ='" & TratarAspasJS(Trim(objRS("Logradouro"))) & "';" & _
						"txtEnd.value ='" & Trim(TratarAspasJS(Trim(objRS("Titulo"))) & " " & Trim(TratarAspasJS(Trim(objRS("Preposicao"))) & " " & TratarAspasJS(Trim(objRS("Rua"))))) & "';" & _
						"cboUFEnd.value ='" & TratarAspasJS(Trim(objRS("Est_Sigla"))) & "';" & _
						"txtBairroEnd.value ='" & TratarAspasJS(Trim(objRS("BairroInicial"))) & "';" & _
						"txtEndCid.value ='" & TratarAspasJS(Trim(objRS("Cid_Sigla"))) & "';" & _
						"txtCepEnd.value ='" & TratarAspasJS(Trim(objRS("Cep"))) & "';" & _
						"txtEndCidDesc.value ='" & TratarAspasJS(Trim(objRS("Cid_Desc"))) & "';" & _
						"}</script>"

						Response.Write "<script language=javascript>parent.spnCEPS.innerHTML = ''</script>"
		End If
	Else
		Response.Write "<script language=javascript>alert('CEP não encontrado.')</script>"
	End if
%>