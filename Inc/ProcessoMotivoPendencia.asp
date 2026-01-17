<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoMotivoPendencia.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Grava motivos de pendência (Status detalhado)
%>
<!--#include file="../inc/data.asp"-->
<%
Dim strAcao
strAcao = Trim(Request.Form("hdnAcao"))
if strAcao = "" then strAcao = Trim(Request.QueryString("strAcao")) End if

Function ResgatarLista(dblSolId,dblPedId)

	Dim strHtml
	Dim blnCor

	strHtml = strHtml &  "<Html><Body topmargin=0 leftmargin=0 bgcolor=#eeeeee class=TA>"
	strHtml = strHtml &  "<Head><link rel=stylesheet type='text/css' href='../css/cla.css'></Head>"
	strHtml = strHtml &  "<table cellpadding=0 cellspacing=1 width=100% >"
	strHtml = strHtml &  "<tr>"
	strHtml = strHtml &  "	<th nowrap>&nbsp;Data</th>"
	strHtml = strHtml &  "	<th>&nbsp;Dias</th>"
	strHtml = strHtml &  "	<th>&nbsp;Freq</th>"
	strHtml = strHtml &  "	<th>&nbsp;Pedido</th>"
	
	strHtml = strHtml &  "	<th>&nbsp;Tecnologia</th>"
	
	strHtml = strHtml &  "	<th>&nbsp;Usuario</th>"
	strHtml = strHtml &  "	<th>&nbsp;Status</th>"
	strHtml = strHtml &  "	<th>&nbsp;Historico</th>"
	strHtml = strHtml &  "</tr>"

	Vetor_Campos(1)="adInteger,2,adParamInput,"
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblSolId
	Vetor_Campos(3)="adInteger,2,adParamInput,1" '->0-Macro, 1-Detalhado, 2-Ultimo Macro, 3-Ultimo Detalhado
	Vetor_Campos(4)="adInteger,2,adParamInput," & dblPedId

	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_StatusSolicitacao",4,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)

	blnCor = true
	While Not objRS.Eof
		if blnCor then
			strHtml = strHtml &  "<tr class=clsSilver >"
			blnCor = false
		Else
			strHtml = strHtml &  "<tr class=clsSilver2>"
			blnCor = true
		End if	
		strHtml = strHtml &  "<td width=15% nowrap >&nbsp;"& Formatar_Data(objRS("StsSol_Data")) &"</td>"		
		strHtml = strHtml &  "<td >&nbsp;" & objRS("stssol_cont") & "</td>"
		strHtml = strHtml &  "<td >&nbsp;" & objRS("stssol_freq") & "</td>"
		strHtml = strHtml &  "<td nowrap >&nbsp;" & objRS("Pedido") & "</td>"
		
		strHtml = strHtml &  "<td nowrap >&nbsp;" & objRS("Tecnologia") & "</td>"
		
		strHtml = strHtml &  "<td width=20% >&nbsp;" & objRS("Usu_UserName") & "</td>"
		strHtml = strHtml &  "<td width=30% >&nbsp;"& objRS("Sts_Desc") & "</td>"
		strHtml = strHtml &  "<td width=35% >&nbsp;"& objRS("StsSol_Historico") &"</td>"
		strHtml = strHtml &  "</tr>"
		objRS.MoveNext
	Wend

	strHtml = strHtml &  "</table>"
	strHtml = strHtml &  "</Html></Body>"
	Response.Write strHtml

End Function

Select Case strAcao
	
	Case "ResgatarLista"
		Call ResgatarLista(Request.QueryString("dblSolId"),Request.QueryString("dblPedId"))

	Case "GravarMotivoPendencia"
		'Gravar Motivo de pendências
		dblSolId = Request.Form("hdnSolId")
		dblPedId = Request.Form("hdnPedId")
		dblAcfId     = Request.Form("hdnAcfId")
		gravarDireto = Request.Form("gravarDireto")
		
		if gravarDireto = "0" and Request.Form("hdnOrigem") <> "Lista" then 'Verifica se
			Set objRSPed = db.Execute("CLA_Sp_Sel_PedidoHistoricoSolicitacao " & dblSolId)
			If not objRSPed.Eof and Not objRSPed.Bof then
				intCount = 0
				While Not objRSPed.Eof
					dblPedId = objRSPed("Ped_Id")
					objRSPed.MoveNext
					intCount= intCount + 1
				Wend
				if intCount > 1 then ' GIC esta alterando status para mais de um pedido
					Response.Write "<script language=javascript>parent.ListarPedidosSolicitacao('" & dblSolId & "')</script>"
					Response.End 
				End if
			End if
		End if

		if Request.Form("hdnOrigem") = "Lista" then

			For Each dblAcfId in Request.Form("chkPedido") 

				Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId 
				Vetor_Campos(2)="adInteger,2,adParamInput," & Request.Form("cboStatusSolic")
				Vetor_Campos(3)="adWChar,30,adParamInput,"  & strUserName
				Vetor_Campos(4)="adDate,10,adParamInput,"
				Vetor_Campos(5)="adWChar,300,adParamInput," & Request.Form("txtMotivo")
				Vetor_Campos(6)="adWChar,1,adParamInput,D" 'Status detalhado
				Vetor_Campos(7)="adInteger,2,adParamOutput,0"
				Vetor_Campos(8)="adInteger,2,adParamInput,"
				Vetor_Campos(9)="adWChar,1,adParamInput,"
				Vetor_Campos(10)="adInteger,2,adParamInput," 
				Vetor_Campos(11)="adInteger,2,adParamInput," & dblAcfId
				

				Call APENDA_PARAM("CLA_sp_ins_StatusSolicitacao",11,Vetor_Campos)

				ObjCmd.Execute'pega dbaction
				DBAction = ObjCmd.Parameters("RET").value
				Call EnviarEmailAlteracaoStatus(Request.Form("hdnSolId"),Request.Form("cboStatusSolic"),Request.Form("txtMotivo"))

			Next

			if DBAction = 1 then
				Response.Write "<script language=javascript>alert('Motivo da pendência gravado com sucesso.')</script>"
				Response.Write "<script language=javascript>parent.AtualizarListaMotivo();window.close();</script>"
			Else
				Response.Write "<script language=javascript>alert('Não foi possível gravar o motivo da pendência.')</script>"
			End if
		Else

			Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId 
			Vetor_Campos(2)="adInteger,2,adParamInput," & Request.Form("cboStatusSolic")
			Vetor_Campos(3)="adWChar,30,adParamInput,"  & strUserName
			Vetor_Campos(4)="adDate,10,adParamInput,"
			Vetor_Campos(5)="adWChar,300,adParamInput," & Request.Form("txtMotivo")
			Vetor_Campos(6)="adWChar,1,adParamInput,D" 'Status detalhado
			Vetor_Campos(7)="adInteger,2,adParamOutput,0"
			Vetor_Campos(8)="adInteger,2,adParamInput," & dblPedId
			Vetor_Campos(9)="adWChar,1,adParamInput,"
			Vetor_Campos(10)="adInteger,2,adParamInput," 
			Vetor_Campos(11)="adInteger,2,adParamInput," & dblAcfId


			Call APENDA_PARAM("CLA_sp_ins_StatusSolicitacao",11,Vetor_Campos)

			ObjCmd.Execute'pega dbaction
			DBAction = ObjCmd.Parameters("RET").value

			if DBAction = 1 then
				Response.Write "<script language=javascript>alert('Motivo da pendência gravado com sucesso.')</script>"
				Response.Write "<script language=javascript>parent.IFrmLista.location.href = 'ProcessoMotivoPendencia.asp?strAcao=ResgatarLista&dblSolId=" & Request.Form("hdnSolId") & "&dblPedId=" & Request.Form("hdnPedId") & "'</script>"
				Call EnviarEmailAlteracaoStatus(Request.Form("hdnSolId"),Request.Form("cboStatusSolic"),Request.Form("txtMotivo"))
			Else
				Response.Write "<script language=javascript>alert('Não foi possível gravar o motivo da pendência.')</script>"
			End if
		End if
End Select
%>
