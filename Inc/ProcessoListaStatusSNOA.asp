<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoMotivoPendencia.asp
'	- Responsável			: Vital
'	- Descrição			: Lista/Grava motivos de pendência (Status detalhado)
%>
<!--#include file="../inc/data.asp"-->
<%

Dim strAcao
strAcao = Trim(Request.Form("hdnAcao"))
if strAcao = "" then strAcao = Trim(Request.QueryString("strAcao")) End if

Function ResgatarLista(dblSolId)

	Dim strHtml
	Dim blnCor

	strHtml = strHtml &  "<Html><Body topmargin=0 leftmargin=0 bgcolor=#eeeeee class=TA>"
	strHtml = strHtml &  "<Head><link rel=stylesheet type='text/css' href='../css/cla.css'></Head>"
	strHtml = strHtml &  "<table cellpadding=0 cellspacing=1 width=100% >"
	strHtml = strHtml &  "<tr>"
	strHtml = strHtml &  "	<th nowrap>&nbsp;Data do Pedido</th>"
	strHtml = strHtml &  "	<th>&nbsp;Número SNOA</th>"
	strHtml = strHtml &  "	<th>&nbsp;Caracteristica EILD</th>"
	strHtml = strHtml &  "	<th>&nbsp;Status</th>"
	strHtml = strHtml &  "	<th>&nbsp;Motivo</th>"
	strHtml = strHtml &  "	<th >&nbsp;Historico SNOA</th>"
	strHtml = strHtml &  "</tr>"

	//dblSolId = 1178545
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_historicoSnoa ",1,Vetor_Campos)
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
		strHtml = strHtml &  "<td width=15% nowrap >&nbsp;"& Formatar_Data(objRS("DATA_REGISTRO")) &"</td>"		
		strHtml = strHtml &  "<td >&nbsp;" & objRS("N_SNOA") & "</td>"
		strHtml = strHtml &  "<td >&nbsp;" & objRS("CARACTERIZACAO_EILD") & "</td>"
		strHtml = strHtml &  "<td nowrap >&nbsp;" & objRS("STATUS_SNOA") & "</td>"
		strHtml = strHtml &  "<td >&nbsp;" & objRS("Motivo_SNOA") & "</td>"
		strHtml = strHtml &  "<td >&nbsp;" & objRS("Historico_SNOA") & "</td>"
		strHtml = strHtml &  "</tr>"
		objRS.MoveNext
	Wend

	strHtml = strHtml &  "</table>"
	strHtml = strHtml &  "</Html></Body>"
	Response.Write strHtml

End Function

Function ConclusaoSNOA(dblSolId)

	'//dblSolId = 1178545

	Dim Rec_IDOrigem
	Dim rec
	Dim Rec_IDDestino
	Dim Rec_ID

	if DBAction = 0 then

		strcboAceitarSNOA = Trim(Request.Form("cboAceitarSNOA"))
		'strMotivo = Request.Form("txtMotivo")
		'response.write "<script>alert('"& Request.Form("txtMotivo") &"')</script>"
	
		'response.write "<script>alert('ConclusaoSNOA')</script>"
		'response.write "<script>alert('"& dblSolId &"')</script>"
		'response.write "<script>alert('"& strcboAceitarSNOA &"')</script>"
		'response.write "<script>alert('"& Request.Form("txtMotivo") &"')</script>"
		
		Vetor_Campos(1)="adInteger,4,adParamInput," & dblSolId
		Vetor_Campos(2)="adWChar,50,adParamInput," & strcboAceitarSNOA
		Vetor_Campos(3)="adWChar,1000,adParamInput," & Request.Form("txtMotivo")

		Vetor_Campos(4)="adInteger,4,adParamOutput,0"
		Vetor_Campos(5)="adWChar,100,adParamOutput,"
		
		on error resume next
		Call APENDA_PARAM("CLA_SP_PUB_1257",5,Vetor_Campos)
		ObjCmd.Execute'pega dbaction
		DBAction = ObjCmd.Parameters("RET").value

		'response.write "<script>alert('"&ObjCmd.Parameters("RET").value&"')</script>"
		'response.write "<script>alert('"&ObjCmd.Parameters("retDesc").value&"')</script>"

	End if

	ResgatarLista(dblSolId)

	ConclusaoSNOA = DBAction

End Function

Select Case strAcao
	
	Case "ResgatarLista"
		Call ResgatarLista(Request.QueryString("dblSolId"))

	Case "ConclusaoSNOA"

		DBAction = ConclusaoSNOA(Request.QueryString("dblSolId"))
		'response.write "<script>alert('Solicitação incluída com sucesso.')</script>"
		Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'')</script>"
end Select
%>
