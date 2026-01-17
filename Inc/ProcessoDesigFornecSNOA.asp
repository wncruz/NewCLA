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
	strHtml = strHtml &  "	<th nowrap>&nbsp;Número SNOA</th>"
	strHtml = strHtml &  "	<th>&nbsp;Designação Fornecedora</th>"
	strHtml = strHtml &  "	<th>&nbsp;Designação Solicitante</th>"
	strHtml = strHtml &  "	<th>&nbsp;Centro Fios A</th>"
	strHtml = strHtml &  "	<th>&nbsp;Centro Fios B</th>"
	strHtml = strHtml &  "</tr>"

	//dblSolId = 1535099
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_designacaoFornecedora ",1,Vetor_Campos)
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
		strHtml = strHtml &  "<td width=15% nowrap >&nbsp;" & objRS("SNOA") & "</td>"
		strHtml = strHtml &  "<td >&nbsp;" & objRS("designacao_recurso_fornecedora") & "</td>"
		strHtml = strHtml &  "<td >&nbsp;" & objRS("DM") & "</td>"
		strHtml = strHtml &  "<td nowrap >&nbsp;" & objRS("id_centro_fios_a") & "</td>"
		strHtml = strHtml &  "<td >&nbsp;" & objRS("id_centro_fios_b") & "</td>"
		strHtml = strHtml &  "</tr>"
		objRS.MoveNext
	Wend

	strHtml = strHtml &  "</table>"
	strHtml = strHtml &  "</Html></Body>"
	Response.Write strHtml

End Function

Select Case strAcao
	
	Case "ResgatarLista"
		Call ResgatarLista(Request.QueryString("dblSolId"))

end Select
%>
