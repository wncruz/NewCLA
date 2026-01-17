<%
'	- Sistema			: CLA
'	- Arquivo			: excelRelaAtivComercialmente.asp
'	- Responsável		: Eduardo Araújo
'	- Descrição			: Recebe SQL para que seja executado sem paginação afim de exportar todos os registros da consulta.
%>
<!--#include file="../inc/data.asp"-->

<HTML>
<HEAD>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
</HEAD>
<Body topmargin=0 leftmargin=0 class=TA>
<Form name=Form1 method=Post onsubmit="return false">
<input type=hidden name=hdnXls>
<title>CLA - Gerando planilha... aguarde...</title>
<center>
<br>
<br>
<b>Gerando planilha... aguarde...</b>
<br>
<br>
* Este processo poderá levar alguns minutos.

<%
Server.ScriptTimeout = 300

strSQL = Request.Form("hdnSQLXLS")


'Response.Write Request.Form("hdnSQLXLS") 
'response.end

Set objRSPag = db.Execute(strSQL)'pega dbaction

intCount=1

if not objRSPag.Eof or not objRSPag.Bof then

	strXls = "<table border=0 cellspacing=1 cellpadding=0 width=760>"
	strXls = strXls  &  "<tr>"
	strXls = strXls  &  "<th width=90>&nbsp;Pedido</th>"
	strXls = strXls  &  "<th>&nbsp;Data</th>"
	strXls = strXls  &  "<th>&nbsp;Dias</th>"
	strXls = strXls  &  "<th nowrap>&nbsp;Nº Acesso</th>"
	strXls = strXls  &  "<th>&nbsp;Cliente</th>"
	strXls = strXls  &  "<th>&nbsp;Localidade</th>"
	strXls = strXls  &  "<th>&nbsp;Velocidade</th>"
	strXls = strXls  &  "<th>&nbsp;Status</th>"
	strXls = strXls  &  "<th>&nbsp;Provedor</th>"
	strXls = strXls  &  "</tr>"
	
	
	strXls = "<table border=0 cellspacing=1 cellpadding=2 width=760 align=center>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<th colspan=16 nowrap align=left>Relatório de Acessos Ativados Comercialmente</th>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<th nowrap>&nbsp;UF</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Provedor</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Número Pedido</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;GLA</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;GIC-L</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;GIC-N</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Dt Carta Provedor</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Dt Aceite</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Cliente</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Velocidade</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Designação Provedor</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Pta A</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Pta B</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Dias Pendentes</th>"
	
	strXls = strXls  & "	<th nowrap>&nbsp;Grupo</th>"
	strXls = strXls  & "	<th nowrap>&nbsp;Status Detalhado</th>"
	
	strXls = strXls  & "</tr>"
		
	Do Until objRSPag.EOF = True

		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		
		strXls = strXls & "<tr class=" & strClass & " >"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("UF")  & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("pro_nome") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("NumeroDoPedido") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("GLA") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("GICL") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("GICN") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("DataPedido") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("DataAceite") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("NomeCliente") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("Velocidade") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("DesignacaoProvedor") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("CnlLocalPTA") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("CnlLocalPTB") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("DiasPendentes") & "</td>"
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("GCli_Descricao") & "</td>"
		
		strXls = strXls & "	<td nowrap align=center>"& objRSPag("StatusDet") & "</td>"
				
		strXls = strXls & "</tr>"

		objRSPag.MoveNext
										
	loop			

	strXls = strXls  &  "</table>"



end if

	if Trim(strXls) <> "" then
		Response.Write "<script language=javascript>document.Form1.hdnXls.value ='" & TratarAspasJS(strXls) & "';</script>"
		strXml = "<?xml version=""1.0"" encoding=""ISO-8859-1""?><xDados>"& strXml & "</xDados>"
		Response.Write "<script language=javascript>parent.objXmlGeral.loadXML('" & strXml & "');</script>"
	Else
		strHtmlRet = strHtmlRet & "<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>"
		strHtmlRet = strHtmlRet & "<tr>"
		strHtmlRet = strHtmlRet & "<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
		strHtmlRet = strHtmlRet & "</tr>"
		strHtmlRet = strHtmlRet & "</table>"
		Response.Write strHtmlRet
	End if
	
%>
</Form>
</html>

<script language=javascript>
function AbrirXls()
{
	with (document.Form1)
	{
		action = "ExcelExport.asp"
		submit()
	}
}
//alert(document.Form1.hdnXls.value)
AbrirXls()
</script>