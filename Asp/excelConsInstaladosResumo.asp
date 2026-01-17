<%
'	- Sistema			: CLA
'	- Arquivo			: excelConsInstaladosResumo.asp
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
Server.ScriptTimeout = 180

dim strNomeProvedor
dim strDataInicio
dim strDataFim


strSQL = Request.Form("hdnSQLXLS")
strNomeProvedor = Request.Form("hdnProvedor")
strDataInicio = Request.Form("hdnDataInicio")
strDataFim = Request.Form("hdnDataFim")

'Response.Write Request.Form("hdnSQLXLS") 
'response.end

Set objRSPag = db.Execute(strSQL)'pega dbaction

intCount=1

if not objRSPag.Eof or not objRSPag.Bof then

	int10 = objRSPag("ate10")
	int20 = objRSPag("ate20")
	int30 = objRSPag("ate30")
	int45 = objRSPag("ate45")
	int60 = objRSPag("ate60")
	int90 = objRSPag("ate90")
	int120 = objRSPag("ate120")
	int180 = objRSPag("ate180")
	int180Mais = objRSPag("mais180")

	intTotal  = objRSPag("Total")
	intTotalDias = objRSPag("TotalDias")
		
	strXls = strXls  & "<table border=0 cellspacing=1 cellpadding=0 width=300 align=center>"
	strXls = strXls  & "<tr >"
	strXls = strXls  & "	<td colspan=2>" & Request.Form("hdnProvedor") & " " & Request.Form("txtDataInicio") & " - " & Request.Form("txtDataFim")   & "</td>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<th colspan=2>&nbsp;Acessos Instalados</th>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<th>&nbsp;Dias</th>"
	strXls = strXls  & "	<th >&nbsp;Quantidade</th>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver >"
	strXls = strXls  & "	<td width=150px >&nbsp;Até 10</td>"
	strXls = strXls  & "	<td  align=right >" & int10 & "&nbsp;</td>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<td>&nbsp;De 11 a 20</td>"
	strXls = strXls  & "	<td align=right>" & int20 & "&nbsp;</td>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<td>&nbsp;De 21 a 30</td>"
	strXls = strXls  & "	<td align=right>" & int30 & "&nbsp;</td>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<td>&nbsp;De 31 a 45</td>"
	strXls = strXls  & "	<td align=right>" & int45 & "&nbsp;</td>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<td >&nbsp;De 46 a 60</td>"
	strXls = strXls  & "	<td align=right>" & int60 & "&nbsp;</td>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<td >&nbsp;De 61 a 90</td>"
	strXls = strXls  & "	<td align=right>" & int90 & "&nbsp;</td>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<td >&nbsp;De 91 a 120</td>"
	strXls = strXls  & "	<td align=right>" & int120 & "&nbsp;</td>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<td >&nbsp;De 121 a 180</td>"
	strXls = strXls  & "	<td align=right>" & int180 & "&nbsp;</td>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver>"
	strXls = strXls  & "	<td>&nbsp;Acima de 180</td>"
	strXls = strXls  & "	<td align=right>" & int180Mais & "&nbsp;</td>"
	strXls = strXls  & "</tr>"
	strXls = strXls  & "<tr class=clsSilver2>"
	strXls = strXls  & "<td>&nbsp;Total</td>"
	strXls = strXls  & "<td align=right>" & intTotal & "&nbsp;</td>"
	strXls = strXls  & "</tr>"

	strXls = strXls  & "<tr class=clsSilver2>"
	strXls = strXls  & "<th>&nbsp;Média em dias</th>"
	if intTotal > 0 then
		strXls = strXls  & "	<th><p align=right>" & Replace(FormatNumber(intTotalDias/intTotal,2),".",",") & "&nbsp;</p></th>"
	Else
		strXls = strXls  & "<th><p align=right>0&nbsp;</p></th>"
	End if
	strXls = strXls  & "</tr>"
	strXls = strXls  & "</table><br>"

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