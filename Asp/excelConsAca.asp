<%
'	- Sistema			: CLA
'	- Arquivo			: excelConsAca.asp
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
	
	Do Until objRSPag.EOF = True

		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

		strIdLogico = objRSPag("Acl_IDAcessoLogico")
		dblSolId	= objRSPag("Sol_Id")
		dblPedId	= objRSPag("Ped_Id")

		if  len(objRSPag("sts_desc")) > 18 then
			strStatus = Left(objRSPag("sts_desc"),18) & "..."
		Else
			strStatus = objRSPag("sts_desc")
		End if	

		strXls = strXls  &  "<tr class="&strClass&">"
		strXls = strXls  &  "<td >" & ucase(objRSPag("Ped_Prefixo")&"-"& right("00000" & objRSPag("Ped_Numero"),5) &"/"& objRSPag("Ped_Ano")) & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & right("0" & day(objRSPag("Ped_Data")),2) & "/" & right("0" & month(objRSPag("Ped_Data")),2) & "/" & year(objRSPag("Ped_Data")) & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Dias_Pend") & "</td>"
		'strXls = strXls  &  "<td >&nbsp;" & strNroAcessoPtaEBT & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("acf_nroacessoptaebt") & "</td>"
		strXls = strXls  &  "<td width=150px>&nbsp;" & objRSPag("Cli_Nome") & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Cid_Sigla") & "</td>"
		'strXls = strXls  &  "<td >&nbsp;" & strVelDescAcessoFisicoInst & "&nbsp;" & TipoVel(intTipoVel) &  "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("vel_desc") & "&nbsp;" & TipoVel(objRSPag("Acf_TipoVel")) &  "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Sts_Desc") & "</td>"
		strXls = strXls  &  "<td >&nbsp;" & objRSPag("Pro_Nome") & "</td>"
		strXls = strXls  &  "</tr>"

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