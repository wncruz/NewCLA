<%
'•EXPERT INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ExcelExportRecebe.asp
'	- Responsável		: PRSS
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
<input type=hidden	name=cboProvedor		value="<%=Request.Form("cboProvedor")%>">
<input type=hidden  name=txtDtInicio		value="<%=Request.Form("txtDtInicio")%>">
<input type=hidden  name=txtDtFim			value="<%=Request.Form("txtDtFim")%>">
<input type=hidden  name=cboTipoProcesso	value="<%=Request.Form("cboTipoProcesso")%>">
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


strSQL = Request.Form("hdnSQLXLS")

Set objRSPag = db.Execute(strSQL)'pega dbaction

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	'Link Xls/Impressão
	strLink =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
					"<a href=""javascript:Form_AbrirXlsrecebe.submit()"" onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Cartas ao Provedor - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
					"</table>"

	strXls = strXls & "<table border=0 width=760 cellspacing=0 cellpadding=1 >"
	strXls = strXls & "<tr><th>Cliente</th><th>Pedido</th><th>Data de Envio</th><th>Processo</th><th>Nº Acesso</th><th>CCTO Provedor</th><th>Id Acesso Lógico (678)</th><th>CNL do Cliente </th></tr>"

	Do Until objRSPag.EOF = True 
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

		'''dblSolId	= objRSPag("Sol_Id")
		'''dblPedId	= objRSPag("Ped_Id")

		'''Vetor_Campos(1)="adInteger,4,adParamInput,"
		'''Vetor_Campos(2)="adInteger,4,adParamInput," & dblPedId
		'''Vetor_Campos(3)="adInteger,4,adParamInput," & dblSolId
		'''Vetor_Campos(4)="adInteger,4,adParamInput,"
		'''Vetor_Campos(5)="adInteger,4,adParamInput,"
		'''Vetor_Campos(6)="adInteger,4,adParamInput,"
		'''Vetor_Campos(7)="adWChar,3,adParamInput,"
		'''Vetor_Campos(8)="adWChar,1,adParamInput,"
		'''Vetor_Campos(9)="adWChar,1,adParamInput,T"
		
		'''strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_ACESSOFISICO",9,Vetor_Campos)
		'''Set objRSFis = db.Execute(strSqlRet)

		'''if Not objRSFis.EOF and not objRSFis.BOF then
			'''strNroAcessoPtaEBT	= Trim(objRSFis("acf_nroacessoptaebt"))
			'''strCctoProvedor		= Trim(objRSFis("Acf_NroAcessoCCTOProvedor"))
			'''dblProId			= Trim(Request.Form("cboProvedor") ) 'objRSFis("Pro_Id"))
		'''End if

		'''Set objRSFis = Nothing

		dblProId			= Trim(Request.Form("cboProvedor") )
		
		strXls = strXls & "<tr class=" & strClass & ">"
		strXls = strXls & "<td nowrap>" & objRSPag("Cli_Nome") & "</td>"
		strXls = strXls & "<td nowrap>" & ucase(objRSPag("Ped_Prefixo")) & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano") & "</td>"
		strXls = strXls & "<td nowrap>" & Trim(split(objRSPag("Doc_DtCriacao")," ")(0)) & "</td>"
		strXls = strXls & "<td nowrap>" & objRSPag("Tprc_Des") & "</td>"
		'strXls = strXls & "<td >" & strNroAcessoPtaEBT & "</td>"
		strXls = strXls & "<td >" & objRSPag("acf_nroacessoptaebt") & "</td>"
		
		'strXls = strXls & "<td >" & strCctoProvedor & "</td>"
		strXls = strXls & "<td >" & objRSPag("Acf_NroAcessoCCTOProvedor") & "</td>"
		
		strXls = strXls & "<td width=100>" & objRSPag("Acl_IDacessoLogico") & "</td>"
		strXls = strXls & "<td width=100>" & objRSPag("Cid_Sigla") & "</td>"

		intCount = intCount+1
		objRSPag.MoveNext
	loop

	strXls = strXls & "</table>"
End if	
		
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