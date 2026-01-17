<%@ CodePage=65001 %>
<%
                Response.ContentType = "text/html; charset=utf-8"
                Response.Charset = "UTF-8"
              %>
<%

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
<meta http-equiv="Content-Type" content="text/html;charset=UTF-8" />
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
Server.ScriptTimeout = 300


strSQL = Request.Form("hdnSQLXLS")

strSQLCampo = Request.Form("hdnCampoSQLXLS")

'response.write "<script>alert('"&strSQLCampo&"')</script>"

Set objRSPag = db.Execute(strSQL)'pega dbaction

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	
	strXls = strXls & " <meta http-equiv=Content-Type content=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet /> "
	strXls = strXls & " <table border=0 width=760 cellspacing=0 cellpadding=1 > "
	
	'strXls = strXls & " <tr><th>Acesso Lógico</th><th>Acesso Físico </th><th>Facilidade </th><th>Tecnologia </th><th>CNL </th><th>UF </th><th>Velocidade do Acesso Físico </th><th>Status   </th><th>Data de Aceite    </th> "

	'strXls = strXls & "<th>Acesso Logico</th><th>Acesso Fisico </th><th>Facilidade </th><th>Tecnologia </th><th>Velocidade do Acesso Fisico </th><th>Status Fisico</th><th>Data Aceite</th><th>CLIENTE</th><th>Tipo</th><th>Logradouro</th><th>Numero</th><th>Bairro</th><th>CNL</th><th>UF</th><th>CEP</th> "

	strXls = strXls & "<th>Acesso Logico</th><th>Acesso Fisico </th><th>Facilidade </th><th>Tecnologia </th><th>Velocidade do Acesso Fisico </th><th>Status Fisico</th><th>Data Aceite</th><th>CLIENTE</th><th>Tipo</th><th>Logradouro</th><th>Numero</th><th>Bairro</th><th>CNL</th><th>UF</th><th>CEP</th><th>PE</th><th>Porta</th><th>SVLAN</th><th>CVLAN</th> "

	'strXls = strXls & " <th>Acesso Logico</th><th>Acesso Fisico </th> "
	Set objRSCampo = db.Execute(strSQLCampo)'pega dbaction

		intCountCampo=1
		if not objRSCampo.Eof and not objRSCampo.Bof then
			Do Until objRSCampo.EOF = True 
				strXls = strXls & " <th> " & ucase(objRSCampo("campo")) & "   </th> "
				intCountCampo = intCountCampo+1
				objRSCampo.MoveNext
			loop

			'strXls = strXls & " </tr> "
		End if	
		
	
	Do Until objRSPag.EOF = True 
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

		
		'dblProId			= Trim(Request.Form("cboProvedor") )
		
		'strXls = strXls & "<tr class=" & strClass & ">"
		'strXls = strXls & "<td nowrap>" & objRSPag("Acl_IDAcessoLogico") & "</td>"
		'strXls = strXls & "<td nowrap>" & ucase(objRSPag("Acf_IDAcessoFisico")) & "</td>"
		'strXls = strXls & "<td nowrap>" & Trim(objRSPag("newfac_nome")) & "</td>"
		'strXls = strXls & "<td nowrap>" & Trim(objRSPag("newtec_nome")) & "</td>"
		'strXls = strXls & "<td nowrap>" & Trim(objRSPag("cid_sigla")) & "</td>"
		'strXls = strXls & "<td nowrap>" & Trim(objRSPag("est_sigla")) & "</td>"
		'strXls = strXls & "<td nowrap>" & Trim(objRSPag("Vel_Desc")) & "</td>"
		'strXls = strXls & "<td nowrap>" & Trim(objRSPag("status")) & "</td>"
		'strXls = strXls & "<td nowrap>" & Trim(objRSPag("Acf_DtAceite")) & "</td>"
		
		
		
		strXls = strXls & "<tr class=" & strClass & ">"
		'strXls = strXls & "<td nowrap>" & objRSPag("sol_id") & "</td>"
		strXls = strXls & "<td nowrap>" & TiraAcento(ucase(objRSPag("acl_idacessologico"))) & "</td>"
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("acf_idacessofisico"))) & "</td>"
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("newfac_nome"))) & "</td>"
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("newtec_nome"))) & "</td>"
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("vel_desc"))) & "</td>"
		
		strXls = strXls & "<td nowrap> " & TiraAcento(Trim(objRSPag("status"))) & " </td>"   'Status Físico
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("Acf_DtAceite"))) & "</td>"   'Data Aceite
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("cli_nome"))) & "</td>"
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("tpl_sigla"))) & "</td>"
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("end_nomelogr"))) & "</td>"
		
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("end_nrologr"))) & "</td>"
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("end_bairro"))) & "</td>"
				
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("cid_sigla"))) & "</td>" 'CNL
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("est_sigla"))) & "</td>" 'UF
		
		strXls = strXls & "<td nowrap>" & TiraAcento(Trim(objRSPag("end_cep"))) & "</td>"
		
		strXls = strXls & "<td nowrap>&#8203;" & TiraAcento(Trim(objRSPag("newfacServico_PE"))) & "</td>"
		strXls = strXls & "<td nowrap>&#8203;" & TiraAcento(Trim(objRSPag("newfacServico_Porta"))) & "</td>"
				
		strXls = strXls & "<td nowrap>&#8203;" & TiraAcento(Trim(objRSPag("newfacServico_SVlan"))) & "</td>" 'CNL
		strXls = strXls & "<td nowrap>&#8203;" & TiraAcento(Trim(objRSPag("newfacServico_Vlan"))) & "</td>" 'UF
		
		Set objRSCampo = db.Execute(strSQLCampo)'pega dbaction

		intCountCampo=1
		if not objRSCampo.Eof and not objRSCampo.Bof then
			Do Until objRSCampo.EOF = True 
				campo = objRSCampo("campo")
				'if 
				'strXls = strXls & "<td nowrap> <script language=javascript> especialCharMask ( '" & Trim(objRSPag(campo)) & "' ) ;</script></td>"
				
				''strXls = strXls & "<td nowrap > " & "'"  & TratarAspasJS(TiraAcento(ucase(Trim(objRSPag(campo))))) & " </td>"
				
				'strXls = strXls & "<td nowrap > =" &  Trim(objRSPag(campo)) & " </td>"
				
				'strXls = strXls & "<td nowrap>=""" & TiraAcento(ucase(Trim(objRSPag(campo)))) & """</td>"
				
				strXls = strXls & "<td nowrap>&#8203;" & TiraAcento(ucase(Trim(objRSPag(campo)))) & "</td>"
				

				intCountCampo = intCountCampo+1
				objRSCampo.MoveNext
			loop

			'strXls = strXls & "</tr>"
		End if	
		

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

//Response.Write strXls
AbrirXls()
</script>