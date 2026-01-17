<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ListasCartasProvedor.ASP
'	- Responsável		: Vital
'	- Descrição			: Listas das cartas que foram enviadas ao provedor
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
<input type=hidden  name=cboHolding	value="<%=Request.Form("cboHolding")%>">
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim strHtmlRet
Dim objXml
'@@JKNUP: Adicionado. Trata o retorno da combo para não vir 0 ou ''.
Dim tpProc

tpProc = Request.Form("cboTipoProcesso")

if tpProc = "0" or tpProc & "" = "" then tpProc = "null" end if
'</@@JKNUP>

Vetor_Campos(1)="adInteger,2,adParamInput,"
Vetor_Campos(2)="adInteger,2,adParamInput," & Request.Form("cboProvedor") 
Vetor_Campos(3)="adInteger,2,adParamInput," & tpProc'Request.Form("cboTipoProcesso") 
Vetor_Campos(4)="adDate,8,adParamInput," & inverte_data(Request.Form("txtDtInicio"))
Vetor_Campos(5)="adDate,8,adParamInput," & inverte_data(Request.Form("txtDtFim"))
Vetor_Campos(6)="adWChar,13,adParamInput," & Request.Form("txtPedido")
Vetor_Campos(7)="adWChar,13,adParamInput,"
Vetor_Campos(8)="adInteger,2,adParamInput," & Request.Form("cboHolding") 

strSql = APENDA_PARAMSTR("CLA_sp_sel_listaCartaProvedor",8,Vetor_Campos)
'strSql = APENDA_PARAMSTR("CLA_sp_sel_relacaodocumento",6,Vetor_Campos)

'PRSS - 12/01/2006 - INICIO
'strSQLXLS = "CLA_sp_sel_relacaodocumento3 null,'"&Request.Form("cboProvedor")&"',"&tpProc&",'"&inverte_data(Request.Form("txtDtInicio"))&"','"&inverte_data(Request.Form("txtDtFim"))&"' , '"&Request.Form("txtPedido")&"', null,"&Request.Form("cboHolding")&""

Response.Write "<script language=javascript>parent.document.forms[0].hdnSQLXLS.value ="&chr(34)&strSql&chr(34)&";</script>"
'Response.Write "<script language=javascript>parent.document.forms[0].hdnSQLXLS.value ="&chr(34)&strSql&chr(34)&";</script>"

Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	'Link Xls/Impressão
	strLink =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
					"<a href=""javascript:AbrirXls()"" onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
					"<a href='javascript:TelaImpressao(800,600,""Cartas ao Provedor - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
					"</table>"
	strHtmlRet = ""
	strHtmlRet = strHtmlRet & "<table border=0 width=758 cellspacing=1 cellpadding=1 >"

	strXls = strXls & "<table border=0 width=760 cellspacing=0 cellpadding=1 >"
	strXls = strXls & "<tr><th>Cliente</th><th>Pedido</th><th>Data de Envio</th><th>Processo</th><th>Nº Acesso</th><th>CCTO Provedor</th><th>Id Acesso Lógico (678)</th><th>CNL do Cliente </th></tr>"

	For intIndex = 1 to objRSPag.PageSize
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

		'dblProId			= Trim(Request.Form("cboProvedor") )

		strHtmlRet = strHtmlRet & "<tr class=" & strClass & ">"
		strHtmlRet = strHtmlRet & "<td width=15px align=center><span id=spn" & intIndex & " style=""cursor:hand"" onClick=""javascript:parent.DetalheCarta(" & objRSPag("Doc_id") & ",'" & objRSPag("Doc_Tipo") & "','" & objRSPag("pro_id") & "')""><img src=""../imagens/Email.gif"" border=0 Alt=""Ver Carta""></span></td>"
		strHtmlRet = strHtmlRet & "<td width=20px><input type=checkbox checked value=" & objRSPag("Doc_Id") & " name=chkCarta onClick=""parent.SelecionarImpressao(" & objRSPag("Doc_Id") & ")""></td>"
		strHtmlRet = strHtmlRet & "<td width=115px>" & objRSPag("Cli_Nome") & "</td>"
		strHtmlRet = strHtmlRet & "<td width=100px>" & ucase(objRSPag("Ped_Prefixo")) & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano") & "</td>"
		strHtmlRet = strHtmlRet & "<td width=65px>" & Trim(split(objRSPag("Doc_DtCriacao")," ")(0)) & "</td>"
		strHtmlRet = strHtmlRet & "<td width=100px>" & objRSPag("Tprc_Des") & "</td>"
		'''strHtmlRet = strHtmlRet & "<td width=85px>" & strNroAcessoPtaEBT & "</td>"
		strHtmlRet = strHtmlRet & "<td width=85px>" & objRSPag("acf_nroacessoptaebt") & "</td>"
		
		'''strHtmlRet = strHtmlRet & "<td width=70px>" & strCctoProvedor & "</td>"		
		strHtmlRet = strHtmlRet & "<td width=70px>" & objRSPag("Acf_NroAcessoCCTOProvedor") & "</td>"
		
		strHtmlRet = strHtmlRet & "<td width=110px>" & objRSPag("Acl_IDacessoLogico") & "</td>"
		strHtmlRet = strHtmlRet & "<td width=45px>" & objRSPag("Cid_Sigla") & "</td>"

		strXls = strXls & "<tr class=" & strClass & ">"
		strXls = strXls & "<td nowrap>" & objRSPag("Cli_Nome") & "</td>"
		strXls = strXls & "<td nowrap>" & ucase(objRSPag("Ped_Prefixo")) & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano") & "</td>"
		strXls = strXls & "<td nowrap>" & Trim(split(objRSPag("Doc_DtCriacao")," ")(0)) & "</td>"
		strXls = strXls & "<td nowrap>" & objRSPag("Tprc_Des") & "</td>"
		'''strXls = strXls & "<td >" & strNroAcessoPtaEBT & "</td>"
		strXls = strXls & "<td >" & objRSPag("acf_nroacessoptaebt") & "</td>"
		
		'''strXls = strXls & "<td >" & strCctoProvedor & "</td>"
		strXls = strXls & "<td >" & objRSPag("Acf_NroAcessoCCTOProvedor") & "</td>"
		
		strXls = strXls & "<td width=100>" & objRSPag("Acl_IDacessoLogico") & "</td>"
		strXls = strXls & "<td width=100>" & objRSPag("Cid_Sigla") & "</td>"


		strXml = strXml & "<Carta>"
		strXml = strXml & "<intIndice>" & objRSPag("Doc_Id") & "</intIndice>"
		strXml = strXml & "<Cliente>" & TratarAspasXML(objRSPag("Cli_Nome")) & "</Cliente>"
		strXml = strXml & "<Pedido>" & ucase(objRSPag("Ped_Prefixo")) & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano") & "</Pedido>"
		strXml = strXml & "<DataCriacao>" & Trim(split(objRSPag("Doc_DtCriacao")," ")(0)) & "</DataCriacao>"
		strXml = strXml & "<Processo>" & objRSPag("Tprc_Des") & "</Processo>"
		strXml = strXml & "<Acao>I</Acao>"
		'''strXml = strXml & "<NroAcesso>" & strNroAcessoPtaEBT & "</NroAcesso>"
		strXml = strXml & "<NroAcesso>" & objRSPag("acf_nroacessoptaebt") & "</NroAcesso>"
		
		strXml = strXml & "<CCTOProv>" & objRSPag("Acf_NroAcessoCCTOProvedor") & "</CCTOProv>"

		strXml = strXml & "<AcessoLogico>" & objRSPag("Acl_IDacessoLogico") & "</AcessoLogico>"
		strXml = strXml & "<Cid_Sigla>" & objRSPag("Cid_Sigla") & "</Cid_Sigla>"

		strXml = strXml & "</Carta>"

		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	Next

	strHtmlRet = strHtmlRet & "</table>"
	strXls = strXls & "</table>"
Else
	strHtmlRet = ""
End if	
		
if Trim(strHtmlRet) <> "" then
	Response.Write strHtmlRet
	Response.Write "<script language=javascript>parent.document.forms[0].hdnXls[0].value ='" & TratarAspasJS(strXls) & "';</script>"
	strXml = "<?xml version=""1.0"" encoding=""ISO-8859-1""?><xDados>"& strXml & "</xDados>"
	Response.Write "<script language=javascript>parent.objXmlGeral.loadXML('" & strXml & "');</script>"

	'Response.Write "<script language=javascript>parent.spnLinks.innerHTML = '" & TratarAspasJS(strLink) & "'</script>"
Else
	strHtmlRet = strHtmlRet & "<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>"
	strHtmlRet = strHtmlRet & "<tr>"
	strHtmlRet = strHtmlRet & "<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
	strHtmlRet = strHtmlRet & "</tr>"
	strHtmlRet = strHtmlRet & "</table>"
	Response.Write strHtmlRet
End if	
%>
</BODY>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</Form>
</HTML>
<%
Set objRSPag = Nothing
DesconectarCla()
%>