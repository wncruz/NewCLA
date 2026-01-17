<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoTransferirCarteira.ASP
'	- Descrição			: Faz a transferência de carteira
%>
<!--#include file="../inc/data.asp"-->
<HTML>
<HEAD>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
</HEAD>
<Body topmargin=0 leftmargin=0 class=TA>
<Form name=Form1 method=Post >
<input type=hidden	name=cboUsuarioDe		value="<%=Request.Form("cboUsuarioDe")%>">
<input type=hidden  name=cboUsuarioPara		value="<%=Request.Form("cboUsuarioPara")%>">
<input type=hidden  name=hdnUsuarioAtual		value="<%=Request.Form("hdnUsuarioAtual")%>">

<%

Select Case Request.Form("hdnAcao") 

	Case "ConsultarCarteira"
		'Chama que realiza consulta 
		call ConsultarTransf
		
	Case "Transferir"

	'Transfere carteira

	Dim intItem
	Dim DBActionAux
	DBActionAux = 0

	if Request.Form("hdnSolId") <> "" then
		Vetor_Campos(1)="adInteger,4,adParamInput," & Request.Form("cboUsuarioDe") 
		Vetor_Campos(2)="adInteger,4,adParamInput," & Request.Form("cboUsuarioPara")
		Vetor_Campos(3)="adInteger,4,adParamInput," & Request.Form("hdnSolId")

		Call APENDA_PARAM("CLA_SP_TransfereCarteira",3,Vetor_Campos)
		Call ObjCmd.Execute()
		Response.Write "<script language=javascript>alert('Carteira Transferida com sucesso.');</script>"

	end if 

	case ""
		'Chama que realiza consulta 
		call ConsultarTransf

End Select	


function ConsultarTransf()

	Dim intIndex
	Dim strSql
	Dim intCount
	Dim strClass
	Dim objRSSol
	Dim dblIdLog
	Dim dblSolId
	Dim strHtmlRet
	Dim strLogico
	Dim strXml
	Dim objDicSol
		
		Set objDicSol = Server.CreateObject("Scripting.Dictionary") 

		strXml = "<xDados>"
			
		'Vetor_Campos(1)="adInteger,4,adParamInput," & Request.Form("hdnUsuarioAtual") 
		'strSql = APENDA_PARAMSTRSQL("CLA_sp_Sel_Carteira",1,Vetor_Campos)
	Vetor_Campos(1)="adInteger,4,adParamInput," 	& Request.Form("cboUsuarioDe") 				'@Usu_ID
	Vetor_Campos(2)="adInteger,4,adParamInput,"		& Request.Form("txtSolId")					'@sol_id
	Vetor_Campos(3)="adWChar,2,adParamInput,"		& mid(Request.form("txtPedNum"), 1, 2)		'@Ped_Prefixo
	Vetor_Campos(4)="adInteger,4,adParamInput,"		& mid(Request.form("txtPedNum"), 4, 5)		'@Ped_Numero
	Vetor_Campos(5)="adInteger,4,adParamInput,"		& mid(Request.form("txtPedNum"), 10, 4)		'@Ped_Ano
	Vetor_Campos(6)="adWChar,60,adParamInput,"		& Request.Form("txtCliente")				'@Cli_Nome
	Vetor_Campos(7)="adWChar,60,adParamInput,"		& Request.Form("txtEndereco") 				'@End_NomeLogr
	Vetor_Campos(8)="adWChar,10,adParamInput,"		& Request.Form("txtNroEnd")					'@End_NroLogr
	Vetor_Campos(9)="adWChar,30,adParamInput,"		& Request.Form("txtComplemento")			'@Aec_Complemento
	Vetor_Campos(10)="adInteger,4,adParamInput,"	& Request.Form("txtFacID")					'@Fac_ID
	Vetor_Campos(11)="adInteger,4,adParamInput," 	& Request.Form("cboCnl")					'@CNL

	strSql = APENDA_PARAMSTRSQL("CLA_sp_Sel_Carteira ",11,Vetor_Campos)		

	Set objRSCarteira = db.Execute(strSql)
	
		Call PaginarRS(0,strSql)
		strHtmlRet = ""

		intCount=1
		if not objRSPag.Eof and not objRSPag.Bof then
			'Link Xls/Impressão
			strLink =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
							"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
							"<a href='javascript:TelaImpressao(800,600,""Pedidos Pendentes - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _ 
							"</table>"
			strHtml = "<table border=0 cellspacing=1 width=760>"

			strXls  = "<table border=0 cellspacing=1 width=760>" & _
							"<tr>" & _
								"<th >&nbsp;Pedido de Acesso</th>" & _
								"<th >&nbsp;Sol</th>" & _
								"<th >&nbsp;Cliente</th>" & _
								"<th >&nbsp;Ação</th>" & _
								"<th nowrap >&nbsp;Nº do Contrato</th>" & _
								"<th >&nbsp;Status Atual</th>" & _
								"<th >&nbsp;Perfil</th>" & _
							"</tr>"
				
			For intIndex = 1 to objRSPag.PageSize
			'For intIndex = 1 to objRSPag.PageSize
			'if intPedId <> dblPedId then
				if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
				dblSolId = objRSPag("Sol_id")
				dblPedId = objRSPag("Ped_id")

				if not isNull(dblPedId) then
					Set objRSSts = db.Execute("CLA_sp_sel_StatusSolicitacao null,null,3," & dblPedId)
				Else
					Set objRSSts = db.Execute("CLA_sp_sel_StatusSolicitacao null,null,3," & dblSolId)
				End if	

				if Not objRSSts.Eof and Not objRSSts.Bof then 
					strStatusDet = objRSSts("Sts_Desc") 
				End if

				strHtml = strHtml & "<tr class='" & strClass & "'>"
				if Not objDicSol.Exists(Cstr(dblSolId)) then
					Call objDicSol.Add(Cstr(dblSolId),Cstr(dblSolId))
					strXml = strXml & "<Sol_Id>"& dblSolId &"</Sol_Id>"
					 strHtml = strHtml & "<td width=20><input type=checkbox name=chkCarteira value="&dblSolId&" onClick=""parent.Transferir(this)""></tr>"
				Else
					strHtml = strHtml & "<td width=20></td>"
				End if	

				if not isNull(objRSPag("Ped_Numero")) then
					strHtml = strHtml & "<td width=96 nowrap>" & ucase(objRSPag("Ped_Prefixo") & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano")) & "</td>"
				Else
					strHtml = strHtml & "<td width=96 nowrap>&nbsp;</td>"
				End if	
				strHtml = strHtml & "<td width=40><a href='javascript:parent.DetalharItem(" & dblSolId & ")'>" & dblSolId & "</a></td>"
				strHtml = strHtml & "<td width=240>" & objRSPag("Cli_Nome") & "</td>"
				strHtml = strHtml & "<td width=100>" & objRSPag("Tprc_Des") & "</td>"
				strHtml = strHtml & "<td width=100><span id=spnGeral onmouseover='showtip(this,event,""" & objRSPag("Acl_NContratoServico") & """);' onmouseout='hidetip();'>" & FormatarCampo(objRSPag("Acl_NContratoServico"),10) & "</span></td>"
				if Len(strStatusDet) > 18 then
					strStatusDet2 = Left(strStatusDet,18) & "..."
				Else
					strStatusDet2 = strStatusDet
				End if	
				strHtml = strHtml & "<td width=140><span id=spnStatus onmouseover=""showtip(this,event,'" & strStatusDet & "');"" onmouseout=""hidetip();"">" & strStatusDet2 & "</span></td>"
				strHtml = strHtml & "<td width=20>" & objRSPag("Age_Desc") & "</td>"
				strHtml = strHtml & "</tr>"

				strXls = strXls & "<tr class='" & strClass & "'>"
				if not isNull(objRSPag("Ped_Numero")) then
					strXls = strXls & "<td >" & ucase(objRSPag("Ped_Prefixo") & "-" & right("00000" & objRSPag("Ped_Numero"),5) & "/" & objRSPag("Ped_Ano")) & "</td>"
				Else	
					strXls = strXls & "<td >" & dblSolId & "</td>"
				End if	
				strXls = strXls & "<td >" & dblSolId & "</td>"
				strXls = strXls & "<td >" & objRSPag("Cli_Nome") & "</td>"
				strXls = strXls & "<td >" & objRSPag("Tprc_Des") & "</td>"
				strXls = strXls & "<td >" & objRSPag("Acl_NContratoServico") & "</td>"
				strXls = strXls & "<td >" & strStatusDet & "</td>"
				strXls = strXls & "<td >" & objRSPag("Age_Desc") & "</td>"
				strXls = strXls & "</tr>"
				
				intCount = intCount+1
				objRSPag.MoveNext
				if objRSPag.EOF then Exit For
			next
			strHtml = strHtml & "</table>"
			strXls = strXls & "</table>"

		Else
			strHtml = strHtml & "<table width=760 border=0 cellspacing=0 cellpadding=0 valign=top>"
			strHtml = strHtml & "<tr>"
			strHtml = strHtml & "<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
			strHtml = strHtml & "</tr>"
			strHtml = strHtml & "</table>"
		End if

		strXml = strXml & "</xDados>"

		Response.Write "<script language=javascript>parent.spnLinks.innerHTML = '" & TratarAspasJS(strLink) & "';parent.document.forms[0].hdnXls[0].value ='" & TratarAspasJS(strXls) & "';</script>"
		Response.Write "<script language=javascript>parent.objXmlCarteira.loadXML('" & strXml & "');</script>"
		
		Response.Write strHtml

end function 

%>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</BODY>

</Form>
</HTML>
<%


Set objRSSol = Nothing
Set objRSCarteira = Nothing
DesconectarCla()
%>
