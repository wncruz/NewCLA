<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<!--#include file="ProcessoEmailProvedor.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: EmailProvedor.ASP
'	- Descrição			: Envia um e-mail ao provedor
%>
<tr>
<td width=760 align=center>
<%
Dim objCarta
Dim intErro
Dim strErro
Dim strProEmail
Dim dblSolId
Dim dblPedId
Dim dblProId
Dim dblEscEntrega
Dim strLink
Dim strXmlSaida
Dim objXml

dblSolId = Request.Form("hdnSolId")
dblPedId = Request.Form("hdnPedId")
dblProId = Request.Form("cboProvedor")
dblSisId = Request.Form("hdnRede")
dblEscEntrega = Request.Form("cboLocalInstala")
intTipoProcesso = Request.Form("hdnTipoProcesso")

if dblPedId = "" then 
	Response.Write "<br><b><p aling=center><font color=red>Pedido não encontrado</font></p></b>"
	Response.End 
End if

Set objRSPed = db.execute("CLA_sp_view_pedido null,null,null,null,null,null," & dblPedId & ",null,null,'T'")
if Not objRSPed.Eof and not objRSPed.Bof then
	if dblSisId = "3" then
		'Set objRSRec = db.execute("CLA_sp_view_recurso " & objRSPed("Rec_IDEntrega")) 
		'if not objRSRec.Eof and not objRSRec.bof then
			'Set objRSPro = db.execute("CLA_sp_sel_estacao " & objRSRec("Esc_Id"))
			Set objRSPro = db.execute("CLA_sp_sel_estacao " & dblEscEntrega)
			strProEmail = Trim(objRSPro("Esc_Email"))  
			strProNome	= Trim(objRSPro("Esc_Contato"))
			'strTemplate = Trim(objRSPro("Esc_PathTemplate"))
			'strPathDoc	= Trim(objRSPro("Esc_PathDoc"))
		'End if
	Else
		Set objRSPro = db.execute("CLA_sp_sel_provedor " & dblProId) 
		if Not objRSPro.Eof and Not objRSPro.bof then
			strProEmail = Trim(objRSPro("pro_email"))
			strProNome	= Trim(objRSPro("Pro_Nome"))
			strTemplate = Trim(objRSPro("Pro_PathTemplate"))
			strPathDoc	= Trim(objRSPro("Pro_PathDoc"))
		End if
	End if
	dblPedId = objRSPed("Ped_Id")
	'Randomize
	strNomeArquivoSaida = UCASE(objRSPed("Ped_Prefixo")) & "-" & right("00000" & objRSPed("Ped_Numero"),5) & "-" & objRSPed("Ped_Ano") & ".doc"

End if

if not isNull(strTemplate) and not isNull(strPathDoc) and strTemplate <> "" and  strPathDoc <> "" then 

		Set objCarta = Server.CreateObject("xCLA_CartaProvedorD.clsCartaProvedor")

		strUserLogado = strUserName
		strXmlEntrada = "<xDados>" & _
						"<strNomeArquivoSaida >" & strNomeArquivoSaida & "</strNomeArquivoSaida >" & _
						"<dblPedId>" & dblPedId & "</dblPedId>" & _
						"<strUserLogado>" & strUserLogado & "</strUserLogado>" & _
						"<dblProId>" & dblProId & "</dblProId>" & _
						"<strRootApp>" & Server.MapPath("..")  & "</strRootApp>" & _
						"</xDados>"

		Call objCarta.CriarDocumento(strXmlEntrada, strXmlSaida, intErro, strErro)

		Set objXml = Server.CreateObject("Microsoft.XMLDOM")
		objXml.loadXml(strXmlSaida)
		'Response.Write "•" & objXml.xml
		Set objNode = objXml.SelectNodes("//strLink")

		'Response.Write objNode.length & " -  " & strLink & " - " & strErro
		if objNode.length > 0 then
			strLink = objNode(0).childNodes(0).text
		End if
		
		if intErro = 0 and strLink <> "" then
			strHtml = strHtml &  "<br><table cellpadding=0 cellspacing=1 width=510 >"
			strHtml = strHtml &  "<tr>"
			strHtml = strHtml &  "	<th nowrap>&nbsp;Enviar Email</th>"
			strHtml = strHtml &  "	<th nowrap>&nbsp;Editar</th>"
			strHtml = strHtml &  "	<th nowrap>&nbsp;Arquivo</th>"
			strHtml = strHtml &  "</tr>"
			strHtml = strHtml &  "<tr class=clsSilver height=17px>"
			strHtml = strHtml &  "<td nowrap align=center>"
			strHtml = strHtml &  "<a href=""javascript:parent.EnviarEmail()""><img src=""../imagens/Email.gif"" border=0 Alt=""Enviar Email ao Provedor""></a>&nbsp;<font color=red><i><span id=spnEnviado></span></i></font></td>"
			strHtml = strHtml &  "<td nowrap align=center><span id=spnRef onClick=""javascript:window.open('../" & strLink & "');"" style=""cursor:hand""><font color=#003388><img src='../imagens/Editar.gif' border=0 alt=Editar></font></span></td>"
			strHtml = strHtml &  "<td nowrap align=center>" & strLink & "</td>"
			strHtml = strHtml &  "</tr></table>"
		Else
			if intErro <> 0 then
				strHtml = strHtml &  "<table cellpadding=0 cellspacing=1 width=300 >"
				strHtml = strHtml &  "<tr>"
				strHtml = strHtml &  "	<th ><p align=center>" & intErro & " - " & strErro & "</p></th>"
				strHtml = strHtml &  "</tr></table>"
			Else	
				strHtml = strHtml &  "<table cellpadding=0 cellspacing=1 width=300 >"
				strHtml = strHtml &  "<tr>"
				strHtml = strHtml &  "	<th ><p align=center>Não foi possível gravar o arquivo.</p></th>"
				strHtml = strHtml &  "</tr></table>"
			End if	
		End if	

		Response.Write strHtml
		%>	
		</td>
		</tr>
		<tr>
		<td align=center height=25px><br>
			<input type="button" class="button" name="btnSair" value="Voltar" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');" onClick="javascript:window.location.replace('facilidade_main.asp')">&nbsp;
			<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
		</td>
		</tr>
		</table>
		<form name=Form1 method=Post>
			<input type=hidden name=hdnAcao>
			<input type=hidden name=hdnEmailProvedor value="<%=strProEmail%>">
			<input type=hidden name=hdnNomeProvedor value="<%=strProNome%>">
			<input type=hidden name=hdnArquivo value="<%=strLink%>">
			<input type=hidden name=hdnPedId value="<%=dblPedId%>">
		</form>
		<link rel=stylesheet type="text/css" href="../css/cla.css">
		<script language='javascript' src="../javascript/cla.js"></script>
		<SCRIPT LANGUAGE=javascript>
		<!--
		function EnviarEmail()
		{
			with (document.forms[0])
			{
				hdnAcao.value = "EnviarEmailProvedor"
				target = "IFrmProcesso"
				action = "ProcessoCartasProv.asp"
				submit()
			}
		}
		
		function EmailEnviado(strNomePro,strProEmail)
		{
			alert("Email enviado com sucesso para " + strNomePro + " (" + strProEmail + ").");
			parent.spnEnviado.innerHTML = '(Email enviado)';
			return true
		}

		//-->
		</SCRIPT>
		<iframe	id			= "IFrmProcesso"
				name        = "IFrmProcesso"
				width       = "0"
				height      = "0"
				frameborder = "0"
				scrolling   = "no" 
				align       = "left">
		</iFrame>
		<%
		Set objCarta = Nothing
		Set objRSPed = Nothing

		Response.End 

End if

'Default
strHtmlEmail = EnviarEmailProvedor(dblSolId,dblPedId,dblProId,strProEmail,"T",intTipoProcesso,dblSisId)
if Trim(strHtmlEmail) <> "" then
	response.write "<table align=center>"
	response.write "<tr><td><font color=red>E-MAIL ENVIADO COM SUCESSO PARA: " & strProEmail & "</font></td></tr>"
	response.write "</table>"
	Response.Write strHtmlEmail
Else
	response.write "<table align=center>"
	response.write "<tr><td><font color=red>Não foi possível enviar o e-mail para o provedor.</font></td></tr>"
	response.write "</table>"
End if
%>
</td>
</tr>
<tr>
<td align=center height=25px>
	<input type="button" class="button" name="btnSair" value="Voltar" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');" onClick="javascript:window.location.replace('facilidade_main.asp')">&nbsp;
	<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
</td>
</tr>
</table>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>