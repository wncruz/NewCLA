<!--#include file="../inc/data.asp"-->
<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: CadastraOS.asp
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: Tela que recebe o pedido e realiza a gravação da OS.
%>
<%
Dim DtAtual
Dim Formato
%>
<HTML>
<HEAD>
<Title>Cadastro de OS
</Title>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
</HEAD>
<Body topmargin=0 leftmargin=0 class=TA>
<input type=hidden name=hdnProvedor value="<%=Request.QueryString("Pro")%>">
<input type=hidden name=hdnPedido value="<%=Request.QueryString("Ped")%>">
<input type=hidden name=hdnAcessoFisico value="<%=Request.QueryString("Acf")%>">
<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
<form method="post" name=Form1 >
<input type=hidden name="hdnAcao">
<tr>
<td >
<input type=hidden name=hdnFormato value ="<%
				if Request.QueryString("Pro") <> "" then
					set objRS = db.execute("CLA_sp_sel_formato_provedor " & Request.QueryString("Pro"))
					if not objRS.Eof then
						Formato = Trim(objRS("Pro_PadraoOS"))
						Response.Write Trim(objRS("Pro_PadraoOS"))
					End if 
				End if
				%>">
<table border=0 cellspacing="1" cellpadding = 0 width="500" >
<tr class=clsSilver>
	<th colspan=2><p align=center>Código OS Operadora</p></th>
</tr>
<tr class=clsSilver>
	<td width="128">&nbsp;Código OS</td>
	<td width="172">
	<input type="text" class="text" name="TxtCodigoOS" value="" maxlength="<%=Len(Formato)%>" size="<%=Len(Formato)+2%>">&nbsp;<%
				if Request.QueryString("Pro") <> "" then
					set objRS = db.execute("CLA_sp_sel_formato_provedor " & Request.QueryString("Pro"))
					if not objRS.Eof then
						Response.Write "(" & Trim(objRS("Pro_PadraoOS")) & ")"
					End if 
				End if
				%>
	</td>
</tr>
<tr class=clsSilver>
	<td width="128">&nbsp;Data Emissão OS</td>
	<td width="142">
	<input type="text" class="text" name="TxtDtEmissao" value="" maxlength="10" size="<%=Len(Formato)+2%>" onChange = "ValidarData(this.value)" onKeyPress="OnlyNumbers();AdicionaBarraData(this);">&nbsp;(dd/mm/aaaa)
	</td>
</tr>
<tr>
<input type=hidden name=hdnDtEnvioEmail value ="<%
				if Request.QueryString("Pro") <> "" then
					set objRS = db.execute("CLA_sp_sel_DtEnvioEmail " & Request.QueryString("Ped"))
					if not objRS.Eof then
						Response.Write Trim(objRS("Ped_DtEnvioEmail"))
						DtAtual = Trim(objRS("DtAtual"))
					End if 
				End if
				%>">
<input type=hidden name=hdnDtAtual value ="<%=DtAtual%>">
	<td colspan=2 align="center" height=30px >
		<input type="button" class="button" name="btnAtualizar" value="Atualizar" style="width:100px" onclick="RealizaCadastroOS(hdnFormato.value,hdnAcessoFisico.value,hdnPedido.value,hdnDtEnvioEmail.value,hdnDtAtual.value);">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.returnValue = 'false';window.close();" style="width:100px" accesskey="X">
	</td>
</tr>
</table>
<%if Request.QueryString("Pro") = "" or Request.QueryString("Ped") = "" or Request.QueryString("Acf") = "" then
	%><script language='javascript'>
	alert("Parametros passados para a página inválidos");
	document.getElementById('btnAtualizar').style.visibility = "hidden";
	return false;
	window.close;
	</script><%
  End if%>
</form>
</BODY>
</HTML>
