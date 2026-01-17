<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: postFile.ASP
'	- Responsável		: Vital
'	- Descrição			: Enviar arquivo para o servidor
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<html>
<head>
	<title>CLA Controle Local de Acesso - Upload Carta Provedor</title>
<SCRIPT LANGUAGE=javascript>
<!--
function Check(){
	with (document.forms[1])
	{
		if (txtFile.value == "")
		{
			alert("Selecione um arquivo!")
			return false
		}
		if (spnLista.innerHTML == "")
		{
			alert("Favor efetuar um busca primeiro")
			return false
		}
			
		document.cookie = "PATH_FILE = ../" + document.forms[0].hdnArquivo.value
		return true
	}
}

function ResgatarArquivo()
{
	with (document.forms[0])
	{
		if (!ValidarDM(txtPedido)) return
		hdnAcao.value = "ResgatarArquivo"
		target = "IFrmProcesso"
		action = "ProcessoCartasProv.asp"
		submit()
	}	
}

function EnviarEmail(dblPedId,strNomePro,strProEmail,strLink,dblDocId)
{
	with (document.forms[0])
	{
		hdnEmailProvedor.value = strProEmail
		hdnNomeProvedor.value = strNomePro
		hdnPedId.value = dblPedId
		hdnArquivo.value = strLink
		hdnDocId.value = dblDocId
		document.cookie = "PATH_FILE = ../" + dblDocId

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
</head>
<body>
<table cellspacing="1" cellpadding="1" border=0 width="760">
	<tr >
		<th colspan=2>&nbsp;•&nbsp;Atualizar Carta Provedor</th>
	</tr>	
	<tr class=clsSilver>
	<td>
		Pedido de Acesso
	</td>
	<form name=Form1 method=Post>
	<input type=hidden name=hdnAcao>
	<input type=hidden name=hdnEmailProvedor>
	<input type=hidden name=hdnNomeProvedor>
	<input type=hidden name=hdnPedId>
	<input type=hidden name=hdnArquivo>
	<input type=hidden name=hdnDocId>
	<td>
		<input type="text" class="text" name="txtPedido" value="DM-" maxlength="13" size="15">
		<input type="button" class="button" name="btnProcurarDM" value="Procurar" onClick="ResgatarArquivo()" accesskey="P" onmouseover="showtip(this,event,'Procurar Arquivo (Alt+P)');">
		</td>
	</tr>
	</form>
	<form action="filePost.asp" enctype="multipart/form-data" method="post" name="f" id="f" target="IFrmProcesso" onSubmit="return Check()">
		<tr class=clsSilver>
			<td>
				<font class="clsObrig">:: </font>Selecione o arquivo
			</td>
			<td>	
				<input class=text type="file" name="txtFile" size="50" />
			</td>	
		<tr>	
			<td colspan=2 align=center>
				<input class=button type="submit" value="Enviar" />
			</td>
		</tr>	
	</form>
</table>
<span id=spnLista></span>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</form>
</body>