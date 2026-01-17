<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: RedirSolicitacao.asp
'	- Responsável		: Vital
'	- Descrição			: Cadatrar/Altera Redirecionamento de solicitação
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblID
Dim strLetra
Dim strUserNameRedir
Dim dblCefId
Dim objRSRedir

dblID = Request.QueryString("id")
if Trim(dblID) = "" then
	dblID = Request.Form("hdnId") 
End if

If Trim(Request.Form("hdnAcao")) = "Gravar" Then

	If dblId="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	End if

	Vetor_Campos(2)="adWChar,1,adParamInput,"& Trim(Request.Form("txtLetra"))
	Vetor_Campos(3)="adInteger,2,adParamInput," & Request.Form("hdnUserId") 
	Vetor_Campos(4)="adInteger,2,adParamInput,"& Request.Form("cboCef")
	Vetor_Campos(5)="adInteger,2,adParamOutput,0"  
	Call APENDA_PARAM("CLA_sp_ins_redirsolicitacao",5,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End if

If dblID <> "" then
	Set objRSRedir = db.execute("CLA_sp_sel_redirsolicitacao " & dblId)
	If Not objRSRedir.Eof and Not objRSRedir.Bof then
		strLetra	= TratarAspasHtml(objRSRedir("Rep_Letra"))
		strUserNameRedir = TratarAspasHtml(objRSRedir("Usu_UserName"))
		dblCefId	= TratarAspasHtml(objRSRedir("Ctfc_id"))
	End if
Else
	strLetra	= TratarAspasHtml(Request.Form("txtLetra"))
	strUserNameRedir = TratarAspasHtml(Request.Form("txtUserName"))
	dblCefId	= TratarAspasHtml(Request.Form("cboCef"))
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="redirsolicitacao.asp" method="post" onSubmit="return checa(this)" name="Form1">
<input type=hidden name=hdnId value="<%=dblID%>">
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtUserName,"O usuário")) return false;
	if (!ValidarCampos(f.txtLetra,"A Letra")) return false;
	if (!ValidarCampos(f.cboCef,"O Centro Funcional")) return false;
	return true;
}

function ProcurarUsuCef()
{
	with (document.forms[0])
	{
		if (!ValidarCampos(txtUserName,"O usuário")) return false;
		hdnAcao.value = "ResgatarUsuCef"
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}

function GravarInfo()
{
	with (document.forms[0])
	{
		if (checa(document.forms[0]))
		{
			hdnAcao.value = "Gravar"
			target = self.name
			action = "redirsolicitacao.asp"
			submit()
		}
	}
}
function LimparLocal()
{
	document.forms[0].hdnId.value = '';
	spnCef.innerHTML = '<Select name=cboCef><Option value=""""></Option></Select>'
}
</script>
<input type=hidden name="hdnAcao">
<input type=hidden name="hdnCefSel">
<input type=hidden name="hdnUserId" value="<%=Request.Form("hdnUserId")%>">
<tr>
	<td>
	<table border="0" cellspacing="1" cellpadding=0 width="760">
	<tr>
		<th colspan=2><p align="center">Cadastro de Redirecionamento de Solicitação</p></th>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Usuário (UserName)</td>
		<td>
			<input type=text class="text" name=txtUserName size=30 maxlength=30 value="<%=strUserNameRedir%>">&nbsp;
			<input type=button name=btnProcurar value=Procurar onclick="ProcurarUsuCef()" class=button>
		</td>
	</tr>

	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Letra de Redir. Pedido</td>
		<td><input type="text" class="text" name="txtLetra" value="<%=strLetra%>" maxlength="1" size="1"></td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Centro Funcional</td>
	<td>
		<span id="spnCef">
			<select name="cboCef">
				<option value=''></option>
			</select>
		</span>
	</td>
	</tr>
	</table>
	<tr>
		<td colspan=2 align="center"><br>
			<input type="button" class="button" name="btnGravar" value="Gravar" onclick="GravarInfo()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
			<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparLocal();LimparForm();setarFocus('txtUserName');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
			<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('redirsolicitacao_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
			<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
		</td>
	</tr>
	<table width="760">
		<tr>
			<td>
				<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
			</td>
		</tr>
	</table>
	<table>

</table>
<iframe	id			= "IFrmProcesso"
		name        = "IFrmProcesso" 
		width       = "0" 
		height      = "0"
		frameborder = "0"
		scrolling   = "no" 
		align       = "left">
</iFrame>
</td>
</tr>
</table>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
<%If Trim(dblID)<> "" then%>
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarUsuCef"
		hdnCefSel.value = "<%=dblCefId%>"
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
<%Else
	if Request.ServerVariables("CONTENT_LENGTH") <> 0 THEN
%>
		with (document.forms[0])
		{
			hdnAcao.value = "ResgatarUsuCef"
			hdnCefSel.value = "<%=Request.Form("cboCef")%>"
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()
		}
<%	
	End if
End if
%>
setarFocus('txtUserName')
//-->
</SCRIPT>
</html>
<%
Set objRSredir = Nothing
DesconectarCla()
%>
