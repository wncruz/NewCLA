<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: usuario.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera usuários do sistema
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblId
Dim strNome		
Dim strEmail	
Dim strRamal	
Dim strUserNameUsuario
Dim objRSUsu

dblId = request("id")
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId") 
End if	

If Trim(Request.Form("btnGravar")) = "Gravar" then
	if dblId="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblId
	end if
	Vetor_Campos(2)="adWChar,30,adParamInput,"& ucase(request.Form("txtNome"))
	Vetor_Campos(3)="adWChar,30,adParamInput,"& lcase(request.Form("txtEmail"))
	Vetor_Campos(4)="adWChar,30,adParamInput,"& request.Form("txtRamal")
	Vetor_Campos(5)="adWChar,30,adParamInput,"& lcase(request.Form("txtUserName"))
	Vetor_Campos(6)="adWChar,10,adParamInput," & strloginrede '-->PSOUTO 12/04/06
	Vetor_Campos(7)="adInteger,2,adParamOutput,0"  
	
	'Response.Write "<script language=javascript>alert('teste1');</script>"	
	Call APENDA_PARAM("CLA_sp_ins_usuario",7,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	
	'Response.Write "<script language=javascript>alert('" & DBAction & "');</script>"
	'Response.Write DBAction
	
	'if DBAction = "1" or DBAction = "2" Then

	'	Response.Write "<script language=javascript>alert('Registro Gravado com Sucesso!');</script>"

	'end if
	
	
End if

If dblId<> "" then
	Set objRSUsu = db.execute("CLA_sp_sel_usuario " & dblId)
	if Not objRSUsu.Eof and Not objRSUsu.bof then
		strNome		= Trim(TratarAspasHtml(objRSUsu("Usu_Nome")))
		strEmail	= Trim(TratarAspasHtml(objRSUsu("Usu_Email")))
		strRamal	= Trim(TratarAspasHtml(objRSUsu("Usu_Ramal")))
		strUserNameUsuario = Trim(TratarAspasHtml(objRSUsu("Usu_Username")))
	End if	
Else
	strNome		= Trim(TratarAspasHtml(Request.Form("txtNome")))
	strEmail	= Trim(TratarAspasHtml(Request.Form("txtEmail")))
	strRamal	= Trim(TratarAspasHtml(Request.Form("txtRamal")))
	strUserNameUsuario = Trim(TratarAspasHtml(Request.Form("txtUserName")))
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="usuario.asp" method="post" onSubmit="return checa(this)" name="Form1" >
<input type=hidden name=hdnId value=<%=dblId%>>
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{

	if (!ValidarCampos(f.txtNome,"O Nome")) return false;
	if (!ValidarCampos(f.txtEmail,"O Email")) return false;
	if (!ValidarEmail(f.txtEmail)) return false;
	if (!ValidarCampos(f.txtRamal,"O Ramal")) return false;
	if (document.forms[0].txtRamal.value.length < 7)
	{ 
		alert("Ramal Incompleto(N7)!")
		document.forms[0].txtRamal.focus()
		return false
	}
	if (!ValidarCampos(f.txtUserName,"O Username")) return false;
	return true;
}
</script>
<tr><td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Usuário</p></th>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Nome</td>
	<td><input type="text" class="text" name="txtNome" value="<%=strNome%>" maxlength="30" size="30"></td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>E-mail</td>
	<td><input type="text" class="text" name="txtEmail" value="<%=strEmail%>" maxlength="30" size="30" onBlur="ValidarEmail(this)"></td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Ramal</td>
	<td><input type="text" class="text" name="txtRamal" value="<%=strRamal%>" maxlength="7" size="9" onkeyUp="ValidarTipo(this,0)"></td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Username</td>
	<td><input type="text" class="text" name="txtUserName" value="<%=strUserNameUsuario%>" maxlength="30" size="30"></td>
</tr>

</table>
<table width="760" border=0>
<tr>
	<td colspan=2 align="center"><br>
		<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('txtNome');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('usuario_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
<table width="760" border=0>
<tr>
<td>
<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('txtNome')
//-->
</SCRIPT>

</html>
<%
Set objRSUsu = Nothing
DesconectarCla()
%>
