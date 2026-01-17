<%
'•EXPERT INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: interface_main.asp
'	- Responsável		: PRSS
'	- Descrição			: Lista/Remove interfaces no sistema utilizadas pela Solicitacao.asp
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
Dim strTplNome
Dim objRSLogr
Dim dblID

dblID = Request.QueryString("ID")
if Trim(dblID) = "" then
	 dblID  = Request.Form("hdnID")
End if	 

if Trim(Request.Form("btnGravar")) = "Gravar" then

	if dblID="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	else
		Vetor_Campos(1)="adInteger,2,adParamInput," & dblID
	end if
	Vetor_Campos(2)="adWChar,15,adParamInput,"& ucase(request("txtNome"))
	Vetor_Campos(3)="adWChar,100,adParamInput,"& ucase(request("txtDesc"))
	Vetor_Campos(4)="adInteger,2,adParamOutput,0"
	Call APENDA_PARAM("CLA_sp_ins_interface",4,Vetor_Campos)
	ObjCmd.Execute'pega dbbtnGravar
	DBAction = ObjCmd.Parameters("RET").value

End if

if dblID<> "" then
	Set objRSLogr = db.execute("CLA_sp_sel_interface null," & Trim(dblID))
	if Not objRSLogr.Eof and Not objRSLogr.Bof then
		strTplNome = TratarAspasHtml(Trim(objRSLogr("ITF_Nome")))
		strTplDesc = TratarAspasHtml(Trim(objRSLogr("ITF_Desc")))
	End if
Else
	strTplNome	= TratarAspasHtml(Request.Form("txtNome"))
	strTplDesc	= TratarAspasHtml(Request.Form("txtDesc"))
end if
%>
<form action="interface.asp" method="post" onSubmit="return checa(this)">
<input type=hidden name=hdnID value="<%=dblID%>" >
<SCRIPT LANGUAGE="JavaScript">
function checa(f)
{
	if (!ValidarCampos(f.txtNome,"O Nome da Interface")) return false;
	return true;
}
</script>
<tr><td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Interfaces</p></th>
</tr>
<tr class=clsSilver>
<td>
<font class="clsObrig">:: </font>Nome
</td>
<td>
<input type="text" class="text" name="txtNome" value="<%=strTplNome%>" maxlength="15" size="20">
</td>
</tr>
<tr class=clsSilver>
<td>
<font class="clsObrig"> </font>Descrição
</td>
<td>
<input type="text" class="text" name="txtDesc" value="<%=strTplDesc%>" maxlength="100" size="50">
</td>
</tr>
</table>
<table width="760" border=0>
<tr>
<td colspan=2 align="center"><br>
	<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
	<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnID.value='';LimparForm();setarFocus('txtNome');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
	<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('interface_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
setarFocus('txtNome');
//-->
</SCRIPT>

</html>
<%
Set objRSLogr = Nothing
DesconectarCla()
%>
