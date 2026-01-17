<!--#include file="../inc/data.asp"-->
<%
Dim dblId
Dim strNome

dblId = Request.QueryString("ID") 
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId") 
End if

If request("btnGravar")="Gravar" Then

	Vetor_Campos(1)="adInteger,2,adParamInput," & dblId
	Vetor_Campos(2)="adWChar,40,adParamInput,"& ucase(Trim(request("txtNome")))
	Vetor_Campos(3)="adWChar,30,adParamInput,"& strloginrede
	Vetor_Campos(4)="adInteger,2,adParamOutput,0"

	Call APENDA_PARAM("CLA_sp_ins_Empreiteira",4,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	
End if

If dblId<> "" then
	Set objRSCid = db.execute("CLA_sp_sel_Empreiteira " & dblId)
	if Not objRSCid.Eof And Not objRSCid.Bof then
		strNome = TratarAspasHtml(Trim(objRSCid("Emp_Nome")))
	End if
Else
	strNome= TratarAspasHtml(Trim(Request.Form("txtNome")))
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="cad_fabricanteONT.asp" method="post" onSubmit="return checa(this)">
<input type=hidden name=hdnId value=<%=dblId%>>
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtNome,"O Nome da Empreiteira")) return false;

	return true;
}
</script>
<tr><td >
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Fabricante ONT</p></th>
</tr>
<tr class=clsSilver>
<td>
<font class="clsObrig">:: </font>Nome
</td>
<td>
<input type="text" class="text" name="txtNome" value="<%=strNome%>" maxlength="40" size="60" onKeyUp="ValidarTipo(this,1)">
</td>
</tr>
</table>

<table width="760">
<tr>
	<td colspan=2 align="center">
		<br>
		<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('txtSigla');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('FabricanteONT_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
<table width="760">
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
Set objRSCid = Nothing
DesconectarCla()
%>
