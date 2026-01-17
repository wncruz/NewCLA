<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: RegimeContrato.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera Regime de contrato
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblId
Dim dblProId
Dim dblTctId
Dim objRSReg

dblId = request("ID")
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId")
End if

if request("btnGravar")="Gravar" then
	if dblId = "" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblId
	end if
	Vetor_Campos(2)="adInteger,3,adParamInput,"&request("cboProvedor")
	Vetor_Campos(3)="adInteger,3,adParamInput,"& request("cboTipoContrato")
	Vetor_Campos(4)="adInteger,2,adParamOutput,0"  
	Call APENDA_PARAM("CLA_sp_ins_regimecontrato",4,Vetor_Campos)
	ObjCmd.Execute'pega dbbtnGravar
	DBAction = ObjCmd.Parameters("RET").value
End if

If dblId<> "" then
	Set objRSReg = db.execute("CLA_sp_sel_regimecontrato "& dblId)
	if Not objRSReg.Eof And Not objRSReg.Bof then
		dblProId	= TratarAspasHtml(objRSReg("Pro_id"))
		dblTctId	= TratarAspasHtml(objRSReg("Tct_id"))
	End if
Else
	dblProId	= TratarAspasHtml(Request.Form("cboProvedor"))
	dblTctId	= TratarAspasHtml(Request.Form("cboTipoContrato"))
End if

%>
<!--#include file="../inc/header.asp"-->
<form action="regimecontrato.asp" method="post" onSubmit="return checa(this)">
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.cboProvedor,"O provedor")) return false;
	if (!ValidarCampos(f.cboTipoContrato,"O Tipo Contrato")) return false;
	return true;
}
</script>

<input type="hidden" name="hdnId" value="<%=dblId%>">
<tr><td >
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Regime Contrato</p></th>
</tr>
<tr class=clsSilver>
<td width="80"><font class="clsObrig">:: </font>Provedor</td>
<td>
	<select name="cboProvedor"> 
		<option value=""></option>
	<%
	Set objRS =db.execute("CLA_sp_sel_provedor 0")
	do while not objRS.eof 
	%> 
	<option value="<%=objRS("pro_id")%>" <%
	if cdbl("0" & objRS("Pro_ID")) = cdbl("0" & dblProId)  then 
		response.write "selected" 
	end if
	%>>
		<%=objRS("pro_nome")&"<br>"%>
	</option>
	<%
	objRS.movenext
	loop 
	%>
	</select>
</td>
</tr>
<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Tipo Contrato</td>
<td>
	<select name="cboTipoContrato"> 
		<option value=""></option>
	<%set objRS =db.execute("CLA_sp_sel_tipocontrato 0")
	do while not objRS.eof 
	%> 
	<option value="<%=objRS("tct_id")%>"<%
	if cdbl("0" & objRS("tct_id")) = cdbl("0" & dblTctId) then
		response.write "Selected"
	end if
	%>>
		<%=objRS("tct_desc")%>
	</option>
<%
objRS.movenext
loop 
%>
</td>
</tr>
</table>
<table width="760">
	<tr>
		<td colspan=2 align="center"><br>
			<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
			<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('cboProvedor');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
			<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('regimecontrato_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
</form>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('cboProvedor')
//-->
</SCRIPT>
</html>
<%
Set objRSReg = Nothing
DesconectarCla()%>
