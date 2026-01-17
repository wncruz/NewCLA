<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: TipoContrato.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera Tipo de Contrato
%>
<!--#include file="../inc/data.asp"-->
<%
Dim objRSTCntr
Dim strDesc		
Dim strQtdeMeses
Dim dblID

dblID = Request.QueryString("ID")
if Trim(dblID) = "" then
	 dblID  = Request.Form("hdnID")
End if	 

if request("btnGravar")="Gravar" then
	if dblID="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	end if
		Vetor_Campos(2)="adWChar,50,adParamInput,"& ucase(request("txtDesc"))
		Vetor_Campos(3)="adInteger,9,adParamInput," & request("cboMeses")
		Vetor_Campos(4)="adInteger,2,adParamOutput,0"  
	Call APENDA_PARAM("CLA_sp_ins_tipocontrato",4,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
End if

If dblID<> "" then
	set objRSTCntr = db.execute("CLA_sp_sel_tipocontrato " & dblID)
	if Not objRSTCntr.Eof and Not objRSTCntr.bof then
		strDesc		= TratarAspasHtml(objRSTCntr("tct_Desc"))
		strQtdeMeses= TratarAspasHtml(objRSTCntr("tct_meses"))
	End if
Else
	strDesc		= TratarAspasHtml(Request.Form("txtDesc"))
	strQtdeMeses= TratarAspasHtml(Request.Form("cboMeses"))
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="tipocontrato.asp" method="post" onSubmit="return checa(this)" name="Form1" >
<input type=hidden name=hdnID value="<%=dblID%>" >
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtDesc,"A Descrição")) return false;
	if (!ValidarCampos(f.cboMeses,"A Quantidade de Meses")) return false;
	return true;
}
</script>
<tr><td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Vigencia de Contrato</p></th>
</tr>
<tr class=clsSilver>
<td width ="80"><font class="clsObrig">:: </font>Descrição</td>
<td>
<input type="text" class="text" name="txtDesc" value="<%=strDesc%>" maxlength="30" size="60" >
</td>
</tr>
<tr class=clsSilver>
<td nowrap><font class="clsObrig">:: </font>Qtde. Meses</td>
<td>
	<select name="cboMeses">
		<option value=""></option>
		<option value="12" <%if trim(strQtdeMeses) = "12" then Response.Write "selected" end if%>>12</option>
		<option value="24" <%if trim(strQtdeMeses) = "24" then Response.Write "selected" end if%>>24</option>
		<option value="36" <%if trim(strQtdeMeses) = "36" then Response.Write "selected" end if%>>36</option>
		<option value="48" <%if trim(strQtdeMeses) = "48" then Response.Write "selected" end if%>>48</option>
		<option value="60" <%if trim(strQtdeMeses) = "60" then Response.Write "selected" end if%>>60</option>
		<option value="74" <%if trim(strQtdeMeses) = "74" then Response.Write "selected" end if%>>74</option>
		<option value="0"  <%if trim(strQtdeMeses) = "0" then Response.Write "selected" end if%>>Indeterminado</option>
		<option value="1"  <%if trim(strQtdeMeses) = "1" then Response.Write "selected" end if%>>Temporário</option>
	</select>
</td>
</tr>
</table>
<table width="760" border=0>
<tr>
	<td colspan=2 align="center"><br>
		<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnID.value='';LimparForm();setarFocus('txtDesc');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('tipocontrato_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
</form>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('txtDesc')
//-->
</SCRIPT>
</html>
<%
Set objRSTCntr = Nothing
DesconectarCla()
%>
