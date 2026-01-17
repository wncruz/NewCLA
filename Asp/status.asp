<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Status.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera Status
%>
<!--#include file="../inc/data.asp"-->
<%
Dim strDesc
Dim strNotifica
Dim strGICN
Dim strGICL
Dim strGLA
Dim strGLAE
Dim strTipoStatus
Dim objRSSts
Dim dblID

dblID = Request.QueryString("ID")
if Trim(dblID) = "" then
	 dblID  = Request.Form("hdnID")
End if	 

if request("btnGravar")="Gravar" then

	if dblID="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,0"
	else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	end if

	Vetor_Campos(2)="adWChar,60,adParamInput,"& ucase(request("txtDesc"))
	Vetor_Campos(3)="adInteger,2,adParamInput,"& request("cboNotifica")
	Vetor_Campos(4)="adInteger,2,adParamInput,"& Cint("0" & Request.Form("chkGICN"))
	Vetor_Campos(5)="adInteger,2,adParamInput,"& Cint("0" & Request.Form("chkGICL"))
	Vetor_Campos(6)="adInteger,2,adParamInput,"& Cint("0" & Request.Form("chkGLA"))
	Vetor_Campos(7)="adInteger,2,adParamInput,"& Cint("0" & Request.Form("chkGLAE"))
	Vetor_Campos(8)="adInteger,2,adParamInput,"& Cint("0" & Request.Form("cboTipoSatus"))
	Vetor_Campos(9)="adInteger,2,adParamOutput,0"
	Vetor_Campos(10)="adInteger,2,adParamInput,"& Cint("0" & Request.Form("chkAVL"))

	Call APENDA_PARAM("CLA_sp_ins_status",10,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End if

if dblID<> "" then
	set objRSSts = db.execute("CLA_sp_sel_Status " & dblID)
	if Not objRSSts.Eof and Not objRSSts.Bof then
		strDesc = TratarAspasHtml(objRSSts("Sts_Desc"))
		if objRSSts("Sts_Notifica") then strNotifica = 1 else strNotifica = 0 End if
		strGICN 		= objRSSts("Sts_GICN")
		strGICL 		= objRSSts("Sts_GICL")	
		strGLA			= objRSSts("Sts_GLA")	
		strGICL 		= objRSSts("Sts_GICL")	
		strGLAE			= objRSSts("Sts_GLAE")
		strAVL			= objRSSts("Sts_AVL")	
		strTipoStatus	= objRSSts("Sts_Tipo")
	End if
Else
	strDesc		= TratarAspasHtml(Request.Form("txtDesc"))
	strNotifica = TratarAspasHtml(Request.Form("cboNotifica"))
	if Request.Form("chkGICN") = 1 then strGICN = true else strGICN = false End if
	if Request.Form("chkGICL") = 1 then strGICL = true else strGICL = false End if
	if Request.Form("chkGLA") = 1 then strGLA = true else strGLA = false End if
	if Request.Form("chkGLAE") = 1 then strGLAE = true else strGLAE = false End if
	if Request.Form("chkAVL") = 1 then strAVL = true else strAVL = false End if
	
	strTipoStatus = Request.Form("cboTipoSatus")
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="status.asp" method="post"onSubmit="return checa(this)" name="Form1" >
<input type=hidden name=hdnID value="<%=dblID%>" >
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtDesc,"A Descrição")) return false;
	if (!ValidarCampos(f.cboNotifica,"A Notificação")) return false;

	if ((f.cboNotifica.value == "1") && (!f.chkGICN.checked && !f.chkGICL.checked && !f.chkGLA.checked && !f.chkGLAE.checked))
	{
		alert("Selecione pelo menos um item de quem deve ser notificado!")
		return false;
	}
	if (!ValidarCampos(f.cboTipoSatus,"O Tipo de Status")) return false;
	return true;
}
</script>
<tr><td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=5><p align="center">Cadastro de Status</p></th>
</tr>
<tr class=clsSilver>
	<td width="80"><font class="clsObrig">:: </font>Descrição</td>
	<td colspan="4"><input type="text" class="text" name="txtDesc" value="<%=strDesc%>" maxlength="60" size="60"></td>
</tr>
<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Notifica</td>
<td colspan="4">
	<select name="cboNotifica">
		<option value=""></option>
		<option value="1" >SIM</option>
		<option value="0" >NAO</option>
	</select>
</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Quem deve ser Notificado</td>
	<td><input type="checkbox" value=1 name="chkGICN">GIC-N</td>
	<td><input type="checkbox" value=1 name="chkGICL">GIC-L</td>
	<td><input type="checkbox" value=1 name="chkGLA">GLA</td>
	<td><input type="checkbox" value=1 name="chkGLAE">GLA-E</td>
	<td><input type="checkbox" value=1 name="chkAVL">AVL</td>
</tr>
<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Tipo de Status</td>
<td colspan="4">
	<select name="cboTipoSatus">
		<option value=""></option>
		<option value="0" <%if strTipoStatus = 0 then Response.Write " selected " End if%>>Status Macro</option>
		<option value="1" <%if strTipoStatus = 1 then Response.Write " selected " End if%>>Status Detalhado</option>
	</select>
</td>
</tr>
</table>
<table width="760" border=0>
<tr>
	<td colspan=2 align="center"><br>
		<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnID.value='';LimparForm();setarFocus('txtDesc');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('status_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
<SCRIPT LANGUAGE=javascript>
<!--
function RetornarSel()
{
	with (document.forms[0])
	{
			cboNotifica.value = <%=lcase(strNotifica)%>
			chkGICN.checked = <%=lcase(strGICN)%>
			chkGICL.checked	= <%=lcase(strGICL)%>
			chkGLA.checked	= <%=lcase(strGLA)%> 
			chkGLAE.checked	= <%=lcase(strGLAE)%>
			chkAVL.checked	= <%=lcase(strAVL)%>
	}
}
setTimeout("RetornarSel();setarFocus('txtDesc')",300)
//-->
</SCRIPT>
</form>
</body>
</html>
<%
Set objRSSts = Nothing
DesconectarCla()
%>
