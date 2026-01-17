<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Velocidade.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera Velocidade
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblId
Dim strDesc
Dim strOrdem 
Dim objRSVel

dblId = request("ID")
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId") 
End if

If Trim(Request.Form("btnGravar")) = "Gravar" then
	If dblId="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblId
	End if
	
	'response.write "<BR>" & ucase(request("txtVel"))  
	'response.write "<BR>" & ucase(request("txtUnid"))  
		'response.write "<BR>" & ucase(request("txtMult")) 
	Vetor_Campos(2)="adWChar,50,adParamInput,"& ucase(request("txtDesc"))
	Vetor_Campos(3)="adInteger,2,adParamInput,"& ucase(request("txtOrdem"))
	Vetor_Campos(4)="addouble,9,adParamInput,"& ucase(request("txtVel")) 
	Vetor_Campos(5)="adWChar,1,adParamInput,"& ucase(request("txtUnid"))  
	Vetor_Campos(6)="adInteger,2,adParamInput,"& ucase(request("txtMult"))  	
	Vetor_Campos(7)="adInteger,2,adParamOutput,0"  	
	
	Call APENDA_PARAM("CLA_sp_ins_velocidade",7,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
End if

If dblId<> "" then
	Set objRSVel = db.execute("CLA_sp_sel_velocidade " & dblId)
	if Not objRSVel.Eof and Not objRSVel.Bof then
		strDesc	= TratarAspasHtml(objRSVel("Vel_Desc"))
		strOrdem = TratarAspasHtml(objRSVel("Vel_Ordem"))
		
		strVel	= TratarAspasHtml(objRSVel("Vel_Velocidade"))
		strUnid = TratarAspasHtml(objRSVel("Vel_Unidade"))
		strMult	= TratarAspasHtml(objRSVel("Vel_Multiplicador"))				
	End if
Else
	strDesc	= TratarAspasHtml(Request.Form("txtDesc"))
	strOrdem = TratarAspasHtml(Request.Form("txtOrdem"))
	
	strVel	= TratarAspasHtml(Request.Form("txtVel"))
	strUnid = TratarAspasHtml(Request.Form("txtUnid"))
	strMult	= TratarAspasHtml(Request.Form("txtMult"))
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="velocidade.asp" method="post" onSubmit="return checa(this)">
<input type=hidden name=hdnId value="<%=dblId%>">

<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtDesc,"A Descrição")) return false;
	if (!ValidarCampos(f.txtOrdem,"A Ordem")) return false;
	if (!ValidarCampos(f.txtVel,"A Velocidade")) return false;
	if (!ValidarCampos(f.txtUnid,"A Unidade")) return false;	
		if (!ValidarCampos(f.txtMult,"O Multiplicador")) return false;	
	return true;
}
</script>

<tr><td >
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Velocidade</p></th>
</tr>
<tr class=clsSilver>
<td width="150"><font class="clsObrig">:: </font>Descrição</td>
<td>
	<input type="text" class="text" name="txtDesc" value="<%=strDesc%>" maxlength="30" size="35" >
</td>
</tr>
<tr class=clsSilver>
<td width="150"><font class="clsObrig">:: </font>Ordenação</td>
<td>
	<input type="text" class="text" name="txtOrdem" value="<%=strOrdem%>" maxlength="4" size="4" onKeyUp="ValidarTipo(this,0)">
</td>
</tr>


<tr class=clsSilver>
<td width="150"><font class="clsObrig">:: </font>Velocidade</td>
<td>
	<input type="text" class="text" name="txtVel" value="<%=strVel%>" maxlength="12" size="10"> <i>(Utilizar vírgula para separação decimal)</i>
</td>
</tr>
<tr class=clsSilver>
<td width="150"><font class="clsObrig">:: </font>Unidade (B, K, M, G)</td>
<td>
	<input type="text" class="text" name="txtUnid" value="<%=strUnid%>" maxlength="1" size="4" onKeyUp="ValidarTipo(this,1)">
</td>
</tr>
<tr class=clsSilver>
<td width="150"><font class="clsObrig">:: </font>Multiplicador</td>
<td>
	<input type="text" class="text" name="txtMult" value="<%=strMult%>" maxlength="4" size="4" onKeyUp="ValidarTipo(this,0)">
</td>
</tr>



</table>
<table width="760" border=0>
<tr>
	<td colspan=2 align="center"><br>
		<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value='';LimparForm();setarFocus('txtDesc');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('velocidade_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
Set objRSVel = Nothing
DesconectarCla()
%>
