<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Combo.asp
'	- Responsável		: EDAR
'	- Descrição			: Cadastra/Altera Combo
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblId
Dim strDesc				
Dim strCodigo			
Dim strSigla			
Dim objRSSer

dblId = request("ID")
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId") 
End if

	'response.write "<script>alert('"&dblId &"')</script>"

if Trim(Request.Form("hdnAcao")) = "Gravar" then

	if dblId="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblId
	end if

	Vetor_Campos(1)="adWChar,4,adParamInput,"&	Trim(Request.Form("txtCodigo"))
	Vetor_Campos(2)="adWChar,30,adParamInput,"&	Trim(Request.Form("txtDesc"))
	Vetor_Campos(3)="adWChar,10,adParamInput," & strloginrede 
	Vetor_Campos(4)="adInteger,2,adParamOutput,0"  

	Call APENDA_PARAM("CLA_sp_ins_newCombo",4,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End if

if dblId<> "" then

	Set objRSSer = db.execute("CLA_sp_sel_newCombo " & dblId)

	'response.write "<script>alert('"&objRSSer("newtec_sigla")&"')</script>"

	if Not objRSSer.Eof and not objRSSer.Bof then
		strDesc				= TratarAspasHtml(objRSSer("newcombo_nome"))
		strCodigo			= TratarAspasHtml(objRSSer("newcombo_id"))
		
	End if
Else
	strDesc					= TratarAspasHtml(Request.Form("txtDesc"))
	strCodigo				= TratarAspasHtml(Request.Form("txtCodigo"))
	
	
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="combo.asp" method="post" >
<input type=hidden name=hdnId value="<%=dblId%>">
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnCboTecnologia>
<input type=hidden name=hdnDesigServ>
<SCRIPT LANGUAGE="JavaScript">
function GravarCombo() 
{
	with (document.forms[0]){
		
		if (!ValidarCampos(txtDesc,"A Descrição")) return false;
		
		
		hdnAcao.value = "Gravar"
		target = self.name
		action = "combo.asp"
		submit()
		return true;
	}

}
	

</script>
<tr><td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Combo</p></th>
</tr>
<!--
<tr class=clsSilver>
	<td width="80"><font class="clsObrig">:: </font>Código</td>
	<td>
		<input type="text" class="text" name="txtCodigo" value="<%=strCodigo%>" maxlength="4" size="4">
	</td>
</tr>
-->

<tr class=clsSilver>
	<td width="80"><font class="clsObrig">:: </font>Descrição</td>
	<td>
		<input type="text" class="text" name="txtDesc" value="<%=strDesc%>" maxlength="30" size="60">
	</td>
</tr>


</table>
<table width="760" border=0>
<tr>
<td colspan=2 align="center"><br>
	
	<input type="button" class="button" name="btnGravar" value="Gravar" onclick="GravarCombo()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
	
	<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('txtDesc')" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
	<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('combo_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
	<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
</td>
</tr>
</table>

</td>
</tr>
</table>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('txtCodigo')
//-->
</SCRIPT>
<iframe	id			= "IFrmProcesso"
		name        = "IFrmProcesso"
		width       = "0"
		height      = "0"
		frameborder = "0"
		scrolling   = "no" 
		align       = "left">
</iFrame>

</html>
<%
Set objRSSer = Nothing
DesconectarCla()
%>
