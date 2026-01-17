<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: CadTipoRadio.asp
'	- Descrição			: Cadastra/Altera Tipos de Radios
strDataAtual =  right("00" & day(now),2) & "/" & right("00" & month(now),2) & "/" & year(now)
%>

<!--#include file="../inc/data.asp"-->
<%
Dim dblId
Dim strDescricao				
Dim strTipoRadio
Dim strDtDesaTivacao
Dim objRSPla

dblId = request("ID")
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId") 
End if
if Trim(Request.Form("hdnAcao")) = "Gravar" then

	if dblId="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblId
	end if
	
	
	Vetor_Campos(2)="adWChar,20,adParamInput,"&	Trim(Request.Form("txtTipoRadio"))
	Vetor_Campos(3)="adWChar,30,adParamInput,"&	Trim(Request.Form("txtDescricao"))
	if Trim(Request.Form("txtDTDesativacao")) = "" then 
		Vetor_Campos(4)="adDate,8,adParamInput,"	
	else
		Vetor_Campos(4)="adDate,8,adParamInput,"&	year(Trim(Request.Form("txtDTDesativacao"))) & "/" & Month (Trim(Request.Form("txtDTDesativacao"))) & "/" & day(Trim(Request.Form("txtDTDesativacao")))
	end if 
			
	Vetor_Campos(5)="adWChar,30,adParamInput,"&	strUserName
	Vetor_Campos(6)="adInteger,2,adParamOutput,0"  
	
	Call APENDA_PARAM("CLA_sp_ins_TipoRadio",6,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End if

if dblId<> "" then

	Set objRSPla = db.execute("CLA_sp_sel_TipoRadio " & dblId)

	if Not objRSPla.Eof and not objRSPla.Bof then
		strTipoRadio	    = TratarAspasHtml(trim(objRSPla("Trd_TipoRadio")))
		strDescricao		= TratarAspasHtml(trim(objRSPla("Trd_Descricao")))		
		strDtDesaTivacao    = TratarAspasHtml(trim(objRSPla("Trd_DtDesativacao")))		
	End if
Else
	strDescricao		= TratarAspasHtml(Request.Form("txtDescricao"))
	strTipoRadio		= TratarAspasHtml(Request.Form("txtTipoRadio"))
	strDtDesaTivacao    = TratarAspasHtml(Request.Form("txtDTDesativacao"))
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="servico.asp" method="post" >
<input type=hidden name=hdnId value="<%=dblId%>">
<input type=hidden name=hdnAcao>
<SCRIPT LANGUAGE="JavaScript">
function GravarTipoRadio() 
{
	with (document.forms[0]){
		if (!ValidarCampos(txtTipoRadio,"O Tipo de Radio")) return false;
		if (!ValidarCampos(txtDescricao,"A Descrição do Radio")) return false;
		if (txtDtDesativacao.value != ""){
			if (!CompararData('<%=strDataAtual%>',txtDtDesativacao.value,1,"A Data de Desativação não deve ser menor que a data de cadastro.")) return false 
		}
				
		hdnAcao.value = "Gravar"
		target = self.name
		action = "CadTipoRadio.asp"
		submit()
		return true;
	}
}
</script>
<tr><td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Tipo de Radio</p></th>
</tr>
<tr class=clsSilver>
	<td width="150"><font class="clsObrig">:: </font>Tipo de Radio</td>
	<td>
		<input type="text" class="text" name="txtTipoRadio" value="<%=strTipoRadio%>" maxlength="20" size="20">
	</td>
</tr>
<tr class=clsSilver>
	<td width="150"><font class="clsObrig">:: </font>Descrição</td>
	<td>
		<input type="text" class="text" name="txtDescricao" value="<%=strDescricao%>" maxlength="30" size="50">
	</td>
</tr>
<tr class=clsSilver>
	<td width="150">&nbsp;&nbsp;&nbsp;&nbsp;Data de Desativacao</td>
	<td>
		<input type="text" class="text" name="txtDtDesativacao" value="<%= strDtDesaTivacao %>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa)
	</td>
</tr>
<td>
</table>
<table width="760" border=0>
<tr>
<td colspan=2 align="center"><br>
	<input type="button" class="button" name="btnGravar" value="Gravar" onclick="GravarTipoRadio()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
	<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('txtTipoRadio')" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
	<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('CadTipoRadio_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
setarFocus('txtTipoRadio')
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
