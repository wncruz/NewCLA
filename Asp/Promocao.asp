<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Promocao.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera Promoção
%>
<!--#include file="../inc/data.asp"-->
<%
Dim strDesc
Dim dblProId 
Dim strVisivel
Dim strSel
Dim objRSPrm
Dim dblID

dblID = Request.QueryString("ID")
if Trim(dblID) = "" then
	 dblID  = Request.Form("hdnID")
End if	 

If Trim(Request.Form("btnGravar")) ="Gravar" then

	If dblID="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,0"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	End if
	Vetor_Campos(2)="adInteger,2,adParamInput," & Request.Form("cboProvedor") 

	Vetor_Campos(3)="adWChar,50,adParamInput,"& ucase(Request.Form("txtDesc"))
	Vetor_Campos(4)="adInteger,2,adParamInput,"& Request.Form("chkVisivel") 
	Vetor_Campos(5)="adInteger,2,adParamOutput,0"  
	
	Call APENDA_PARAM("CLA_sp_ins_promocaoprovedor ",5,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End if

If dblID<> "" then
	set objRSPrm = db.execute("CLA_sp_sel_promocaoprovedor " & dblID)
	if Not objRSPrm.Eof and Not objRSPrm.Bof then
		strDesc		= TratarAspasHtml(objRSPrm("prm_desc"))
		dblProId = TratarAspasHtml(objRSPrm("pro_id"))
		strVisivel	= ""
		if Cint("0" & objRSPrm("prm_visivel")) = 1 then strVisivel = " checked " End if
	End if
Else
	strDesc		= TratarAspasHtml(Request.Form("txtDesc"))
	dblProId = TratarAspasHtml(Request.Form("cboProvedor"))
	strVisivel	= ""
	if Cint("0" & Trim(Request.Form("chkVisivel"))) = 1 then strVisivel = " checked " End if
End if

%>
<!--#include file="../inc/header.asp"-->
<form action="Promocao.asp" method="post" name="Form1">
<input type=hidden name=hdnID value="<%=dblID%>" >

<SCRIPT LANGUAGE="JavaScript">
function checa(f)
{

	if (!ValidarCampos(f.cboProvedor,"Provedor")) return false;
	if (!ValidarCampos(f.txtDesc,"Descrição")) return false;

	return true;
}
</script>
<tr><td>
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align=center>Cadastro de Promoção</p></th>
</tr>
<tr class=clsSilver>
	<td>
		<font class="clsObrig">:: </font>Provedor
	</td>
	<td>
		<select name="cboProvedor">
			<option value=""></option>
			<%	Set objRS = db.execute("CLA_sp_sel_provedor") 
				While Not objRS.Eof
					strSel = ""
					if Cdbl("0" & objRS("Pro_id")) = Cdbl("0" & dblProId) then strSel = " Selected " End if
					Response.Write "<Option value="& objRS("Pro_id") & strSel & ">" & objRS("Pro_Nome") & "</Option>"
					objRS.MoveNext
				Wend
				Set objRS = Nothing
			%>
		</select>
	</td>
</tr>

<tr class=clsSilver>
	<td>
		<font class="clsObrig">:: </font>Promoção
	</td>
	<td>
		<input type="text" class="text" name="txtDesc" value="<%=strDesc%>" maxlength="30" size="60">
	</td>
</tr>

<tr class=clsSilver>
	<td>
		&nbsp;&nbsp;&nbsp;&nbsp;Visível
	</td>
	<td>
		<input type=checkbox name=chkVisivel value=1 <%=strVisivel%> >
	</td>
</tr>
</table>
<table width=760>
<tr>
<td colspan=2 align="center"><br>
	<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
	<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnID.value='';LimparForm();setarFocus('cboProvedor');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
	<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('Promocao_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
	setarFocus('cboProvedor')
//-->
</SCRIPT>
</html>
<%
Set objRSPrm = Nothing
DesconectarCla()
%>
