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
	Vetor_Campos(5)="adInteger,2,adParamInput," & ucase(Trim(request("cboTecnologia")))

	Call APENDA_PARAM("CLA_sp_ins_FabricanteONT",5,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	
End if

If dblId<> "" then
	Set objRSCid = db.execute("CLA_sp_sel_FabricanteONT " & dblId)
	if Not objRSCid.Eof And Not objRSCid.Bof then
		strNome = TratarAspasHtml(Trim(objRSCid("Font_Nome")))
		dblATFId = TratarAspasHtml(Trim(objRSCid("ATF_id")))
	End if
Else
	strNome= TratarAspasHtml(Trim(Request.Form("txtNome")))
	dblATFId = TratarAspasHtml(Trim(request("cboTecnologia")))
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="cad_fabricanteONT.asp" method="post" onSubmit="return checa(this)">
<input type=hidden name=hdnId value=<%=dblId%>>
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtNome,"O Nome do Fabricante")) return false;
	if (!ValidarCampos(f.cboTecnologia,"Tecnologia")) return false;

	return true;
}
</script>
<tr><td >
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Fabricante ONT | EDD</p></th>
</tr>
<tr>
	<td>
	<font class="clsObrig">:: </font>Tecnologia
	</td>
	<td nowrap colspan=2>
		<Select name=cboTecnologia>
			<Option value="">:: TECNOLOGIA EBT</Option>
			<%
			set objRS = db.execute("CLA_sp_sel_tecnologiaFab 0")
			While not objRS.Eof
				strItemSel = ""
				if Trim(dblATFId) = Trim(objRS("ATF_ID")) then strItemSel = " Selected " End if
				Response.Write "<Option value=" & objRS("ATF_ID") & strItemSel & ">" & objRS("Sigla") & "</Option>"
				objRS.MoveNext
			Wend
			strItemSel = ""
			%>
			</Select>
		
	</td>
</tr>
<tr class=clsSilver>
<td>
<font class="clsObrig">:: </font>Nome
</td>
<td>
<input type="text" class="text" name="txtNome" value="<%=strNome%>" maxlength="40" size="60" onKeyUp="ValidarTipo(this,10)">
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
