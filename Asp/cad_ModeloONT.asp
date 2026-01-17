<!--#include file="../inc/data.asp"-->
<%
Dim dblId
Dim strNome

dblId = Request.QueryString("ID") 
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId") 
End if

If request("hdnAcao")="Gravar" Then
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblId
	Vetor_Campos(2)="adInteger,8,adParamInput,"& ucase(Trim(request("cboFabricante")))
	Vetor_Campos(3)="adWChar,40,adParamInput,"& ucase(Trim(request("txtModelo")))
	Vetor_Campos(4)="adInteger,2,adParamInput,"& ucase(Trim(request("txtQtdPorta1")))'GEth
	Vetor_Campos(5)="adInteger,2,adParamInput,"& ucase(Trim(request("txtQtdPorta2")))'FEth
	Vetor_Campos(6)="adInteger,2,adParamInput,"& ucase(Trim(request("txtQtdPorta3")))'E1
	Vetor_Campos(7)="adWChar,30,adParamInput,"& strloginrede
	Vetor_Campos(8)="adInteger,2,adParamOutput,0"
	Vetor_Campos(9)="adInteger,2,adParamInput,"& ucase(Trim(request("txtQtdPorta4")))'VDSL,

	Call APENDA_PARAM("CLA_sp_ins_TipoONT",9,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
End if

If dblId<> "" then
	Set objRSCid = db.execute("CLA_sp_sel_TipoONT " & dblId & ",0")
	if Not objRSCid.Eof And Not objRSCid.Bof then
		strFont = TratarAspasHtml(Trim(objRSCid("Font_ID")))
		strModelo = TratarAspasHtml(Trim(objRSCid("Tont_Modelo")))
		
		strPortaGEth = TratarAspasHtml(Trim(objRSCid("GEth")))
		strPortaFEth = TratarAspasHtml(Trim(objRSCid("FEth")))
		strPortaE1 = TratarAspasHtml(Trim(objRSCid("E1")))
		strPortaVDSL = TratarAspasHtml(Trim(objRSCid("VDSL")))
		
	End if
Else
	strFont = TratarAspasHtml(Trim(Request.Form("cboFabricante")))
	strModelo = TratarAspasHtml(Trim(Request.Form("txtModelo")))
	strPortaGEth = TratarAspasHtml(Trim(Request.Form("txtQtdPorta1")))
	strPortaFEth = TratarAspasHtml(Trim(Request.Form("txtQtdPorta2")))
	strPortaE1 = TratarAspasHtml(Trim(Request.Form("txtQtdPorta3")))
	strPortaVDSL = TratarAspasHtml(Trim(Request.Form("txtQtdPorta4")))
	
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="cad_ModeloONT.asp" method="post" onSubmit="return checa(this)">
<input type=hidden name=hdnId value=<%=dblId%>>
<input type=hidden name=hdnacao>
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtNome,"O Nome do Fabricante")) return false;


	return true;
}

function Gravar()
{
	document.forms[0].hdnacao.value = 'Gravar';
	document.forms[0].submit();
}

function visualizarportas()
{
	document.forms[0].hdnacao.value = 'Visualizar';
	document.forms[0].submit();
}
</script>
<tr><td >
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Modelo ONT | EDD</p></th>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Fabricante</td>
	<td>
		<select name="cboFabricante" >
			<Option value=""></Option>
			<% 
			set objRS = db.execute("CLA_sp_sel_FabricanteONT 0")
			While not objRS.Eof 
				strSel = ""
				if Trim(objRS("Font_ID")) = Trim(strFont) then strSel = " Selected " End if
				Response.Write "<Option value=" & objRS("Font_ID")& strSel & ">" & objRS("Font_Nome") & " / " & objRS("sigla") &  "</Option>"
				objRS.MoveNext
			Wend
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Modelo</td>
	<td><input type="text" class="text" name="txtModelo" value="<%=strModelo%>" maxlength="40" size="60" onKeyUp="ValidarTipo(this,10)"></td>
</tr>
<tr class=clsSilver>	
	<td><font class="clsObrig">:: </font>Qtd portas Giga Ethernet</td>
	<td><input type="text" class="text" name="txtQtdPorta1" value="<%=strPortaGEth%>" maxlength="2" size="4" onKeyUp="ValidarTipo(this,0)"></td>
</tr>
<tr class=clsSilver>	
	<td><font class="clsObrig">:: </font>Qtd portas Fast Ethernet</td>
	<td><input type="text" class="text" name="txtQtdPorta2" value="<%=strPortaFEth%>" maxlength="2" size="4" onKeyUp="ValidarTipo(this,0)"></td>
</tr>
<tr class=clsSilver>	
	<td><font class="clsObrig">:: </font>Qtd portas E1</td>
	<td><input type="text" class="text" name="txtQtdPorta3" value="<%=strPortaE1%>" maxlength="2" size="4" onKeyUp="ValidarTipo(this,0)"></td>
</tr>
<tr class=clsSilver>	
	<td><font class="clsObrig">:: </font>Qtd portas VDSL</td>
	<td><input type="text" class="text" name="txtQtdPorta4" value="<%=strPortaVDSL%>" maxlength="2" size="4" onKeyUp="ValidarTipo(this,0)"></td>
</tr>
</table>

<table width="760">
<tr>
	<td colspan=2 align="center">
		<br>
		<%if 1=0 then%>
			<input type="button" class="button" name="btnVisualizar" value="Visualizar" accesskey="I" onmouseover="showtip(this,event,'Listar Portas (Alt+I)');" onClick="javascript:visualizarportas();">&nbsp;
		<%end if%>
		<input type="button" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');" onClick="javascript:Gravar();">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('txtSigla');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('ModeloONT_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
	setarFocus('cboFabricante');
//-->
</SCRIPT>
</html>
<%
Set objRSCid = Nothing
DesconectarCla()
%>
