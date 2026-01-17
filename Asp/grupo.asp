<%
'•EXPERT INFORMATICA
'	- Sistema				: CLA
'	- Arquivo				: Grupo.asp
'	- Responsável		: LPEREZ
'	- Descrição			: Cadastra/Altera GRUPO
%>
<!--#include file="../inc/data.asp"-->
<%
Dim strDesc
Dim strID
Dim objRSGrp
Dim dblID

dblID = Request.QueryString("ID")
if Trim(dblID) = "" then
	 dblID  = Request.Form("hdnID")
End if	 

If dblID<> "" then
	Set objRSGrp = db.execute("CLA_sp_sel_GrupoCliente " & dblID)
	if Not objRSGrp.Eof and Not objRSGrp.Bof then
		strDesc = TratarAspasHtml(objRSGrp("GCli_Descricao"))	
		strID = TratarAspasHtml(objRSGrp("GCli_ID"))
	End if
Else
	strDesc = TratarAspasHtml(Request.Form("txtdesc"))
	strID = TratarAspasHtml(Request.Form("txtID"))
End if


if Request.Form("btnGravar")="Gravar" then
	If dblID="" then
		Vetor_Campos(1)="adInteger,4,adParamInput,0"
	Else
		Vetor_Campos(1)="adInteger,4,adParamInput,"& dblID
	End if
	Vetor_Campos(2)="adWChar,100,adParamInput,"& ucase(request("txtDesc"))
	Vetor_Campos(3)="adInteger,4,adParamInput,"& strID
	Vetor_Campos(4)="adInteger,4,adParamOutput,0"  
	
	Call APENDA_PARAM("CLA_sp_ins_grupo",4,Vetor_Campos)
	ObjCmd.Execute'pega dbbtnGravar
	DBAction = ObjCmd.Parameters("RET").value
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="grupo.asp" method="post" onSubmit="return checa(this)" name="Form1">
<input type=hidden name=hdnID value="<%=dblID%>" >
<SCRIPT LANGUAGE="JavaScript">
	function checa(f) 
	{
		if (!ValidarCampos(f.txtDesc,"Grupo")) return false;
	}
</script>
	<tr>
		<td>
		<table border="0" cellspacing="1" cellpadding=0 width="760">
			<tr>
				<th colspan=2><p align="center">Cadastro de Grupos</p></th>
			</tr>
			<tr class=clsSilver>
				<td>
					<font class="clsObrig"></font>&nbsp;&nbsp;&nbsp;Código
				</td>
				<td align='left'>
					<input type="text" class="text" name="txtID" value="<%=trim(strID)%>" language=javascript onKeyUp="ValidarNTipo(this,1,3)" maxlength="5" size="3" disabled>
				</td>
			</tr>
			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Grupo
				</td>
				<td align='left'>
					<input type="text" class="text" name="txtDesc" value="<%=trim(strDesc)%>" language=javascript onKeyUp="ValidarNTipo(this,1,3)" maxlength="20" size="8">
				</td>
			</tr>
		</table>
		<table width="760" border=0>
		<tr>
			<td colspan=2 align="center"><br>
				<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
				<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnID.value='';LimparForm();setarFocus('txtDesc');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
				<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('grupo_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
setarFocus('txtDesc')
</SCRIPT>
</body>
</html>
<%
Set objRSGrp = Nothing
DesconectarCla()
%>
