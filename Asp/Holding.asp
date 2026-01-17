<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Holding.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera Holding
%>
<!--#include file="../inc/data.asp"-->
<%
Dim strDesc
Dim objRSHol
Dim dblID

dblID = Request.QueryString("ID")
if Trim(dblID) = "" then
	 dblID  = Request.Form("hdnID")
End if	 

If Trim(Request.Form("btnGravar")) = "Gravar" then

	If dblID="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	End if

	Vetor_Campos(2)="adWChar,50,adParamInput,"& ucase(request("txtDesc"))
	Vetor_Campos(3)="adInteger,2,adParamOutput,0"  
	
	Call APENDA_PARAM("CLA_sp_ins_holding ",3,Vetor_Campos)
	ObjCmd.Execute
	DBAction = ObjCmd.Parameters("RET").value

End if

If Trim(dblID) <> "" then
	set objRSHol = db.execute("CLA_sp_sel_holding " & dblID)
	If Not objRSHol.Eof and Not objRSHol.Bof then
		strDesc = TratarAspasHtml(Trim(objRSHol("hol_desc")))
	End if
Else	
	strDesc = TratarAspasHtml(Trim(Request.Form("txtDesc")))
End if

%>
<!--#include file="../inc/header.asp"-->
<form action="holding.asp" method="post" onSubmit="return checa(this)">
<input type=hidden name=hdnID value="<%=dblID%>" >
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtDesc,"Descrição")) return false;
	return true;
}
</script>
<tr>
	<td >
		<table border=0 cellspacing="1" cellpadding="0" width="760">
		<tr>
			<th colspan=2><p align="center">Cadastro de Holding</p></th>
		</tr>
		<tr class=clsSilver>
			<td>
				<font class="clsObrig">:: </font>Nome
			</td>
			<td>
				<input type="text" class="text" name="txtDesc" value="<%=strDesc%>" maxlength="30" size="60">
			</td>
		</tr>
		</table>
		<table width="760">
			<tr>
				<td colspan=2 align="center">
					<br>
					<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
					<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnID.value='';LimparForm();setarFocus('txtDesc');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
					<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('Holding_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
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
setarFocus('txtDesc')
//-->
</SCRIPT>

</html>
<%
Set objRSHol = Nothing
DesconectarCla()
%>
