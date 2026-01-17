<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AssocUserCef_main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Associação de usuário com centro funcional
%>
<!--#include file="../inc/data.asp"-->
<%
If request("hdnAcao")="Excluir" then
	Call ExcluirRegistro("CLA_sp_del_usuarioctfc")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post"  onSubmit="return ConfirmarRemocao()">
<script language="JavaScript">
function ProcurarUsuario()
{
	with (document.forms[0])
	{
		if (txtBusca.value != "")
		{
			action = "AssocUserCef_Main.asp"
			hdCurrentPage.value = 1
			submit()
		}
		else
		{
			alert("Informe o Username!")
			txtBusca.focus()
			return
		}
	}
}
</script>
<tr><td width=100%>
		<table border="0" cellspacing="1" cellpadding=0 width="760">
			<tr>
				<th colspan=2><p align="center">Associação de Usuário com Centro Funcional</p>
				</th>
			</tr>
			<tr class=clsSilver>
				<td >User Name&nbsp;&nbsp;</td>
				<td ><input type=text name=txtBusca maxlength=30 class="text" value="<%=Server.HTMLEncode(Request.Form("txtBusca"))%>">&nbsp;
				<input type="button" class="button" name=btnProcurar value=Procurar onclick="ProcurarUsuario()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');"></td>
			</tr>
		</table>
	</td>
</tr>

<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th>&nbsp;Usuário</th>
	<th>&nbsp;Centro Funcional</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass

if Trim(Request.Form("txtBusca")) <> ""  then
	strSql = "CLA_sp_sel_usuarioctfc null,null,'" & TratarAspasSQL(Trim(Request.Form("txtBusca"))) & "'"

	Call PaginarRS(0,strSql)
	intCount=1
	if not objRSPag.Eof and not objRSPag.Bof then
		For intIndex = 1 to objRSPag.PageSize
			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
			%>
			<tr class=<%=strClass%>>
				<td ><a href="AssocUserCef.asp?ID=<%=objRSPag("UsuCtfc_ID") %>">&nbsp;<%=objRSPag("Usu_UserName")%></a></td>
				<td >&nbsp;<%=objRSPag("Ctf_AreaFuncional") & " - " & objRSPag("Cid_Sigla") & " " & objRSPag("Esc_Sigla") & " - " & objRSPag("Age_Sigla") & " - " & objRSPag("Age_Desc")%></td>
				<td ><input type="checkbox" name="excluir" value="<%=objRSPag("UsuCtfc_ID")%>" onClick="AddSelecaoChk(this)"></td>
			</tr>
			<%
			intCount = intCount+1
			objRSPag.MoveNext
			if objRSPag.EOF then Exit For
		Next
	End if
End if
%>
</table>
</td>
</tr>
<tr>
	<td align=center>
		<input type="hidden" name="hdnAcao" >  
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('AssocUserCef.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<input type="button" class="button" name="btnExcluir" value="Excluir" onClick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');" >
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
<tr>
<td>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
</td>
</tr>
</table>
</body>
</html>
<%
Set objRSPag = Nothing
DesconectarCla()
%>
