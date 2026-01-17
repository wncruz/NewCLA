<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: usuario_main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remover usuários do sistema
%>
<!--#include file="../inc/data.asp"-->
<%
If Trim(Request.Form("hdnAcao")) = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_usuario")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post" >
<input type=hidden name=hdnAcao>
<script language="JavaScript">
function ProcurarUsuario()
{
	with (document.forms[0])
	{
		if (txtBusca.value != "")
		{
			hdCurrentPage.value = 1
			submit()
		}
		else
		{
			alert("Informe o username!")
			txtBusca.focus()
			return
		}
		
	}
}
</script>
<tr>
	<td width=100% >
		<table border="0" cellspacing="1" cellpadding=0 width="760">
			<tr>
				<th colspan=2><p align="center">Cadastro de Usuário</p></th>
			</tr>
			<tr class=clsSilver>
				<td >Busca (Username)&nbsp;&nbsp;</td>
				<td ><input type=text name=txtBusca maxlength=30 size=30 class="text" value="<%=TratarAspasHtml(Trim(Request.Form("txtBusca")))%>" >&nbsp;
				<input type="button" class="button" name=btnProcurar value=Procurar onclick="ProcurarUsuario()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th>&nbsp;Nome</th>
	<th>&nbsp;E-mail</th>
	<th>&nbsp;Ramal</th>
	<th>&nbsp;Username</th>
	<th>&nbsp;Status</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
</td>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass

if Trim(Request.Form("txtBusca")) <> ""  then
	strSql = "CLA_sp_view_usuario 0,'" & TratarAspasSQl(Trim(Request.Form("txtBusca"))) & "'"

Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td width="200"><a href="usuario.asp?ID=<%=objRSPag("Usu_ID") %>">&nbsp;<%=TratarAspasHtml(objRSPag("Usu_Nome"))%></a> </td>
			<td width="70">&nbsp;<%=TratarAspasHtml(objRSPag("Usu_Email"))%></td>
			<td width="70">&nbsp;<%=TratarAspasHtml(objRSPag("Usu_Ramal"))%></td>
			<td width="70">&nbsp;<%=TratarAspasHtml(objRSPag("Usu_Username"))%></td>
			<td width="70">&nbsp;
			<% strStatus=TratarAspasHtml(objRSPag("Usu_Inativo"))
			  if  strStatus = "S" then
				Response.Write "INATIVO"			
			  else
				Response.Write "ATIVO"	
			  end if
			%>			
			</td>
			<td ><input  type="checkbox" name="excluir" value="<%=objRSPag("Usu_ID")%>" onClick="AddSelecaoChk(this)"></td>
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
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('usuario.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<input type="button" class="button" name="btnExcluir" value="Excluir" onclick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
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
