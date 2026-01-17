<%
'•EXPERT INFORMATICA
'	- Sistema				: CLA
'	- Arquivo				: Grupo_Main.asp
'	- Responsável		: LPEREZ
'	- Descrição			: Lista/Remove Grupos
%>
<!--#include file="../inc/data.asp"-->
<%
if Trim(Request.Form("hdnAcao")) = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_GrupoCliente")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post"  onSubmit="return ConfirmarRemocao()">
<input type=hidden name=hdnAcao>
<tr>
<td >
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=3><p align="center">Cadastro de Grupos</p></th>
</tr>
<tr>
<th>&nbsp;Codigo</th>
<th>&nbsp;Descrição</th>
<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
</td>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_sel_GrupoCliente 0"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td width="150"><a href="grupo.asp?ID=<%=objRSPag("GCli_ID") %>">&nbsp;<%=TratarAspasHtml(objRSPag("GCli_ID"))%></a></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("GCli_Descricao"))%></td>
			<td ><input  type="checkbox" name="excluir" onClick="AddSelecaoChk(this)" value="<%=objRSPag("GCli_ID")%>"></td>
		</tr>
		<%
		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	Next
End if
%>
</table>
<tr>
	<td align=center>
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('grupo.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<input type="button" class="button" name="btnExcluir" value="Excluir" onClick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
</body>
</html>
<%
Set objRSPag = Nothing
DesconectarCla()
%>
