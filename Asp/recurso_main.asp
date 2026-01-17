<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Recurso_main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Recurso
%>
<!--#include file="../inc/data.asp"-->
<%
If Trim(Request.Form("hdnAcao")) = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_recurso")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post" >
<input type=hidden name=hdnAcao>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=6><p align="center">Cadastro de Recurso</p></th>
</tr>
<tr>
	<th>&nbsp;Sistema</th>
	<th>&nbsp;Estação</th>
	<th>&nbsp;Distribuidor</th>
	<th width="250">&nbsp;Provedor</th>
	<th width="70">&nbsp;Plataforma</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_view_recurso 0"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td >&nbsp;<a href="recurso.asp?ID=<%=objRSPag("rec_id")%>"><%=TratarAspasHtml(objRSPag("Sis_Desc"))%></a></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("Cid_Sigla")) & " "%> <%=TratarAspasHtml(objRSPag("Esc_Sigla"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("Dst_Desc"))%></a></td>
			<td  width="300">&nbsp;<%=TratarAspasHtml(objRSPag("Pro_Nome"))%></td>
			<td  width="70">&nbsp;<%=TratarAspasHtml(objRSPag("Pla_TipoPlataforma"))%></td>
			<td ><input  type="checkbox" name="excluir" value="<%=objRSPag("rec_ID")%>" onClick="AddSelecaoChk(this)"></td>
		</tr>
		<%
		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	Next
End if
%>
</table>
</td>
</tr>
<tr>
	<td align=center>
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('recurso.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<input type="button" class="button" name="btnExcluir" value="Excluir" onClick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
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
