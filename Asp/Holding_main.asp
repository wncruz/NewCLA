<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Holding_Main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Holding
%>
<!--#include file="../inc/data.asp"-->
<%
If request("hdnAcao")="Excluir" then
	Call ExcluirRegistro("CLA_sp_del_holding")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post" >
<input type=hidden name=hdnAcao>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align=center>Cadastro de Holding</p></th>
</tr>
<tr>
	<th>&nbsp;Nome</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_sel_holding"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class=<%=strClass%>>
			<td ><a href="holding.asp?ID=<%=objRSPag("Hol_id") %>">&nbsp;<%=TratarAspasHtml(objRSPag("Hol_Desc"))%></a> </td>
			<td ><input type="checkbox" name="excluir" value="<%=objRSPag("hol_id")%>" onClick="AddSelecaoChk(this)"></td>
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
	<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('holding.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
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
