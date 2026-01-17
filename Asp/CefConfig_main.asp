<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: CefConfig_main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Configurações CEF
%>
<!--#include file="../inc/data.asp"-->
<%

If Trim(Request.Form("hdnAcao"))="Excluir" then
	Call ExcluirRegistro("CLA_sp_del_ConfigCtf")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post" >
<input type=hidden name=hdnAcao>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=4><p align="center">Parâmetros do Centro Funcional</p></th>
</tr>
<tr>
	<th>&nbsp;Centro Funcional</th>
	<th>&nbsp;Redirecionamento de Carteira</th>
	<th>&nbsp;Avaliador de Acesso</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_sel_ConfigCtf"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td ><a href="CefConfig.asp?ID=<%=objRSPag("Ctf_ID") %>">&nbsp;<%=TratarAspasHtml(objRSPag("Ctf_AreaFuncional"))%> - <%=TratarAspasHtml(objRSPag("Cid_Sigla"))%></a> </td>
			<td >&nbsp;<%if objRSPag("Cfg_RedirecionamentoCarteira") = "1" then Response.Write "SIM" else Response.Write "NÃO" End if%></td>
			<td >&nbsp;<%if objRSPag("Cfg_Avaliador") = "1"   then Response.Write "SIM" else Response.Write "NÃO" End if %></td>
			
			<td ><input type="checkbox" name="excluir" value="<%=objRSPag("Ctf_ID")%>" onClick="AddSelecaoChk(this)"></td>
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
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('CefConfig.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
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
