<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AssocServVel_main.asp
'	- Responsável		: Vital
'	- Descrição			: Associaçção de servoço com velocidade
%>
<!--#include file="../inc/data.asp"-->
<%
If Request.Form("hdnAcao") = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_AssocServVeloc")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post"  onSubmit="return ConfirmarRemocao()">
<input type=hidden name=hdnAcao>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=3><p align=center>Associação de Serviço com Velocidade</p></th>
</tr>
<tr>
<th>&nbsp;Serviço</th>
<th>&nbsp;Velocidade</th>
<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass

strSql = "CLA_sp_sel_AssocServVeloc"

Call PaginarRS(0,strSql)
intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class=<%=strClass%>>
			<td ><a href="AssocServVel.asp?ID=<%=objRSPag("Asv_ID") %>">&nbsp;<%=TratarAspasHtml(objRSPag("Ser_Desc"))%></a></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("Vel_Desc"))%></td>
			<td ><input type="checkbox" name="excluir" value="<%=objRSPag("Asv_ID")%>" onClick="AddSelecaoChk(this)"></td>
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
	<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('AssocServVel.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
	<input type="button" class="button" name="btnExcluir" value="Excluir" onClick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">
	<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
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
