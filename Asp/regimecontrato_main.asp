<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: RegimeContrato_main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Regime de contrato
%>
<!--#include file="../inc/data.asp"-->
<%
if Trim(Request.Form("hdnAcao")) = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_regimecontrato")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post" >
<input type=hidden name=hdnAcao>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=3><p align="center">Cadastro de Regime Contrato</p></th>
</tr>
<tr>
<th>&nbsp;Provedor</th>
<th>&nbsp;Tipo Contrato</th>
<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
</td>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim strTipoContrato
Dim objRSTipoCntr

strSql = "CLA_sp_sel_regimecontrato 0"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		Set objRSTipoCntr =db.execute("CLA_sp_sel_tipocontrato " & objRSPag("tct_id"))
		if not objRSTipoCntr.Eof and not objRSTipoCntr.bof then
			strTipoContrato = TratarAspasHtml(objRSTipoCntr("tct_desc"))
		End if
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td >&nbsp;<a href="regimecontrato.asp?ID=<%=objRSPag("Reg_ID") %>"><%=TratarAspasHtml(objRSPag("pro_nome"))%></a></td>
			<td >&nbsp;<%=strTipoContrato%></td>
			<td ><input  type="checkbox" name="excluir" value="<%=objRSPag("Reg_ID")%>" onClick="AddSelecaoChk(this)"></td>
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
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('regimecontrato.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
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
Set objRSTipoCntr = Nothing
Set objRSPag = Nothing
DesconectarCla()
%>
