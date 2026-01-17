<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Distribuicao_Main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Distribuidor
%>
<!--#include file="../inc/data.asp"-->
<%
If Trim(Request.Form ("hdnAcao")) = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_distribuicao")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post" >
<input type=hidden name=hdnAcao>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align="center">Cadastro de Distribuidor</p></th>
</tr>

<tr class=clsSilver>
	<th>&nbsp;Descrição</th>
	<th width="20" >
		<input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo
	</th>
</tr>
</td>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_sel_distribuicao 0"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class=<%=strClass%>>
			<td ><a href="distribuicao.asp?ID=<%=objRSPag("dst_id") %>">&nbsp;<%=TratarAspasHtml(objRSPag("dst_Desc"))%></a> </td>
			<td ><input  type="checkbox" name="excluir" value="<%=Trim(objRSPag("dst_id"))%>" onClick="AddSelecaoChk(this)"></td>
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
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('distribuicao.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
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
