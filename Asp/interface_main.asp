<%
'•EXPERT INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: interface_main.asp
'	- Responsável		: PRSS
'	- Descrição			: Lista/Remove interfaces no sistema utilizadas pela Solicitacao.asp
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
If request("hdnAcao")="Excluir" then
	Call ExcluirRegistro("CLA_sp_del_interface")
End if
%>
<form name="Form1" method="post" >
<input type=hidden name=hdnAcao>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=3><p align=center>Cadastro de Interfaces</p></th>
</tr>
<tr>
<th>&nbsp;Nome</th>
<th>&nbsp;Descrição</th>
<th width="20px"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_sel_interface"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td width="20%"><a href="interface.asp?ID=<%=objRSPag("ITF_Id") %>">&nbsp;<%=TratarAspasHtml(objRSPag("ITF_Nome"))%></a></td>
			<td width="70%">&nbsp;<%=TratarAspasHtml(objRSPag("ITF_Desc"))%></td>
			<td width="20" ><input type="checkbox" name="excluir" value="<%=objRSPag("ITF_id")%>" onClick="AddSelecaoChk(this)"></td>
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
	<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('interface.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
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
