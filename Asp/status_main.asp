<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Status_main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Status
%>
<!--#include file="../inc/data.asp"-->
<%
If Trim(Request.Form("hdnAcao")) ="Excluir" Then
	Call ExcluirRegistro("CLA_sp_del_status")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post" >
<input type=hidden name=hdnAcao>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=9><p align="center">Cadastro de Status</p></th>
</tr>
<tr>
	<th>&nbsp;Descrição</th>
	<th>&nbsp;Notifica</th>
	<th>&nbsp;GIC-N</th>
	<th>&nbsp;GIC-L</th>
	<th>&nbsp;GLA</th>
	<th>&nbsp;GLA-E</th>
	<th>&nbsp;AVL</th>
	<th>&nbsp;Tipo</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_sel_Status"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td ><a href="status.asp?ID=<%=objRSPag("Sts_ID") %>">&nbsp;<%=TratarAspasHtml(objRSPag("Sts_Desc"))%></a> </td>
			<td  width="70">&nbsp;<%if objRSPag("Sts_Notifica") = true then response.write "SIM" else response.write "NÃO" end if%></td>
			<td  width="70">&nbsp;<%if objRSPag("Sts_GICN") = true then response.write "SIM" else response.write "NÃO" end if%></td>
			<td  width="70">&nbsp;<%if objRSPag("Sts_GICL") = true then response.write "SIM" else response.write "NÃO" end if%></td>
			<td  width="70">&nbsp;<%if objRSPag("Sts_GLA") = true then response.write "SIM" else response.write "NÃO" end if%></td>
			<td  width="70">&nbsp;<%if objRSPag("Sts_GLAE") = true then response.write "SIM" else response.write "NÃO" end if%></td>
			<td  width="70">&nbsp;<%if objRSPag("Sts_AVL") = true then response.write "SIM" else response.write "NÃO" end if%></td>			
			<td  width="70">&nbsp;<%if objRSPag("Sts_Tipo") = 0 then response.write "Macro" else response.write "Detalhado" end if%></td>
			<td ><input type="checkbox" name="excluir" value="<%=objRSPag("Sts_ID")%>" onClick="AddSelecaoChk(this)"></td>
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
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('status.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
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
