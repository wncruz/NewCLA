<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: CadTipoRadio_main.asp
'	- Descrição			: Lista/Remove Tipos de Radios
%>
<!--#include file="../inc/data.asp"-->
<%
If Trim(Request.Form("hdnAcao")) = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_TipoRadio")
End if
%>
<!--#include file="../inc/header.asp"-->
<script language = "JavaScript">

function ProcurarTipoRadio()
{
	with (document.forms[0])
	{
			submit()
	}
}


</script>
<form name="Form1" method="post">
<input type=hidden name=hdnAcao>
<tr>
	<td width=100%>
		<table border="0" cellspacing="1" cellpadding=0 width=760 >
			<tr>
				<th colspan=2><p align="center">Cadastro de Tipos de Radio</p></th>
			</tr>
			<tr class=clsSilver>
				<td >Tipo de Radio&nbsp;&nbsp;</td>
				<td >
				<input type=text name=txtBusca maxlength=25 size = 45 class="text" value="<%=TratarAspasHtml(Request.Form("txtBusca"))%>" >&nbsp;
				<input type="button" class="button" name=btnProcurar value=Procurar onclick="ProcurarTipoRadio()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
<th width="130">&nbsp;Tipo</th>
<th width="220">&nbsp;Descrição</th>
<th width="100">&nbsp;Data Cadastro</th>
<th width="100">&nbsp;Data Desativação</th>
<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass

if (trim(Request.Form("txtBusca")) = "") then
	strSql = "CLA_sp_sel_TipoRadio"
else
	strSql = "CLA_sp_sel_TipoRadio null ,'" &  Request.Form("txtBusca") & "'"
end if 

Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td ><a href="CadTipoRadio.asp?ID=<%=objRSPag("Trd_ID") %>">&nbsp;<%=TratarAspasHtml(objRSPag("Trd_TipoRadio"))%></a> </td>
			<td > <% = TratarAspasHtml(objRSPag("Trd_Descricao")) %>  </td>
			<% 	if objRSPag("Trd_DtCadastro") <> "" then %>
				<td > <% = TratarAspasHtml( right("0" & day(objRSPag("Trd_DtCadastro")),2) & "/" & right("0" & month(objRSPag("Trd_DtCadastro")),2) & "/" & year(objRSPag("Trd_DtCadastro")) ) %>  </td>
			<%else%>
				<td></td>
			<%end if%>
			<% 	if objRSPag("Trd_DtDesativacao") <> "" then %>
				<td > <% = TratarAspasHtml( right("0" & day(objRSPag("Trd_DtDesativacao")),2) & "/" & right("0" & month(objRSPag("Trd_DtDesativacao")),2) & "/" & year(objRSPag("Trd_DtDesativacao")) ) %>  </td>
			<%else%>
				<td></td>
			<%end if%>
			<td><input  type="checkbox" name="excluir" value="<%=objRSPag("Trd_id")%>" onClick="AddSelecaoChk(this)"></td>
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
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('CadTipoRadio.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
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
