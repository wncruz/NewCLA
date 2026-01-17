<!--#include file="../inc/data.asp"-->
<%
If Trim(Request.Form("hdnAcao"))="Excluir" then
	Call ExcluirRegistro("CLA_sp_del_FabricanteONT")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post"  action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type=hidden name=hdnAcao>
<script language="JavaScript">
function ProcurarEstacao()
{
	with (document.forms[0])
	{
		if (txtBusca.value != "")
		{
			submit()
		}
		else
		{
			alert("Informe a sigla da Estacao!")
			txtBusca.focus()
			return
		}
		
	}
}
</script>
<tr>
	<td width=100%>
		<table border="0" cellspacing="1" cellpadding=0 width=760 >
			<tr>
				<th colspan=2><p align="center">Cadastro de Fabricante ONT | EDD</p></th>
			</tr>
			<tr class=clsSilver>
				<td >Busca &nbsp;&nbsp;</td>
				<td >
				<input type=text name=txtBusca maxlength=8 class="text" value="<%=TratarAspasHtml(Request.Form("txtBusca"))%>" >&nbsp;
				<input type="button" class="button" name=btnProcurar value=Procurar onclick="ProcurarEstacao()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th width="40">Tecnologia</th>
	<th width="30">Fabricante</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim objAryParam

if Trim(Request.Form("txtBusca")) <> ""  then
	objAryParam = split(Trim(Request.Form("txtBusca"))," ")
	if Ubound(objAryParam) > 0 then
		strSql = "CLA_sp_sel_FabricanteONT 0,'" & TratarAspasSql(objAryParam(0)) & "','" & TratarAspasSql(objAryParam(1)) & "'"
	Else
		strSql =  "CLA_sp_sel_FabricanteONT 0,'" & TratarAspasSql(objAryParam(0)) & "'"
	End if
Else
	strSql = "CLA_sp_sel_FabricanteONT 0"
End if

Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class=<%=strClass%>>
			<td>&nbsp;<%=TratarAspasHtml(objRSPag("Sigla"))%></td>
			<td>&nbsp;<a href="cad_fabricanteONT.asp?ID=<%=objRSPag("Font_ID") %>"><%=TratarAspasHtml(objRSPag("Font_Nome"))%></a></td>
			<td><input type="checkbox" name="excluir" value="<%=objRSPag("Font_ID")%>" onClick="AddSelecaoChk(this)"></td>
		</tr>
		<%
		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	Next
End if
'End if
%>
</table>
</td>
</tr>
<tr>
	<td align=center>
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('cad_fabricanteONT.asp?TpProc=INC')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
<!--
@@ LPEREZ - 25/05/2006
<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('estacao_incluir.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');"> -->
<!--	LP<input type="button" class="button" name="btnExcluir" value="Excluir" onclick="ExlcuirRegistro()"  accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">-->
		<input type="button" class="button" name="btnExcluir" value="Excluir" onclick="ExlcuirRegistro()"  accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">
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
