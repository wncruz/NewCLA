<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: CentroFuncional_Main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove centro funcional
%>
<!--#include file="../inc/data.asp"-->
<%
If request("hdnAcao")="Excluir" then
	Call ExcluirRegistro("CLA_sp_del_centrofuncional")
End if
%>
<!--#include file="../inc/header.asp"-->
<tr>
<td>
<!-- PRSS - em implemantação inicio-->
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
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=5><p align=center>Cadastro de Centro Funcional</p></th>
</tr>
<tr class=clsSilver>
  <td >Busca (CNL Sigla. Ex.: SPO ou SPO IG)&nbsp;&nbsp;</td>
  <td colspan=4>
	<input type=text name=txtBusca maxlength=8 class="text" value="<%=TratarAspasHtml(Request.Form("txtBusca"))%>" >&nbsp;
	<input type="button" class="button" name=btnProcurar value=Procurar onclick="ProcurarEstacao()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">
  </td>
</tr>
<!-- PRSS - em implemantação fim-->
<tr>
	<th>&nbsp;Centro Funcional</th>
	<th>&nbsp;Compl Pre</th>
	<th>&nbsp;Compl Ger</th>
	<th>&nbsp;Área Tecno</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass

if Trim(Request.Form("txtBusca")) <> ""  then
	objAryParam = split(Trim(Request.Form("txtBusca"))," ")
	if Ubound(objAryParam) > 0 then
		strSql = "CLA_sp_sel_centrofuncionalFull null,'" & TratarAspasSql(objAryParam(0)) & "','" & TratarAspasSql(objAryParam(1)) & "'"
	Else
		strSql =  "CLA_sp_sel_centrofuncionalFull null,'" & TratarAspasSql(objAryParam(0)) & "'"
	End if	
Else
	strSql = "CLA_sp_sel_centrofuncionalFull"
End if





Call PaginarRS(0,strSql)
intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
		<td ><a href="CentroFuncional.asp?ID=<%=objRSPag("Ctfc_ID") %>"><%=TratarAspasHtml(objRSPag("Ctf_AreaFuncional"))%> - <%=TratarAspasHtml(objRSPag("Cid_Sigla"))%></a></td>
		<td >&nbsp;<%=TratarAspasHtml(objRSPag("Esc_Sigla"))%> </td>
		<td >&nbsp;<%=TratarAspasHtml(objRSPag("Age_Sigla"))%> </td>
		<td >&nbsp;<%=TratarAspasHtml(objRSPag("Ctfc_AreaTecnologica"))%></td>
		<td ><input type="checkbox" name="excluir" value="<%=objRSPag("Ctfc_ID")%>" onClick="AddSelecaoChk(this)"></td>
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
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('CentroFuncional.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<input type="button" class="button" name="btnExcluir" value="Excluir" onClick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
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