<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AssocTecnologiaFacilidade_main.asp
'	- Responsável		: Vital
'	- Descrição			: Associação de Tecnologia com Facilidade
%>
<!--#include file="../inc/data.asp"-->
<%
If Request.Form("hdnAcao") = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_AssocTecnologiaFacilidade")
End if
%>
<!--#include file="../inc/header.asp"-->
<!--<form name="Form1" method="post"  onSubmit="return ConfirmarRemocao()" > -->
<form name="Form1" method="post"  action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type=hidden name=hdnAcao>
<script language="JavaScript">
function Procurar()
{
	with (document.forms[0])
	{
		if (txtBusca.value != "" || txtBusca2.value != "")
		{
			submit()
		}
		else
		{
			alert("Informe a Tecnologia ou a Facilidade!")
			txtBusca.focus()
			return
		}
		
	}
}
</script>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=11><p align=center>Associação de Tecnologia com Facilidade</p></th>
</tr>
<tr class=clsSilver2>
  <td nowrap>Busca Facilidade &nbsp;&nbsp;</td>
  <td colspan=11>
	<input type=text name=txtBusca2 maxlength=50 class="text" value="<%=TratarAspasHtml(Request.Form("txtBusca2"))%>" >&nbsp;
	<input type="button" class="button" name=btnProcurar value=Procurar onclick="Procurar()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">
  </td>
</tr>
<tr class=clsSilver>
  <td nowrap >Busca Tecnologia &nbsp;&nbsp;</td>
  <td colspan=11>
	<input type=text name=txtBusca maxlength=50 class="text" value="<%=TratarAspasHtml(Request.Form("txtBusca"))%>" >&nbsp;
	
  </td>
</tr>

<tr>



<th>&nbsp;Facilidade</th>
<th>&nbsp;Tecnologia</th>
<th>&nbsp;Dados Serviço</th>
<th>&nbsp;Fase 1 Automática</th>
<th>&nbsp;Fase Ativação Automática</th>
<th>&nbsp;Fase Alteração Automática</th>
<th>&nbsp;Fase Cancelamento Automática</th>
<th>&nbsp;Fase Desativação Automática</th>
<th>&nbsp;Compartilha acesso</th>
<th>&nbsp;Compartilha Cliente</th>
<th>&nbsp;SAIP</th>

<!--<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th> -->
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass

if Trim(Request.Form("txtBusca")) <> "" or Trim(Request.Form("txtBusca2")) <> "" then
	strSql = "CLA_sp_sel_AssocTecnologiaFacilidade null,null,null,'" & Trim(Request.Form("txtBusca")) & "','" & Trim(Request.Form("txtBusca2")) & "'"
	//strSql = "CLA_sp_sel_centrofuncionalFull null,'" & TratarAspasSql(objAryParam(0)) & "','" & TratarAspasSql(objAryParam(1)) & "'"
Else
	//strSql = "CLA_sp_sel_AssocTecnologiaFacilidade"

End if



Call PaginarRS(0,strSql)
intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class=<%=strClass%>>
			<td ><a href="AssoctecnologiaFacilidade.asp?ID=<%=objRSPag("assoc_tecfac_id") %>">&nbsp;<%=TratarAspasHtml(objRSPag("newfac_nome"))%></a></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("newtec_nome"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("dados_servico"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("fase_1_automatico"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("fase_ativacao_automatico"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("fase_alteracao_automatico"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("fase_cancelamento_automatico"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("fase_desativacao_automatico"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("compartilha_acesso"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("compartilha_cliente"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("fase_config_saip"))%></td>
			<!--<td ><input type="checkbox" name="excluir" value="<%=objRSPag("assoc_tecfac_id")%>" onClick="AddSelecaoChk(this)"></td>-->
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
	<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('AssocTecnologiaFacilidade.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
	<!--<input type="button" class="button" name="btnExcluir" value="Excluir" onClick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');"> -->
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
