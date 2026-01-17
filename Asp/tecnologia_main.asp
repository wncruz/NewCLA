<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: tecnologia_main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Tecnologia
%>
<!--#include file="../inc/data.asp"-->
<%
If Trim(Request.Form("hdnAcao")) = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_newtecnologia")
End if
%>
<!--#include file="../inc/header.asp"-->

<form name="Form1" method="post"  action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type=hidden name=hdnAcao>
<script language="JavaScript">
function Procurar()
{
	with (document.forms[0])
	{
		if (txtBusca.value != "")
		{
			submit()
		}
		else
		{
			alert("Informe a Tecnologia!")
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
	<th colspan=4><p align="center">Cadastro de Tecnologia</p></th>
</tr>
<tr class=clsSilver>
  <td >Busca (Tecnologia Ex.: Fo ou Fo EDD)&nbsp;&nbsp;</td>
  <td colspan=4>
	<input type=text name=txtBusca maxlength=30 size = "60" class="text" value="<%=TratarAspasHtml(Request.Form("txtBusca"))%>" >&nbsp;
	<input type="button" class="button" name=btnProcurar value=Procurar onclick="Procurar()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">
  </td>
</tr>
<tr>
<th>&nbsp;Descrição</th>
<th>&nbsp;Sigla</th>
<th>&nbsp;Status</th>
<!--<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>-->
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass



if Trim(Request.Form("txtBusca")) <> ""  then
	strSql = "CLA_sp_sel_newTecnologia null,null,null,'" & Trim(Request.Form("txtBusca")) & "'" 
Else
	'strSql = "CLA_sp_sel_newTecnologia null,null,null"
End if


Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td ><a href="tecnologia.asp?ID=<%=objRSPag("newtec_id") %>">&nbsp;<%=TratarAspasHtml(objRSPag("newtec_nome"))%></a> </td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("newtec_sigla"))%> </td> 
			<td> 
			<% strStatus=TratarAspasHtml(objRSPag("newtec_ativo"))
			  if  strStatus = "S" then
				Response.Write "ATIVO"				
			  else
				
				Response.Write "INATIVO"
			  end if
			%>		
			</td>
			<!--<td><input  type="checkbox" name="excluir" value="<%=objRSPag("newtec_id")%>" onClick="AddSelecaoChk(this)"></td>-->
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
		
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('tecnologia.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<!--<input type="button" class="button" name="btnExcluir" value="Excluir" onClick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');"> -->
		
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
