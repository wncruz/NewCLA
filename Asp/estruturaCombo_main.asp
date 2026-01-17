<%
'	- Sistema			: CLA
'	- Arquivo			: estruturaCombo_main.asp
'	- Responsável		: EDAR
'	- Descrição			: Estrutura de Combo
%>
<!--#include file="../inc/data.asp"-->
<%
If Request.Form("hdnAcao") = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_AssocCombo")
End if
%>
<!--#include file="../inc/header.asp"-->
<form name="Form1" method="post"  onSubmit="return ConfirmarRemocao()">
<input type=hidden name=hdnAcao>
<script language="JavaScript">
function Procurar()
{
	with (document.forms[0])
	{
		if (txtBusca.value != "" )
		{
			submit()
		}
		else
		{
			alert("Informe o Combo!")
			txtBusca.focus()
			return
		}
		
	}
}
</script>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">

<tr class=clsSilver>
  <td nowrap >Busca Combo &nbsp;&nbsp;</td>
  <td colspan=8>
	<input type=text name=txtBusca maxlength=50 class="text" value="<%=TratarAspasHtml(Request.Form("txtBusca"))%>" >&nbsp;
	<input type="button" class="button" name=btnProcurar value=Procurar onclick="Procurar()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">
  
	
  </td>
</tr>

<tr>
<!--<th>&nbsp;Tecnologia | Facilidade</th> -->
<th colspan=8 >&nbsp;Combo</th>


<!-- <th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th> -->
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass

if Trim(Request.Form("txtBusca")) <> ""  then
	//strSql = "CLA_sp_sel_AssocCombo null,null,null,'" & Trim(Request.Form("txtBusca")) & "'
	
	strSql = "CLA_sp_sel_newCombo null,null,'" & Trim(Request.Form("txtBusca")) & "'" 
	
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
			
			<td colspan=8><a href="cst_estruturaCombo.asp?ID=<%=objRSPag("newcombo_id") %>">&nbsp;<%=TratarAspasHtml(objRSPag("newcombo_nome"))%> </a></td>
			
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
	<!--<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('estruturaTecnologia.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">-->
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
