<%@ CodePage=65001 %>
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: estruturaTecnologia_main.asp
'	- Responsável		: Vital
'	- Descrição			: Estrutura de Tecnologia
%>
<!--#include file="../inc/data.asp"-->
<%
If Request.Form("hdnAcao") = "Excluir" then
	Call ExcluirRegistro("CLA_sp_del_AssocTecnologiaFacilidade")
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
	<!--
	<th colspan=8><p align=center>Estrutura de Tecnologia</p></th>
	
		<font class="clsObrig">:: </font>Tecnologia + Facilidade
		<select name="cboAssocTecnologiaFacilidade">
			<option value=""></option>
			<% set objRS = db.execute("CLA_sp_sel_AssocTecnologiaFacilidade  ")
			
				While Not objRS.Eof
			 		strSel = ""
					if Cdbl("0" & objRS("assoc_tecfac_id")) = Cdbl("0" & dblIDAtual) then strSel = " disabled "
					Response.Write "<Option value="& Trim(objRS("assoc_tecfac_id")) & strSel & ">" & Trim(objRS("newtec_nome"))  + " - " + Trim(objRS("newfac_nome")) & "</Option>"
					objRS.MoveNext
				Wend
				Set objRS = Nothing
			%>
					</select> 
        -->
		<!--<input type="button" class="button" name=btnProcurar value=Procurar onclick="ProcurarUsuario()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');"></td>-->
	<!-- </tr> -->

</tr>
<tr class=clsSilver2>
  <td nowrap>Busca Facilidade   </td>
  <td colspan=8>
	<input type=text name=txtBusca2 maxlength=50 class="text" value="<%=TratarAspasHtml(Request.Form("txtBusca2"))%>" > 
	</td>
</tr>
<tr class=clsSilver>
  <td nowrap >Busca Tecnologia   </td>
  <td colspan=8>
	<input type=text name=txtBusca maxlength=50 class="text" value="<%=TratarAspasHtml(Request.Form("txtBusca"))%>" > 
	<input type="button" class="button" name=btnProcurar value=Procurar onclick="Procurar()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">
  
	
  </td>
</tr>

<tr>
<!--<th> Tecnologia | Facilidade</th> -->
<th colspan=8 > Facilidade | Tecnologia</th>


<!-- <th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th> -->
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
			<!--<td ><a href="cst_estruturaTecnologia.asp?ID=<%=objRSPag("assoc_tecfac_id") %>"> <%=TratarAspasHtml(objRSPag("newtec_nome"))%>       |        <%=TratarAspasHtml(objRSPag("newfac_nome"))%></a></td> -->
			
			<td colspan=8><a href="cst_estruturaTecnologia.asp?ID=<%=objRSPag("assoc_tecfac_id") %>"> <%=TratarAspasHtml(objRSPag("newfac_nome"))%>       |        <%=TratarAspasHtml(objRSPag("newtec_nome"))%></a></td>
			
			<!--<td ><input type="checkbox" name="excluir" value="<%=objRSPag("assoc_tecfac_id")%>" onClick="AddSelecaoChk(this)"></td> -->
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
