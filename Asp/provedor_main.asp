<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Provedor_main.asp
'	- Descrição			: Lista/Remove Provedor
%>
<!--#include file="../inc/data.asp"-->
<%
If Trim(Request.Form("hdnAcao")) = "Excluir" Then
	Call ExcluirRegistro("CLA_sp_del_provedor")
End if
%>
<!--#include file="../inc/header.asp"-->
<%
For Each Perfil in objDicCef
	if Perfil = "PST" then dblCtfcIdPst = objDicCef(Perfil)
Next
%>
<SCRIPT LANGUAGE=javascript>
<!--
function AutorizaEdicao(intID,strProvedor)
{
	with (document.forms[0])
	{
		if (hdnPerfil.value != "")
		{
			action = "provedor.asp?ID="+intID
			target = self.name 
			submit()
		}
		else
		{
			var strProvedorAux = new String(strProvedor)
			if (strProvedorAux.toUpperCase().indexOf("EMBRATEL") == -1)
			{
				alert("O perfil do usuário permite editar\nsomente provedores Embratel.")
				return
			}else
			{
				action = "provedor.asp?ID="+intID
				target = self.name 
				submit()
			}
		}
	}
}
//-->
</SCRIPT>
<form name="Form1" method="post">
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnPerfil value="<%=dblCtfcIdPst%>">
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=8><p align=center>Cadastro de Provedor</p></th>
</tr>
<tr>
	<th>&nbsp;Código</th>
	<th>&nbsp;Nome</th>
	<th>&nbsp;Contato</th>
	<th>&nbsp;Padrão de Designação</th>
	<th>&nbsp;Holding</th>
	<th>&nbsp;Visível</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%

Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_sel_provedor 0"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("pro_Cod"))%></td>

			<%if dblCtfcIdPst <> "" then%>
				<td >&nbsp;<a href="javascript:AutorizaEdicao(<%=objRSPag("pro_id")%>,'<%=objRSPag("pro_Nome")%>')">&nbsp;<%=TratarAspasHtml(objRSPag("pro_Nome"))%></a> </td>
			<%Else
				if inStr(1,Ucase(objRSPag("pro_Nome")),"EMBRATEL") > 0 then
			%>
					<td >&nbsp;<a href="javascript:AutorizaEdicao(<%=objRSPag("pro_id")%>,'<%=objRSPag("pro_Nome")%>')">&nbsp;<%=TratarAspasHtml(objRSPag("pro_Nome"))%></a> </td>
				<%Else%>
					<td >&nbsp;<%=TratarAspasHtml(objRSPag("pro_Nome"))%></td>
				<%End if%>
			<%End if%>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("pro_contato"))%></td>
			<td  nowrap><%=TratarAspasHtml(objRSPag("Pro_PadraoDesigMin"))%><br><%=TratarAspasHtml(objRSPag("Pro_PadraoDesigMax"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("Hol_Desc"))%></td>
			<td >&nbsp;<%if objRSPag("pro_Visivel") = 1 then Response.Write "SIM" else Response.Write "NÃO" End if%></td>
			<td ><input type="checkbox" name="excluir" value="<%=objRSPag("pro_id")%>" onClick="AddSelecaoChk(this)"
				<%if dblCtfcIdPst = "" and inStr(1,Ucase(objRSPag("pro_Nome")),"EMBRATEL") = 0 then Response.Write " disabled " %>
				 ></td>
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
	<!-- ***** Causa-raíz: CH-56898XNW -->		
	<%if strLoginRede = "SCESAR" OR strLoginRede = "EDAR" OR strLoginRede = "T3FRRP" OR strLoginRede = "RCCARD" OR strLoginRede = "LUIZCMP" OR strLoginRede = "CARLAGP"  OR strLoginRede = "MARCELN"  OR strLoginRede = "ANDRELF" OR strLoginRede = "Z164075" OR strLoginRede = "AFSEDA" OR strLoginRede = "ARIELLA" OR strLoginRede = "NATHPIM" then %> 
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('provedor.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<input type="button" class="button" name="btnExcluir" value="Excluir" onClick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">
	<%end if%>	
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
