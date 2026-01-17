<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Velocidade_main.asp
'	- Responsável		: Vital
'	- Descrição			: Lista/Remove Velocidade
%>
<!--#include file="../inc/data.asp"-->
<%

                                                                                                                    


If Trim(Request.Form("hdnAcao"))="Excluir" then
	Call ExcluirRegistro("CLA_sp_del_velocidade")
End if
%>
<!--#include file="../inc/header.asp"-->

<%
IF strLoginRede <> "SCESAR" and strLoginRede <> "EDAR" and strLoginRede <> "JCARTUS" and strLoginRede <> "T3FRRP" THEN
	msg = "<br><p align=center><b><font color=#000080 face=Arial Black size=5>Informativo CLA</font></b></p>"                         
	msg = msg & "<p align=center><b><font color=#000080 face=Arial Black size=4>O cadastro de Velocidade no CLA está restrito à equipe de suporte.<br>Por favor abrir OI.</font></b></p>"           
	Response.write msg                                                                                                           
	response.end                                                                                                                 
END IF  
%>
<form name="Form1" method="post" >
<input type=hidden name=hdnAcao>
<tr>
<td>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=3><p align="center">Cadastro de Velocidade</p></th>
</tr>
<tr>
	<th>&nbsp;Descrição</th>
	<th>&nbsp;Ordenação</th>
	<th width="20"><input type="checkbox" name="excluirtudo" onclick="seleciona_tudo(this);AddSelecaoChk(this)">Tudo</th>
</tr>
<%  
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_sel_velocidade"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td ><a href="velocidade.asp?ID=<%=objRSPag("vel_id") %>">&nbsp;<%=TratarAspasHtml(objRSPag("Vel_Desc"))%></a> </td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("vel_Ordem"))%></td>
			<td ><input type="checkbox" name="excluir" value="<%=objRSPag("vel_id")%>" onClick="AddSelecaoChk(this)"></td>
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
		<input type="button" class="button" name="Incluir" value="Incluir" onClick="javascript:window.location.replace('velocidade.asp')" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">
		<input type="button" class="button" name="btnExcluir" value="Excluir" onClick="ExlcuirRegistro()" accesskey="R" onmouseover="showtip(this,event,'Excluir (Alt+R)');">
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
