<%
	strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<link rel=stylesheet type="text/css" href="../css/cla.css">
</HEAD>
<BODY topmargin=0 leftmargin=0>
<table width="760" border="0" cellspacing="0" cellpadding="0">
<tr > 
	<td valign=top>
		<img name="embratel" src="../imagens/topo_embratel.jpg" width=760px height=80px border="0">
	</td>
</tr>	
<tr>
	<td background="../imagens/marca.gif" height=350 align=center valign=center>
		<img name="embratel" src="../imagens/Erro.jpg" border="0">
		O usuário <font color=red><%=strLoginRede%> </font> não esta cadastrado no sistema CLA.
	</td>	
</tr>		
</table>
</BODY>
</HTML>
