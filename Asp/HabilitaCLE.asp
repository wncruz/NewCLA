<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->

<script>
function HabilitarCLE()
{  
		with (document.forms[0])
	  {
	  	acao.value="atualizar"
			action = "HabilitaCLE.asp"
			submit()
		}
}
</script>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<body>
<br><br>
<table width="760" border="0" cellspacing="0" cellpadding="0"><tr><td>
<%
IF strLoginRede <> "EDAR" and strLoginRede <> "FEMAG" and strLoginRede <> "JCARTUS" and strLoginRede <> "SCESAR" THEN
	response.write "<br><p align=center><b><font color=red>***************************************************************************</font></b></p>"
	response.write "<p align=center><b><font color=red>Usuário " & strLoginRede & " não está autorizado nesta tarefa.</font></b></p>"
	response.write "<p align=center><b><font color=red>***************************************************************************</font></b></p><br><br>"
else
hab_CLE = request.form("hab_CLE")
if isnull(hab_CLE) or trim(hab_CLE)="" then
	set rs = db.execute("select Conn_Tipo from CLA_ConexaoInterf where conn_Sistema='CLE'")
	hab_CLE = rs("Conn_Tipo")
end if
acao    = request.form("acao")
%>
<form name="frm" method="post">
<%
if acao="atualizar" then
	db.execute("update CLA_ConexaoInterf set Conn_Tipo='" & hab_CLE & "' where conn_Sistema='CLE'")
	if hab_CLE="S" then
		response.write "<p align=center><b><font color=blue>Interface CLE habilitada com sucesso</font></b></p>"
	else
		response.write "<p align=center><b><font color=red>Interface CLE desabilitada com sucesso</font></b></p>"
	end if
else
%>
<table align=center cellspacing="2" cellpadding="2" border=0>
<TR><TD align=center><b>Interface CLE</b></TD></TR>
<TR><TD><input type="hidden" name="acao">Habilitar ?  
<select name="hab_CLE" onchange="limpa()">
<option value="S" <%if hab_CLE="S" then%>selected<%end if%>>SIM</option>
<option value="N" <%if hab_CLE="N" then%>selected<%end if%>>NÃO</option>
</select>
</td></tr>
<TR><TD align=center> 
<INPUT TYPE=BUTTON OnClick="HabilitarCLE();" class="button" VALUE="Habilitar CLE" style="color:darkred;;font-weight:bold;width:180px">
</TD></TR>
</table>
</form>
<%
end if
end if' if strLoginRede
%>
<span id="spn"><p align=center>Para realizar testes de conectividade com o CLE, <a href="ConsultaCLE.asp">Clique aqui</a></p></span>
</tr></table>
</body>
</html>