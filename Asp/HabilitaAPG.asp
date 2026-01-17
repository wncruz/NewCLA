<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->

<script>
function HabilitarAPG()
{  
		with (document.forms[0])
	  {
	  	acao.value="atualizar"
			action = "HabilitaAPG.asp"
			submit()
		}
}
</script>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<body>
<br><br>
<table width="760" border="0" cellspacing="0" cellpadding="0"><tr><td>
<%
IF strLoginRede <> "EDAR" and strLoginRede <> "MSCAPRI" and strLoginRede <> "JCARTUS" and strLoginRede <> "SCESAR" THEN
	response.write "<br><p align=center><b><font color=red>***************************************************************************</font></b></p>"
	response.write "<p align=center><b><font color=red>Usuário " & strLoginRede & " não está autorizado nesta tarefa.</font></b></p>"
	response.write "<p align=center><b><font color=red>***************************************************************************</font></b></p><br><br>"
else
hab_APG = request.form("hab_APG")
if isnull(hab_APG) or trim(hab_APG)="" then
	set rs = db.execute("select Conn_Tipo from CLA_ConexaoInterf where conn_Sistema='APG'")
	hab_APG = rs("Conn_Tipo")
end if
acao    = request.form("acao")
%>
<form name="frm" method="post">
<%
if acao="atualizar" then
	db.execute("update CLA_ConexaoInterf set Conn_Tipo='" & hab_APG & "' where conn_Sistema='APG'")
	if hab_APG="S" then
		response.write "<p align=center><b><font color=blue>Interface APG <u>habilitada</u> com sucesso</font></b></p>"
	else
		response.write "<p align=center><b><font color=red>Interface APG <u>desabilitada</u> com sucesso</font></b></p>"
	end if
else
%>
<table align=center cellspacing="2" cellpadding="2" border=0>
<TR><TD align=center><b>Interface APG</b></TD></TR>
<TR><TD><input type="hidden" name="acao">Habilitar ?  
<select name="hab_APG" onchange="limpa()">
<option value="S" <%if hab_APG="S" then%>selected<%end if%>>SIM</option>
<option value="N" <%if hab_APG="N" then%>selected<%end if%>>NÃO</option>
</select>
</td></tr>
<TR><TD align=center> 
<INPUT TYPE=BUTTON OnClick="HabilitarAPG();" class="button" VALUE="Habilitar APG" style="color:darkred;;font-weight:bold;width:180px">
</TD></TR>
</table>
</form>
<%
end if
end if' if strLoginRede
%>
</td></tr></table>
</body>
</html>