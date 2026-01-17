<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

IF strLoginRede <> "EDAR" and strLoginRede <> "PRSSILV" and strLoginRede <> "SCESAR" and strLoginRede <> "MSCAPRI" and strLoginRede <> "T3FRRP" THEN
	msg = "<p align=center><b><font color=#000080 face=Arial Black size=6>Sistema NewCLA</font></b></p>"
	msg = msg & "<p align=center><b><font color=#000080 face=Arial Black size=4>Em Testes</font></b></p>"
	Response.write msg
	response.end
END IF
%>

<html>
<title>Formulário para teste da interface CLA x APG.</title>
<center>
<br>
<br>

<form name="Form_1" method="post" action="../../newcla/access_interf/Asp/Solicitar_Acesso_APG_CLA.asp">
  <textarea name="hdnxml"></textarea> hdnxml<br><br>
  <input type="Submit" name="btnok" value="Solicitar_Acesso_APG_CLA.asp"><br>
</form>


<form name="Form_1" method="post" action="../../newcla/access_interf/Asp/Construir_Acesso_APG_CLA.asp">
  <textarea name="hdnxml"></textarea> hdnxml<br><br>
  <input type="Submit" name="btnok" value="Construir_Acesso_APG_CLA.asp"><br>
</form>

<form name="Form_1" method="post" action="../../newcla/access_interf/Asp/Entregar_Acesso_APG_CLA.asp">
  <textarea name="hdnxml"></textarea> hdnxml<br><br>
  <input type="Submit" name="btnok" value="Entregar_Acesso_APG_CLA.asp"><br>
</form>

<form name="Form_1" method="post" action="../../newcla/access_interf/Asp/Encerrar_Acesso_APG_CLA.asp">
  <textarea name="hdnxml"></textarea> hdnxml<br><br>
  <input type="Submit" name="btnok" value="Encerrar_Acesso_APG_CLA.asp"><br>
</form>