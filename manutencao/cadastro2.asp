<%
dim tamanhoLogon
dim usuario

If Request.ServerVariables("LOGON_USER") = "" Then
 	Response.Status = "401 access denied"
else
	tamanhoLogon=len(Request.ServerVariables("LOGON_USER"))
	usuario=ucase(trim(mid(Request.ServerVariables("LOGON_USER"),10,tamanhoLogon)))
end if

%>
<% Server.ScriptTimeOut = 3600 %>
<html>
<head>
	<title>Atualização da base via SQL - CLA Produção</title>
</head>

<body>
<%if usuario = "PRSS" OR  usuario = "PRSSILV" OR usuario = "EDAR" OR usuario = "FEMAG" THEN %>
<form action="cadastro_2.asp" method="post" name="frm_sql" id="frm_sql">
<h2>selecione o tipo de operação que deseja realizar:</h2>
<%if usuario = "PRSS" OR usuario = "PRSSILV" then%>
<INPUT type="radio" id=radio1 name=radio1 value=Tipo1>Atualização<br>
<%END IF%>
<INPUT type="radio" id=radio1 name=radio1 value=Tipo2 Checked>Select<br><br>

Entre com o código SQL (sem aspas) a ser executado na base: <font color=#ff0000>CLA - Produção</font> <br>
<textarea cols="90" rows="15" name="str_sql">
</textarea><br>
<input type="submit" name="btn_executar" value="Executar">
<input type="reset" name="reset" value="Limpar">
</form>
<%ELSE
  Response.write "Sro(a) " & usuario & " esta pagina esta em manutenção."
END IF%>
</body>
</html>
