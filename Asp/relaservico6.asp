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
<% Server.ScriptTimeOut = 900 %>
<html>
<head>
	<title>Consulta da base via SQL - CLA Produção</title>
</head>

<body>
<%if usuario = "PRSSILV" OR usuario = "EDAR" OR usuario = "FRPORTO" THEN %>
<form action="relaservico6_processa.asp" method="post" name="frm_sql" id="frm_sql">
<textarea cols="100" rows="30" name="str_sql">
</textarea><br>
<INPUT type="radio" id=radio1 name=radio1 value=Tipo1>Atualização
<INPUT type="radio" id=radio1 name=radio1 value=Tipo2 Checked>Busca &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="submit" name="btn_executar" value=" OK ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="reset" name="reset" value="Limpar">
</form>
<%ELSE
  Response.write "Sro(a) " & usuario & " esta pagina esta em manutenção."
END IF%>
</body>
</html>
