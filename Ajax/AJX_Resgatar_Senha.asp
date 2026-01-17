<!--#include file="../inc/data.asp"-->
<%
Response.Expiresabsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

set objRS = db.execute("CLA_sp_Gera_Senhas 'PIN',6")
While Not objRS.eof
  var_senha = objRS("Senha")
  objRS.MoveNext
Wend
%>
<%=var_senha%>
