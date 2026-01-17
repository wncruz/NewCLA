<%@ CodePage=65001 %>
<%
Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Teste Simples</title>
</head>
<body>
    <h1>Teste 1: Página ASP básica funcionou!</h1>
    
    <hr>
    
    <h2>Teste 2: Tentando include header.asp</h2>
    <%
    On Error Resume Next
    %>
    <!--#include file="../inc/header.asp"-->
    <%
    If Err.Number <> 0 Then
        Response.Write "<p style='color:red;'>ERRO no header.asp: " & Err.Description & "</p>"
    Else
        Response.Write "<p style='color:green;'>Header.asp incluído com sucesso!</p>"
    End If
    %>
    
    <hr>
    
    <h2>Teste 3: Tentando include footer.asp</h2>
    <%
    On Error Resume Next
    %>
    <!--#include file="../inc/footer.asp"-->
    <%
    If Err.Number <> 0 Then
        Response.Write "<p style='color:red;'>ERRO no footer.asp: " & Err.Description & "</p>"
    Else
        Response.Write "<p style='color:green;'>Footer.asp incluído com sucesso!</p>"
    End If
    %>
    
    <hr>
    
    <h2>Teste 4: Formulário Básico</h2>
    <form method="post" action="">
        <input type="hidden" name="teste" value="ok">
        <button type="submit">Enviar Teste</button>
    </form>
    
    <%
    If Request.Form("teste") = "ok" Then
        Response.Write "<p style='color:green; font-weight:bold;'>✓ POST funcionou!</p>"
    End If
    %>
    
</body>
</html>