<%
Set objFSO = server.CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile(server.MapPath("..\") & "\CartasProvedor\tstEmail.htm",  true)
objFile.Writeline("<html><body>Teste de email</body></html>")
objFile.Close

Set mail = Server.CreateObject("CDONTS.NewMail")
mail.from = "eduardo.nascimento@claro.com.br"
'mail.to = "jcartus@embratel.com.br"
mail.to = "eduardo.nascimento@claro.com.br"
mail.cc = "eduardo.nascimento@claro.com.br"
mail.subject = "Teste de Email - CLA" 
mail.MailFormat = 0
mail.BodyFormat = 0
mail.body = "Este email foi enviado em " & now
mail.AttachFile Server.mappath("..\CartasProvedor\tstEmail.htm")

mail.send
set mail = nothing
response.write "<script>alert('Email enviado !')</script>"
%>
