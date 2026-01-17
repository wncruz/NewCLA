<%
Set mail = Server.CreateObject("CDONTS.NewMail")
mail.from = "acessos@embratel.com.br"
'mail.to = "prss@embratel.com.br"
mail.to = "joaofns@embratel.com.br"
mail.cc = "osvaper@embratel.com.br"
mail.subject = "Solicitaчуo de Serviчo: Desativar Acesso -  DM-38356/2006"
mail.MailFormat = 0
mail.BodyFormat = 0
mail.body = "segue em anexo carta de solicitaчуo de serviчo referente: Desativar Acesso  -  COMPANHIA ENERGETICA DE GOIAS  -  DM-38356/2006"
mail.AttachFile Server.mappath("emailprovedor.htm")

mail.send
set mail = nothing
%>