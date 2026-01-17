<!--#include file="../inc/data.asp"-->
<%
strProEmail = Request.Form("strProEmail")
hdnstrHTML = Request.Form("hdnstrXML")
ASSUNTO = Request.Form("ASSUNTO")
strFromEmail = Request.Form("strFromEmail")
'NdAcfid = Request.Form("NdAcfid")
dblProId = Request.Form("dblProId")
dblPedId = Request.Form("dblPedId")
intTipoProcesso = Request.Form("intTipoProcesso")
strUserName = Request.Form("strUserName")
strCaminho = server.MapPath("..\")

hdnstrHTML = "<html>" & hdnstrHTML

Vetor_Campos(1)="adInteger,4,adParamInput," & dblProId
Vetor_Campos(2)="adInteger,4,adParamInput," & dblPedId
Vetor_Campos(3)="adInteger,4,adParamInput," & intTipoProcesso 
Vetor_Campos(4)="adWChar,30,adParamInput," & strUserName
Vetor_Campos(5)="adDate,8,adParamInput,"	  'Data de envio getdate()
Vetor_Campos(6)="adWChar,4000,adParamInput," & trim(hdnstrHTML)
strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_ins_documento",6,Vetor_Campos)

Call db.Execute(strSqlRet)

'if Ucase(strLoginRede) = "PRSSILV" or Ucase(strLoginRede) = "EDAR" then
   'IF Desenvolvedor:
  ' strProEmail = "edar@embratel.com.br"
'End if

Set objFSO = server.CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile(strCaminho & "\CartasProvedor\emailprovedor.htm",  true)
objFile.Writeline(hdnstrHTML)
objFile.Close

Set ObjMail	= Server.CreateObject("CDONTS.NewMail")
ObjMail.From = strFromEmail '"acessos@embratel.com.br"
ObjMail.To	 = strProEmail  'impleme@embratel.com.br
ObjMail.Subject = ASSUNTO
ObjMail.BodyFormat = 0
ObjMail.MailFormat = 0
ObjMail.AttachFile ( strCaminho & "\CartasProvedor\emailprovedor.htm") 
ObjMail.Body = "segue em anexo carta de solicitação de serviço referente: " &  ASSUNTO
ObjMail.Send

Set ObjMail = Nothing
Set objFSO = Nothing 
Set objFile = Nothing 
%>
<script language="javascript">
	parent.Envio.btnEnviar.disabled = true
	alert('Email enviado');
</script>