<!--#include file="../inc/data.asp"-->
<%
strProEmail = Request.Form("strProEmail")
hdnstrHTML = Request.Form("hdnstrXML")
ASSUNTO = Request.Form("ASSUNTO")
strFromEmail = Request.Form("strFromEmail")
NdAcfid = Request.Form("NdAcfid")
dblProId = Request.Form("dblProId")
dblPedId = Request.Form("dblPedId")
intTipoProcesso = Request.Form("intTipoProcesso")
strUserName = Request.Form("strUserName")
strCaminho = server.MapPath("..\")

SET objRS = Server.CreateObject("ADODB.Recordset")
strSql = "select ped_prefixo,ped_numero,ped_ano from cla_pedido where ped_id = " & dblPedId
set objRS = db.execute(strSql)

Pedido_antigo = ucase(objRS("ped_prefixo")) & "-" & right("00000" & objRS("ped_numero"),5) & "/" & objRS("ped_ano")

hdnstrHTML = "<html>" & hdnstrHTML

Vetor_Campos(1)="adInteger,4,adParamInput," & NdAcfid
Vetor_Campos(2)="adWChar,30,adParamInput," & strUserName
Vetor_Campos(3)="adInteger,4,adParamOutput,"
Vetor_Campos(4)="adInteger,4,adParamOutput,"
Vetor_Campos(5)="adInteger,4,adParamOutput,"

Call APENDA_PARAM("CLA_sp_DesativacaoLote",5,Vetor_Campos)
ObjCmd.Execute'pega dbaction

dblSolId = ObjCmd.Parameters("RET").value
dblPedId = ObjCmd.Parameters("RET2").value
dblSisId = ObjCmd.Parameters("RET3").value

SET objRS = Server.CreateObject("ADODB.Recordset")
strSql = "select ped_prefixo,ped_numero,ped_ano from cla_pedido where ped_id = " & dblPedId
set objRS = db.execute(strSql)

Pedido_novo = ucase(objRS("ped_prefixo")) & "-" & right("00000" & objRS("ped_numero"),5) & "/" & objRS("ped_ano")

ASSUNTO = replace(ASSUNTO,Pedido_antigo,Pedido_novo)
hdnstrHTML = replace(hdnstrHTML,Pedido_antigo,Pedido_novo)

Vetor_Campos(1)="adInteger,4,adParamInput," & dblPedId
Vetor_Campos(2)="adWChar,30,adParamInput," & strUserName

Call APENDA_PARAM("CLA_sp_ins_agentesolicitacaoLiberacao",2,Vetor_Campos)
ObjCmd.Execute

Vetor_Campos(1)="adInteger,4,adParamInput," & dblProId
Vetor_Campos(2)="adInteger,4,adParamInput," & dblPedId
Vetor_Campos(3)="adInteger,4,adParamInput,2" 
Vetor_Campos(4)="adWChar,30,adParamInput," & strUserName
Vetor_Campos(5)="adDate,8,adParamInput,"	  'Data de envio getdate()
Vetor_Campos(6)="adWChar,4000,adParamInput," & trim(hdnstrHTML)
strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_ins_documento",6,Vetor_Campos)

Call db.Execute(strSqlRet)

if Ucase(strLoginRede) = "PRSSILV" or Ucase(strLoginRede) = "EDAR" then
   'IF Desenvolvedor:
   strProEmail = "prssilv@embratel.com.br"
End if

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