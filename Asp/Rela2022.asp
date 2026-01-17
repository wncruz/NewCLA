<html>
<title>Checa de Serviço 0800.</title>
<center>
<br>
<br>
<%
strXml = Request("txtAreaSolicitar")
stracao = Request("hdnacao")
%>
<script>
function enviar(acao)
{
	if(Form_1.txtAreaSolicitar.value != "")
	  {
		Form_1.hdnacao.value=acao;
		Form_1.submit();
	  }
	else
	  {
		alert("Favor preencher o XML!")
	  }
}
</script>

<form name="Form_1" method="post" action="Rela2022.asp">
   <textarea name="txtAreaSolicitar"><%=strXml%></textarea><br><br>
   <input type="button" name="btnok1" value="Solicitar Acesso" onclick="enviar('solicitar');">
   <input type="button" name="btnok2" value="Solicitar Acesso ASMS" onclick="enviar('solicitarAsms');">
   <input type="button" name="btnok6" value="Solicitar ASMS Schema" onclick="enviar('solicitarAsmsSchema');">
   <input type="button" name="btnok3" value="Notificar Término ASMS" onclick="enviar('notificarAsms');">
   <input type="button" name="btnok4" value="Notificar Término" onclick="enviar('notificar');">
   <input type="button" name="btnok5" value="Mud TIT" onclick="enviar('MudTit');">
   <input type="button" name="btnok6" value="Mud TIT ASMS" onclick="enviar('MudTitASMS');">
   <input type="button" name="btnlimp" value="Limpar" onclick="Form_1.txtAreaSolicitar.value=''">
   <input type="hidden" name="hdnacao">
</form>
<%
if strXml <> "" then
	if stracao = "solicitar" then
		sUrl = "http://localhost/newcla/access_interf/Asp/Aprov_Solicitar_Acesso.asp"
	elseif stracao = "solicitarAsms" then
		sUrl = "http://localhost/newcla/access_interf/Asp/Aprov_Solicitar_Acesso_ASMS.asp"
	elseif stracao = "solicitarAsmsSchema" then
		sUrl = "http://localhost/newcla/access_interf/Asp/Aprov_Solicitar_Acesso_ASMS_Schemas.asp"
	elseif stracao = "notificarAsms" then
		sUrl = "http://localhost/newcla/access_interf/Asp/Aprov_Notificar_Termino_ASMS.asp"
	elseif stracao = "MudTitASMS" then
		sUrl = "http://localhost/newcla/access_interf/Asp/Aprov_Mudanca_Titularidade_ASMS.asp"
	elseif stracao = "MudTit" then
		sUrl = "http://localhost/newcla/access_interf/Asp/Aprov_Mudanca_Titularidade.asp"
	elseif stracao = "notificar" then
		sUrl = "http://localhost/newcla/access_interf/Asp/Aprov_Notificar_Termino.asp"
	else
	    sUrl = "http://localhost/newcla/access_interf/Asp/Aprov_Notificar_Termino.asp"
	end if
	
	'StrLogin = "PRSSILV"
	
	Set doc = server.CreateObject("Microsoft.XMLDOM")
	'doc.loadXml(strXml)
	'doc.save(Server.MapPath("TesteConectividade.xml"))

	Set xmlhttp = Server.CreateObject("Microsoft.XMLHTTP")
	Set xmlhttp = CreateObject("Msxml2.XMLHTTP")

	xmlhttp.Open "POST", sUrl, False
	xmlhttp.Send strXml
	
	response.write xmlhttp.responseText
else
	response.write "<font color=red>Favor preencher o XML</font>"
End if
%>