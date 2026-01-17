<!--#include file="../inc/data.asp"-->
<%
strProEmail = Request.Form("hdnEmailProvedor")
strNomePro	= Request.Form("hdnNomeProvedor")
dblPedId	= Request.Form("hdnPedId")

strCaminho = server.MapPath("..\")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(strCaminho & "\CartasProvedor\reenvio_emailprovedor.htm",  true)
objFile.WriteLine( request("stremail"))
objFile.Close

Set objRSPed = db.execute("CLA_sp_view_pedido null,null,null,null,null,null," & dblPedId & ",null,null,'T'")

if not objRSPed.Eof and not objRSPed.Bof then

  Set objRSSol = db.execute("CLA_sp_view_solicitacaomin " & objRSPed("Sol_Id"))

  Set ObjMail	= Server.CreateObject("CDONTS.NewMail")

  ObjMail.From = "acessos@embratel.com.br"
  
  '@@DEBUG
  'SELECT CASE STRLOGINREDE
  'CASE "ISAV"
    'strProEmail = "isav@embratel.com.br"
  'CASE "T3PRSS"
    'strProEmail = "paulo.silva@bull.com.br"
  'CASE ELSE
    'strProEmail = "fmag@embratel.com.br"
  'END SELECT
  
  Set objRSConf = db.Execute("select * from cla_config where Config_ID = 4 and Config_Estado = 0 and Config_Data > getdate()")
  If Not objRSConf.eof and  not objRSConf.Bof Then
  	ObjMail.To	 = strProEmail
  	ObjMail.bcc = "embratel@asteriscos.com.br" 'Teste de função. 
  	ObjMail.Subject = "REENVIO - " & AcaoPedidoEmail(ucase(objRSPed("tprc_id"))) & "  -  " & trim(objRSSol("Cli_nome")) & "  -  " & ucase(objRSPed("Ped_Prefixo")) & "-" & right("00000" & objRSPed("Ped_Numero"),5) & "/" & objRSPed("Ped_Ano")
  	ObjMail.BodyFormat = 0
  	ObjMail.MailFormat = 0
  	'Response.Write strArquivo
  	ObjMail.AttachFile strCaminho & "\CartasProvedor\reenvio_emailprovedor.htm"
  	ObjMail.Send
  	Set ObjMail = Nothing
  	Set objRSPed = Nothing
  end if

  Response.Write "<script language=javascript>parent.EmailEnviado('" & strNomePro & "','" & strProEmail & "');</script>"
Else
  Set objRSPed = Nothing
End if

Function AcaoPedidoEmail(intTipo)
	Select case intTipo
		case 1
			AcaoPedidoEmail = "Instalar Acesso"
		case 2			
			AcaoPedidoEmail = "Retirar Acesso"
		case 3		
			AcaoPedidoEmail = "Alterar Acesso"
		case 4
			AcaoPedidoEmail = "Cancelar Acesso"
	End Select
End Function
%>