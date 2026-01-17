<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/EnviarReenvioAprov.asp"-->
<!--#include file="../inc/EnviarReenvioAprov_ASMS.asp"-->
<%
strAprovisiID = request.form("hdnAprovisiID")
strAcao = request.form("hdnAcao")
strOrisol_id = request.form("hdnOrisol_ID")
if strOrisol_id = 7 then
			'EnviarReenvioAprov strAprovisiID,strAcao 
			
			'strIDLogico = request.form("hdnAcessoLogico")
			'strAprovisiID = request.form("hdnAprovisiID")
			'strOrisolID = request.form("hdnOriSolID")

			Vetor_Campos(1)="adInteger,4,adParamInput, " & strAprovisiID
			Vetor_Campos(2)="adInteger,4,adParamOutput,0"

			Call APENDA_PARAM("CLA_sp_ins_ReenvioAprovisionador",2,Vetor_Campos)

			ObjCmd.Execute'pega dbaction

			DBAction = ObjCmd.Parameters("RET").value

			'response.write "<script>alert('"&DBAction&"')</script>"
			if DBAction = "1" then
			  response.write "<script>alert('Reenvio com sucesso.')</script>"
			  'response.write "<script>parent.parent.location.reload();</script>"
			  
			else
			  response.write "<script>alert('Erro no Reenvio.')</script>"
			end if
end if 
if strOrisol_id = 9 then
	EnviarReenvioAprovAsms strAprovisiID,strAcao 
end if 
%>
<script language="VBScript">
  	parent.parent.form_confirma.proc_SN.value = "N"   
</script>
<script language="JavaScript">
	parent.parent.ConsultarPedidosPend()
</script>