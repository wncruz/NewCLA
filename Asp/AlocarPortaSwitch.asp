<!--#include file="../inc/data.asp"-->
<%

'response.write "<script>alert('"&Request.Form("cboLocalInstala")&"')</script>"
'response.write "<script>alert('"&Request.Form("cboDistLocalInstala")&"')</script>"
'response.write "<script>alert('"&Request.Form("hdnPedId")&"')</script>"
'response.write "<script>alert('"&Request.Form("hdnRede")&"')</script>"
'response.write "<script>alert('"&Request.Form("hdnRede")&"')</script>"
'response.write "<script>alert('"&Request.Form("hdnProvedor")&"')</script>"


Vetor_Campos(1)="adInteger,2,adParamInput," & Request.Form("cboLocalInstala")
Vetor_Campos(2)="adInteger,2,adParamInput," & Request.Form("cboDistLocalInstala")
Vetor_Campos(3)="adInteger,2,adParamInput," & Request.Form("hdnProvedor")
Vetor_Campos(4)="adInteger,2,adParamInput," & Request.Form("hdnRede")
Vetor_Campos(5)="adInteger,2,adParamInput," & strPlataforma
Vetor_Campos(6)="adInteger,2,adParamOutput,0"

strSqlRet = APENDA_PARAMSTR("CLA_sp_check_recurso2",6,Vetor_Campos)
Set objRSRec = db.execute(strSqlRet)
Set DBAction = objRSRec("ret")

dblRecId = ""
If DBAction = 0 then
	dblRecId = objRSRec("Rec_ID")
End if
if Request.Form("hdnTipoProcesso") <> "4" and dblRecId = "" then
	Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
	Response.End 
End if


if Request.Form("rdoPortaSwitchID") = "" then
	PortaId = Request.Form("hdnrdoPortaSwitchID")
else 
	PortaId = Request.Form("rdoPortaSwitchID")
end if
Vetor_Campos(1)="adInteger,8,adParamInput, " & Request.Form("hdnSwitchID")
Vetor_Campos(2)="adWChar,10,adParamInput, " & Request.Form("hdnIdLog")
Vetor_Campos(3)="adWChar,5,adParamInput, "  & Request.Form("hdnvlanSwitch")
Vetor_Campos(4)="adWChar,20,adParamInput, "  & Request.Form("hdnportaoltSwitch")
Vetor_Campos(5)="adWChar,13,adParamInput, "  & Request.Form("hdnpeSwitch")
Vetor_Campos(6)="adInteger,8,adParamInput, "  & PortaId 'Request.Form("rdoPortaID")
Vetor_Campos(7)="adWChar,30,adParamInput, "  & Request.Form("hdndesigRadioIP")
Vetor_Campos(8)="adInteger,8,adParamInput, " & Request.Form("hdnAcfIdRadio")
Vetor_Campos(9)="adWChar,5,adParamInput, "  & Request.Form("hdnSvlanSwitch")
Vetor_Campos(10)="adInteger,8,adParamInput, " & Request.Form("hdnProvedor")
Vetor_Campos(11)="adInteger,8,adParamInput, " & Request.Form("hdnPedId")
Vetor_Campos(12)="adInteger,8,adParamInput, " & dblRecId
Vetor_Campos(13)="adInteger,4,adParamOutput,0 "
Vetor_Campos(14)="adWChar,40,adParamInput, "  & Server.HTMLEncode(Request.Form("hdnportaSwitchLadoMetro"))



Call APENDA_PARAM("CLA_sp_ins_SwitchPorta",14,Vetor_Campos)
'response.write APENDA_PARAMstr("CLA_sp_ins_ONTPorta",10,Vetor_Campos)
'response.end
ObjCmd.Execute'pega dbaction
DBAction = ObjCmd.Parameters("RET").value
%>
<script language=javascript>
<%
If DBAction <> "0" Then
	%>
		alert('<%=DBAction%> - Facilidade não alocada. Verifique os campos obrigatórios.');	
	<%
ELSE
	%>
		alert('Facilidade Alocada com Sucesso!');
		parent.window.close();
	<%
END IF
%>
</script>

