<!--#include file="../inc/data.asp"-->
<%


if Request.Form("rdoPortaPEID") = "" then
	PortaId = Request.Form("hdnrdoPortaPEID")
else 
	PortaId = Request.Form("rdoPortaPEID")
end if

dblRecId = ""	

'response.write "<script>alert('"&Request.Form("hdnCVLAN_ETHERNET")&"')</script>"
'response.write "<script>alert('"&Request.Form("hdnSVLAN_ETHERNET")&"')</script>"


Vetor_Campos(1)="adWChar,10,adParamInput, " & Request.Form("hdnIdLog")
Vetor_Campos(2)="adInteger,8,adParamInput, "  & PortaId 'Request.Form("rdoPortaID")
Vetor_Campos(3)="adInteger,8,adParamInput, " & Request.Form("hdnAcfIdRadio")
Vetor_Campos(4)="adWChar,30,adParamInput, "  & Request.Form("hdnUplink")
Vetor_Campos(5)="adWChar,30,adParamInput, "  & dblRecId
Vetor_Campos(6)="adWChar,30,adParamInput, "  & Request.Form("hdnSolId")
Vetor_Campos(7)="adInteger,4,adParamOutput,0 "
Vetor_Campos(8)="adWChar,30,adParamInput, "  & Request.Form("hdnCVLAN_ETHERNET")
Vetor_Campos(9)="adWChar,30,adParamInput, "  & Request.Form("hdnSVLAN_ETHERNET")
Vetor_Campos(10)="adWChar,30,adParamInput, "  & Request.Form("hdnVLAN_PortaOLT")


Call APENDA_PARAM("CLA_SP_INS_UPLINK",10,Vetor_Campos)
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
		
		parent.ResgatarEthernetPE();
		//parent.window.close();
	<%
END IF
%>
</script>

