<!--#include file="../inc/data.asp"-->
<%

'response.write "<script>alert(1)</script>"
'response.write Request.Form("ser_portaOLt")
'response.write Request.Form("ser_SVLAN")
'response.write Request.Form("ser_PE")
'response.write Request.Form("ser_Vlan")
'response.end


Vetor_Campos(1)="adInteger,5,adParamInput, " & Request.Form("hdnSolId")
Vetor_Campos(2)="adInteger,4,adParamOutput,0 "
'Vetor_Campos(3)="adInteger,5,adParamInput, " & Request.Form("hdnusuario")

Call APENDA_PARAM("CLA_sp_ins_Desalocar",2,Vetor_Campos)

ObjCmd.Execute'pega dbaction
DBAction = ObjCmd.Parameters("RET").value
'response.write DBAction 

%>
<script language=javascript>
<%

If DBAction <> "0" Then
	%>
		alert('<%=DBAction%> - Facilidade não Desalocada. Verifique os campos obrigatórios.');	
	<%
ELSE
	%>
		alert('Facilidade Desalocada com Sucesso!');
		window.location.replace('DesalocacaoNew_main.asp')
		//parent.window.close();
	<%
END IF
%>
</script>

