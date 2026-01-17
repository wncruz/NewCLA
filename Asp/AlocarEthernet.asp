<!--#include file="../inc/data.asp"-->
<%


Vetor_Campos(1)="adWChar,20,adParamInput, " & Request.Form("cboPropriedadeInter")
Vetor_Campos(2)="adWChar,30,adParamInput, " & Request.Form("cboPropriedadeEDD")
Vetor_Campos(3)="adWChar,10,adParamInput, " '& Request.Form("cboPropriedadePE")

Vetor_Campos(4)="adWChar,20,adParamInput, " & Request.Form("hdnAcfIdRadio")
Vetor_Campos(5)="adWChar,30,adParamInput, " & Request.Form("hdnIdLog")
Vetor_Campos(6)="adWChar,30,adParamInput, " & Request.Form("txtNroAcessoEbtEthernet")

Vetor_Campos(7)="adInteger,4,adParamOutput,0 "


Call APENDA_PARAM("CLA_SP_INS_ETHERNET",7,Vetor_Campos)
'response.write APENDA_PARAMstr("CLA_sp_ins_Switch",4,Vetor_Campos)
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

