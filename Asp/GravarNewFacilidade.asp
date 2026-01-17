<!--#include file="../inc/data.asp"-->
<%


Response.ContentType = "text/html; charset=utf-8"
	Response.Charset = "UTF-8"

'response.write "<script>alert(1)</script>"
'response.write "<script>alert('"&Request.Form("hdnSolId")&"')</script>" 
'response.write "<script>alert('"&Request.Form("ser_Vlan")&"')</script>" 
'response.write "<script>alert('"&Request.Form("ser_SVLAN")&"')</script>" 
'response.write "<script>alert('"&Request.Form("ser_PE")&"')</script>" 
'response.write "<script>alert('"&Request.Form("ser_portaOLt")&"')</script>" 
'response.write "<script>alert('"&Request.Form("campo_1")&"')</script>" 
'response.write "<script>alert('"&Request.Form("campo_2")&"')</script>" 
'response.write "<script>alert('"&Request.Form("campo_3")&"')</script>" 
'response.write"<script>alert('"&Request.Form("campo_4")&"')</script>" 
'response.write "<script>alert('"&Request.Form("campo_5")&"')</script>" 
'response.write "<script>alert('"&Request.Form("campo_6")&"')</script>" 
'response.write "<script>alert('"&Request.Form("campo_7")&"')</script>" 
'response.write "<script>alert('"&Request.Form("campo_8")&"')</script>" 
'response.write "<script>alert('"&Request.Form("campo_9")&"')</script>" 
'response.end

'Response.Write "<script language='javascript'>alert('entrei')</script>"

'-----------------------------
'-- Good início
'-----------------------------
Vetor_Campos(1)="adInteger,5,adParamInput, " & Request.Form("hdnSolId")
Vetor_Campos(2)="adWChar,100,adParamInput, " & Request.Form("ser_Vlan")
Vetor_Campos(3)="adWChar,100,adParamInput, " & Request.Form("ser_SVLAN")
Vetor_Campos(4)="adWChar,100,adParamInput, " & Request.Form("ser_PE")
Vetor_Campos(5)="adWChar,100,adParamInput, " & Request.Form("ser_portaOLt")
Vetor_Campos(6)="adWChar,100,adParamInput, " & Request.Form("campo_1")
Vetor_Campos(7)="adWChar,100,adParamInput, " & Request.Form("campo_2")
Vetor_Campos(8)="adWChar,100,adParamInput, " & Request.Form("campo_3")
Vetor_Campos(9)="adWChar,100,adParamInput, " & Request.Form("campo_4")
Vetor_Campos(10)="adWChar,100,adParamInput, " & Request.Form("campo_5")
Vetor_Campos(11)="adWChar,100,adParamInput, " & Request.Form("campo_6")
Vetor_Campos(12)="adWChar,100,adParamInput, " & Request.Form("campo_7")
Vetor_Campos(13)="adWChar,100,adParamInput, " & Request.Form("campo_8")
Vetor_Campos(14)="adWChar,100,adParamInput, " & Request.Form("campo_9")
Vetor_Campos(15)="adWChar,100,adParamInput, " & Request.Form("campo_10")
Vetor_Campos(16)="adWChar,100,adParamInput, " & Request.Form("campo_11")
Vetor_Campos(17)="adWChar,100,adParamInput, " & Request.Form("campo_12")
Vetor_Campos(18)="adWChar,100,adParamInput, " & Request.Form("campo_13")

'-----------------------------
Vetor_Campos(19)="adWChar,100,adParamInput,  " '& Request.Form("hdnusuario")

Vetor_Campos(20)="adWChar,100,adParamInput, " & Request.Form("campo_14")
Vetor_Campos(21)="adWChar,100,adParamInput, " & Request.Form("campo_15")
Vetor_Campos(22)="adWChar,100,adParamInput, " & Request.Form("campo_16")
Vetor_Campos(23)="adWChar,100,adParamInput, " & Request.Form("campo_17")
Vetor_Campos(24)="adWChar,100,adParamInput, " & Request.Form("campo_18")
Vetor_Campos(25)="adWChar,100,adParamInput, " & Request.Form("campo_19")
Vetor_Campos(26)="adWChar,100,adParamInput, " & Request.Form("campo_20")
Vetor_Campos(27)="adWChar,100,adParamInput, " & Request.Form("campo_21")
Vetor_Campos(28)="adWChar,100,adParamInput, " & Request.Form("campo_22")
Vetor_Campos(29)="adWChar,100,adParamInput, " & Request.Form("campo_23")
Vetor_Campos(30)="adWChar,100,adParamInput, " & Request.Form("campo_24")
Vetor_Campos(31)="adWChar,100,adParamInput, " & Request.Form("campo_25")
Vetor_Campos(32)="adWChar,100,adParamInput, " & Request.Form("campo_26")

Vetor_Campos(33)="adWChar,100,adParamInput, " & Request.Form("campo_27")
Vetor_Campos(34)="adWChar,100,adParamInput, " & Request.Form("campo_28")
Vetor_Campos(35)="adWChar,100,adParamInput, " & Request.Form("campo_29")
Vetor_Campos(36)="adWChar,100,adParamInput, " & Request.Form("campo_30")
Vetor_Campos(37)="adWChar,100,adParamInput, " & Request.Form("campo_31")
Vetor_Campos(38)="adWChar,100,adParamInput, " & Request.Form("campo_32")
Vetor_Campos(39)="adWChar,100,adParamInput, " & Request.Form("campo_33")
Vetor_Campos(40)="adWChar,100,adParamInput, " & Request.Form("campo_34")
Vetor_Campos(41)="adWChar,100,adParamInput, " & Request.Form("campo_35")
Vetor_Campos(42)="adWChar,100,adParamInput, " & Request.Form("campo_36")
Vetor_Campos(43)="adWChar,100,adParamInput, " & Request.Form("campo_37")
Vetor_Campos(44)="adWChar,100,adParamInput, " & Request.Form("campo_38")
Vetor_Campos(45)="adWChar,100,adParamInput, " & Request.Form("campo_39")

Vetor_Campos(46)="adInteger,5,adParamInput, " & Request.Form("cboLocalInstala")
Vetor_Campos(47)="adInteger,5,adParamInput, " & Request.Form("cboLocalConfig")
Vetor_Campos(48)="adInteger,7,adParamInput, " & Request.Form("hdnAcfId")
Vetor_Campos(49)="adInteger,4,adParamOutput,0 "


'Response.Write "<script language='javascript'>alert('param')</script>"
'
'Vetor_Campos(48)="adInteger,7,adParamInput, " & Request.Form("hdnAcfId")


'Vetor_Campos(19)="adInteger,4,adParamOutput,0 "
'Vetor_Campos(20)="adWChar,100,adParamInput,  " '& Request.Form("hdnusuario")

'Vetor_Campos(21)="adWChar,100,adParamInput, " & Request.Form("campo_14")
'Vetor_Campos(22)="adWChar,100,adParamInput, " & Request.Form("campo_15")
'Vetor_Campos(23)="adWChar,100,adParamInput, " & Request.Form("campo_16")
'Vetor_Campos(24)="adWChar,100,adParamInput, " & Request.Form("campo_17")
'Vetor_Campos(25)="adWChar,100,adParamInput, " & Request.Form("campo_18")
'Vetor_Campos(26)="adWChar,100,adParamInput, " & Request.Form("campo_19")
'Vetor_Campos(27)="adWChar,100,adParamInput, " & Request.Form("campo_20")
'Vetor_Campos(28)="adWChar,100,adParamInput, " & Request.Form("campo_21")
'Vetor_Campos(29)="adWChar,100,adParamInput, " & Request.Form("campo_22")
'Vetor_Campos(30)="adWChar,100,adParamInput, " & Request.Form("campo_23")
'Vetor_Campos(31)="adWChar,100,adParamInput, " & Request.Form("campo_24")
'Vetor_Campos(32)="adWChar,100,adParamInput, " & Request.Form("campo_25")
'Vetor_Campos(33)="adWChar,100,adParamInput, " & Request.Form("campo_26")

'Vetor_Campos(34)="adWChar,100,adParamInput, " & Request.Form("campo_27")
'Vetor_Campos(35)="adWChar,100,adParamInput, " & Request.Form("campo_28")
'Vetor_Campos(36)="adWChar,100,adParamInput, " & Request.Form("campo_29")
'Vetor_Campos(37)="adWChar,100,adParamInput, " & Request.Form("campo_30")
'Vetor_Campos(38)="adWChar,100,adParamInput, " & Request.Form("campo_31")
'Vetor_Campos(39)="adWChar,100,adParamInput, " & Request.Form("campo_32")
'Vetor_Campos(40)="adWChar,100,adParamInput, " & Request.Form("campo_33")
'Vetor_Campos(41)="adWChar,100,adParamInput, " & Request.Form("campo_34")
'Vetor_Campos(42)="adWChar,100,adParamInput, " & Request.Form("campo_35")
'Vetor_Campos(43)="adWChar,100,adParamInput, " & Request.Form("campo_36")
'Vetor_Campos(44)="adWChar,100,adParamInput, " & Request.Form("campo_37")
'Vetor_Campos(45)="adWChar,100,adParamInput, " & Request.Form("campo_38")
'Vetor_Campos(46)="adWChar,100,adParamInput, " & Request.Form("campo_39")

'Vetor_Campos(47)="adInteger,5,adParamInput, " & Request.Form("cboLocalInstala")
'Vetor_Campos(48)="adInteger,5,adParamInput, " & Request.Form("cboLocalConfig")

'-----------------------------
'-- Good início
'-----------------------------
'Vetor_Campos(48)="adInteger,7,adParamInput, " & Request.Form("hdnAcfId")

'-----------------------------
'-- Good fim
'-----------------------------
'vparam=""
'for i = 1 to 49
'vparam = vparam & Vetor_Campos(i) & ","
'Response.Write "<script language='javascript'>alert('i: " & i & " - " & Vetor_Campos(i) & "')</script>"
'next

'Response.Write "<script language='javascript'>alert('i: " & vparam & "')</script>"
'respponse.end
		

Call APENDA_PARAM("CLA_sp_ins_NewAlocarFacilidade2",49,Vetor_Campos)

ObjCmd.Execute'pega dbaction
DBAction = ObjCmd.Parameters("RET").value

'Response.Write "<script language='javascript'>alert('DBAction: " & DBAction & "')</script>"

'strSqlRet = APENDA_PARAMSTR("CLA_sp_ins_NewAlocarFacilidade2",48,Vetor_Campos)
'response.write strSqlRet 

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
		window.location.replace('AlocacaoNew_main.asp');
		//parent.window.close();
	<%
END IF
%>
</script>

