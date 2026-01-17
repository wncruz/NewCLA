<!--#include file="../inc/data.asp"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
</HEAD>
<BODY topmargin=0 leftmargin=0 class=TA>
<form name=Form1 method=Post>
<table width=100% cellspacing="1" cellpadding="0">
<tr >
	<th width=3%>&nbsp;Sel&nbsp;</th>
	<th width=15%>&nbsp;Acesso Físico</th>
	<th width=20%>&nbsp;Facilidade</th>
	<th width=20%>&nbsp;Tecnologia</th>
	<!--<th width=16%>&nbsp;Compl End</th> -->
	<th width=14%>&nbsp;Provedor</th>
	<!--<th width=36%>&nbsp;Cliente</th>-->
	<th width=7% nowrap>&nbsp;Vel Fis</th>	
	<th width=33%>&nbsp;Cliente</th>
	<!--<th width=15% nowrap>&nbsp;Nº Acesso</th> -->
</tr>

<%

Function in_array(element, arr)
  in_array = False
  For i=0 To Ubound(arr)
     If Trim(arr(i)) = Trim(element) Then
        in_array = True
        Exit Function      
     End If
  Next
End Function
 
provedor = Array("143","145","154","136","137")
sistema = Array("12")
tecnologia = Array("1" , "4")
provedorRip = Array("145")
tecnologiaRIP = Array( "3" , "7")
 


'JCARTUS@ - Aceitar endereço para o APG (CH-32221HVV)
'IF Request.Form("strOrigemAPG") = "APG" then
'	intEnd = Request.QueryString("intEnd") 
'	strCidSigla =  Trim(Request.Form("txtEndCidAPG"))
'	strUF		=  Trim(Request.Form("cboUFEndAPG"))
'	'strTplSigla	=  Trim(Request.Form("cboLogrEndAPG"))
'	strEnd		=  Trim(Request.Form("txtEndAPG"))
'	strNroEnd	=  Trim(Request.Form("txtNroEndAPG"))
'	strCompl	=  Trim(Request.Form("txtComplEndAPG"))
'	strCep		=  Trim(Request.Form("txtCepEndAPG"))
'else
intEnd = Request.QueryString("intEnd") 
strCidSigla =  Trim(Request.Form("txtEndCid"))
strUF		=  Trim(Request.Form("cboUFEnd"))
'strTplSigla	=  Trim(Request.Form("cboLogrEnd"))
strEnd		=  Trim(Request.Form("txtEnd"))
strNroEnd	=  Trim(Request.Form("txtNroEnd"))
strCompl	=  Trim(Request.Form("txtComplEnd"))
strCep		=  Trim(Request.Form("txtCepEnd"))
	strOrisol_id		=  Trim(Request.Form("hdnOriSol_ID"))
	strAcf_id			=  Trim(Request.Form("hdnStrAcfId"))

strNroSev =  Trim(Request.Form("hdnNroSev2"))

strtxtFacilidade = Trim(Request.Form("txtFacilidade"))
strcboTecnologia = Trim(Request.Form("cboTecnologia"))
strtxtFantasia = Trim(Request.Form("txtFantasia"))
strtxtCC = Trim(Request.Form("txtCC"))

strfac_id = Trim(Request.Form("hdnfac_id"))
if trim(strtxtFacilidade) <>"" Then
  strfac_id = split(trim(strtxtFacilidade),"__")(0) 
  strfac_des = split(trim(strtxtFacilidade),"__")(1) 
end if 

'response.write "<script>alert('"& Trim(Request.Form("txtFacilidade")) &"')</script>"
'response.write "<script>alert('"& trim(strtxtFacilidade) &"')</script>"
'response.write "<script>alert('"& trim(strfac_id) &"')</script>"
'response.write "<script>alert('"& trim(strcboTecnologia) &"')</script>"
'response.write "<script>alert('"& trim(strtxtFantasia) &"')</script>"
'response.write "<script>alert('"& trim(strtxtCC) &"')</script>"

	strAcf_IDAcessoFisico = Trim(Request.QueryString("hdnStrAcfIDAcessoFisico"))
'end if

if isnull(strAcf_id) then
'Response.Write "<script language=javascript>alert('" & intEnd & strCidSigla & strUF & strTplSigla & strEnd & strNroEnd & strCep & "')</script>"
if Trim(strCidSigla) = ""  or Trim(strUF) = "" or Trim(strEnd) = ""  or Trim(strNroEnd) = "" or Trim(strCep) = "" then
	Response.Write "<script language=javascript>alert('Os Campos uf,cnl,\nlogradouro,número e cep são obrigatórios!')</script>"
	Response.Write "<script language=javascript>parent.divIDFis" & intEnd & ".style.display = 'none';" & _
				   "</script>"
	Response.End 
End if
end if

' Williams
'response.write "<script>alert('"& trim(strfac_des) &"')</script>"
achou_FAC = 0 '0 = Não encontrado e 1 = Encontrado
Vetor_Campos(1)="adInteger,2,adParamInput," & strNroSev
Vetor_Campos(2)="adInteger,2,adParamOutput,0"	
Call APENDA_PARAM("CLA_sp_sel_historico_solucao_ssa",2,Vetor_Campos)
Set objRSCli = ObjCmd.Execute
if DBAction = 0 then
   if Not objRSCli.eof and  Not objRSCli.bof then
	   Do While Not objRSCli.eof			  
 		      if UCASE(trim(objRSCli("FAC_DES"))) = UCASE(trim(strfac_des)) then
			     achou_FAC = 1
		         Exit Do		         
	      end if
	   objRSCli.MoveNext
	   loop
   end if
end if

''--ch ''Vetor_Campos(1)="adInteger,2,adParamInput,"
''--ch ''Vetor_Campos(2)="adWChar,100,adParamInput,"
''--ch ''Vetor_Campos(3)="adWChar,30,adParamInput,"
''--ch ''Vetor_Campos(4)="adWChar,60,adParamInput,"
''--ch ''Vetor_Campos(5)="adWChar,4,adParamInput,"	& strCidSigla	'Cid Sigla
''--ch ''Vetor_Campos(6)="adWChar,2,adParamInput,"	& strUF			'UF
''--ch ''Vetor_Campos(7)="adWChar,15,adParamInput,"	'& strTplSigla	'Tlp Sigla
''--ch ''Vetor_Campos(8)="adWChar,60,adParamInput,"	& strEnd
''--ch ''Vetor_Campos(9)="adWChar,10,adParamInput,"	& strNroEnd	 
''--ch ''Vetor_Campos(10)="adWChar,30,adParamInput,"	'& strCompl	 
''--ch ''Vetor_Campos(11)="adWChar,9,adParamInput,"	& strCep		'CEP
''--ch ''Vetor_Campos(12)="adWChar,1,adParamInput,F"
''--ch ''Vetor_Campos(13)="adInteger,2,adParamInput," & strOrisol_id		'Origem solicitacao 6 sgap , 7 sgav , 9 asms
''--ch ''Vetor_Campos(14)="adInteger,2,adParamInput," & strAcf_id		'Origem solicitacao 6 sgap , 7 sgav , 9 asms
''--ch ''strSql = APENDA_PARAMSTR("CLA_sp_sel_AcessoFisico2",14,Vetor_Campos)
'Response.Write strSql
'response.write "<script>Trim(Request.Form("txtEndCid"))</script>"
'response.write "<script>alert('"& trim(strSql) &"')</script>"
''--ch ''Set objRSSolic = db.Execute(strSql)



Vetor_Campos(1)="adWChar,4,adParamInput,"	& strCidSigla	'Cid Sigla
Vetor_Campos(2)="adWChar,2,adParamInput,"	& strUF			'UF
Vetor_Campos(3)="adWChar,60,adParamInput,"	& strEnd
Vetor_Campos(4)="adWChar,10,adParamInput,"	& strNroEnd	 
Vetor_Campos(5)="adWChar,9,adParamInput,"	& strCep		'CEP
Vetor_Campos(6)="adWChar,1,adParamInput,F"
Vetor_Campos(7)="adInteger,2,adParamInput," & strOrisol_id		'Origem solicitacao 6 sgap , 7 sgav , 9 asms
Vetor_Campos(8)="adInteger,2,adParamInput," & strAcf_id		'Origem solicitacao 6 sgap , 7 sgav , 9 asms

Vetor_Campos(9)="adWChar,15,adParamInput,"	& strAcf_IDAcessoFisico

Vetor_Campos(10)="adInteger,2,adParamInput," & strNroSev

Vetor_Campos(11)="adInteger,2,adParamInput," & strfac_id
Vetor_Campos(12)="adInteger,2,adParamInput," & strcboTecnologia
Vetor_Campos(13)="adWChar,50,adParamInput,"	& strtxtFantasia
Vetor_Campos(14)="adWChar,11,adParamInput,"	& strtxtCC

'for i = 1 to 14
'Response.Write "<script language=javascript>alert('i: " & i & " - " & Vetor_Campos(i) & "')</script>"
'next

strSql = APENDA_PARAMSTR("CLA_sp_sel_new_comp_acessofisico",14,Vetor_Campos)
'Response.Write strSql
'response.write "<script>alert('"& strAcf_id &"')</script>"
'response.write "<script>alert('"& trim(strSql) &"')</script>"
 set objRSSolic = db.Execute(strSql)
blnCor = true
qtd_reg = 0
if Not objRSSolic.Eof and Not objRSSolic.Bof then 
	While Not objRSSolic.Eof 
	
		exibir_compart = "S"
				
		'response.write "<script>alert('" & objRSSolic("Acf_DtAceite") & "')</script>"
		'response.write "<script>alert('" & exibir_compart & "')</script>"
		'response.write "<script>alert('" & objRSSolic("Acf_IDAcessoFisico") & "')</script>"

 'if (exibir_compart = "S") then
 		qtd_reg = qtd_reg + 1
		if strClass = "clsSilver" then strClass = "clsSilver2" else strClass = "clsSilver" End if
	%>
		<tr class="<%=strClass%>" width=100% >
			<% if (strOrisol_id = "9" and objRSSolic("Acf_IDAcessoFisico") = "") then
			%>
				<td width=3%><input type=radio name=rdoIDFis<%=intEnd%> value="" Aec_IdFis="<%=objRSSolic("Aec_IdFis")%>" prop="<%=objRSSolic("Acf_Proprietario")%>" onClick="parent.SelIDFisComp_(this,<%=intEnd%>,<%=objRSSolic("Acf_ID")%>,<%=achou_FAC%>)" <%if Request.Form("hdnIdAcessoFisico") = objRSSolic("Acf_IDAcessoFisico") then Response.Write " checked " End if%>></td>
			<% else
			%>
				<td width=3%><input type=radio name=rdoIDFis<%=intEnd%> value="<%=objRSSolic("Acf_IDAcessoFisico")%>" Aec_IdFis="<%=objRSSolic("Aec_IdFis")%>" prop="<%=objRSSolic("Acf_Proprietario")%>" onClick="parent.SelIDFisComp_(this,<%=intEnd%>,<%=objRSSolic("Acf_ID")%>,<%=achou_FAC%>)" <%if Request.Form("hdnIdAcessoFisico") = objRSSolic("Acf_IDAcessoFisico") then Response.Write " checked " End if%>></td>
			<% end if %>	
			<td width=15% nowrap><%=objRSSolic("Acf_IDAcessoFisico") & " " & objRSSolic("Acf_Proprietario")%></td>
			<!-- <td width=17%><%=objRSSolic("Aec_Complemento")%></td> -->
			<td width=20%><%=objRSSolic("newfac_nome")%></td>
			<td width=20%><%=objRSSolic("newtec_nome")%></td>
			<td width=12%><%=objRSSolic("Pro_Nome")%></td>
			<!--<td width=30%><%=objRSSolic("Cli_Nome")%></td>-->
			<td width=7%><%=objRSSolic("Vel_Desc")%>&nbsp;<%=TipoVel(objRSSolic("Acf_TipoVel"))%></td>
			
			<td width=33%><%=objRSSolic("cli_nome")%></td>
			<!--<td width=12% nowrap><%=objRSSolic("Acf_NroAcessoPtaEbt")%></td> -->
		</tr>
	<%
' End If
	objRSSolic.MoveNext
	Wend
	
	if (qtd_reg = 0) then
	
		if ( exibir_compart = "N" and strOrisol_id =  "9" ) then
			Response.Write "<script language=javascript>alert('Não é possível compartilhar o acesso físico antes do Aceite!')</script>"
		end if
		if (strOrisol_id <> "9") then
			Response.Write "<script language=javascript>alert('ID(s) Físico(s) não Encontrados!')</script>"
    end if
	  Response.Write "<script language=javascript>parent.divIDFis" & intEnd & ".style.display = 'none';" & _
					"try{parent.spnBtnLimparIdFis" & intEnd & ".innerHTML ='';}catch(e){}" & _
				   "</script>"				   
		response.End
	end if
	
	Response.Write "<script language=javascript>parent.divIDFis" & intEnd & ".style.display = ''</script>"
	response.write "<script language=javascript>parent.document.Form2.btnAddAcesso.disabled=false</script>"
	response.write "<script language=javascript>parent.document.Form2.btnRemAcesso.disabled=false</script>"
	response.write "<script language=javascript>parent.document.Form2.btnIDFis1.disabled=false</script>"
	response.write "<script language=javascript>parent.document.Form2.btnLimparAcesso.disabled=false</script>"
	
	response.write "<script language=javascript>parent.document.Form2.cboInterFaceEnd.disabled=false</script>"
	response.write "<script language=javascript>parent.document.Form2.cboInterFaceEndFis.disabled=false</script>"

	strRet = "<script language=javascript>parent.divIDFis" & intEnd & ".style.display = '';"
	strRet = strRet & "try{parent.spnBtnLimparIdFis" & intEnd & ".innerHTML ='<input type=button class=button name=btnLimparIdFis" & intEnd & " value=""Limpar Id Físico"" onClick=""ReenviarSolicitacao("
	if intEnd = 1 then 
		strRet = strRet & "138,2"
	Else
		strRet = strRet & "139,2"
	End if
	strRet = strRet & ")"">'}catch(e){}</script>"
	Response.Write strRet

Else

if (strOrisol_id <> "9") then

	Response.Write "<script language=javascript>alert('ID(s) Físico(s) não Encontrados!')</script>"
end if
	Response.Write "<script language=javascript>parent.divIDFis" & intEnd & ".style.display = 'none';" & _
					"try{parent.spnBtnLimparIdFis" & intEnd & ".innerHTML ='';}catch(e){}" & _
				   "</script>"
End if				   
%>
</table>	
</Form>
</BODY>
</HTML>
