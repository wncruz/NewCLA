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
<tr>
	<th width=3%>&nbsp;Sel&nbsp;</th>
	<th width=15%>&nbsp;ID Físico</th>
	<th width=16%>&nbsp;Compl End</th>
	<th width=14%>&nbsp;Provedor</th>
	<th width=36%>&nbsp;Cliente</th>
	<th width=15%>&nbsp;Vel Fis</th>
	<th width=15% nowrap>&nbsp;NºAcesso</th>
</tr>

<%
intEnd = Request.QueryString("intEnd") 

if intEnd = 1 then
	strCidSigla =  Trim(Request.Form("txtEndCid"))
	strUF		=  Trim(Request.Form("cboUFEnd"))
	strTplSigla	=  Trim(Request.Form("cboLogrEnd"))
	strEnd		=  Trim(Request.Form("txtEnd"))
	strNroEnd	=  Trim(Request.Form("txtNroEnd"))
	strCompl	=  Trim(Request.Form("txtComplEnd"))
	strCep		=  Trim(Request.Form("txtCepEnd"))
Else
	strCidSigla =  Trim(Request.Form("txtEndCidInterme"))
	strUF		=  Trim(Request.Form("cboUFEndInterme"))
	strTplSigla	=  Trim(Request.Form("cboLogrEndInterme"))
	strEnd		=  Trim(Request.Form("txtEndInterme"))
	strNroEnd	=  Trim(Request.Form("txtNroEndInterme"))
	strCompl	=  Trim(Request.Form("txtComplEndInterme"))
	strCep		=  Trim(Request.Form("txtCepEndInterme"))
End if	

if Trim(strCidSigla) = ""  or Trim(strUF) = "" or Trim(strTplSigla) = "" or  Trim(strEnd) = ""  or Trim(strNroEnd) = "" or Trim(strCep) = "" then
	Response.Write "<script language=javascript>alert('Os Campos uf,cnl,sigla do logradouro,\nlogradouro,número e CEP são obrigatórios!')</script>"
	Response.Write "<script language=javascript>parent.divIDFis" & intEnd & ".style.display = 'none';" & _
				   "</script>"
	Response.End
End if

Vetor_Campos(1)="adInteger,2,adParamInput,"
Vetor_Campos(2)="adWChar,100,adParamInput,"
Vetor_Campos(3)="adWChar,30,adParamInput,"
Vetor_Campos(4)="adWChar,60,adParamInput,"
Vetor_Campos(5)="adWChar,4,adParamInput,"	& strCidSigla	'Cid Sigla
Vetor_Campos(6)="adWChar,2,adParamInput,"	& strUF			'UF
Vetor_Campos(7)="adWChar,15,adParamInput,"	& strTplSigla	'Tlp Sigla
Vetor_Campos(8)="adWChar,60,adParamInput,"	& strEnd
Vetor_Campos(9)="adWChar,10,adParamInput,"	& strNroEnd	 
Vetor_Campos(10)="adWChar,30,adParamInput,"	& strCompl	 
Vetor_Campos(11)="adWChar,9,adParamInput,"	& strCep		'CEP
Vetor_Campos(12)="adWChar,1,adParamInput,F"

strSql = APENDA_PARAMSTR("CLA_sp_sel_AcessoFisico",12,Vetor_Campos)
Set objRSSolic = db.Execute(strSql)

blnCor = true
if Not objRSSolic.Eof and Not objRSSolic.Bof then 
	While Not objRSSolic.Eof 
		if strClass = "clsSilver" then strClass = "clsSilver2" else strClass = "clsSilver" End if
	%>
		<tr class="<%=strClass%>" width=100% >
			<td width=3%><input type=radio name=rdoIDFis<%=intEnd%> value="<%=objRSSolic("Acf_IDAcessoFisico")%>" prop="<%=objRSSolic("Acf_Proprietario")%>" onClick="parent.SelIDFisComp(this,<%=intEnd%>)" ></td>
			<td width=15% nowrap><%=objRSSolic("Acf_IDAcessoFisico") & " " & objRSSolic("Acf_Proprietario")%></td>
			<td width=17%><%=objRSSolic("Aec_Complemento")%></td>
			<td width=12%><%=objRSSolic("Pro_Nome")%></td>
			<td width=36%><%=objRSSolic("Cli_Nome")%></td>
			<td width=12%><%=objRSSolic("Vel_Desc")%></td>
			<td width=12%><%=objRSSolic("Acf_NroAcessoPtaEbt")%></td>
		</tr>
	<%objRSSolic.MoveNext
	Wend
	Response.Write "<script language=javascript>parent.divIDFis" & intEnd & ".style.display = ''</script>"

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
	Response.Write "<script language=javascript>alert('ID(s) Físico(s) não Encontrados!')</script>"
	Response.Write "<script language=javascript>parent.divIDFis" & intEnd & ".style.display = 'none';" & _
				"</script>"
End if
%>
</table>	
</Form>
</BODY>
</HTML>
