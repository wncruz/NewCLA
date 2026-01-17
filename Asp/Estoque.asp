<!--#include file="../inc/data.asp"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
</HEAD>
<BODY topmargin=0 leftmargin=0 class=TA>
<form name=Form1 method=Post>
<table width=100% border=0 cellspacing="1" cellpadding="1">
<tr>
	<th width=3%>&nbsp;Sel</th>
	<th width=15%>&nbsp;ID Físico</th>
	<th width=42%>&nbsp;Endereço</th>
	<th width=14% nowrap>&nbsp;Nº Acesso</th>
	<th width=15%>&nbsp;Vel Fis</th>
	<th width=15% nowrap>&nbsp;Pz Contr.</th>
	<th width=15% nowrap>&nbsp;Pz Rest.</th>
</tr>
<%
if Trim(Request.Form("cboProvedor")) = ""  then
	Response.Write "<script language=javascript>alert('Provedor é um Campo Obrigatório!')</script>"
	Response.Write "<script language=javascript>try{parent.divIDFis" & intEnd & ".style.display = 'none';}catch(e){}</script>"
	Response.End 
End if

intEnd = Request.QueryString("intEnd") 

if intEnd = 1 then
	strCidSigla =  Trim(Request.Form("txtEndCidInstala"))
	strUF		=  Trim(Request.Form("cboUFEndInstala"))
	strTplSigla	=  Trim(Request.Form("cboLogrEndInstala"))
	strEnd		=  Trim(Request.Form("txtEndInstala"))
	strNroEnd	=  Trim(Request.Form("txtNroEndInstala"))
	strCompl	=  Trim(Request.Form("txtComplEndInstala"))
	strCep		=  Trim(Request.Form("txtCepEndInstala"))
Else
	strCidSigla =  Trim(Request.Form("txtEndCidInterme"))
	strUF		=  Trim(Request.Form("cboUFEndInterme"))
	strTplSigla	=  Trim(Request.Form("cboLogrEndInterme"))
	strEnd		=  Trim(Request.Form("txtEndInterme"))
	strNroEnd	=  Trim(Request.Form("txtNroEndInterme"))
	strCompl	=  Trim(Request.Form("txtComplEndInterme"))
	strCep		=  Trim(Request.Form("txtCepEndInterme"))
End if	


dblProId = Request.Form("cboProvedor")
if dblProId = "" then dblProId = "null"

if Trim(Request.Form("hdnCtfcIdGLA")) <> "" then
	dblCtfcId = Request.Form("hdnCtfcIdGLA")
Else
	dblCtfcId = "null"
End if	

Set objRSSolic = db.Execute("CLA_SP_sel_Estoque " & dblProId & "," & dblCtfcId)

blnCor = true
if Not objRSSolic.Eof and Not objRSSolic.Bof then 
	While Not objRSSolic.Eof 
	strEndereco = Trim(objRSSolic("Tpl_Sigla")) & "&nbsp;" & Trim(objRSSolic("End_NomeLogr")) & ",&nbsp;" & Trim(objRSSolic("End_NroLogr")) & "&nbsp;" & Trim(objRSSolic("Aec_Complemento")) & " - " & Trim(objRSSolic("End_Cep")) & " " & Trim(objRSSolic("Cid_Sigla"))  & " " & Trim(objRSSolic("Est_Sigla")) 
	if strClass = "clsSilver" then strClass = "clsSilver2" else strClass = "clsSilver" End if
	%>
		<tr class="<%=strClass%>" width=100% >
			<td width=3%> <input type=radio name=rdoIDFis<%=intEnd%> value="<%=objRSSolic("Acf_IDAcessoFisico")%>" prop="<%=objRSSolic("Acf_Proprietario")%>" onClick="parent.SelIDFisComp(this,2,'')" ></td>
			<td width=15% nowrap><%=objRSSolic("Acf_IDAcessoFisico") & " " & objRSSolic("Acf_Proprietario")%></td>
			<td width=17%><%=strEndereco%></td>
			<td width=12%><%=objRSSolic("Acf_NroAcessoPtaEbt")%></td>
			<td width=12%><%=objRSSolic("Vel_Desc")%></td>
			<td width=12%><%=objRSSolic("Tct_Desc")%></td>
			<td width=12%><%=objRSSolic("PrazoRestante")%></td>
		</tr>
	<%objRSSolic.MoveNext
	Wend
	Response.Write "<script language=javascript>parent.divIDFis" & intEnd & ".style.display = '';" & _
				   "</script>"
Else
	Response.Write "<script language=javascript>alert('Estoque não Encontrado!')</script>"
	Response.Write "<script language=javascript>parent.divIDFis" & intEnd & ".style.display = 'none';" & _
				   "</script>"
End if				   
%>
</table>	
</Form>
</BODY>
</HTML>