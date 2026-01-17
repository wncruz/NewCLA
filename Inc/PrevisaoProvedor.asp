<!--#include file="../inc/data.asp"-->
<Html>
<Head>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
</Head>
<Body topmargin=0 leftmargin=0>
<SCRIPT LANGUAGE=javascript>
<!--
function GravarPrevisaoProvedor()
{
	with (document.forms[0])
	{
		
		
		if (!ValidarCampos(Form1.txtDtPrev,"Previsão de Entrega pelo Fornecedor")) return
		
		if (!ValidarCampos(Form1.cboProjetoEspecial,"Tipo de EILD")) return
		
		if (hdnEILD.value == "S")
		{
			if (!ValidarCampos(Form1.txtProjEsp_NumeroProposta,"Número de compra SNOA:")) return
		}
		
		
		if (cboProjetoEspecial.value == "Especial")
		  {
			if (!ValidarCampos(Form1.txtCusto,"Custo")) return
		  }
		
		if (!ValidarTipoInfo(Form1.txtDtPrev,1,"Previsão de Entrega pelo Fornecedor")) return;
		
		if (!ValidarTipoInfo(Form1.txtDtReaprPrev,1,"Previsão Reaprazada de Entrega pelo Fornecedor")) return;
		
		hdnAcao.value = "GravarPrevisaoProvedor"
		hdnGravouPrevisao.value = "Sim"
		target = "IFrmPrevisaoProvedor"
		action = "PrevisaoProvedor.asp"
		submit()
	}
}

//-->
</SCRIPT>


<form name=Form1 method=Post >
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnGravouPrevisao>
<input type=hidden name=hdnAcf_ID value="<%=Request.QueryString("dblAcfId")%>">
<input type=hidden name=hdnSol_ID value="<%=Request.QueryString("dblSolId")%>">
<input type=hidden name=hdnEILD value="<%=Request.QueryString("dblEild")%>">
<input type=hidden name=hdnPonta value="<%=Request.QueryString("dblPonta")%>">

<%
Dim AcfID
Dim strDtPrevEntregaProvedor
Dim strDtReaprPrevEntregaProvedor
Dim StrProjetoEspecial
Dim StrProjEsp_NumeroProposta
Dim StrProjEsp_Prazo

AcfID = Request.QueryString("dblAcfId")

SolId = Request.QueryString("dblSolId")
strDtPrevEntregaProvedor = ""
strDtReaprPrevEntregaProvedor = ""
StrProjEsp_NumeroProposta = ""
StrProjEsp_Prazo = ""


If	Trim(Request.form("hdnAcao")) = "GravarPrevisaoProvedor" And Request.Form("hdnAcf_ID") > 0 Then 

	strDtPrevEntregaProvedor = Mid(Trim(Request.Form("txtDtPrev")),7,4) + Mid(Trim(Request.Form("txtDtPrev")),4,2) + Mid(Trim(Request.Form("txtDtPrev")),1,2)
	StrProjetoEspecial = Trim(Request.Form("cboProjetoEspecial"))
	StrProjEsp_NumeroProposta = Trim(Request.Form("txtProjEsp_NumeroProposta"))
	StrProjEsp_PrazoEst = Trim(Request.Form("txtProjEsp_Prazo"))
	StrCusto = Trim(Request.Form("txtCusto"))

	If Trim(Request.Form("txtDtReaprPrev")) <> "" And Not IsNull(Trim(Request.Form("txtDtReaprPrev"))) Then
	
		strDtReaprPrevEntregaProvedor = Mid(Trim(Request.Form("txtDtReaprPrev")),7,4) + Mid(Trim(Request.Form("txtDtReaprPrev")),4,2) + Mid(Trim(Request.Form("txtDtReaprPrev")),1,2)	
	
	Else

		strDtReaprPrevEntregaProvedor = Null
		
	End If
	
	''@@ Davif - Atualiza Informações na Tabela Cla_AcessoFisico
	
	StrSql = "Update Cla_AcessoFisico "
	StrSql = StrSql & " Set Acf_DtPrevistaEntregaAcesso  = '" &	strDtPrevEntregaProvedor & "', "
	
	If  Not IsNull(Trim(strDtReaprPrevEntregaProvedor)) Then
		
		StrSql = StrSql & " Acf_DtReaprPrevistaEntregaAcesso = '" & strDtReaprPrevEntregaProvedor & "', " 
	
	End If
	
	StrSql = StrSql & " Acf_ProjetoEspecial = '" & StrProjetoEspecial & "', "
	StrSql = StrSql & " Acf_ProjEsp_NumeroProposta = '" & StrProjEsp_NumeroProposta & "', " 
	StrSql = StrSql & " Acf_ProjetoEspecialCustoEst = '" & StrCusto & "', " 
	StrSql = StrSql & " Acf_ProjEsp_Prazo = '" & StrProjEsp_Prazo & "' " 
	
	StrSql = StrSql & " Where Acf_ID = " & Request.Form("hdnAcf_ID")

	'Response.Write "<script language=javascript>alert(" & StrSql & ");</script>"
	db.Execute(strsql)
	
	DBErro = 0
	
	if Request.Form("hdnEILD") = "S" then
	
		Vetor_Campos(1)="adInteger,10,adParamInput, " & Request.Form("hdnSol_ID")
		Vetor_Campos(2)="adWChar,20,adParamInput, " & trim(StrProjEsp_NumeroProposta)
		Vetor_Campos(3)="adWChar,20,adParamInput, " &  Request.Form("hdnPonta")
		Vetor_Campos(4)="adInteger,10,adParamInput, " & Request.Form("hdnAcf_ID")
		Vetor_Campos(5)="adInteger,10,adParamOutput,0"
	
		Call APENDA_PARAM("CLA_SP_INS_ASSOCLOGICOSNOA",5,Vetor_Campos)
	
		ObjCmd.Execute'pega dbaction
		
		DBErro = ObjCmd.Parameters("RET").value
		
	end if 
	
	
	
	if DBErro = 0 then
		Response.Write "<script language=javascript>alert('Registro Gravado com Sucesso !');</script>"
		
	end if
	
	
	
	
End If


If AcfID > 0 Then

	strsql = "Select Acf_DtPrevistaEntregaAcesso, Acf_DtReaprPrevistaEntregaAcesso, Acf_ProjetoEspecial, "
	strsql = StrSql & " Acf_ProjEsp_NumeroProposta, Acf_ProjEsp_Prazo, Acf_ProjetoEspecialCustoEst "
	strsql = StrSql & " From Cla_AcessoFisico Where Acf_id = " & AcfID
	
	Set objRS = db.Execute(strsql)
	
	If Not ObjRS.eof and Not ObjRS.bof Then 
		strDtPrevEntregaProvedor = Formatar_Data(ObjRS("Acf_DtPrevistaEntregaAcesso"))
		strDtReaprPrevEntregaProvedor = Formatar_Data(ObjRS("Acf_DtReaprPrevistaEntregaAcesso"))
		StrProjetoEspecial = ObjRS("Acf_ProjetoEspecial")
		StrProjEsp_NumeroProposta = ObjRS("Acf_ProjEsp_NumeroProposta")
		StrProjEsp_Prazo = ObjRS("Acf_ProjEsp_Prazo")
		StrCusto = ObjRS("Acf_ProjetoEspecialCustoEst")
	End If

End If 



%>
	
<iframe	id			= "IFrmPrevisaoProvedor"
	    name        = "IFrmPrevisaoProvedor"
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "auto"
	    align       = "left">
</iFrame>

<table border=0 cellspacing="1" cellpadding="0"width="100%">
<tr><th colspan=4 >&nbsp;•&nbsp;Previsões de Entrega do Provedor</th></tr>
 <tr class=clsSilver>
	 <td width=170px rowspan="2" ><font class="clsObrig">:: </font>Data prevista de entrega pelo provedor</td>
	<td width=170px rowspan="2">

<% 'Se a data já foi preenchida não poderá ser alterada

	If Trim(strDtPrevEntregaProvedor) = "" or isnull(strDtPrevEntregaProvedor) Then 

%>
	
		<input type="text" class="text" size="10" name="txtDtPrev" VALUE="<%=strDtPrevEntregaProvedor%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa) 

<%Else %>

		<input type="text" class="text" readonly size="10" name="txtDtPrev" VALUE="<%=strDtPrevEntregaProvedor%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa) 

<%End IF%>	

	</td>
	
	<td colspan="2"><font class="clsObrig">:: </font> Tipo de EILD:  
		<select name="cboProjetoEspecial" >
			<option value=""></option>
			<option value="Padrao" <%if strProjetoEspecial = "Padrao" then Response.Write " selected " End If%> >Padrão </option>
			<option value="Especial" <%if strProjetoEspecial = "Especial" then Response.Write " selected " End If%> >Especial </option>
		</select> &nbsp;&nbsp;Prazo: <input type="text" class="text" size="5" name="txtProjEsp_Prazo" VALUE="<%=StrProjEsp_Prazo%>" maxlength="3" size="3" onKeyPress="OnlyNumbers()" > (Dias)
	</td>

</tr>


 <tr class=clsSilver>
	
	<td width=170x><font class="clsObrig">:: </font> Custo (R$):</td>

	<td ><input type="text" class="text" name="txtCusto" value="<%=StrCusto%>" maxlength="14" size="20" tabindex="3" onKeyPress="return(MascaraMoeda(this,'.',',',event,14))"></td>
	
</tr>


<tr class=clsSilver>
	<td width=170>Data Reaprazada prevista de entrega pelo provedor</td>
	<td>
	<input type="text" class="text" size="10" name="txtDtReaprPrev" VALUE="<%=strDtReaprPrevEntregaProvedor%>" maxlength="10" size="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaaa) 
	</td>
	
	<!--<td width=170px><b>Número de compra SNOA:</b> </td> -->
	<td width=170px> </td>

	<td > 

		<!--<input type="text" class="text" size="15" name="txtProjEsp_NumeroProposta" VALUE="<%=StrProjEsp_NumeroProposta%>" maxlength="15"  onKeyPress="OnlyNumbers()"> -->
		
<%'If (Trim(strDtPrevEntregaProvedor) = "" or isnull(strDtPrevEntregaProvedor) Or Trim(strDtReaprPrevEntregaProvedor) = "" or isnull(strDtReaprPrevEntregaProvedor)) Then  %>

			<input type=button name=btnGravar value="Gravar Previsão" class=button onclick="GravarPrevisaoProvedor()" accesskey="H" onmouseover="showtip(this,event,'Gravar Previsão (Alt+H)');">

<%'End If %>


	</td>
</tr>
</Form>


</table>
</Body>
</Html>