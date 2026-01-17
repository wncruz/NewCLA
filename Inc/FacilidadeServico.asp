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
	
<iframe	id			= "IFrmFacilidadeServico"
	    name        = "IFrmFacilidadeServico"
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "auto"
	    align       = "left">
</iFrame>

<table border=0 cellspacing="1" cellpadding="0"width="100%">
		<tr><th colspan=4 >&nbsp;•&nbsp;Facilidade de Serviço</th></tr>
		<tr class=clsSilver>
			<td width="170px" nowrap><font class="clsObrig">:: </font>VLAN </td>
			<td colspan="9" >
				<input type=text class=text name=Vlan  size='5' maxlength='5'  onKeyUp="ValidarTipo(this,0)">
			</td>
		</tr>
		<tr class=clsSilver2> 
			<td width="170px" nowrap><font class="clsObrig">:: </font>Porta</td>
			<td colspan="9" >
				<input type=text class=text name=portaOLt  size='16' maxlength='16' > 
			</td>
		</tr>
		<tr class=clsSilver>
			<td width="170px" nowrap><font class="clsObrig">:: </font>SVLAN</td>
			<td colspan="9" >
				<input type=text class=text name=SVLAN  size='5' maxlength='5' onKeyUp="ValidarTipo(this,0)">
			</td>
		</tr>
		<tr class=clsSilver2>
			<td width="170px" nowrap><font class="clsObrig">:: </font>PE</td>
			<td colspan="9" >
				<input type=text class=text name=PE  size='40' maxlength='40' >
			</td>
		</tr>
		<tr class=clsSilver>
			<td colspan="9" >
				<input type=button name=btnGravar value="Gravar Previsão" class=button onclick="GravarPrevisaoProvedor()" accesskey="H" onmouseover="showtip(this,event,'Gravar Previsão (Alt+H)');">

			</td>
		</tr>
</Form>


</table>
</Body>
</Html>