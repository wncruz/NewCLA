<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Acao_main.asp
'	- Descrição			: Lista/Remove Ação
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblId
Dim strPro
Dim strEst
Dim strCidade
Dim strContratadaContrato
Dim strNomeContratada
Dim strContratadaEnd
Dim strContratadaCep
Dim strContratadaTel
Dim strContratadaFax
Dim strContratadaContato
Dim strContratadaDepto
Dim strContratadaEmail
Dim strContratadaEst
Dim strContratadaCidade
Dim strContratadaNumAgregado
Dim StrCNPJContratante
Dim strNomeContrante
Dim strContratanteEnd
Dim strContratanteCep
Dim strContratanteTel
Dim strContratanteFax
Dim strContratanteContato
Dim strContratanteDepto
Dim strContratanteEmail
Dim strContratanteEst
Dim strContratanteCidade
dim strTemplate
Dim objRSPro
Dim strSel

'<!-- CH-28482YPU - Inicio -->
dim strIEContratante
'<!-- CH-28482YPU - Fim -->


dblId = Request.QueryString("ID")
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId")
End if

if Trim(Request.Form("hdnAcao"))="Gravar" then

	'Response.write(request.Form("txtIEContratante"))

	If dblId="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"	& dblId
	End if

	Vetor_Campos(2)="adInteger,2,adParamInput,"		& 	request.Form("cboProvedor")
	Vetor_Campos(3)="adWChar,2,adParamInput,"		& 	request.Form("cboUF")
	Vetor_Campos(4)="adWChar,4,adParamInput,"		& 	request.Form("cboLocalidade")
	Vetor_Campos(5)="adWChar,40,adParamInput,"		& 	request.Form("txtContratadaContrato")
	Vetor_Campos(6)="adWChar,80,adParamInput,"		&	request.Form("txtNome")
	Vetor_Campos(7)="adWChar,80,adParamInput,"		&	request.Form("txtEnd")
	Vetor_Campos(8)="adWChar,2,adParamInput,"		& 	request.Form("cboUFContratada")
	Vetor_Campos(9)="adWChar,4,adParamInput,"		& 	request.Form("cboLocalidadeContratada")
	Vetor_Campos(10)="adWChar,9,adParamInput,"		& 	request.Form("txtCep")
	Vetor_Campos(11)="adWChar,20,adParamInput,"		&	request.Form("txtTel")
	Vetor_Campos(12)="adWChar,20,adParamInput,"		&	request.Form("txtFax")
	Vetor_Campos(13)="adWChar,80,adParamInput,"		&	request.Form("txtContatoContratada")
	Vetor_Campos(14)="adWChar,80,adParamInput,"		&	request.Form("txtDeptoContratada")
	Vetor_Campos(15)="adWChar,300,adParamInput,"	& 	request.Form("txtContratadaEmail")
	Vetor_Campos(16)="adWChar,80,adParamInput,"		&	request.Form("txtNomeContratante")
	Vetor_Campos(17)="adWChar,80,adParamInput,"		&	request.Form("txtEndContratante")
	Vetor_Campos(18)="adWChar,2,adParamInput,"		& 	request.Form("cboUFContratante")
	Vetor_Campos(19)="adWChar,4,adParamInput,"		& 	request.Form("cboLocalidadeContratante")
	Vetor_Campos(20)="adWChar,9,adParamInput,"		& 	request.Form("txtCepContratante")
	Vetor_Campos(21)="adWChar,20,adParamInput,"		&	request.Form("txtTelContratante")
	Vetor_Campos(22)="adWChar,20,adParamInput,"		&	request.Form("txtFaxContratante")
	Vetor_Campos(23)="adWChar,80,adParamInput,"		&	request.Form("txtContatoContratante")
	Vetor_Campos(24)="adWChar,80,adParamInput,"		&	request.Form("txtDeptoContratante")
	Vetor_Campos(25)="adWChar,80,adParamInput,"		& 	request.Form("txtContratanteEmail")
	Vetor_Campos(26)="adWChar,30,adParamInput,"		& 	request.Form("txtTemplate")
	
	Vetor_Campos(27)="adInteger,3,adParamOutput,0"

	Vetor_Campos(28)="adWChar,20,adParamInput,"		&	request.Form("txtCNPJContratante")

	Vetor_Campos(29)="adWChar,20,adParamInput,"		&	request.Form("txtNumAgregadoContratada")
	Vetor_Campos(30)="adWChar,10,adParamInput," 	& 	strloginrede '-->PSOUTO 24/04/06
	
	
	Vetor_Campos(31)="adWChar,50,adParamInput,"		& 	request.Form("txtRepresentanteEmail")
	Vetor_Campos(32)="adWChar,50,adParamInput,"		&	request.Form("txtNomeRepresentanteComercial")
	Vetor_Campos(33)="adWChar,100,adParamInput,"	&	request.Form("txtEndRepresentante")
	Vetor_Campos(34)="adWChar,2,adParamInput,"		& 	request.Form("cboUFRepresentante")
	Vetor_Campos(35)="adWChar,4,adParamInput,"		& 	request.Form("cboLocalidadeRepresentante")
	Vetor_Campos(36)="adWChar,8,adParamInput,"		& 	request.Form("txtCepRepresentante")
	Vetor_Campos(37)="adWChar,20,adParamInput,"		&	request.Form("txtTelRepresentante")
	Vetor_Campos(38)="adWChar,20,adParamInput,"		&	request.Form("txtFaxRepresentante")

	
	'<!-- CH-28482YPU - Inicio -->
	'Obs: Alterar a procedure para receber o campo de incrição estadual 
	Vetor_Campos(39)="adWChar,20,adParamInput,"	&	request.Form("txtIEContratante")
	'<!-- CH-28482YPU - Fim -->

	on error resume next
	Call APENDA_PARAM("Cla_sp_ins_ContratoProvedor",39,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

	'Response.Write(ERR.number  &  "! "  & ERR.Description)
	'Response.End()
End if

if dblId<> "" then
	Set objRSPro = db.execute("CLA_sp_sel_ContratoProvedor " & dblId)
	if Not objRSPro.Eof and not objRSPro.Bof then
		 strPro					= TratarAspasHtml(Trim(objRSPro("Pro_ID")))
		 strEst					= TratarAspasHtml(Trim(objRSPro("Est_Sigla")))
		 strCidade				= TratarAspasHtml(Trim(objRSPro("Cid_Sigla")))
		 strContratadaContrato	= TratarAspasHtml(Trim(objRSPro("Cpro_ContratadaContrato")))
		 strNomeContratada		= TratarAspasHtml(Trim(objRSPro("CPro_ContratadaRazao")))
		 strContratadaEnd		= TratarAspasHtml(Trim(objRSPro("CPro_ContratadaEnd")))
		 strContratadaEst		= TratarAspasHtml(Trim(objRSPro("CPro_ContratadaEstSigla")))
		 strContratadaCidade	= TratarAspasHtml(Trim(objRSPro("CPro_ContratadaCidSigla")))
		 strContratadaCep		= TratarAspasHtml(Trim(objRSPro("CPro_ContratadaCEP")))
		 strContratadaTel		= TratarAspasHtml(Trim(objRSPro("CPro_ContratadaTelefone")))
		 strContratadaFax		= TratarAspasHtml(Trim(objRSPro("CPro_ContratadaFax")))
		 strContratadaContato	= TratarAspasHtml(Trim(objRSPro("CPro_ContratadaContato")))
		 strContratadaDepto		= TratarAspasHtml(Trim(objRSPro("CPro_ContratadaDepto")))
		 strContratadaEmail		= TratarAspasHtml(Trim(objRSPro("Cpro_ContratadaEmail")))
		 strContratadaNumAgregado = TratarAspasHtml(Trim(objRSPro("Cpro_Numero_Agregado")))

		 strCNPJContratante		= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteCGC_CNPJ")))

		 strNomeContratante		= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteRazao")))
		 strContratanteEnd		= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteEnd")))
		 strContratanteEst		= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteEstSigla")))
		 strContratanteCidade	= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteCidSigla")))
		 strContratanteCep		= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteCEP")))
		 strContratanteTel		= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteTelefone")))
		 strContratanteFax		= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteFax")))
		 strContratanteContato	= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteContato")))
		 strContratanteDepto	= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteDepto")))
		 strContratanteEmail	= TratarAspasHtml(Trim(objRSPro("Cpro_ContratanteEmail")))
		 strTemplate			= TratarAspasHtml(Trim(objRSPro("Cpro_Template")))
		 
		 
		 strNomeRepresentante	= TratarAspasHtml(Trim(objRSPro("CPro_RepresentanteNome")))
		 strRepresentanteEnd	= TratarAspasHtml(Trim(objRSPro("CPro_RepresentanteEnd")))
		 strRepresentanteEst	= TratarAspasHtml(Trim(objRSPro("CPro_RepresentanteEstSigla")))
		 strRepresentanteCidade	= TratarAspasHtml(Trim(objRSPro("CPro_RepresentanteCidSigla")))
		 strRepresentanteCep	= TratarAspasHtml(Trim(objRSPro("CPro_RepresentanteCEP")))
		 strRepresentanteTel	= TratarAspasHtml(Trim(objRSPro("CPro_RepresentanteTelefone")))
		 strRepresentanteFax	= TratarAspasHtml(Trim(objRSPro("CPro_RepresentanteFax")))
		 strRepresentanteEmail	= TratarAspasHtml(Trim(objRSPro("CPro_RepresentanteEmail")))

		'<!-- CH-28482YPU - Inicio -->
		 strIEContratante		= TratarAspasHtml(Trim(objRSPro("CPro_ContratanteIE")))
		 '<!-- CH-28482YPU - Fim -->'

	End if
Else

	strPro					= TratarAspasHtml(Trim(Request.Form("cboProvedor")))
	strEst					= TratarAspasHtml(Trim(Request.Form("cboUF")))
	strCidade				= TratarAspasHtml(Trim(Request.Form("cboLocalidade")))
	strContratadaContrato	= TratarAspasHtml(Trim(Request.Form("txtContratadaContrato")))
	strNomeContratada		= TratarAspasHtml(Trim(Request.Form("txtNome")))
	strContratadaEnd		= TratarAspasHtml(Trim(Request.Form("txtEnd")))
	strContratadaEst		= TratarAspasHtml(Trim(Request.Form("cboUFContratada")))
	strContratadaCidade		= TratarAspasHtml(Trim(Request.Form("cboLocalidadeContratada")))
	strContratadaCep		= TratarAspasHtml(Trim(Request.Form("txtCep")))
	strContratadaTel		= TratarAspasHtml(Trim(Request.Form("txtTel")))
	strContratadaFax		= TratarAspasHtml(Trim(Request.Form("txtFax")))
	strContratadaContato	= TratarAspasHtml(Trim(Request.Form("txtContatoContratada")))
	strContratadaDepto		= TratarAspasHtml(Trim(Request.Form("txtDeptoContratada")))
	strContratadaEmail		= TratarAspasHtml(Trim(Request.Form("txtContratadaEmail")))

	strContratadaNumAgregado = TratarAspasHtml(Trim(Request.Form("txtNumAgregadoContratada")))

	strCNPJContratante		= TratarAspasHtml(Trim(Request.Form("txtCNPJContratante")))

	strNomeContratante		= TratarAspasHtml(Trim(Request.Form("txtNomeContratante")))
	strContratanteEnd		= TratarAspasHtml(Trim(Request.Form("txtEndContratante")))
	strContratanteEst		= TratarAspasHtml(Trim(Request.Form("cboUFContratante")))
	strContratanteCidade	= TratarAspasHtml(Trim(Request.Form("cboLocalidadeContratante")))
	strContratanteCep		= TratarAspasHtml(Trim(Request.Form("txtCepContratante")))
	strContratanteTel		= TratarAspasHtml(Trim(Request.Form("txtTelContratante")))
	strContratanteFax		= TratarAspasHtml(Trim(Request.Form("txtFaxContratante")))
	strContratanteContato	= TratarAspasHtml(Trim(Request.Form("txtContatoContratante")))
	strContratanteDepto		= TratarAspasHtml(Trim(Request.Form("txtDeptoContratante")))
	strContratanteEmail		= TratarAspasHtml(Trim(Request.Form("txtContratanteEmail")))
	strTemplate				= TratarAspasHtml(Trim(Request.Form("txtTemplate")))
	
	
	strNomeRepresentante	= TratarAspasHtml(Trim(Request.Form("txtNomeRepresentanteComercial")))
	strRepresentanteEnd		= TratarAspasHtml(Trim(Request.Form("txtEndRepresentante")))
	strRepresentanteEst		= TratarAspasHtml(Trim(Request.Form("cboUFRepresentante")))
	strRepresentanteCidade	= TratarAspasHtml(Trim(Request.Form("cboLocalidadeRepresentante")))
	strRepresentanteCep		= TratarAspasHtml(Trim(Request.Form("txtCepRepresentante")))
	strRepresentanteTel		= TratarAspasHtml(Trim(Request.Form("txtTelRepresentante")))
	strRepresentanteFax		= TratarAspasHtml(Trim(Request.Form("txtFaxRepresentante")))
	strRepresentanteEmail	= TratarAspasHtml(Trim(Request.Form("txtRepresentanteEmail"))) 

	'<!-- CH-28482YPU - Inicio -->
	strIEContratante		= TratarAspasHtml(Trim(Request.Form("txtIEContratante")))
	'<!-- CH-28482YPU - Fim -->

End if
%>
<!--#include file="../inc/header.asp"-->


<form action="Estacao.asp" method="post" >
<input type=hidden name=hdnId value="<%=dblId%>" >

<SCRIPT LANGUAGE="JavaScript">
function ResgatarCidade(obj,intCid)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarCidade"
		hdnUFAtual.value = obj.value
		if (intCid == 0 ){
			hdnNomeCboCid.value = "Localidade"
		}
		if (intCid == 1 ){
			hdnNomeCboCid.value = "LocalidadeContratada"
		}
		if (intCid == 2 ){
			hdnNomeCboCid.value = "LocalidadeContratante"
		}
		if (intCid == 3 ){
			hdnNomeCboCid.value = "LocalidadeRepresentante"
		}

		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}



function GravarInfo()
{
	with (document.forms[0])
	{
		try{
		if (!ValidarCampos(cboProvedor,"O Provedor")) return false;
		if (!ValidarCampos(cboUF,"O UF")) return false
		if (!ValidarCampos(txtContratadaContrato,"O Contrato da Contratada")) return false
		if (!ValidarCampos(txtNome,"O Nome da Contratada")) return false
		if (!ValidarCampos(txtNomeContratante,"O Nome da Contratante")) return false
		if (!ValidarCampos(txtContratadaEmail,"O E-mail da Contratada")) return false
		if (!ValidarCampos(txtContratanteEmail,"O E-mail da Contratante")) return false
		}
		catch(e)
		{
			alert(e.description)
			return
		}

		hdnAcao.value = "Gravar"
		target = self.name
		action = "provedorcontrato.asp"
		submit()
	}
}

function CarregarDoc()
{
	document.resolveExternals = false;
}
CarregarDoc()

</script>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnNomeCboCid>
<input type=hidden name=hdnCidSel>
<input type=hidden name=hdnCTFResp>

<tr><td>
<table border=0 cellspacing="1" cellpadding="0" width="760">


<!--<tr><th colspan=5><p align="center">Cadastro de Informações Contratuais</p></th></tr>-->

<tr><th colspan=6><p style="FONT-SIZE: 14px" align=center>Cadastro de Informações Contratuais</p></th></tr>


<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Provedor</td>
<td colspan = 3>
<%
'seleciona os tipos de logradouro
set objRS = db.execute("CLA_sp_sel_provedor 0")
%>
	<select name="cboProvedor">
	<Option value=""></Option>
	<%do while not objRS.eof%>
		<option value="<%=trim(objRS("Pro_ID"))%>"
		<%If Trim(strPro) = Trim(objRS("Pro_ID")) then Response.write "selected" End if
		%>><%=trim(objRS("Pro_Nome"))%></option>
	<%
		objRS.movenext
	loop
	%>
	</select>
</td>
</tr>
<tr class=clsSilver>
	<td>
		<font class="clsObrig">:: </font>UF
	</td>
	<td colspan = 3>
		<select name="cboUF" onChange="ResgatarCidade(this,0)">
			<option value=""></option>
			<% set objRS = db.execute("CLA_sp_sel_Estado ''")
				While Not objRS.Eof
					strSel = ""
					if Trim(objRS("Est_Sigla")) = Trim(strEst) then strSel = " selected "
					Response.Write "<Option value="& objRS("Est_Sigla") & strSel & ">" & objRS("Est_Sigla") & "</Option>"
					objRS.MoveNext
				Wend
				Set objRS = Nothing
			%>
		</select>
	&nbsp;&nbsp;CNL Provedor&nbsp;&nbsp;
	<span id=spnLocalidade>
			<select name="cboLocalidade">
				<option value=""></option>
				<%
					if  strEst  <> "" then
						set objRS = db.execute("CLA_sp_view_estadocidade '" & strEst & "'")
						While Not objRS.Eof
							strSel = ""
							if Trim(objRS("Cid_Sigla")) = Trim(strCidade) then strSel = " selected "
							Response.Write	"<Option value="& objRS("Cid_Sigla") & strSel & ">" & objRS("Cid_Sigla") & " - " & objRS("Cid_Desc") & "</Option>"
							objRS.MoveNext
						Wend
						Set objRS = Nothing
					end if
				%>
			</select>
		</span>
	</td>
</tr>
<%if strloginrede = "RCCARD " OR strloginrede = "EDAR" OR strloginrede = "T3FRRP" OR strloginrede = "JOAOFNS" OR strloginrede = "FLAVIAB" then%>
<tr class=clsSilver>
<td>
	&nbsp;&nbsp;&nbsp;Formulário de envio
</td>
<td colspan = 3>
	<input type="text" class="text" name="txtTemplate" value="<%=strTemplate%>" maxlength="30" size="35" ID="Text1" <%if strloginrede = "JOAOFNS" then%>disabled<%end if%>>
</td>
</tr>
<%else%>
  <input type="hidden" class="text" name="txtTemplate" value="<%=strTemplate%>" ID="Text1">
<%end if%>

<tr>
	<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Dados da Contratada</th>
</tr>


<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Contrato</td>
	<td colspan = 3><input type="text" class="text" name="txtContratadaContrato" value="<%=strContratadaContrato%>" maxlength="80" size="35" ></td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Nome</td>
	<td colspan = 3><input type="text" class="text" name="txtNome" value="<%=strNomeContratada%>" maxlength="80" size="85" ></td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Endereço</td>
	<td colspan = 3><input type="text" class="text" name="txtEnd" value="<%=strContratadaEnd %>" maxlength="80" size="85" ></td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;UF</td>
	<td colspan = 3>
		<select name="cboUFContratada" onChange="ResgatarCidade(this,1)">
			<option value=""></option>
			<% set objRS = db.execute("CLA_sp_sel_Estado ''")
				While Not objRS.Eof
					strSel = ""
					if Trim(objRS("Est_Sigla")) = Trim(strContratadaEst) then strSel = " selected "
					Response.Write "<Option value="& objRS("Est_Sigla") & strSel & ">" & objRS("Est_Sigla") & "</Option>"
					objRS.MoveNext
				Wend
				Set objRS = Nothing
			%>
		</select>
	&nbsp;&nbsp;CNL
	<span id=spnLocalidadeContratada>
			<select name="cboLocalidadeContratada">
				<option value=""></option>
				<%
					if  strContratadaEst  <> "" then
						set objRS = db.execute("CLA_sp_view_estadocidade '" & strContratadaEst & "'")
						While Not objRS.Eof
							strSel = ""
							if Trim(objRS("Cid_Sigla")) = Trim(strContratadaCidade) then strSel = " selected "
							Response.Write	"<Option value="& objRS("Cid_Sigla") & strSel & ">" & objRS("Cid_Sigla") & " - " & objRS("Cid_Desc") & "</Option>"
							objRS.MoveNext
						Wend
						Set objRS = Nothing
					end if
				%>
			</select>
		</span>
	</td>
</tr>

<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;CEP</td>
	<td><input type="text" class="text" name="txtCep" value="<%=strContratadaCEP %>" maxlength="9" size="15" ></td>
	<td>&nbsp;&nbsp;&nbsp;Número de Agregado</td>
	<td><input type="text" class="text" name="txtNumAgregadoContratada" value="<%=strContratadaNumAgregado %>" maxlength="9" size="15" ></td>
</tr>

<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Tel</td>
	<td><input type="text" class="text" name="txtTel" value="<%=strContratadaTEL %>" maxlength="20" size="35" >	</td>
	<td>&nbsp;&nbsp;&nbsp;Fax</td>
	<td><input type="text" class="text" name="txtFax" value="<%=strContratadaFAX %>" maxlength="20" size="25" ></td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Email</td>
	<td colspan = 3><input type="text" class="text" name="txtContratadaEmail" value="<%=strContratadaEmail%>" maxlength="300" size="35" alfatipo = "min"  >&nbsp;&nbsp;&nbsp;(E-mail de destino da carta)</td>
</tr>

<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Contato</td>
	<td><input type="text" class="text" name="txtContatoContratada" value="<%=strContratadaContato%>" maxlength="80" size="35" ></td>
	<td>&nbsp;&nbsp;&nbsp;Depto</td>
	<td><input type="text" class="text" name="txtDeptoContratada" value="<%=strContratadaDepto%>" maxlength="80" size="25" ></td>
</tr>

<tr>
	<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Dados do Contratante</th>
</tr>


<!-- CH-28482YPU - Inicio -->
<!--
<tr class=clsSilver>
	<td><font class="clsObrig"> </font>Contratante CNPJ</td>
	<td colspan = 3><input type="text" class="text" name="txtCNPJContratante" value="<%=strCNPJContratante%>" maxlength="20" size="25" ></td>
</tr>
-->

<tr class=clsSilver>
	<td><font class="clsObrig"> </font>&nbsp;&nbsp;&nbsp;CNPJ</td>
	<td colspan = 1><input type="text" class="text" name="txtCNPJContratante" value="<%=strCNPJContratante%>" maxlength="20" size="25" ></td>

	<td><font class="clsObrig"> </font>&nbsp;&nbsp;&nbsp;Inscrição Estadual</td>
	<td colspan = 1><input type="text" class="text" name="txtIEContratante" value="<%=strIEContratante%>" maxlength="20" size="25" ></td>

</tr>

<!-- CH-28482YPU - Fim -->

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Nome</td>
	<td colspan = 3><input type="text" class="text" name="txtNomeContratante" value="<%=strNomeContratante%>" maxlength="80" size="85" ></td>
</tr>

<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Endereço</td>
	<td colspan = 3><input type="text" class="text" name="txtEndContratante" value="<%=strContratanteEnd %>" maxlength="80" size="85" ></td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;UF</td>
	<td colspan = 3>
		<select name="cboUFContratante" onChange="ResgatarCidade(this,2)">
			<option value=""></option>
			<% set objRS = db.execute("CLA_sp_sel_Estado ''")
				While Not objRS.Eof
					strSel = ""
					if Trim(objRS("Est_Sigla")) = Trim(strContratanteEst) then strSel = " selected "
					Response.Write "<Option value="& objRS("Est_Sigla") & strSel & ">" & objRS("Est_Sigla") & "</Option>"
					objRS.MoveNext
				Wend
				Set objRS = Nothing
			%>
		</select>
		&nbsp;&nbsp;CNL Contratante
		<span id=spnLocalidadeContratante>
			<select name="cboLocalidadeContratante">
				<option value=""></option>
				<%
					if  strContratanteEst  <> "" then
						set objRS = db.execute("CLA_sp_view_estadocidade '" & strContratanteEst & "'")
						While Not objRS.Eof
							strSel = ""
							if Trim(objRS("Cid_Sigla")) = Trim(strContratanteCidade) then strSel = " selected "
							Response.Write	"<Option value="& objRS("Cid_Sigla") & strSel & ">" & objRS("Cid_Sigla") & " - " & objRS("Cid_Desc") & "</Option>"
							objRS.MoveNext
						Wend
						Set objRS = Nothing
					end if
				%>
			</select>
		</span>
	</td>
	</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;CEP</td>
	<td colspan = 3><input type="text" class="text" name="txtCepContratante" value="<%=strContratanteCep %>" maxlength="9" size="15" ></td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Tel</td>
	<td><input type="text" class="text" name="txtTelContratante" value="<%=strContratanteTel %>" maxlength="20" size="35" ></td>
	<td>&nbsp;&nbsp;&nbsp;Fax</td>
	<td><input type="text" class="text" name="txtFaxContratante" value="<%=strContratanteFax %>" maxlength="20" size="25" ></td>	</td>

</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Email</td>
	<td colspan = 3><input type="text" class="text" name="txtContratanteEmail" value="<%=strContratanteEmail%>" maxlength="80" size="35" alfatipo = "min" >&nbsp;&nbsp;&nbsp;(E-mail de origem da carta)</td>
</tr>
<tr class=clsSilver>
	<td>&nbsp;&nbsp;&nbsp;Contato</td>
	<td><input type="text" class="text" name="txtContatoContratante" value="<%=strContratanteContato%>" maxlength="80" size="35" ></td>
	<td>&nbsp;&nbsp;&nbsp;Depto</td>
	<td><input type="text" class="text" name="txtDeptoContratante" value="<%=strContratanteDepto%>" maxlength="80" size="25" ></td>
</tr>

<!-- resolucao 600 -->

<tr>
	<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Dados do Representante Comercial</th>
</tr>


	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Nome</td>
		<td colspan = 3><input type="text" class="text" name="txtNomeRepresentanteComercial" value="<%=strNomeRepresentante%>" maxlength="80" size="85" ></td>
	</tr>

	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Endereço</td>
		<td colspan = 3><input type="text" class="text" name="txtEndRepresentante" value="<%=strRepresentanteEnd %>" maxlength="80" size="85" ></td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>UF</td>
		<td colspan = 3>
			<select name="cboUFRepresentante" onChange="ResgatarCidade(this,3)">
				<option value=""></option>
				<% set objRS = db.execute("CLA_sp_sel_Estado ''")
					While Not objRS.Eof
						strSel = ""
						if Trim(objRS("Est_Sigla")) = Trim(strRepresentanteEst) then strSel = " selected "
						Response.Write "<Option value="& objRS("Est_Sigla") & strSel & ">" & objRS("Est_Sigla") & "</Option>"
						objRS.MoveNext
					Wend
					Set objRS = Nothing
				%>
			</select>
			<font class="clsObrig">&nbsp;&nbsp;&nbsp;:: </font>CNL
			<span id=spnLocalidadeRepresentante>
				<select name="cboLocalidadeRepresentante">
					<option value=""></option>
					<%
						if  strRepresentanteEst  <> "" then
							set objRS = db.execute("CLA_sp_view_estadocidade '" & strRepresentanteEst & "'")
							While Not objRS.Eof
								strSel = ""
								if Trim(objRS("Cid_Sigla")) = Trim(strRepresentanteCidade) then strSel = " selected "
								Response.Write	"<Option value="& objRS("Cid_Sigla") & strSel & ">" & objRS("Cid_Sigla") & " - " & objRS("Cid_Desc") & "</Option>"
								objRS.MoveNext
							Wend
							Set objRS = Nothing
						end if
					%>
				</select>
			</span>
		</td>
		</td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>CEP</td>
		<td colspan = 3><input type="text" class="text" name="txtCepRepresentante" value="<%=strRepresentanteCep %>" maxlength="9" size="15" ></td>
	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Tel</td>
		<td><input type="text" class="text" name="txtTelRepresentante" value="<%=strRepresentanteTel %>" maxlength="20" size="35" ></td>
		<td>&nbsp;&nbsp;&nbsp;Fax</td>
		<td><input type="text" class="text" name="txtFaxRepresentante" value="<%=strRepresentanteFax %>" maxlength="20" size="25" ></td>	</td>

	</tr>
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Email</td>
		<td colspan = 3><input type="text" class="text" name="txtRepresentanteEmail" value="<%=strRepresentanteEmail%>" maxlength="80" size="35" alfatipo = "min" ></td>
	</tr>

<!-- resolucao 600 -->

</table>
<table width="760" border=0>
<tr >
	<td colspan=2 align="center"><br>
<!-- ***** Causa-raíz: CH-42480KAR -->		


		<input type="button" class="button" name="btnGravar" value="Gravar" onclick="GravarInfo()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('txtSigla');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		
		<!--
		<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('provedorcontrato_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
		-->
	</td>
</tr>
</table>
<table width="760" border=0>
<tr>
<td>
<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
</td>
</tr>
</table>
</td>
</tr>
</table>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso"
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no"
	    align       = "left">
</iFrame>
</form>
</body>
</html>