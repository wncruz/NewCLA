<%
'	- Sistema			: CLA
'	- Arquivo			: CorrigirFacMig.asp
'	- Descrição		: Corrige uma Facilidade migrada com erro
%>

<!--#include file="../inc/data.asp"-->

<%

Dim strProId
Dim strRede
Dim strLocalInstala
Dim strDistrib
dim strUF
Dim intCount
Dim strCepConsulta



REM Dados da facilidade migrado, cliente, e acesso
Dim strAcf_ID
Dim strAcf_IDAcessoFisico
Dim strAcl_IDAcessoLogico
Dim strCli_Nome
Dim strDesignacaoServico
Dim strEstacao
Dim strProvedor
Dim strDistribuidor
Dim strSistema
Dim strDominio
Dim strNo
Dim strSlot
Dim strPorta
Dim strTime_Slot
Dim strFila
Dim strBastidor
Dim strRegua
Dim strPosicao
Dim strTronco
Dim strProNome
Dim strPar
Dim strReferencia
Dim strAcf_NroAcessoPtaEbt '-- Designação do Provedor
Dim intIdFis
Dim intIDLog
Dim intEstacao
Dim intRede
Dim DM
Dim FiltroDM


intRede = Request("hdnRede")
intEstacao = Request("hdnEstacao")

intIdFis = request("IDFis")
intIDLog = request("IDLog")

REM Obtém o FiltroDM -------------------------------------------------
FiltroDM = Request.Form("hdnFiltroDM") 
REM -----------------------------------------------------------------------


strCboRet = ""

Set ObjRsMig = db.execute("CLA_SP_Obter_Dados_Fac_Acesso_Mig " & intIdFis & ", " & intIDLog )

if Not ObjRsMig.Eof and not ObjRsMig.Bof then
	strAcf_ID				= TratarAspasHtml(Trim(ObjRsMig("Acf_ID")))
	strAcl_IDAcessoLogico	= TratarAspasHtml(Trim(ObjRsMig("Acl_IDAcessoLogico")))
	strAcf_IDAcessoFisico	= TratarAspasHtml(Trim(ObjRsMig("Acf_IDAcessoFisico")))
	strCli_Nome				= TratarAspasHtml(Trim(ObjRsMig("Cli_Nome")))
	strDesignacaoServico	= TratarAspasHtml(Trim(ObjRsMig("DesignacaoServico")))
	strEstacao				= TratarAspasHtml(Trim(ObjRsMig("Estacao")))
	strProvedor				= TratarAspasHtml(Trim(ObjRsMig("Provedor")))
	strDistribuidor			= TratarAspasHtml(Trim(ObjRsMig("Distribuidor")))
	strSistema				= TratarAspasHtml(Trim(ObjRsMig("Sistema")))
	strDominio				= TratarAspasHtml(Trim(ObjRsMig("Dominio")))
	strNo					= TratarAspasHtml(Trim(ObjRsMig("No")))
	strSlot					= TratarAspasHtml(Trim(ObjRsMig("Slot")))
	strPorta				= TratarAspasHtml(Trim(ObjRsMig("Porta")))
	strTime_Slot			= TratarAspasHtml(Trim(ObjRsMig("Time_Slot")))
	strFila					= TratarAspasHtml(Trim(ObjRsMig("Fila")))
	strBastidor				= TratarAspasHtml(Trim(ObjRsMig("Bastidor")))
	strRegua				= TratarAspasHtml(Trim(ObjRsMig("Regua")))
	strPosicao				= TratarAspasHtml(Trim(ObjRsMig("Posicao")))
	strTronco				= TratarAspasHtml(Trim(ObjRsMig("Tronco")))
	strPar					= TratarAspasHtml(Trim(ObjRsMig("Par")))
	strProId				= TratarAspasHtml(Trim(ObjRsMig("Pro_ID")))
	strRede					= TratarAspasHtml(Trim(ObjRsMig("Sis_ID")))
	strLocalInstala			= TratarAspasHtml(Trim(ObjRsMig("Esc_ID")))
	strDistrib				= TratarAspasHtml(Trim(ObjRsMig("Dst_ID")))
	strUF					= TratarAspasHtml(Trim(ObjRsMig("Est_Sigla")))
	strProNome				= TratarAspasHtml(Trim(ObjRsMig("Pro_Nome")))
	strReferencia			= TratarAspasHtml(Trim(ObjRsMig("Ped_Referencia")))
	strAcf_NroAcessoPtaEbt  = TratarAspasHtml(Trim(ObjRsMig("Acf_NroAcessoPtaEbt")))
	DM						= TratarAspasHtml(Trim(ObjRsMig("DM")))
End if

%>


<SCRIPT LANGUAGE=javascript>
<!--
function ResgatarDistribuicao(obj,intDistrib)
{
	with (document.forms[0])
	{
		if (obj.value != "")
		{
			hdnAcao.value = "ResgatarDistribuicao"
			hdnEstacaoAtual.value = obj.value

			switch (intDistrib)
			{
				case 1:
					hdnNomeLocal.value = "DistLocalInstala"
					break
				case 2:
					hdnNomeLocal.value = "DistLocalConfig"
					break
				case 3:
					hdnNomeLocal.value = "DistLocalInstalaInter"
					break
			}
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()
		}	
	}
}

function ValidarRede()
{
	with (document.forms[0])
	{
	
		switch (parseInt(0+cboSistema.value))
		{
			case 1:	
				if (!ValidarCampos(txtFila,"Fila")) return false
				if (!ValidarCampos(txtBastidor,"Bastidor")) return false
				if (!ValidarCampos(txtRegua,"Régua")) return false
				if (!ValidarCampos(txtPosicao,"Posicao")) return false
				if (!ValidarCampos(txtTimeSlot,"TimeSlot")) return false
				if (!ValidarPar(txtTimeSlot,"Timeslot")) return false
				var objAryPar = txtTimeSlot.value.split("-")
				if (objAryPar.length > 1){
					if ((parseFloat(objAryPar[1])-parseFloat(objAryPar[0]))+1 > 100){alert("A quatidade máxima permitida são 100 timeslots.");return false}
				}
				if (!ValidarCampos(txtDominio,"Dominio")) return false
				if (!ValidarCampos(txtNO,"NO")) return false
				if (!ValidarCampos(txtSlot,"Slot")) return false
				if (!ValidarCampos(txtPorta,"Porta")) return false
				//if (!ValidarCampos(txtLink,"Link")) return false
				break
			case 2:
				if (!ValidarCampos(txtTronco,"Tronco")) return false
				if (!ValidarCampos(txtPar,"Par")) return false
				var objAryPar = txtPar.value.split("-")
				if (objAryPar.length > 0){
					if ((parseFloat(objAryPar[1])-parseFloat(objAryPar[0]))+1 > 100){alert("A quatidade máxima permitida são 100 pares.");return false}
				}
				break
			case 3:
				if (!ValidarCampos(txtTronco,"Cabo")) return false
				if (txtTronco.value.length != 2 ) {alert("Verifique o padrão do tipo do Cabo (N2)!");txtTronco.focus();return false}
				if (!ValidarCampos(txtPar,"Par")) return false
				var objAryPar = txtPar.value.split("-")
				if (objAryPar.length > 0){
					if ((parseFloat(objAryPar[1])-parseFloat(objAryPar[0]))+1 > 100){alert("A quatidade máxima permitida são 100 troncos.");return false}
				}
				if (!ValidarPar(txtPar,"Par")) return false
				//if (!ValidarCampos(cboTipoCabo,"Tipo do Cabo")) return false
				//if (!ValidarCampos(txtLateral,"Derivação")) return false
				//if (!ValidarCampos(txtCaixaEmenda,"PADE")) return false
				break
	
	}
	return true
  }
}


function VoltarMig()
{

	with (document.forms[0])
	{
		action = "ListarFacMig.asp?uf=<%=strUF%>&hdnRede=<%=intRede%>&hdnEstacao=<%=intEstacao%>&hdnDM=<%=DM%>&hdnFiltroDM=<%=FiltroDM%>"
		submit()
	}
}


function ValidarPar(obj,strCampo)
{
	if (obj.value.length < 4)
	{
		alert("Verifique o padrão para o campo " + strCampo + " (min.:N4, max.:N4-N4)!")
		obj.focus()
		return false
	}

	if (obj.value.length > 4 && obj.value.length != 9 )
	{
		alert("Verifique o padrão para o campo " + strCampo + " (min.:N4, max.:N4-N4)!")
		obj.focus()
		return false
	}

	return true

}

function Gravar(intAcao)
{
	with (document.forms[0])
	{
		
		if (!ValidarCampos(cboProvedor,"Provedor")) return
		if (!ValidarCampos(cboSistema,"Rede")) return
		if (!ValidarCampos(cboLocalInstala,"Local de Instalação")) return
		if (!ValidarCampos(cboDistLocalInstala,"Distribuidor do Local de Instalação")) return

		if (!ValidarRede())
		{ 
			return 
		}
		else
		{
			if (intAcao==0){
				hdnAcao.value = "GravarFacilidade"
				target = "IFrmProcesso"
				action = "ProcessoAlocaFacMig.asp"
				submit()
			}	
		}
	}
}

function ResgatarCamposFac(obj)
{
	obj="E";  // Forca a Facilidade de Entrada
	with (document.forms[0])
	{
		if (obj=="E")
		{
			divInstala.style.display = ''
			hdnSistema.value = cboSistema.value
		}
		else
		{
			try{
				if (spnCamposFac.innerHTML != "" && txtParInicial1 == '[object]') return
			}catch(e){}	
			hdnSistema.value = 4
			divInstala.style.display = 'none'  

		}
		hdnAcao.value = "ResgatarCamposFacilidade"
		target = "IFrmProcesso"
		action = "ProcessoAlocaFacMig.asp"
		submit()
	}
}

function ResgatarCamposFac2(obj)
{
	with (document.forms[0])
	{
		ResgatarCamposFac("E") // E - indica facilidade de entrada
	}	
}
function RetornaCboPlataforma(sisID, PlaID)
{
	if (sisID != 1) {
		spnPlataforma.innerHTML = ""
		return 
	}
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var strXML
	
	strXML = "<root>"
	strXML = strXML + "<plaid>" +  PlaID + "</plaid>"
	strXML = strXML + "<funcao>AtualizaFormtato(document.forms[0].cboPlataforma(document.forms[0].cboPlataforma.selectedIndex).tipoPla)</funcao>"
	strXML = strXML + "</root>" 
	
	xmlDoc.loadXML(strXML);
	xmlhttp.Open("POST","RetornaPlataforma.asp" , false);
	xmlhttp.Send(xmlDoc.xml);
	
	
	strXML = xmlhttp.responseText;
	spnPlataforma.innerHTML = strXML

	document.forms[0].hdnPlataforma.value = document.forms[0].cboPlataforma(document.forms[0].cboPlataforma.selectedIndex).tipoPla
	
}

function AtualizaFormtato(strTipoPlataforma){

	document.forms[0].hdnPlataforma.value = strTipoPlataforma

	if (strTipoPlataforma == "LUCENT"){
			spnDominio.innerHTML = "<input type=text class=text name=txtDominio		maxlength=10  size=8 >(A10)"
			spnNo.innerHTML = "<input type=text class=text name=txtNO			maxlength=2   size=2 onKeyUp='ValidarTipo(this,0)' >(N2)"
			spnSlot.innerHTML = "<input type=text class=text name=txtSlot		maxlength=2   size=2 onBlur= SearchDom(this,\'01,02,03,04,05,06,07,08,09,10,11,12\') onKeyUp='ValidarTipo(this,2)'>(N2)"
			spnPorta.innerHTML = "<input type=text class=text name=txtPorta		maxlength=2   size=2 onBlur= SearchDom(this,\'01,02,03,04,05,06,07,08,09,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63\') onKeyUp='ValidarTipo(this,2)'>(N2)" 
	}
	else {
			spnDominio.innerHTML = "<input type=text class=text name=txtDominio		maxlength=3  size=4 onKeyUp='ValidarTipo(this,0)' >(N3)"
			spnNo.innerHTML = "<input type=text class=text name=txtNO				maxlength=3  size=3 onKeyUp='ValidarTipo(this,0)'>(N3) "
			spnSlot.innerHTML = "<input type=text class=text name=txtSlot		maxlength=2   size=2 onBlur=SearchDom(this,\'A1,A2,A3,A4,A5,A6,A7,A8,B1,B2,B3,B4,B5,B6,B7,B8\') onKeyUp='ValidarTipo(this,2)'>(A2)"
			spnPorta.innerHTML = "<input type=text class=text name=txtPorta		maxlength=1   size=1 onBlur=SearchDom(this,\'A,B\')>(A1)</span>"
	}
		
}
//-->
</SCRIPT>

<!--#include file="../inc/headermig.asp"-->

<form name="Form1" method="post" action="">
<input type=hidden name=hdnIdFis value="<%=intIdFis%>" >
<input type=hidden name=hdnIdLog value="<%=intIdLog%>" >

<input type=hidden name=hdnAcao>
<input type="hidden" name="hdnEstacaoAtual" >
<input type="hidden" name="hdnNomeLocal" >
<input type="hidden" name="hdnSistema">
<input type="hidden" name="hdnPlataforma">

<input type=hidden name=hdnUF value=<%=strUF%>>
<input type=hidden name=hdnRede value=<%=intRede%>>
<input type=hidden name=hdnEstacao value=<%=intEstacao%>>
<input type=hidden name=hdnDM value=<%=DM%>>
<input type=hidden name=hdnFiltroDM value="<%=FiltroDM%>">

<tr><td>
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=10><p align="center">Correção de Facilidade</p></th>
</tr>

<!-- Exibe o endereço migrado, juntamente com os dados do cliente e acesso -->
<tr class=clsSilver>
<td>:. Cliente</td>
<td><input type="text" class="text" name="txtNomeClienteMig"  maxlength="40" size="26" readonly value="<%=strCli_Nome%>"></td>
<!--/tr>

</tr>
<tr class=clsSilver-->
<td>:. Designação do Serviço</td>
<td><input type="text" class="text" name="txtDesignacaoServicoMig" maxlength="40" size="26" readonly value="<%=strDesignacaoServico%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>:. Designação do Acesso Lógico</td>
<td><input type="text" class="text" name="txtIdLogicoMigradoMig" maxlength="40" size="26" readonly value="<%=strAcl_IDAcessoLogico%>"></td>
<!--/tr>

</tr>
<tr class=clsSilver-->
<td>:. Designação do Acesso Físico</td>
<td><input type="text" class="text" name="txtIdFisicoMigradoMig" maxlength="40" size="26" readonly value="<%=strAcf_IDAcessoFisico%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>:. Estação</td>
<td><input type="text" class="text" name="txtEstacaoMig" maxlength="40" size="26" readonly value="<%=strEstacao%>"></td>
<!--/tr>

</tr>
<tr class=clsSilver-->
<td>:. Provedor</td>
<td><input type="text" class="text" name="txtProvedorMig" maxlength="40" size="26" readonly value="<%=strProNome%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>:. Distribuidor</td>
<td><input type="text" class="text" name="txtDistribuidorMig" maxlength="40" size="26" readonly value="<%=strDistribuidor%>"></td>
<!--/tr>

</tr>
<tr class=clsSilver-->
<td>:. Sistema</td>
<td><input type="text" class="text" name="txtSistemaMig" maxlength="40" size="26" readonly value="<%=strSistema%>"></td>
</tr>

<% if strSistema = "Nao Deterministico" then %>
</tr>
<tr class=clsSilver>
<td>:. Tronco</td>
<td><input type="text" class="text" name="txtTroncoMig" maxlength="20" size="26" readonly value="<%=strTronco%>"></td>
<!--/tr>

</tr>
<tr class=clsSilver-->
<td>:. Par</td>
<td><input type="text" class="text" name="txtParMig" maxlength="20" size="26" readonly value="<%=strPar%>"></td>
</tr>

<tr class=clsSilver>
<td>:. Referência</td>
<td><input type="text" class="text" name="txtReferenciaMig" maxlength="40" size="26" readonly value="<%=strReferencia%>"></td>

<td>:. Desig. Provedor</td>
<td><input type="text" class="text" name="txtDesignacaoProMig" maxlength="40" size="26" readonly value="<%=strAcf_NroAcessoPtaEbt%>"></td>
</tr>


<% else %>
</tr>
<tr class=clsSilver>
<td>:. Domínio</td>
<td><input type="text" class="text" name="txtDominioMig" maxlength="20" size="26" readonly value="<%=strDominio%>"></td>
<!--/tr>

</tr>
<tr class=clsSilver-->
<td>:. Nó</td>
<td><input type="text" class="text" name="txtNoMig" maxlength="20" size="26" readonly value="<%=strNo%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>:. Slot</td>
<td><input type="text" class="text" name="txtSlotMig" maxlength="20" size="26" readonly value="<%=strSlot%>"></td>
<!--/tr>

</tr>
<tr class=clsSilver-->
<td>:. Porta</td>
<td><input type="text" class="text" name="txtPortaMig" maxlength="20" size="26" readonly value="<%=strPorta%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>:. Time Slot</td>
<td><input type="text" class="text" name="txtTimeSlotMig" maxlength="20" size="26" readonly value="<%=strTime_Slot%>"></td>
<!--/tr>

</tr>
<tr class=clsSilver-->
<td>:. Fila</td>
<td><input type="text" class="text" name="txtFilaMig" maxlength="20" size="26" readonly value="<%=strFila%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>:. Bastidor</td>
<td><input type="text" class="text" name="txtBastidorMig" maxlength="20" size="26" readonly value="<%=strBastidor%>"></td>
<!--/tr>

</tr>
<tr class=clsSilver-->
<td>:. Regua</td>
<td><input type="text" class="text" name="txtReguaMig" maxlength="20" size="26" readonly value="<%=strRegua%>"></td>
</tr>

<tr class=clsSilver>
<td>:. Posicao</td>
<td><input type="text" class="text" name="txtPosicaoMig" maxlength="20" size="26" readonly value="<%=strPosicao%>"></td>
<!--/tr>

</tr>
<tr class=clsSilver-->
<td>:. Referência</td>
<td><input type="text" class="text" name="txtReferenciaMig" maxlength="40" size="26" readonly value="<%=strReferencia%>"></td>
</tr>

<tr class=clsSilver>
<td>:. Desig. Provedor</td>
<td colspan=3><input type="text" class="text" name="txtDesignacaoProMig" maxlength="40" size="26" readonly value="<%=strAcf_NroAcessoPtaEbt%>"></td>
</tr>
<%end if%>

<p>
</table>

<div id=divInstala style="DISPLAY: ''">
<table border=0 cellspacing="1" cellpadding="0" width="760">

<tr>
	<th colspan=2 ><p align="center">Nova Facilidade</p></td>
</tr>
<tr>
	<th colspan=2 >&nbsp;•&nbsp;Recurso</td>
</tr>

<tr class=clsSilver2>
	<td width=150px ><font class="clsObrig">:: </font>Provedor</td>
	<td >
		<select name="cboProvedor" >
			<option value=""></option>
			<%	set objRS = db.execute("CLA_sp_sel_provedor 0")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strProId) = Trim(objRS("Pro_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "'" & strItemSel & ">" & objRS("Pro_Nome") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>

<tr class=clsSilver2>
	<td><font class="clsObrig">:: </font>Rede</td>
	<td >
		<select name="cboSistema" onChange="ResgatarCamposFac2(this);RetornaCboPlataforma(cboSistema.value,0)" >
			<option value=""></option>
			<%
				Set objRS = db.execute("CLA_sp_sel_sistema 0")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strRede) = Trim(objRS("Sis_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value='" & objRS("Sis_ID")& "'" & strItemSel & ">" & Ucase(objRS("Sis_Desc")) & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""

			%>
		</select>
		<span ID = spnPlataforma>
		</span>
	</td>
</tr>
<tr class=clsSilver2>
	<td width=150px nowrap><font class="clsObrig">:: </font>Local de Instalação</td>
	<td  >
		<select name="cboLocalInstala" onChange="ResgatarDistribuicao(this,1)">
			<option value=""></option>
			<%set objRS = db.execute("CLA_sp_sel_estacao_mig " & strUF)
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strLocalInstala) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value=" & objRS("Esc_ID") & strItemSel & ">" & objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>

<tr class=clsSilver2>
	<td><font class="clsObrig">:: </font>Distribuidor</td>
	<td >
		<span id="spnDistLocalInstala">
		<select name="cboDistLocalInstala" >
			<option value=""></option>
			<%	if Trim(strLocalInstala) <> "" then 
					set objRS = db.execute("CLA_sp_view_recursodistribuicao " & strLocalInstala)
					While not objRS.Eof 
						strItemSel = ""
						if Trim(strDistrib) = Trim(objRS("Dst_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value=" & objRS("Dst_ID") & strItemSel & ">" & objRS("Dst_Desc") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				End if	
			%>
		</select>
		</span>
	</td>
</tr>
</table>
</div>

<table width="760" border=0>
<tr>
	<td >
		<span id=spnCamposFac></span>
	</td>
</tr>
</table>


<table cellspacing="1" cellpadding="0" width="760" border=0>
<tr >
	<td colspan=2 align="center"><br>
		<input type="button" class="button" name="btnGravar" value="Gravar" onclick="Gravar(0)" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
		<input type="button" class="button" name="Voltar" value="Voltar" onclick="VoltarMig()" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
	</td>
</tr>
</table>
<table width="760" border=0 cellspacing="1" cellpadding="0">
<tr>
	<td>
	<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
	<!--input type="text" class="text" name="txtSql" maxlength="20" size="26" readonly value="<%=strRegua%>"-->
	</td>
</tr>
<tr>
	<td>
		<font class="clsObrig">:: </font>Legenda: A - Alfanumérico;  N - Numérico;  L - Letra
	</td>
</tr>
</table>

<table width="760" border=0 cellspacing="1" cellpadding="0">
<iframe	id			= "IFrmProcesso"
		name        = "IFrmProcesso" 
		width       = "100%" 
		height      = "100%"
		frameborder = "0"
		scrolling   = "no" 
		align       = "left">
</iFrame>
</table>
</form>
</body>
<SCRIPT LANGUAGE=javascript>
setarFocus('cboProvedor');
ResgatarCamposFac2(document.forms[0].cboSistema);
RetornaCboPlataforma(document.forms[0].cboSistema.value,0);

</Script>
</html>






