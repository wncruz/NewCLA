<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsultarFacilidadeRec.ASP
'	- Descrição			: Consulta de Facilidades
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function DetalharFacilidade(intFacId){
	var objNode = objXmlGeral.selectNodes("//Facilidade[@Fac_Id="+intFacId+"]")
	if (objNode.length>0){
		var intRet = window.showModalDialog('MessageConsFac.asp',objNode,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	}
}

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

function ResgatarCamposFac(obj)
{
	with (document.forms[0])
	{
		if (obj.value == "E")
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
		spnPosicoes.innerHTML = "" 
		hdnAcao.value = "ResgatarCamposFacilidade"
		target = "IFrmProcesso"
		action = "ProcessoConsFac.asp"
		submit()
	}
}


function ValidarRede()
{
   var blnHoriz = false
   var blnVert = false
   with (document.forms[0])
   {
	 switch (parseInt(0+cboSistema.value))
	 {
		case 1:	
			if (!ValidarCampos(txtBastidor,"Bastidor")) return false
			break

		case 2:
			if (rdoInter[0].checked)
			{
				if (!ValidarCampos(txtTronco,"Tronco")) return false
			}
			if (rdoInter[1].checked)
			{
				if (!ValidarCampos(txtCoordenada1,"PADE/PAC")) return false
			}
			break
		case 3:
			if (rdoInter[0].checked)
			{
				if (!ValidarCampos(txtTronco,"Cabo")) return false
				if (txtTronco.value.length != 2 ) {alert("Verifique o padrão do tipo do Cabo (N2)!");txtTronco.focus();return false}
			}	

			if (rdoInter[1].checked)
			{
				if (!ValidarCampos(txtCoordenada1,"PADE/PAC")) return false
			}
			break
	}	

	return true
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

function Procurar()
{
	with (document.forms[0])
	{

		if (!rdoInter[0].checked && !rdoInter[1].checked)
		{
			alert("Selecione o Tipo de Facilidade.")
			return
		}
		
		if (rdoInter[0].checked)
		{
			if (!ValidarCampos(cboProvedor,"Provedor")) return
			if (!ValidarCampos(cboSistema,"Rede")) return
		}else{
			if (!ValidarCampos(txtCoordenada1,"Coordenada")) return
		}	
		if (!ValidarCampos(txtQtdeRegistros,"Quantidade de registros")) return
		//if (parseInt(txtQtdeRegistros.value) > 100 )
		//{
		//	alert("A quantidade máxima permitida por consulta são 100 registros!")
		//	txtQtdeRegistros.value = 100
		//	return
		//}

		if (rdoStatusFac[1].checked || rdoStatusFac[2].checked)
		{
			hdnAcao.value = "ResgatarPosicoesOcupadasRet"
		}
		else
		{
			hdnAcao.value = "ResgatarPosicoesLivresRet"
		}	
		hdnRede.value = cboSistema.value
		target = "IFrmProcesso"
		action = "ProcessoConsFac.asp"
		submit()
	}
}
function DetalharSolicitacao(dblSolId)
{
	with (document.forms[0])
	{
		//Enviar para pagina de Detalhamento do pedido
		hdnSolId.value = dblSolId
		DetalharFac()
	}	
}

function ResgatarCamposFac2(obj)
{
	with (document.forms[0])
	{
		if (rdoInter[0].checked) ResgatarCamposFac(document.forms[0].rdoInter[0])
		if (rdoInter[1].checked) ResgatarCamposFac(document.forms[0].rdoInter[1])
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

function CheckEstacaoInstala(objCNL,objCompl,usu,origemEst)
{
	with (document.forms[0])
	{
	
		if (objCNL.value != "" && objCompl.value != "")
		{
			hdnCNLEstUsu.value = objCNL.value
			hdnComplEstUsu.value = objCompl.value
			hdnOrigemEst.value = origemEst
			hdnUsuario.value = usu
			hdnAcao.value = "CheckEstacaoInstala"
			target = "IFrmProcesso"
			action = "ProcessoSolic.asp"
			submit()
		}
	}
}

//-->
</SCRIPT>
<form method="post" name="Form1">
<input type="hidden" name="hdnAcao">
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type="hidden" name="hdnEstacaoAtual" >
<input type="hidden" name="hdnNomeLocal" >
<input type="hidden" name="hdnRede" >
<input type="hidden" name="txtNroReg" value=50 >
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnSistema>
<input type=hidden name=hdnNomeCons value="ConsFac">
<input type=hidden name=hdnPlataforma value="ConsFac">
<tr>
<td >
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2 ><p align="center">Consulta de Facilidades</p></td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Status da Facilidade</td>
	<td>
		<input type=radio value=0 name=rdoStatusFac checked>Livres&nbsp;
		<input type=radio value=1 name=rdoStatusFac>Ocupadas&nbsp;
		<input type=radio value=2 name=rdoStatusFac>Todos
	</td>
</tr>
<tr>
	<th colspan=2 >&nbsp;•&nbsp;Recurso</td>
</tr>
<tr class=clsSilver>
	<td width=170 ><font class="clsObrig">:: </font>Tipo de Facilidade</td>
	<td>
		<input type=radio value=E name=rdoInter onClick="ResgatarCamposFac(this)">&nbsp;Entrada&nbsp;&nbsp;
		<input type=radio value=I name=rdoInter onClick="ResgatarCamposFac(this)">&nbsp;Interligação
	</td>
</tr>
</table>
<div id=divInstala style="DISPLAY: ''">
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr class=clsSilver>
	<td width=170 ><font class="clsObrig">:: </font>Provedor</td>
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
<tr class=clsSilver>
	<td width=170><font class="clsObrig">:: </font>Rede</td>
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
<tr class=clsSilver>
	<td width=170 nowrap>&nbsp;&nbsp;&nbsp;Estação</td>
	<td>
		<input type="Hidden" name="cboLocalInstala">
		<input type="text" class="text" name="txtCNLLocalInstala"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		&nbsp;<input type="text" class="text" name="txtComplLocalInstala"  maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoInstala(document.Form1.txtCNLLocalInstala,document.Form1.txtComplLocalInstala,<%=dblUsuId%>,1);" TIPO="A">
	Ex.: SPO IG</td>
</tr>
<tr class=clsSilver>
	<td width=170>&nbsp;&nbsp;&nbsp;Distribuidor</td>
	<td >
		<span id="spnDistLocalInstala">
		<select name="cboDistLocalInstala" >
			<option value=""></option>
			<%	if Trim(strLocalConfig) <> "" then 
					set objRS = db.execute("CLA_sp_view_recursodistribuicao " & strLocalConfig)
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
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr class=clsSilver>
	<td width=170><font class="clsObrig">:: </font>Quantidade de Registros</td>
	<td ><input type=text name=txtQtdeRegistros onkeyup="ValidarTipo(this,0)" value=50 size=5 maxlength=4 class=text></td>
</tr>
</table>
<table width="760" cellspacing="0" cellpadding="0">
<tr>
	<td>
		<span id=spnCamposFac></span>
	</td>
</tr>
</table>
<table width="760">
	<tr>
		<td colspan=2 align="center">
		<input type="button" class="button" name="btnGravar" value="Procurar" onclick="Procurar()" >&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparForm()">&nbsp;
		<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.location.replace('main.asp')">
		</td>
	</tr>
</table>
<table width="760">
<tr>
	<td align="center" width=100%>
		<span id=spnPosicoes></span>
	</td>
</tr>
</table>
<table width="760">
<tr>
	<td>
	<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
	</td>
</tr>
<tr>
	<td>
		<font class="clsObrig">:: </font>Legenda: A - Alfanumérico;  N - Numérico;  L - Letra
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
</body>
</html>