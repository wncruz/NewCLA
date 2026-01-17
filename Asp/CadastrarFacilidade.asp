<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: CadastrarFacilidade.asp
'	- Descrição			: Cadastrar Facilidade
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
Dim strItemSel
Dim strProId
Dim strRede
Dim strLocalInstala
Dim strDistrib
Dim strLocalConfig
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


function ResgatarTipoSwitch() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboTipoSwitch').value
		//rede = document.getElementById('cboTipoSwitch').value
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoTipoSwitch;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_TipoSwitch.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		//strXML = "<dados><param>"+param+"</param><param1>14</rede></dados>"
		strXML = "<dados><param>"+param+"</param></dados>"
        xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoTipoSwitch () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spncboTipoSwitch").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}

function EsconderTipoSwitch()
{
	document.getElementById("spncboTipoSwitch").innerHTML = "";
}


function ResgatarTipoDSLAM() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboTipoDSLAM').value
		
		//alert (param)
		//rede = document.getElementById('cboTipoSwitch').value
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoTipoDSLAM;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_TipoDSLAM.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		//strXML = "<dados><param>"+param+"</param><param1>14</rede></dados>"
		strXML = "<dados><param>"+param+"</param></dados>"
        xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoTipoDSLAM () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spncboTipoDSLAM").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}

function EsconderTipoDSLAM()
{
	document.getElementById("spncboTipoDSLAM").innerHTML = "";
}



function ResgatarTipoONT() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboFabricanteONT').value
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoTipoONT;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_TipoONT.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><texto>S</texto></dados>"
        xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoTipoONT () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spncboTipoONT").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}

function ValidarItemOE(campo)
  {
    if (campo.value == "0")
    {
	  campo.value = "001"
	}							    
}

function EsconderFabONT()
{
	document.getElementById("spncboFabricante").innerHTML = "";
}

function EsconderTipoONT()
{
	document.getElementById("spncboTipoONT").innerHTML = "";
}

function ValidarRede()
{
	with (document.forms[0])
	{
	
	if (rdoInter[0].checked)
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
				
			case 12:	
				/**
				Radio IP
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
				**/
				break
		}
	}
	else
	{
		if (!ValidarCampos(txtQtdePares,"Quantidade de Pares")) return false
		if (parseFloat(txtQtdePares.value) > 100){alert("A quatidade máxima permitida são 100 pares.");return false}
		if (!ValidarCampos(txtParInicial1,"Par Inicial 1")) return false
		if (!ValidarCampos(txtCoordenada1,"Coordenada 1")) return false
		if (!ValidarCampos(cboLocalInstalaInter,"Local de Instalação")) return false
		if (!ValidarCampos(cboDistLocalInstalaInter,"Distribuidor do Local de Instalação")) return false
		if (!ValidarCampos(txtParInicial2,"Par Inicial 2")) return false
		if (!ValidarCampos(txtCoordenada2,"Coordenada 2")) return false
		if (!ValidarCampos(cboLocalConfig,"Local de Configuração")) return false
		if (!ValidarCampos(cboDistLocalConfig,"Distribuidor do Local de Configuração")) return false
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

function JanelaConfirmacao(objXmlGeral){
	var intRet = window.showModalDialog('ConfirmacaoCadFac.asp',objXmlGeral,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	if (intRet == 1){
		intRet = 0
		Gravar(1) 
	}
} 

function Gravar(intAcao)
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
			if (!ValidarCampos(cboLocalInstala,"Local de Instalação")) return
			if (!ValidarCampos(cboDistLocalInstala,"Distribuidor do Local de Instalação")) return
		}

		if (!ValidarRede())
		{ 
			return 
		}
		else
		{
			
			if ( cboSistema.value == "12" || cboSistema.value == "14" || cboSistema.value == "15" || cboSistema.value == "16") {
				//alert (cboSistema.value)
				//return
				if(cboSistema.value == "12"){
					
						if (!ValidarCampos(cboFabricanteONT,"Fabricante")) return
						if (!ValidarCampos(cboTipoONT,"Modelo")) return
						if (!ValidarCampos(txtPortaPE,"Porta PE")) return
						if (!ValidarCampos(txtPE,"PE")) return
				
				}
				
				if(cboSistema.value == "14"){
					
						if (!ValidarCampos(cboTipoSwitch,"Tipo do Equipamento")) return
						if (!ValidarCampos(cboPropriedade,"Propriedade do Equipamento")) return
						if (!ValidarCampos(txtDesignacao,"Designação")) return
						if (!ValidarCampos(cboFabricanteONT,"Fabricante")) return
						if (!ValidarCampos(cboTipoONT,"Modelo")) return
				
				}
				
				if(cboSistema.value == "15"){
					
						if (!ValidarCampos(cboTipoSwitch,"Tipo do Equipamento")) return
						if (!ValidarCampos(cboPropriedade,"Propriedade do Equipamento")) return
						if (!ValidarCampos(txtDesignacao,"Designação")) return
						if (!ValidarCampos(cboFabricanteONT,"Fabricante")) return
						if (!ValidarCampos(cboTipoONT,"Modelo")) return
				
				}
				
				if(cboSistema.value == "16"){
					
						if (!ValidarCampos(cboTipoDSLAM,"Tipo do Equipamento")) return
						if (!ValidarCampos(cboPropriedade,"Propriedade do Equipamento")) return
						if (!ValidarCampos(txtDesignacao,"Designação")) return
						if (!ValidarCampos(cboFabricanteONT,"Fabricante")) return
						if (!ValidarCampos(cboTipoONT,"Modelo")) return
				
				}
				
				hdnAcao.value = "GravarFacilidade"
				target = "IFrmProcesso"
				action = "ProcessoCadFac.asp"
				submit()
			}
			else{
			if (intAcao==0){
				hdnAcao.value = "ConfirmarFacilidade"
				target = "IFrmProcesso"
				action = "ProcessoCadFac.asp"
				submit()
			}else{
				hdnAcao.value = "GravarFacilidade"
				target = "IFrmProcesso"
				action = "ProcessoCadFac.asp"
				submit()
			}	
		}
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
		hdnAcao.value = "ResgatarCamposFacilidade"
		target = "IFrmProcesso"
		action = "ProcessoCadFac.asp"
		submit()
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
	if(strTipoPlataforma == "E1"){
			spnDominio.innerHTML = "<input type=text class=text name=txtDominio		maxlength=3  size=4 onKeyUp='ValidarTipo(this,0)' >(N3)"
			spnNo.innerHTML = "<input type=text class=text name=txtNO				maxlength=3  size=3 onKeyUp='ValidarTipo(this,0)'>(N3) "
			spnSlot.innerHTML = "<input type=text class=text name=txtSlot		maxlength=2   size=2 onBlur=SearchDom(this,\'A1,A2,A3,A4,A5,A6,A7,A8,B1,B2,B3,B4,B5,B6,B7,B8\') onKeyUp='ValidarTipo(this,2)'>(A2)"
			spnPorta.innerHTML = "<input type=text class=text name=txtPorta		maxlength=1   size=1 onBlur=SearchDom(this,\'A,B\')>(A1)</span>"
	}
	if(strTipoPlataforma == "DATACOM") {
			spnDominio.innerHTML = "<input type=text class=text name=txtDominio		maxlength=12  size=14 >(A12)"
			spnNo.innerHTML = "<input type=text class=text name=txtNO				maxlength=3  size=3 onKeyUp='ValidarTipo(this,0)'>(N3) "
			spnSlot.innerHTML = "<input type=text class=text name=txtSlot		maxlength=2   size=2 onBlur=SearchDom(this,\'A1,A2,A3,A4,A5,A6,A7,A8,B1,B2,B3,B4,B5,B6,B7,B8\') onKeyUp='ValidarTipo(this,2)'>(A2)"
			spnPorta.innerHTML = "<input type=text class=text name=txtPorta		maxlength=1   size=1 onBlur=SearchDom(this,\'A,B\')>(A1)</span>"
	}
	if(strTipoPlataforma == "IP") {
			spnDominio.innerHTML = "<input type=text class=text name=txtDominio		maxlength=12  size=14 >(A12)"
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
<input type=hidden name=hdnUsuario>
<input type=hidden name=hdnOrigemEst>
<input type=hidden name=hdnCNLEstUsu>
<input type=hidden name=hdnComplEstUsu>
<input type="hidden" name="hdnAcao">
<input type="hidden" name="hdnEstacaoAtual" >
<input type="hidden" name="hdnNomeLocal" >
<input type="hidden" name="hdnSistema">
<input type="hidden" name="hdnPlataforma">
<tr>
<td >
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2 ><p align="center">Cadastro de Facilidade</p></td>
</tr>
<tr>
	<th colspan=2 >&nbsp;•&nbsp;Recurso</td>
</tr>
<tr class=clsSilver>
	<td width=150px ><font class="clsObrig">:: </font>Tipo de Facilidade</td>
	<td>
		<input type=radio value=E name=rdoInter onClick="ResgatarCamposFac(this)">&nbsp;Entrada&nbsp;&nbsp;
		<input type=radio value=I name=rdoInter onClick="ResgatarCamposFac(this)">&nbsp;Interligação
	</td>
</tr>
</table>
<div id=divInstala style="DISPLAY: ''">
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr class=clsSilver>
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

<tr class=clsSilver>
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
<tr class=clsSilver>
	<td width=150px nowrap><font class="clsObrig">:: </font>Local de Instalação</td>
	<td>
		<input type="Hidden" name="cboLocalInstala">
		<input type="text" class="text" name="txtCNLLocalInstala"  maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		&nbsp;<input type="text" class="text" name="txtComplLocalInstala"  maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoInstala(document.Form1.txtCNLLocalInstala,document.Form1.txtComplLocalInstala,<%=dblUsuId%>,1);" TIPO="A">
	Ex.: SPO IG</td>
</tr>

<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Distribuidor</td>
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
</td>
</tr>
<tr>

<table width="760">
<tr>
	<td >
		<span id=spnCamposFac></span>
	</td>
</tr>
</table>

<table width="760">
<tr>
	<td colspan=2 align="center">
		<input type="button" class="button" name="btnGravar" value="Gravar" onclick="Gravar(0)" accesskey="G" onmouseover="showtip(this,event,'Incluir (Alt+G)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparForm();setarFocus('cboProvedor');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
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
		width       = "100%" 
		height      = "100%"
		frameborder = "0"
		scrolling   = "no" 
		align       = "left">
</iFrame>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('cboProvedor')
//-->
</SCRIPT>

</html>
<%DesconectarCla()%>
