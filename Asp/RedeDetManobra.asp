<%
'•Accenture
'	- Sistema			: CLA
'	- Arquivo			: RedeDetManobra.asp
'	- Descrição			: Tela copia da RedeDet.asp só alterando o necessario para a Manobra

Dim nomeEstacao
Dim idEstacao

if Request.Form("hdnProvedor2") = "" then 
	strPlaID = Request.Form("hdnPlataforma")
else
	strPlaID = Request.Form("hdnProvedor2")
end if 

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXmlReturn") <> "" then
	objXmlDados.loadXml(Request.Form("hdnXmlReturn"))
Else
	objXmlDados.loadXml("<xDados/>")
End if
%>
<!--#include file="../inc/data.asp"-->
<html>
<head>
<title>CLA - Controle Local de Acesso</title>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlReturn = window.dialogArguments
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")




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
		if (hdnProvedor2.value != "")
		{
			if (!ValidarCampos(hdnProvedor2,"Provedor")) return
			if (!ValidarCampos(hdnLocalInstala,"Estação")) return

			hdnAcao.value = "ResgatarTimeSlots"
			target = "IFrmProcesso"
			action = "ProcessoConsRedeDetManobra.asp"
			submit()
		}	
	}
}

function ResgatarDominioNO(obj)
{
	with (document.forms[0])
	{
		if (obj.value != "")
		{
			if (hdnProvedor2.value == "" || hdnLocalInstala.value == "")
			{
				alert("Provedor/Estação são obrigatórios para resgatar Domínio-NO.")
				hdnLocalInstala.value = ""
				return
			}		  
			hdnAcao.value = "ResgatarDominioNO"
			target = "IFrmProcesso"
			action = "ProcessoConsRedeDetManobra.asp"
			submit()
		}	
	}
}

function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		hdnSolId.value = dblSolId
		var strNome = "Facilidade" + dblSolId 
		var objJanela = window.open("about:blank",null,"status=no,toolbar=no,enubar=no,location=no,scrollbars = Yes,resizable=Yes")
		//null, null, "status=no,toolbar=no,menubar=no,location=no,resizable=Yes,scrollbars = Yes"
		//var intRet = window.showModalDialog('facilidadeDet.asp', dblSolId,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
		objJanela.name = strNome
		target = strNome
		action = "facilidadeDet.asp?RedeDet=1"
		submit()
		//DetalharFac()
		/*target = window.top.name
		action = "ConsultaGeralDet.asp"
		submit() */
	}	
}

function DetalharFacilidade(intFacId){
	if (intFacId != ""){
		var objNode = objXmlGeral.selectNodes("//Facilidade[@Fac_Id="+intFacId+"]")
	}else{
		var objNode = objXmlGeral.selectNodes("//Facilidade")
	}	
	if (objNode.length>0){
		var intRet = window.showModalDialog('MessageConsFac.asp',objNode,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
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
	strXML = strXML + "<funcao></funcao>"
	strXML = strXML + "</root>" 
	
	xmlDoc.loadXML(strXML);
	xmlhttp.Open("POST","RetornaPlataforma.asp" , false);
	xmlhttp.Send(xmlDoc.xml);
	
	strXML = xmlhttp.responseText;
	parent.spnPlataforma.innerHTML = strXML
	
}


//-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function window_onunload(){
	PopularXml(objXmlReturn)
	window.returnValue = objXmlReturn
}
//-->
</SCRIPT>
</head>
<body class=TA LANGUAGE=javascript onunload="return window_onunload()">
<form method="post" name="Form1">
<input type="hidden" name="hdnAcao">
<input type="hidden" name="hdnSolId">
<input type="hidden" name="hdnPedId">
<input type="hidden" name="hdnFacId">
<input type="hidden" name="hdnPlataforma" = "<%= strPlaID %>">

<input type="hidden" name="txtRedDetTimeslot">
<input type="hidden" name="txtRedDetBastidor">
<input type="hidden" name="txtRedDetRegua">
<input type="hidden" name="txtRedDetPosicao">
<input type="hidden" name="txtRedDetFila">
<input type="hidden" name="txtRedDetEstacao">
<input type="hidden" name="txtRedDetDistribuidor">
<input type="hidden" name="txtRedDetProvedor">
<input type="hidden" name="txtRedDetPlataforma">

<input type="hidden" name="hdnFacDetid" >
<input type="hidden" name="hdnEstacaoAtual" >
<input type="hidden" name="hdnNomeLocal" >
<input type="hidden" name="hdnRede" >
<input type="hidden" name="hdnXmlReturn">
<input type="hidden" name="hdnJSReturn">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<tr>
<td >
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2 ><p align="center">Controle de Rede Determinística</p></td>
</tr>

<tr class=clsSilver>
	<td width=150px ><font class="clsObrig">:: </font>Provedor</td>
	<td>
		<input type="text" name="txtProvedor" readonly maxlength=30 size=30  value="" class="text"></td>
		<td><input type="hidden" name="hdnProvedor2" value="" class="text"></td>
	</td>
</tr>
	
<tr class=clsSilver>
	<td width=150px nowrap><font class="clsObrig">:: </font>Estação</td>
	<td>
		<input type="text" readonly name="txtLocalEntrega" value="" class="text">
		<input type="hidden" name="hdnLocalInstala" value="">
	</td>
</tr>
<tr class=clsSilver>
<!-- @@ LPEREZ - 03/04/2006	-->
	<td><font class="clsObrig">:: </font>Status da Facilidade</td>
	<td>
		<input type=radio value=0 name=rdoStatusFac checked 
		onClick="ResgatarDominioNO('stsFac')">Livres&nbsp;
		<input type=radio value=1 name=rdoStatusFac 
		onClick="ResgatarDominioNO('stsFac')">Ocupadas&nbsp;
		<input type=radio value=2 name=rdoStatusFac 
		onClick="ResgatarDominioNO('stsFac')">Todos
	</td>
<!-- @@ LP -->
<!--	
	<td><font class="clsObrig">:: </font>Status da Facilidade</td>
	<td>
		<input type=radio value=0 name=rdoStatusFac checked>Livres&nbsp;
		<input type=radio value=1 name=rdoStatusFac>Ocupadas&nbsp;
		<input type=radio value=2 name=rdoStatusFac>Todos
	</td>
-->	
</tr>
<tr class=clsSilver>
	<td width=150px nowrap>&nbsp;&nbsp;&nbsp;&nbsp;Domínio - NO</td>
	<td>
		<span id=spnDominioNO>
			<select name="cboDominioNO">
				<option value=""></option>
			<%	if Trim(strProId) <> "" and  Trim(strEstacao) <> "" then
					strDominioNO = Request.Form("cboDominioNO")
					if strDominioNO = "" then
						set objNode = objXmlDados.getElementsByTagName("cboDominioNO")
						if objNode.length > 0 then
							strDominioNO = objNode(0).childNodes(0).text
						End if
					End if	

					set objRS = db.execute("CLA_sp_sel_facilidade_entrada_Agrupado " & strProId & "," & strEstacao)
					While not objRS.Eof 
						strItemSel = ""
						if Trim(strDominioNO) = Trim(objRS("Fac_Dominio")) & "•" & Trim(objRS("Fac_NO")) then strItemSel = " Selected " End if
						Response.Write "<Option value=""" & Trim(objRS("Fac_Dominio")) & "•" & Trim(objRS("Fac_NO")) & """ " & strItemSel & ">" & Trim(objRS("Fac_Dominio")) & " - " & Trim(objRS("Fac_NO")) & "</Option>"
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
<table width="760">
	<tr>
		<td colspan=2 align="center">
		<input type="button" class="button" name="btnGravar" value="Procurar" onclick="Procurar()" >&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparForm()">&nbsp;
		<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.close()">
		</td>
	</tr>
</table>
<table width="760" cellspacing="0" cellpadding="0">
<tr>
	<td align="center" width=100%>
		<iframe	id			= "IFrmProcesso"
			    name        = "IFrmProcesso" 
			    width       = "100%" 
			    height      = "170px"
			    frameborder = "0"
			    scrolling   = "overflow" 
			    align       = "left">
		</iFrame>
	</td>
</tr>
</table>
<table cellspacing=1 width=760 cellpadding=0 border=0>
	<tr>
		<td colspan=2 class=clsSilver2>&nbsp;•&nbsp;Legenda
		</td>
	</tr>
	<tr class=clsSilver>
		<td width=5px bgcolor=blue>&nbsp;&nbsp;</td>
		<td width=755px nowrap>&nbsp;&nbsp;Status do pedido "Aceito/Instalado" (time-slot reservado - acesso entregue)</td>
	</tr>	
	<tr class=clsSilver>
		<td width=5px bgcolor=red>&nbsp;&nbsp;</td>
		<td width=755px nowrap>&nbsp;&nbsp;Status do pedido "Pendente" (time-slot reservado - acesso não entregue)</td>
	</tr>	
	<tr class=clsSilver>
		<td width=5px bgcolor=#33CC33>&nbsp;&nbsp;</td>
		<td width=755px nowrap>&nbsp;&nbsp;Em estoque</td>
	</tr>	
	<tr class=clsSilver>
		<td width=5px bgcolor=white>&nbsp;&nbsp;</td>
		<td nowrap width=755px>&nbsp;&nbsp;Vago</td>
	</tr>	
	<tr class=clsSilver>
		<td>
			<font class="clsObrig" align=center>&nbsp;::&nbsp;</font>
		</td>
		<td>
			&nbsp;&nbsp;Campos de preenchimento obrigatório.
		</td>
	</tr>
</table>
</td>
</tr>
</table>
<SCRIPT LANGUAGE=javascript>
<!--
function AlocarFac()
{
	with (document.forms[0])
	{
		//Parametriza
		hdnPedId.value = RequestNode(objXmlReturn,"hdnPedId")
		hdnFacId.value = arguments[0]
		//Popula
		/*txtRedDetBastidor.value		= arguments[0]
		txtRedDetRegua.value		= arguments[1]
		txtRedDetPosicao.value		= arguments[2]
		txtRedDetTimeslot.value		= arguments[3]
		txtRedDetFila.value			= arguments[4]
		txtRedDetEstacao.value		= arguments[5]
		txtRedDetDistribuidor.value	= arguments[6] */
		txtRedDetProvedor.value	= hdnProvedor2.value
		//alert(arguments[0])
		hdnAcao.value = "AlocarFacConsRedeDetManobra"
		target = "IFrmProcesso2"
		action = "ProcessoFac.asp"
		submit()
	}
}

function LimparFacilidade()
{
	with (document.forms[0])
	{
		txtRedDetBastidor.value		= ""
		txtRedDetRegua.value		= ""
		txtRedDetPosicao.value		= ""
		txtRedDetTimeslot.value		= ""
		txtRedDetFila.value			= ""
		txtRedDetEstacao.value		= ""
		txtRedDetDistribuidor.value	= ""
	}
}

function CarregarDocLog()
{
	document.onreadystatechange = CheckStateDocLog;
	document.resolveExternals = false;
}

function CheckStateDocLog()
{
  var state = document.readyState;
  if (state == "complete")
  {
	PopularForm(objXmlReturn)
	with (document.forms[0])
	{
		if (hdnProvedor2.value != "" && hdnLocalInstala.value != "")
		{
			if (RequestNode(objXmlReturn,"hdnLocalInstala") != "")
			{
				//hdnJSReturn.value = "parent.PopularForm(parent.objXmlReturn);parent.Procurar();parent.document.forms[0].hdnJSReturn.value = ''"
				ResgatarDominioNO(hdnLocalInstala)
			}else
			{
				hdnJSReturn.value = ""
			}
		}
	}	
  }
}

CarregarDocLog()
//-->
</SCRIPT>
<iframe	id			= "IFrmProcesso2"
	    name        = "IFrmProcesso2" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</body>
</Html>
