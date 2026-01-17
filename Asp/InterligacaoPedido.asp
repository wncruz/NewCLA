<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/AlocacaoFac.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: InterligacaoPedido.asp
'	- Descrição			: Alteração de PADE/PAC

strCamposInter = "<table cellspacing=1 cellpadding=0 width=100% border=0>"
strCamposInter = strCamposInter & "<th width=200px>&nbsp;Nº do Acesso</th>"
strCamposInter = strCamposInter & "<th width=150px>&nbsp;Origem</th>"
strCamposInter = strCamposInter & "<th width=150px>&nbsp;Destino</th>"
strCamposInter = strCamposInter & "</table>"

Dim intTipoProcesso
Dim strAcfObs
Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
objXmlDados.loadXml("<xDados/>")

dblPedId = Request("hdnPedId")

Set objRS = db.Execute("CLA_SP_Sel_Facilidade " & dblPedId)
Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,strAcfObs)
Set objNode = objXmlDados.SelectNodes("//Facilidade")

intFac = objNode.length
strXmlFac = FormatarXml(objXmlDados)
%>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
var objAry = window.dialogArguments
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
var intIndice = 0
var objAryObjs = new Array()
var objAryFac = new Array()
var objAryFacRet
var strPagina = "ManobraInterligacao"

//Campos DET
objAryObjs[1] = new Array("Int_CorOrigem")
//Campos NDET

objAryFac[0] = new Array("","")

var strDet = new String('<%=strDet%>')

var strCamposInter = new String('<%=strCamposInter%>')

function CarregarLista()
{
	objXmlGeral.onreadystatechange = CheckStateXml;
	objXmlGeral.resolveExternals = false;
	if (parseInt(<%=intFac%>) != 0){
		objXmlGeral.loadXML("<%=strXmlFac%>") 
		document.forms[0].Ped_Id.value = objAry[2]

	}else{
		var objXmlRoot = objXmlGeral.createNode("element","xDados","")
		objXmlGeral.appendChild (objXmlRoot)
		document.forms[0].Ped_Id.value = objAry[2]
	}	
}
//Verifica se o Xml já esta carregado
function CheckStateXml()
{
  var state = objXmlGeral.readyState;
  
  if (state == 4)
  {
    var err = objXmlGeral.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    } 
    else 
    {
		ListarInterligacoes(<%=dblPedId%>)
	}
  }
}
function JanelaConfirmacaoFac(objXmlGeral){
	var intRet = window.showModalDialog('ConfirmacaoAlocacaoFac.asp',objXmlGeral,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	if (intRet == 1){
		intRet = 0
		//Alterar(1) 
	}
} 

function CarregarDoc()
{
	document.onreadystatechange = CheckStateDoc;
	document.resolveExternals = false;
}

function CheckStateDoc()
{
  var state = document.readyState;
  
  if (state == "complete")
  {
    var err = objXmlGeral.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    } 
    else 
    {
		//Carrega depois que o documento foi esta completo
		ResgatarInfoRede()
	}
  }
}

CarregarDoc()
//-->
</SCRIPT>
<Title>CLA - Controle Local de Acesso</Title>
<body scroll=no>
<form method="post">
<input type=hidden name=hdnXml>
<input type=hidden name=hdnAcao>
<input type="hidden" name="hdnIntIndice">
<input type="hidden" name="cboRede">
<input type="hidden" name="hdnProvedor">
<input type="hidden" name="Acf_Id">
<input type="hidden" name="Ped_Id">
<input type="hidden" name="hdnRecId">

<table border="0"  cellspacing="1" cellpadding="0" width=100%>
<tr>
	<th colspan=2><p align=center><font color=#ffffff>PADE/PAC Alocadas</font></p></th>
</tr>
</table>
<span id=spnCampos></span>
<table cellspacing=1 cellpadding=0 width=100% border=0>
	<tr class=clsSilver>
		<td >
			<iframe id=IFrmFacilidade 
					name=IFrmFacilidade 
					align=left 
					src="ListaFacilidades.asp" 
					frameBorder=0 
					width="100%" 
					BORDER=0
					height=80>
			</iframe>
		</td>
	</tr>
</table>
<span id=spnDet></span>
<span id=spnNDet></span>
<span id=spnAde></span>
<table border=0 cellspacing="0" cellpadding="0" width=100%  >
	<tr>
		<td >
			<iframe	id			= "IFrmProcesso1"
				    name        = "IFrmProcesso1" 
				    width       = "100%" 
				    height      = "18px"
				    frameborder = "0"
				    scrolling   = "no" 
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>
<table width="100%" border=0>
	<tr >
		<td align=center>
			<input type="button" class="button" name="btnImprimir" value="Imprimir" onClick="Imprimir()" accesskey="W" onmouseover="showtip(this,event,'Imprimir (Alt+W)');">&nbsp;
			<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.close()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
		</td>
	</tr>	
</table>
</form>
<SCRIPT LANGUAGE=javascript>
<!--
function Imprimir()
{
	IFrmFacilidade.focus() 
	window.print()
}

function ResgatarInfoRede()
{
	with (document.forms[0])
	{
		spnDet.innerHTML = strDet 
		spnCampos.innerHTML = strCamposInter 
	}
}
//-->
</SCRIPT>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<iframe	id			= "IFrmProcesso2"
	    name        = "IFrmProcesso2" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</body>
</html>