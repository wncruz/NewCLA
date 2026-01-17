<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AlteracaoTaxaServico.asp
'	- Responsável		: Vital
'	- Descrição			: Alteração da taxa de serviço para o processo de alteração
%>
<!--#include file="../inc/data.asp"-->
<HTML>
<HEAD>
<Title>CLA - Controle Local de Acesso</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function window_onunload() {
	window.returnValue = objXmlGeral
}
//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var objAryParam = window.dialogArguments
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function Gravar()
{
	var objNode = objXmlGeral.selectNodes("//TaxaServico")
	var blnAchou = false
	if (objNode.length > 0)
	{
		for (var intIndex=0;intIndex<objNode.length;intIndex++)
		{
			if (objNode[intIndex].childNodes[2].text!="")
			{
				blnAchou = true
			}
		}
		if (!blnAchou && objNode.length == 1)
		{
			alert("Não é possível desativar o único acesso físico disponível.")
			return
		}
		else
		{
			if (!blnAchou)
			{
				alert("Favor informar pelo menos uma taxa de serviço.")
				return
			}	
		}
	}

	var objNode = objXmlGeral.selectNodes("//xDados")
	objNodeFilho = objXmlGeral.createNode("element", "Ret", "")
	objNodeFilho.text = "0"
	objNode[0].appendChild (objNodeFilho)

	window.returnValue = objXmlGeral
	window.close();
}

function Sair()
{
	var objNode = objXmlGeral.selectNodes("//xDados")
	objNodeFilho = objXmlGeral.createNode("element", "Ret", "")
	objNodeFilho.text = "999"
	objNode[0].appendChild (objNodeFilho)
	window.returnValue = objXmlGeral
	window.close();
}

function AtualizarTaxaServico(obj)
{
	var objNode = objXmlGeral.selectNodes("//TaxaServico[Acf_Id="+obj.Acf_Id+"]")
	if (objNode.length > 0)
	{
		objNode[0].childNodes[2].text = obj.value
	}
}
//-->
</SCRIPT>
</HEAD>
<BODY  class=TA leftmargin=3 LANGUAGE=javascript onunload="return window_onunload()">
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
<form name=Form1 method=Post onSubmit="return false">
<input type=hidden name=hdnAcao value="AlterarStatus">
<input type=hidden name=hdnSubAcao >
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnStsId>
<input type=hidden name=hdnUserName value="<%=strUserName%>">
<input type=hidden name=hdnHistorico>
<input type=hidden name=hdnIdLog>
<input type=hidden name=hdnIdFis>
<input type=hidden name=hdnPropAcesso>
<input type=hidden name=hdnDesigServ>
<input type=hidden name=hdnDtEntrega>
<input type=hidden name=hdnCboServico>
<input type=hidden name=hdnPadraoDesignacao>
<input type=hidden name=hdnTipoProcesso>
<input type=hidden name=hdnSubSubAcao>
<input type=hidden name=hdnXml>
<input type=hidden name=hdnAcfId>
<input type=hidden name=hdnIdAcessoFisLista>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<script language=javascript>
function ListarTaxaServico()
{
	with (document.forms[0])
	{
		hdnIdLog.value = objAryParam[0] 
		hdnSolId.value = objAryParam[1] 
		hdnTipoProcesso.value = objAryParam[2]
		hdnIdAcessoFisLista.value = objAryParam.join(",")

		hdnAcao.value = "ListarTaxaServico"
		target = "IFrmProcesso"
		action = "ProcessoAlteracao.asp"
		submit()
	}	
}
ListarTaxaServico()
</script>
<table border=0 width=100% cellspacing=0 cellpadding=0>
	<tr>
		<th><p align=center>Alteração da Taxa de Serviço</p></td>
	</tr>
</table>
<span id=spnListaTaxaServico></span>
<table border=0 width=100% >
<tr>
	<td align=center>
		<input type="button" class="button" name="btnGravar" value="Gravar"  onclick="Gravar()" accesskey="I" onmouseover="showtip(this,event,'Alterar a taxa de serviço (Alt+I)');" >&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:Sair();" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>		
</form>
</BODY>
</HTML>