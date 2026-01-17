<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language='javascript' src="../javascript/claMsg.js"></script>
</HEAD>
<BODY >
<%
hdnAcao				= request("hdnAcao")
hdnTipoAcao			= request("hdnTipoAcao")
hdnIdAcessoLogico	= request("hdnIdAcessoLogico")
hdnSolId			= request("hdnSolId")
hdn678				= request("hdn678")
hdnTipoProcesso		= request("hdnTipoProcesso")
%>
<form method="post" name="Form4"  >
<input type=hidden name=hdnAcao				value="<%=hdnAcao%>">
<input type=hidden name=hdnTipoAcao			value="<%=hdnTipoAcao%>" >
<input type=hidden name=hdnXml				value="<%=hdnIdAcessoLogico%>">
<input type=hidden name=hdnIdAcessoLogico	value="<%=hdnSolId%>">
<input type=hidden name=hdnSolId			value="<%=hdnSolId%>">
<input type=hidden name=hdn678				value="<%=hdn678%>">
<input type=hidden name=hdnTipoProcesso		value="AlterarInfoAcesso">
</form>
<iframe	id			= "IFrmProcesso"
		name        = "IFrmProcesso"
		width       = "0"
		height      = "0"
		frameborder = "0"
		scrolling   = "no"
		align       = "left">
</iFrame>

<script >


/*
	function resposta(DBAction,valor){
		alert(DBAction)
		//parent.resposta(DBAction,valor);
		
	}
*/
	function Message(objXmlRet){
	
	var intRet = window.showModalDialog('Message.asp',objXmlRet,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	var objNode = objXmlRet.selectNodes("//CLA_RetornoTmp")
	if(objNode[0].attributes[0].value == 2){
		window.returnValue = true
	}else{
		window.returnValue = false
	}
	self.close();
	
}
	var objAryFac = window.dialogArguments
	Form4.hdnXml.value = objAryFac.xml
	with(Form4){
		target = "IFrmProcesso"
		action = "ProcessoSolic.asp"
		submit();
	}

</script>
</BODY>
</HTML>

