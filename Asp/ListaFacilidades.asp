<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel=stylesheet type="text/css" href="../css/cla.css">
</HEAD>
<BODY topmargin=0 leftmargin=0 bgColor=#eeeeee class=TA >
<Form name=Form1 method=Post>
	<span id=spnFacilidade></span>
</Form>	
</BODY>
</HTML>
<SCRIPT LANGUAGE=javascript>
<!--
//Verifica se o Xml já esta carregado
function CheckStateIframe()
{
  var state = document.readyState;
  if (state == "complete")
  {
	parent.CarregarLista()
  }
}

function CarregarIframe()
{
	document.onreadystatechange = CheckStateIframe;
	document.resolveExternals = false;
}
CarregarIframe()
//-->
</SCRIPT>
