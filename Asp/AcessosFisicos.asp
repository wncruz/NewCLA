
<HTML>
<HEAD>

<link rel=stylesheet type="text/css" href="../css/cla.css">
</HEAD>
<BODY topmargin=0 leftmargin=0 bgColor=#eeeeee class=TA>
<Form name=Form1 method=Post>
	<span id=spnAcessoFis></span>
</Form>	
</BODY>
</HTML>
<SCRIPT LANGUAGE=javascript>

function CheckStateIframe()
{
	
  var state = document.readyState;
  //alert(state);
  if (state == "complete")
  {
	//alert("2");
	parent.CarregarLista();
	
	
  }
}

function CarregarIframe()
{
	document.onreadystatechange = CheckStateIframe;
	document.resolveExternals = false;
}
CarregarIframe()

</SCRIPT>