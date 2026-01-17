<SCRIPT LANGUAGE=javascript>
<!--
var objXmlMsg = new ActiveXObject("Microsoft.XMLDOM")

function CarregarMsg()
{
	objXmlMsg.onreadystatechange = CheckStateXmlHeader;
	objXmlMsg.resolveExternals = false;
	objXmlMsg.load("../xml/claMsg.xml")
}
//Verifica se o Xml já esta carregado
function CheckStateXmlHeader()
{
  var state = objXmlMsg.readyState;
  
  if (state == 4)
  {
    var err = objXmlMsg.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    } 
    else 
    {
		resposta(<%=Cint("0" & DBAction)%>,'');
	}
  }
}
CarregarMsg()
//-->
</SCRIPT>