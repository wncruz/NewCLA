<%
Response.Expires = -1
Response.CacheControl = "no-cache"
%>
<HTML>
<HEAD>
<TITLE>CLA - Controle Local de Acesso&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--
function Retornar(objBtn)
{
	top.returnValue = objBtn.name
	window.close()
}
//-->
</SCRIPT>
</HEAD>
<BODY topmargin=5 leftmargin=10 background="../imagens/bgMsgbox.gif" >
<span id=spnLista></span>
<SCRIPT LANGUAGE=javascript>
<!--
var objAryParams = window.dialogArguments
var strTable = new String("<table width='98%' border='0' cellspacing='1' cellpadding='2' align='center'>")
for (intIndex=0;intIndex<objAryParams.length;intIndex++)
{
	if (intIndex==0)
	{
		strTable += "<tr><td align=center height=51px valign=center><img src='../imagens/interrogacao.gif' border=0></td><td align=center valign=top><font face=Arial size=2>"+objAryParams[intIndex]+"</font></td></tr><tr><td nowrap align=center colspan=2 valign=button>"
	}
	else
	{
		strTable += "<input type=button name='"+intIndex+"' value='"+objAryParams[intIndex]+"' onclick='Retornar(this)' style='width:75px;font-size: 9px;font-weight: normal;font-family: Verdana, Arial;'>&nbsp;</b>"
	}
}
strTable += "</td></tr></table>"
	
spnLista.innerHTML = strTable

//-->
</SCRIPT>
</BODY>
</HTML>

