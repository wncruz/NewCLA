<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: ProcessoListaSenha.ASP
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: Página que recebe e exibe o html do Grid das páginas ConsultaOSProvedor e CadastraOSProvedor.
%>

<SCRIPT LANGUAGE=javascript>
//Marca/Desmarca todos os checkboxs do grid
function CheckTodosSenha()
{
	var strCheck = ""
	for(var i =1;;i++)
	{
		//alert(document.getElementById('chkLinha' + i))
		if (document.getElementById('chkLinha' + i) == "[object]")
		{
			document.getElementById('chkLinha' + i).checked = document.getElementById('chkTudo').checked;
			if(document.getElementById('chkLinha' + i).checked == true)
			{
				if(document.getElementById('hdnUtilizado' + i).value == "SIM")
				{
					document.getElementById('chkLinha' + i).checked = false
				}
				else
					strCheck = strCheck + document.getElementById('hdnLinha' + i).value + ",";
			}
		}
		else 
			break;
	}
	if(strCheck != "")
		strCheck = strCheck.substring(0,strCheck.length -1)
	parent.parent.document.getElementById('hdnCheck').value = strCheck

}

//Atualiza quais foram marcados
function CheckUmSenha()
{
	var strCheck = ""
	for(var i =1;;i++)
	{
		if (document.getElementById('chkLinha' + i) == "[object]")
		{
			if(document.getElementById('chkLinha' + i).checked == true)
			{
				if(document.getElementById('hdnUtilizado' + i).value == "SIM")
				{
					alert("Senha PIN já utilizada pela Solicitação " + document.getElementById('hdnSol' + i).value);
					document.getElementById('chkLinha' + i).checked = false
				}
				else
					strCheck = strCheck + document.getElementById('hdnLinha' + i).value + ",";
			}
		}
		else 
			break;
	}
	if(strCheck != "")
		strCheck = strCheck.substring(0,strCheck.length -1)
	parent.parent.document.getElementById('hdnCheck').value = strCheck
}

</SCRIPT>
<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
<HTML>
<HEAD>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
</HEAD>
<BODY>
<Form name=Form1 method=Post onsubmit="return false">
<input type=hidden name=hdnAprovID value =""> 
<input type=hidden name=hdnAprov_Utilizado value =""> 
<%
	Response.Write Request.Form("hdstrHtmlRet")
%>
</HTML>
