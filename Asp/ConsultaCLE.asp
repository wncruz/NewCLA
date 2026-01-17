<script>
function isNumber(n) {   return !isNaN(parseFloat(n)) && isFinite(n); } 

//Complementa numericos com zeros a esqueda e alfa como espaços a direita
function CompletarCampo(obj)
{
	
	if (obj.value != "" && obj.value != 0 )
	{
		var intLen = parseInt(obj.size) - parseInt(obj.value.length)
	
		switch (obj.TIPO.toUpperCase())
		{
			case "N":
				for (var intIndex=0;intIndex<intLen;intIndex++)
				{
					obj.value = "0" + obj.value
				}
				break
			default :
				for (var intIndex=0;intIndex<intLen;intIndex++)
				{
					obj.value = obj.value + " "
				}
		}
	}	
}

function consultaCLE()
{   spnConta15.innerHTML="" 
    if (isNumber(document.getElementById("txtConta15").value))
    {
      var retConta15 = "<table border=1><tr><td colspan=2 align=center>RESULTADO</td></tr>";
			var strConta15 = document.getElementById("txtConta15").value
			var xmlDoc  = new ActiveXObject("Microsoft.XMLDOM");
			var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
			
			var strXML
			strXML      = "<root>"
			strXML      = strXML + "<conta15>" +  document.getElementById("txtConta15").value + "</conta15>"
			strXML      = strXML + "</root>"
			xmlDoc.loadXML(strXML);
			xmlhttp.Open("POST","retornaContaCLE.asp" , false);
			xmlhttp.Send(xmlDoc.xml);
      strXML      = xmlhttp.responseText;				
      xmlDoc.loadXML(strXML);
			var ndCodRetorno     = xmlDoc.getElementsByTagName("codRetorno")[0].firstChild.nodeValue
			var ndRazaoSocial    = xmlDoc.getElementsByTagName("razaoSocial")[0].firstChild.nodeValue;
			var ndNomeFantasia   = xmlDoc.getElementsByTagName("nomeFantasia")[0].firstChild.nodeValue;	
//			alert(ndCodRetorno);		 

			if (ndCodRetorno=="*"){			
				alert("Não foi possível consultar o CLE. Tente novamente.")
			}
			else if (ndCodRetorno=="2"){		
				alert("SubConta não encontrada no CLE.")
				document.getElementById("txtConta15").focus();
			}
			else if(ndRazaoSocial!="*"){  
				retConta15 = retConta15 + "<tr><td>Razão Social</td><td>" + ndRazaoSocial + "</td></tr>";
				retConta15 = retConta15 + "<tr><td>Nome Fantasia</td><td>" + ndNomeFantasia + "</td></tr>";				
				retConta15 = retConta15 + "</table>";
				spnConta15.innerHTML = retConta15;
			}
			else
			{			
				alert("Conta Corrente [" + strConta15 + "] não encontrada no CLE.")
			}
		}
		else{
				alert("Informe uma Conta Corrente válida.");
				document.getElementById("txtConta15").focus();
		}							    
}
</script>
:. Informe a Conta Corrente (N15): <input type=text size=15 maxlength=15 id=txtConta15 name=txtConta15  TIPO="N" onblur="CompletarCampo(this)">
<INPUT TYPE=BUTTON OnClick="consultaCLE();" VALUE="Consultar CLE">
<br><br><br>
<span id=spnConta15></span>
</body>
</html>