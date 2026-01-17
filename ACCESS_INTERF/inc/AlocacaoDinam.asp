<script language='javascript' src="../javascript/cla.js"></script>

<script language='javascript' >



function ResgatarTecnologia() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param  = document.getElementById('cboNewFacilidade').value
		param2 = document.getElementById('cboNewTecnologia').value
		param3 = document.getElementById('hdnSolId').value
		param4 = document.getElementById('hdnAcf_ID2').value
		param5 = document.getElementById('hdnAcl_IDAcessoLogico').value

		//alert (param)
		//alert (param2)
		//alert (param3)
		//alert (param4)
		//alert (param5)

		if (param ==0){param=0}
		if (param2 ==0){param2 =0}
		if (param3 ==0){param3 =0}
		if (param4 ==0){param4 =0}
		if (param5 ==0){param5 =0}

        xmlhttp.onreadystatechange = processadorMudancaEstadoTecnologia;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_Tecnologia.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<Dados><param>"+param+"</param><param2>"+param2+"</param2><param3>"+param3+"</param3><param4>"+param4+"</param4><param5>"+param5+"</param5></Dados>"
	  // alert (strXML)     
	xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoTecnologia () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spnTecnologia").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}

	function Trim(str){return str.replace(/^\s+|\s+$/g,"");}
	
	
	function validarNumerico(input) {
	  var valor = input.value;

	  // Express o regular para validar n meros inteiros ou decimais
	  var regex = /^[+-]?\d*\.?\d*$/;

	  // Se o valor n o for um n mero v lido, removemos o  ltimo caractere digitado
	  if (!regex.test(valor)) {
		input.value = valor.slice(0, -1);  // Remove o  ltimo caractere digitado
	  }
	}
	
	//Fun  o para a valida  e de tipos
function ValidarTipo4(Campo,intTipo)
{
		//alert(document.getElementsByName(Campo))
		
	//var 	intTipo = 0;
	//alert(Campo.value)
	//if (Campo == '[object]')
	//{
		//alert(1)
		var checkStr = Campo.value;
	//}
	//else
	//{
	//	alert(2)
	//	var checkStr = Campo;
//	}	
	
	
	//var checkStr = document.getElementsByName(Campo);
	//alert(checkStr.length)
	var allValid = true;
	var decPoints = 0;
	var allNum = "";
	switch (intTipo)
	{
		case 0:
			var checkOK = "0123456789" //int,smallint,bit
			break
		case 1:
			var checkOK = "QWERTYUIOPASDFGHJKL ZXCVBNMqwertyuiopasdfghjkl zxcvbnm " 
			break
		case 2:
			var checkOK = "QWERTYUIOPASDFGHJKL ZXCVBNMqwertyuiopasdfghjkl zxcvbnm01234546789 "
			break
		case 4:
			var checkOK = "-" //Tra o do cep
			break
		case 5:
			var checkOK = " " //Em banco
			break
		case 6:
			var checkOK = "0123456789 " //int,smallint,bit com espa o
			break
		
		//@@Davif - Incluido para aceitar o caracter (*) na Designa  o do Servi o
		case 7:
			var checkOK = "QWERTYUIOPASDFGHJKL ZXCVBNMqwertyuiopasdfghjkl zxcvbnm01234546789*/:.-_|\ "
			break
		case 8:
			var checkOK = "0123456789." //int,smallint,bit
			break

		default:
			var checkOK = intTipo //Recebe o pr prio valor
			break
	}

	//alert(checkOK)
	//alert(checkStr.length)
	for (var i = 0;  i < checkStr.length;  i++)
	{
		
		ch = checkStr.charAt(i);
		//alert(ch)
		for (var j = 0;  j < checkOK.length;  j++)
			if (ch == checkOK.charAt(j))
			break;
		if (j == checkOK.length)
		{
			allValid = false;
			break;
		}
		allNum += ch;
		//alert(allNum)
	}
	//alert(allValid)
	if (!allValid)
	{
		if (Campo == '[object]')
		{
			alert("Tipo de campo incorreto.") 
			Campo.value=allNum
			if (Campo.disabled == false) Campo.focus();
		}
		else
		{
			alert("Campo fora do padr o.") 
			Campo.value=allNum
			if (Campo.disabled == false) Campo.focus();
		}	
		return (false);
	}
	return (true);
}


	function GravarNewFacilidade2 ()
	{
		//alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
		//alert(document.forms[0].elements.length);

		
		for (var intIndex=0;intIndex<document.forms[0].elements.length;intIndex++){
			var elemento = document.forms[0].elements[intIndex];
/**
			if (elemento.name == "1" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "2" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "3" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "4" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "5" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "6" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "7" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "8" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "9" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "10" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "11" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "12" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "13" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "14" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "15" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "16" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "17" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "18" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "19" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "20" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "21" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "22" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "23" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "24" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "25" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "26" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "27" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "28" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "29" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "30" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "31" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "32" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "33" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "34" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "35" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "36" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "37" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "38" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}
			if (elemento.name == "39" && elemento.value == "S" )
			{

				strCampo = 'campo_' + elemento.name;
				//alert(strCampo );
				var check = document.getElementsByName(strCampo);
				for(var i = 0; i < check.length; i++){
				 	
				//	alert(check[i].value);
					if ( IsEmpty(check[i].value) ){
						alert("Favor informar o campo obrigat rio da Facilidade do Acesso ");
						return;
					}
						
					 
				}
			}

**/

            // Williams (Z518145) - 22/05/2025
            // Desativado para n o criticar obrigatoriedade quando o campo n o tiver preenchido. Esse campo ja vem bloqueado
			//

			//if (elemento.name == "hdnfacilidadeServico")
			//{
			//	//alert(elemento.name);
			//	//alert(elemento.type);
			//	//alert(elemento.value);
			//	if ( elemento.value == 1 )
			//	{
			//		if ( IsEmpty(document.forms[0].ser_Vlan.value) ){
			//			alert("Favor informar a VLAN");
			//			return;
			//		}
			//		if ( IsEmpty(document.forms[0].ser_portaOLt.value) ){
			//			alert("Favor informar a Porta");
			//			return;
			//		}
			//		if ( IsEmpty(document.forms[0].ser_PE.value) ){
			//			alert("Favor informar o Eqpto Agregado");
			//			return;
			//		}
				
				
			//	}
			//}
		
		}


		with (document.forms[0])
		{
			target = "IFrmProcesso"
			action = "GravarNewFacilidade.asp"
			submit()
			
			
		}

	}
</script>

<!--
<table border=0 cellspacing="0" cellpadding="0" width="760">
	<tr>
		<td>
			<iframe	id			= "IFrmEntregaProv"
				    name        = "IFrmEntregaProv"
				    width       = "100%"
				    height      = "83"
				    src			= "../inc/PrevisaoProvedor.asp?dblAcfId=<%=DblAcf_ID%>&dblSolId=<%=dblSolId%>&dblEild=<%=strEild%>&dblPonta=<%=strPonta%>"
					frameborder = "0"
				    scrolling   = "no"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>
-->

<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
	<tr>
		<th colspan=10>&nbsp; &nbsp;Recurso</th>
	</tr>

<!--JCARTUS-->
	<tr class="clsSilver">
		<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Entrega</td>
		<td colspan="8">
	    	<%
			set objRS = db.execute("CLA_sp_sel_estacao " & Trim(strLocalInstala))
			%> 
			<input type="Hidden" name="cboLocalInstala" value="<%=strLocalInstala%>">
		  <input type="text" class="text" disabled name="txtCNLLocalEntrega" value=<%=objRS("Cid_Sigla")%> maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A"  >&nbsp;
		  &nbsp;<input type="text" class="text" disabled name="txtComplLocalEntrega" value=<%=objRS("Esc_Sigla")%> maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsuFac(document.Form2.txtCNLLocalEntrega,document.Form2.txtComplLocalEntrega,<%=dblUsuId%>,1);" TIPO="A" >
		</td> 
		<td colspan="1">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Configura  o</td>
		<td colspan="8">
			<%
			set objRS = db.execute("CLA_sp_sel_estacao " & Trim(strLocalConfig))
			%>
			<input type="Hidden" name="cboLocalConfig" value="<%=strLocalConfig%>">
		  <input type="text" class="text" disabled name="txtCNLLocalConfig" value=<%=objRS("Cid_Sigla")%> maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		  &nbsp;<input type="text" class="text" disabled name="txtComplLocalConfig" value=<%=objRS("Esc_Sigla")%> maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsuFac(document.Form2.txtCNLLocalConfig,document.Form2.txtComplLocalConfig,<%=dblUsuId%>,2);" TIPO="A">
		</td>
		<td colspan="1">&nbsp;</td>
	</tr>
<!--JCARTUS-->

	

	

	<tr class=clsSilver>
		<td width=170 ><font class="clsObrig">:: </font>Provedor</td>
		<td colspan="9" >
			<%
			set objRS = db.execute("CLA_sp_sel_provedor " & Trim(strProId))
			%>
			<select name="cboProvedor" style="width:250px" disabled>
				<%Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "' tag_provedor=" & strCartaProv & strItemSel & ">" & objRS("Pro_Nome") & "</Option>"%>
			</select>
		</td>
	</tr>
	<%
		set objRS = db.execute("CLA_sp_sel_newconsultaTecnologiaFacilidade " & dblSolId )
		newfac_id = Trim(objRS("newfac_id")) 
	%>

	<tr class=clsSilver>
		<td width=170 ><font class="clsObrig">:: </font>Facilidade</td>
		<td colspan="9" >
			
			<select name="cboNewFacilidade" style="width:250px" disabled>
				<%Response.Write "<Option value='" & Trim(objRS("newfac_id")) & "' >" & objRS("newfac_Nome") & "</Option>"%>
			</select>
		</td>
	</tr>
	
	<tr class=clsSilver>
		<td width=170 ><font class="clsObrig">:: </font>Tecnologia</td>
		<td colspan="9" >
			<%
				set objRS2 = db.execute("CLA_sp_sel_AssocTecnologiaFacilidade null, null,  " & newfac_id   )
			%>
			<select name="cboNewTecnologia" style="width:250px" disabled >
				

					<%
					While not objRS2.Eof
						strItemSel = ""
						if Trim(objRS("newTec_id")) = Trim(objRS2("newTec_id")) then strItemSel = " Selected " End if
						Response.Write "<Option value=" & objRS2("newTec_id") & strItemSel & ">" & objRS2("newTec_Nome") & "</Option>"
						objRS2.MoveNext
					Wend
					strItemSel = ""
					%>
			</select>
			<!-- <input type="Button" class="button" name="btnAlt" value="Alterar Tecnologia" onclick="ResgatarTecnologia()"> -->
		</td>
		
	</tr>
	
	
	
	
</table>
<span ID=spnTecnologia>
<%
	strFacilidadeServico = "0"
		
			Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
			Vetor_Campos(2)="adInteger,2,adParamInput," & DblAcf_ID
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_newfacilidadeServico",2,Vetor_Campos)
			'Response.Write strSqlRet
			Set objRS = db.Execute(strSqlRet)
			objRS.Close
			objRS.CursorLocation = adUseClient
			objRS.Open

intCount=1
if not objRS.Eof and not objRS.Bof then
	
	
		if Trim(objRS("orisol_id"))= "10" then 

			strVlan =  "VLAN "
			strSVlan = "SVLAN "
			strPorta = "Porta "
			strPe =    "Eqpto Agregador "


		else

			strVlan =  "VLAN"
			strSVlan = "SVLAN"
			strPorta = "Porta"
			strPe =    "PE"


		end if
	
	strFacilidadeServico = "1"

%>					
	<table cellspacing=0 cellpadding=0 width=760 border=0>
		<tr><th colspan=10>&nbsp; &nbsp;Facilidade do Servi o</th></tr>
		
		<tr class=clsSilver2>
			<td width="170px" nowrap><font class="clsObrig">:: </font><%=strPE %></td>
			<td colspan="9" >
				<input type=text class=text name='ser_PE'  size='40' maxlength='15' value="<%=Trim(objRS("newfacservico_pe"))%>"  disabled>
			</td>
		</tr>
		<tr class=clsSilver> 
			<td width="170px" nowrap><font class="clsObrig">:: </font><%=strPorta %></td>
			<td colspan="9" >
				<input type=text class=text name='ser_portaOLt'  size='40' maxlength='30' value="<%=Trim(objRS("newfacservico_porta"))%>" disabled> 
			</td>
		</tr>

		<tr class=clsSilver2>
			<td width="170px" nowrap><font class="clsObrig">:: </font><%=strVlan %> </td>
			<td colspan="9" >
			<input type=text class=text name='ser_Vlan'  size='5' maxlength='4'  onKeyUp="ValidarTipo(this,0)" value="<%=objRS("newfacservico_vlan")%>" disabled>
			</td>
		</tr>
		
		<tr class=clsSilver>
			<td width="170px" nowrap> &nbsp;&nbsp; <%=strSVlan %> </td>
			<td colspan="9" >
				<input type=text class=text name='ser_SVLAN'  size='5' maxlength='4' onKeyUp="ValidarTipo(this,0)" value="<%=Trim(objRS("newfacservico_svlan"))%>" disabled>
			</td>
		</tr>
		

	</table>
<% end if %>


<table cellspacing=0 cellpadding=0 width=760 border=0>
	<tr><th colspan=10>&nbsp; &nbsp;Facilidade do Acesso</th></tr>
		
			
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim strTipo 

'strSql = "CLA_sp_sel_AssocTecnologiaFacilidade"

strFacilidadeAcesso = "false"

'Call PaginarRS(0,strSql)
			Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_newAlocacaoAcesso",1,Vetor_Campos)
			'Response.Write strSqlRet
			Set objRS = db.Execute(strSqlRet)
			objRS.Close
			objRS.CursorLocation = adUseClient
			objRS.Open
intCount=1
'response.write "<script>alert('"&objRS.PageSize&"')</script>"
if not objRS.Eof and not objRS.Bof then
	strFacilidadeAcesso = "true"
	'For intIndex = 1 to objRS.PageSize
	While Not objRS.Eof
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

			if trim(objRS("formato")) = "TEXTO" 	then strTipo = "7" end if
			if trim(objRS("formato")) = "NUMERICO"  then strTipo = "0" end if 
			'response.write "<script>alert('"&trim(objRS("formato"))&"')</script>"
			'response.write "<script>alert('"&trim(strTipo)&"')</script>"
			'response.write "<script>alert('"&trim(objRS("newcombo_id")&"')</script>"
			'trim(objRS("newcombo_id")
		%>
		<tr class=<%=strClass%>>
			<td width="170px" nowrap><% if trim(objRS("obrigatorio")) = "S" 	then  %>  <font class="clsObrig">:: </font>  <% end if  %>

				<%=TratarAspasHtml(objRS("label"))%></td>
			<% if trim(objRS("formato")) = "COMBO"  then %>
			
				<%
					set objRS2 = db.execute("CLA_sp_sel_EstruturaCombo null,  " & trim(objRS("newcombo_id")) )
					estrutura_combo_id = Trim(objRS2("estrutura_combo_id")) 
				%>

					<td colspan="9" >
						
						<select name='campo_<%=intCount%>' style="width:250px" >
								<option value=""></option>
								<%
									While not objRS2.Eof
										strItemSel = ""
										if Trim(objRS("conteudo")) = Trim(objRS2("label")) then strItemSel = " Selected " End if
										Response.Write "<Option value=" & objRS2("estrutura_combo_id") & strItemSel & ">" & objRS2("label") & "</Option>"
										objRS2.MoveNext
									Wend
									strItemSel = ""
								%>
						
						</select>
					</td>
				
			<% else %>
				<td colspan="9" >
					<input type=text class=text name='campo_<%=intCount%>' oninput="ValidarTipo4(this, <%=strTipo%>)" size="<%=trim(objRS("tamanho"))%>" maxlength="<%=trim(objRS("tamanho"))%>" value="<%=trim(objRS("conteudo"))%>" >
				
				</td>
				
			<% end if %>
				<input type="Hidden" name='<%=intCount%>' value="<%=trim(objRS("obrigatorio"))%>">
				
		</tr>
		<%
		intCount = intCount+1
		objRS.MoveNext
	Wend
		'objRS.MoveNext
		'if objRS.EOF then Exit For
	'Next
End if
%>
		<!--</td>
	</tr>-->
</table>

<input type="Hidden" name="facilidadeAcesso" value="<%=strFacilidadeAcesso%>">
<input type="Hidden" name="hdnfacilidadeServico" value="<%=strFacilidadeServico%>">
<!--<input type="Hidden" name="hdnSolId" value="<%=dblSolId%>"> -->

<input type="Hidden" name="hdnAcl_IDAcessoLogico" value="<%=strIdLogico%>">
<input type="Hidden" name="hdnAcf_ID2" value="<%=DblAcf_ID%>">	

</span>
<span id=spnDet></span>
<span id=spnNDet></span>
<span id=spnAde></span>
<span id=spnBsodNet></span>
<span id=spnBsodVia></span>
<span id=spnBsod></span>
<span id=spnBsodLight></span>
<span id=spnFoEtherNet></span>
<span id=spnSwitchRadioIP></span>

<table border=0 cellspacing="0" cellpadding="0" width=760 >
	<tr>
		<td >
			<iframe	id			= "IFrmProcesso1"
				    name        = "IFrmProcesso1"
				    width       = "760"
				    height      = "18px"
				    frameborder = "0"
				    scrolling   = "no"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>



<table width="760" border=0>
	<tr><td>
	<table width=50% border=0 align=center cellspacing=1 cellpadding=1>
		<tr class=clsSilver2>
			
			<td colspan=4 align=center><input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Alterar Facilidade(s)" onclick="return GravarNewFacilidade2()" accesskey="I" onmouseover="showtip(this,event,'Alterar Facilidade(s)(Alt+I)');"></td>
					
			
		</tr>
			
			
		
	</table>
	</td>
	</tr>
</table>

</Form>
<Form name="Form10" method="Post">
	

</Form>
<table width="760">
	<tr>
		<td>
			<font class="clsObrig">:: </font> Campos de preenchimento obrigat rio.
		</td>
	</tr>
	<tr>
		<td>
			&nbsp;&nbsp;&nbsp;&nbsp;Legenda: A - Alfanum rico;  N - Num rico;  L - Letra
		</td>
	</tr>
 
</table>

