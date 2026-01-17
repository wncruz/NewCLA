/*
•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
	- Sistema			: CLA
	- Arquivo			: Cla.js
	- Responsável		: Vital
	- Descrição			: Funções JAVASCRIPTf genéricas utilizadas no sistema cla
*/

function isInteger(s){
	var i;
    for (i = 0; i < s.length; i++){   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
	var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){   
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
	// February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
		if (i==2) {this[i] = 29}
   } 
   return this
}

function ValidarData(dtStr){
	var dtCh= "/";
	var minYear=1900;
	var maxYear=2200;
	var daysInMonth = DaysArray(12)
	var pos1=dtStr.indexOf(dtCh)
	var pos2=dtStr.indexOf(dtCh,pos1+1)
	var strDay=dtStr.substring(0,pos1)
	var strMonth=dtStr.substring(pos1+1,pos2)
	var strYear=dtStr.substring(pos2+1)
	strYr=strYear
	if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
	if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
	for (var i = 1; i <= 3; i++) {
		if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
	}
	month=parseInt(strMonth)
	day=parseInt(strDay)
	year=parseInt(strYr)
	if (pos1==-1 || pos2==-1){
		alert("O formato da data deve ser: dd/mm/aaaa")
		return false
	}
	if (strMonth.length<1 || month<1 || month>12){
		alert("Mês inválido")
		return false
	}
	if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
		alert("Dia inválido")
		return false
	}
	if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
		alert("Entre com um ano com 4 digitos entre "+minYear+" e "+maxYear)
		return false
	}
	if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
		alert("Data inválida")
		return false
	}
return true
}

//Executa a busca de OS
function ProcurarCadastroOS(){
	with (document.forms[0])
	{
		if (!ValidarDM(txtPedido)) return;
		target = "IFrmLista"
		action = "ListaCadastraOSProvedor.asp"
		submit()
	}
}

function ValidarVirgula(Campo)
{
	var allValid = true;
	var allNum = "";
	if (Campo == '[object]')
	{
		var checkStr = Campo.value;
	}
	else
	{
		var checkStr = Campo;
	}	
	for (var i = 0;  i < checkStr.length;  i++)
	{
		ch = checkStr.charAt(i);
		if(ch == "&")
		{
			allValid = false;
			break;
		}
	}
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
			alert("Campo fora do padrão.") 
			Campo=allNum
		}	
		return (false);
	}
	return (true);
	
}

//Função para a validaçõe de tipos
function ValidarTipo(Campo,intTipo)
{
		//alert(Campo)
	if (Campo == '[object]')
	{
		var checkStr = Campo.value;
	}
	else
	{
		var checkStr = Campo;
	}	
	var allValid = true;
	var decPoints = 0;
	var allNum = "";
	switch (intTipo)
	{
		case 0:
			var checkOK = "0123456789" //int,smallint,bit
			break
		case 1:
			var checkOK = "QWERTYUIOPASDFGHJKLÇZXCVBNMqwertyuiopasdfghjklçzxcvbnm " 
			break
		case 2:
			var checkOK = "QWERTYUIOPASDFGHJKLÇZXCVBNMqwertyuiopasdfghjklçzxcvbnm01234546789 "
			break
		case 4:
			var checkOK = "-" //Traço do cep
			break
		case 5:
			var checkOK = " " //Em banco
			break
		case 6:
			var checkOK = "0123456789 " //int,smallint,bit com espaço
			break
		
		//@@Davif - Incluido para aceitar o caracter (*) na Designação do Serviço
		case 7:
			var checkOK = "QWERTYUIOPASDFGHJKLÇZXCVBNMqwertyuiopasdfghjklçzxcvbnm01234546789*/:.-_|\ "
			break
		case 8:
			var checkOK = "0123456789." //int,smallint,bit
			break

		default:
			var checkOK = intTipo //Recebe o próprio valor
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
			alert("Campo fora do padrão.") 
			
			Campo=allNum
		}	
		return (false);
	}
	return (true);
}


//Função para a verificação de varios tipo em um mesmo textbox
function ValidarNTipo()
{
	//Argumentos
	//ObjTextbox,Tipo,Tam,ProxTipo,Tam,ProxTipo,Tam,...
	//var checkOK = str;
	//var checkStr = Campo.value;
	var allValid = true;
	var decPoints = 0;
	var allNum = "";
	
	if (arguments.length > 0 )
	{
		var Campo = arguments[0]
		var strValue = arguments[0].value
		var intTamIni = 0
		var intTam = 0
		var intTipo = 0
		
		var intIndex=1

		while (intIndex < arguments.length)
		{
			intTipo	= arguments[intIndex]
			intTamIni = intTamIni + intTam
			intTam	= arguments[intIndex+1]
			
			checkStr = strValue.substring(intTamIni,intTamIni+intTam)
			
			if (!ValidarTipo(checkStr,intTipo))
			{
				intIndex = arguments.length
				Campo.value =  strValue.substring(0,intTamIni)
				Campo.focus()
				return false;
			}
			else
			{
				intIndex = intIndex + 2
			}	
		}
	}
	
	return true;
}

//Verifica o tipo atual do campo
function RetornarTipoAtual(intPos,objAryTipo)
{
	
	for (var intIndex=0;intIndex<objAryTipo.length;intIndex++)
	{
		if (intPos >= objAryTipo[intIndex][0] && intPos <= objAryTipo[intIndex][1] )
		{
			switch (parseInt(objAryTipo[intIndex][2]))
			{
				case 1:
					var checkOK = "QWERTYUIOPASDFGHJKLÇZXCVBNMqwertyuiopasdfghjklçzxcvbnm " 
					break
				case 2:
					var checkOK = "QWERTYUIOPASDFGHJKLÇZXCVBNMqwertyuiopasdfghjklçzxcvbnm01234546789 "
					break
				default:
					var checkOK = "0123456789 " //int,smallint,bit
					break
			}
		}
	}	
	
	return checkOK
}

//Valida um determinado range ex. números de 1 a 9
function ValidarRange(obj,intIni,iniFim)
{
	if (parseInt(obj.value)	< intIni || parseInt(obj.value)	> iniFim)
	{
		alert("Valor fora do intervalo " + intIni + " a " + iniFim + ".")
		obj.value = ""
		return false
	}
	else
	{
		return true
	}
}

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

//Faz submit quando paginamos
function ValidarPaginacao(strBtnPaginacao)
{
	with (document.forms[0])
	{
		if (strBtnPaginacao=="PagNro")
		{
			if (TbNroPag.value == "" || parseInt(TbNroPag.value) < 1 || parseInt(TbNroPag.value) > parseInt(hdTotalPages.value))
			{
				alert("Número de página inválida.")
				TbNroPag.focus()
				return
			}
		}
		target = self.name
		action = hdnPagina.value+"?btn="+strBtnPaginacao
		submit()
	}
}

//Limpa todos os campos do formulário
function LimparForm()
{
	for(var iForm=0; iForm < document.forms.length; iForm++)
	{
			
		for(var iElement=0; iElement < document.forms[iForm].elements.length; iElement++)
		{
			sTipo = document.forms[iForm].elements[iElement].type; 
			if( sTipo== "select-one" ||sTipo == "select-multiple" || sTipo == "text" || sTipo == "textarea")
			{
				document.forms[iForm].elements[iElement].value = ""
			}
			if (sTipo == "radio" || sTipo == "checkbox")
			{
				document.forms[iForm].elements[iElement].checked = false
			}
		}
	}
}

//Valida formato de e-mail
function ValidarEmail(campo)
{
	return true;

	if (campo.value!="")
	{ 
		var exclude=/[^@\-\.\w]|^[_@\.\-]|[\._\-]{2}|[@\.]{2}|(@)[^@]*\1/;
		var check=/@[\w\-]+\./;
		var checkend=/\.[a-zA-Z]{2,3}$/;

		if(((campo.value.search(exclude) != -1)||(campo.value.search(check)) == -1)||(campo.value.search(checkend) == -1))
		{
			alert("Email fora do padrão.");
			//campo.focus()
			return false;
		}
	}
	return true;
}

//Envia para exclusão de registro nos cadostros básicos
function ExlcuirRegistro()
{
	with (document.forms[0])
	{
		if (ConfirmarRemocao())
		{
			action = hdnPagina.value
			hdnAcao.value = "Excluir"
			submit()
			return true;
		}	
	}
}

//Adicina barras automaticamente em campos do tipo data
function AdicionaBarraData(objObjeto)
{
	if(objObjeto.value.length == 2 || objObjeto.value.length == 5)
	{
		objObjeto.value += "/";
	}
}

//Adiciona o - no campo CEP
function AdicionaBarraCep(objObjeto)
{
	if(objObjeto.value.length == 5)
	{
		objObjeto.value += "-";
	}
}

//Permite somente números
function OnlyNumbers()
{
	//Se for caracter de controle retorna
	if (event.keyCode < 32)
	{
		event.returnValue = true;
		return;
	}

	//Verifica se foi digitado um número
	if ((String.fromCharCode(event.keyCode) < '0') || (String.fromCharCode(event.keyCode) > '9'))
		event.returnValue = (false)
	else
		event.returnValue = (true);
}

// Faz replace de um caracter especifico em uma string
function Replace(strString,strSearch,strReplace)
{
	var strAux = new String(strString)
	while (strAux.indexOf(strSearch) != -1)
	{
		strAux = strAux.replace(strSearch, strReplace)
	}
	return strAux
}

//Verifica em um campos esta em banco ou não possui valor
function IsEmpty(psString)
{
	/*
	** Caracteres Inválidos
	*/
	var lsTab   = '\t', // Tab Char
		lsSpace = ' ' , // Space
        lsCRLF  = '\n', // CR LF
		lsCR    = '\r'; // CR
	
	/*
	** Procura por caracteres válidos
	*/
	for (var liPos = 0; liPos < psString.length; liPos++)
	{
		var lsChar = psString.charAt(liPos);
		if (lsChar != lsTab   &&
			lsChar != lsSpace && 
			lsChar != lsCRLF  && 
			lsChar != lsCR )
			return (false);
	}
	
	return (true);
}

//Validação para campos obrigatórios
function ValidarCampos(obj,strMsg)
{
	var blnAchou = false
	if ( obj.type == "radio" || obj.type == "checkbox")
	{
		for (var intIndex=0;intIndex<obj.length;intIndex++)
		{
			if (obj[intIndex].checked)
			{
				blnAchou = true
			}
		}
		if (!blnAchou)
		{
			alert(strMsg + " é um campo obrigatório.")	
		}
	}

	
	else
	{
		if (obj.value == "" || IsEmpty(obj.value))
		{
			alert(strMsg + " é um campo obrigatório.")
			obj.focus()
			return false
		}
	}	
	return true 
}

function getCheckedRadioValue(radioObj) {
	if(!radioObj)
		return "";
	var radioLength = radioObj.length;
	if(radioLength == undefined)
		if(radioObj.checked)
			return radioObj.value;
		else
			return "";
	for(var i = 0; i < radioLength; i++) {
		if(radioObj[i].checked) {
			return radioObj[i].value;
		}
	}
	return "";
}

//Envia para a página de detalhamento de pedido
function DetalharFac()
{
	with (document.forms[0])
	{
		target = window.top.name
		action = "facilidadeDet.asp"
		submit()
	}	
}

//Faz verificações do CPF/CNPJ
function VerificarCpfCnpj(str, tipo)
{
	//1 - CPF, 2-CNPJ, 3 - os 2
	var lbObject = (str == '[object]');
		
	if (tipo == 1)
	{
		if (!ValidarCPF(str))
		{
			if (lbObject)
			{
				alert("CPF "+ str.value + " inválido.");
				str.focus();
			}
			return false; 
		}
		return true;
	}

	else if (tipo == 2)
	{
		if (!ValidarCNPJ(str.value))
		{
			if (lbObject) 
			{
				alert("CNPJ "+ str.value + " inválido.");
				str.focus();
			}
			return false; 
		}
		return true;
	}
	
	else if(tipo == 3)
	{
		if (!ValidarCPF(str) && !ValidarCNPJ(str))
		{
			if (lbObject) 
			{
				alert("CPF/CNPJ "+ str.value + " inválido.");
				str.focus();
			}
			return false; 
		}
		return true;
	}

}

//Valida CPF
function  ValidarCPF(psCPF) 
{
	var lsAux    = '';
	var lsCPF    = '';
	var liPeso   = 2;
	var liSoma   = 0;
	var liTemp   = 0;
	var liDigito = 0;
	
	if (psCPF == '[object]')
		lsCPF = psCPF.value
	else
		lsCPF = psCPF;
	
	var liPos  = 0;

	for (liPos = lsCPF.length - 1; liPos >= 0; liPos--)
		if (!isNaN(lsCPF.charAt(liPos)))
			lsAux = lsCPF.charAt(liPos) + lsAux;
			
	for (liPos = lsAux.length - 3; liPos >= 0; liPos--)
	{
		liSoma += parseInt(lsCPF.charAt(liPos)) * liPeso;
		liPeso++;
	}
	
	liTemp   = (liSoma % 11);
	liDigito = IIf((liTemp < 2), 0, (11 - liTemp));
	
	if (parseInt(lsCPF.charAt(lsCPF.length - 2)) != liDigito)
		return (false);
	
	liPeso = 2;
	liSoma = 0;

	for (liPos = lsAux.length - 2; liPos >= 0; liPos--)
	{
		liSoma += parseInt(lsCPF.charAt(liPos)) * liPeso;
		liPeso++;
	}
	
	liTemp   = (liSoma % 11);
	liDigito = IIf((liTemp < 2), 0, (11 - liTemp));
	
	if (parseInt(lsCPF.charAt(lsCPF.length - 1)) != liDigito)
		return (false);
	
	return (true);
}

//Valida CNPJ
function ValidarCNPJ(psCNPJ) 
{
	var liPeso = 2;
	var liSoma = 0;
	var lsAux  = '';
	var liTemp = 0;
	var liDigito = 0;
	var lsCNPJ = '';
	
	if (psCNPJ == '[object]')
		lsCNPJ = psCNPJ.value
	else
		lsCNPJ = psCNPJ;
	
	var liPos  = 0;

	for (liPos = lsCNPJ.length - 1; liPos >= 0; liPos--)
		if (!isNaN(lsCNPJ.charAt(liPos)))
			lsAux = lsCNPJ.charAt(liPos) + lsAux;
			
	for (liPos = lsAux.length - 3; liPos >= 0; liPos--)
	{
		liSoma += parseInt(lsCNPJ.charAt(liPos)) * liPeso;
		liPeso  = IIf((liPeso == 9), 2, liPeso + 1);
	}
	
	liTemp   = (liSoma % 11);
	liDigito = IIf((liTemp < 2), 0, (11 - liTemp));
	
	if (parseInt(lsCNPJ.charAt(lsCNPJ.length - 2)) != liDigito)
		return (false);
	
	liPeso = 2;
	liSoma = 0;

	for (liPos = lsAux.length - 2; liPos >= 0; liPos--)
	{
		liSoma += parseInt(lsCNPJ.charAt(liPos)) * liPeso;
		liPeso  = IIf((liPeso == 9), 2, liPeso + 1);
	}
	
	liTemp   = (liSoma % 11);
	liDigito = IIf((liTemp < 2), 0, (11 - liTemp));
	
	if (parseInt(lsCNPJ.charAt(lsCNPJ.length - 1)) != liDigito)
		return (false);
	
	return (true);
}

//Imprementa a função Iff no JS
function IIf(pbCond, pvValueTrue, pvValueFalse)
{
	if (pbCond)
		return (pvValueTrue)
	else
		return (pvValueFalse);
}

//Valida um determinado tipo de dado
function ValidarTipoInfo(obj,intTipo,strMsg)
{
	var strValor
	if (obj == '[object]')
		strValor = obj.value
	else
		strValor = obj;

	switch (parseInt(intTipo))
	{
		case 0: //Númerico
			if (!IsNumber(strValor))
			{
				alert(strMsg+" com tipo inválido.")
				if (obj == '[object]') obj.focus()
				return false;
			}
			break
		case 1: //Data
			if (!IsDate(strValor))
			{
				alert(strMsg+" inválida.")
				if (obj == '[object]') obj.focus()
				return false;
			}
			break
		case 2: //Cep
			if (!IsCep(strValor))
			{
				alert(strMsg+" fora do padrão.")
				if (obj == '[object]') obj.focus()
				return false;
			}
			break
	}
	return (true);
}

//Valida tipo numérico
function IsNumber(strNumber)
{
	var strInvalidChar = IsEmpty(strNumber);
	
	for (var i = 0; i < strNumber.length; i++)
	{
		var strChar = strNumber.charAt(i);
		if (strChar != "." && strChar != "," && strChar != "-")
			if (isNaN(parseInt(strChar)))
				strInvalidChar = true  || strInvalidChar
			else
				strInvalidChar = false || strInvalidChar;
	}
	
	return (!strInvalidChar);
}

//Valida tipo data
function IsDate(strData)
{
	var objDiasMes = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
	
	if (strData == "" || IsEmpty(strData)) return true
	var objData    = strData.split('/');

	if (objData.length != 3)
		return false;
	
	if (!(objData[1] >= 1 && objData[1] <= 12))
		return false;

	if (objData[2].length != 4 || objData[2] < 1900) 
		return false;

	if (Math.floor(objData[2] / 4) * 4 == objData[2]) 
		objDiasMes[1] = 29;

	for(var intCont = 0; intCont < objData.length; intCont++)
	{
		if (IsEmpty(objData[intCont]))
			return (false);
		if (!IsNumber(objData[intCont]))
			return (false)
		else if (parseFloat(objData[intCont]) < 0)
			return (false);
	}

	if (!((objData[0] >= 1) && (objData[0] <= objDiasMes[objData[1] - 1])))
		return (false);

	return (true);
}

//Valida tipo CEP
function IsCep(strCep)
{
	var objCep = strCep.split('-');

	if (objCep.length != 2)
		return false;
	
	if (objCep[0].length != 5) 
		return false;

	if (objCep[1].length != 3) 
		return false;

	return (true);
}

//Implementa Mid no js
function Mid(strValor,intIni,intTam)
{
	var strValorAux = ""
	if (intIni > 0) intIni = parseInt(intIni)
	if (intTam > 0) intTam = parseInt(intIni+intTam-1)

	for (var intIndex = 0;intIndex<strValor.length;intIndex++)
	{
		if (intIndex >= intIni && intIndex <= intTam)
		{
			strValorAux += strValor.charAt(intIndex);
		}
	}
	
	return strValorAux
	
}

//Faz maxlength em run-time para campos de textbos e scrollbox
function MaxLength(poObject, plSize)
{
	if (event.keyCode > 46)
		event.returnValue = (poObject.value.length < plSize);
	
}

//Formata tipo money
function FormatMoney(dblNumber, intDecimal)
{
	var strFormat = new String("")
	var strFormatDecimal = new String()
	var strFinal = new String("")
	var strNumber = new String(dblNumber)
	var intDecimais = 1

	strNumber		 = Replace(strNumber,".","")
	strNumber		 = Replace(strNumber,",","")
	
	strFormatDecimal = strNumber.substring(strNumber.length-intDecimal,strNumber.length)
	strNumber		 = strNumber.substring(0,strNumber.length-intDecimal)
	if (strNumber == "") strNumber = "0" 

	for (var intIndex=strNumber.length-1;intIndex>=0;intIndex--)
	{
		if ((intDecimais%3) == 0 && intIndex != 0)
		{
			strFormat = strFormat + strNumber.charAt(intIndex) + "."
		}	
		else
		{
			strFormat = strFormat + strNumber.charAt(intIndex)
		}
		intDecimais = intDecimais + 1
	}

	
	for (var intIndex=strFormat.length-1;intIndex>=0;intIndex--)
	{
		strFinal = strFinal + strFormat.charAt(intIndex)
	}

	if (strFormatDecimal != "0" && strFinal != "0" )
	{
		return strFinal	+ "," + strFormatDecimal
	}
	else
	{
		return "0,00"
	}	
}

//Mostra modal como um alert e botões em português
function alertbox(){
	return window.showModalDialog("../INC/alertbox.asp",arguments,"dialogHeight: 119px; dialogWidth: 318px; edge: Raised; center: Yes; help: No; resizable: No; status: No; scroll: No;");
}

//Resgata o padrão de designação do serviço
function ResgatarServico(obj)
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarPadraoServico"
		hdnCboServico.value = obj.value
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}	
}


//Monta a designação do serviço
function MontarDesigServico()
{
	try //Pode não existir esses campos
	{
		if (arguments.length > 0){
			var blnValidarObrig = arguments[0]
		}else{
			var blnValidarObrig = false
		}
		with (document.forms[0])
		{
			hdnDesigServ.value = ""
			if (!txtPadrao.length)
			{
				if (txtPadrao.blnObrig == 1 && IsEmpty(txtPadrao.value) && blnValidarObrig){
					alert("Padrão de Designação do Serviço é obrigatório.")
					txtPadrao.focus()
					return false
				}
				else{
					if (!IsEmpty(txtPadrao.value)){
						hdnDesigServ.value = txtPadrao.value
					}else{
						for(var intIndexII=0;intIndexII<txtPadrao.size;intIndexII++){
							hdnDesigServ.value += "$"
						}	
					}	
				}
			}
			else{
				for (var intIndex=0;intIndex<txtPadrao.length;intIndex++){
					if (txtPadrao[intIndex].blnObrig == 1 && IsEmpty(txtPadrao[intIndex].value) && blnValidarObrig){
						alert("O "+parseInt(intIndex+1)+"º item do Padrão de Designação do Serviço é obrigatório.")
						txtPadrao[intIndex].focus()
						return false
					}
					else{
						if (!IsEmpty(txtPadrao[intIndex].value)){
							hdnDesigServ.value += txtPadrao[intIndex].value
						}else{
							for(var intIndexII=0;intIndexII<txtPadrao[intIndex].size;intIndexII++){
								hdnDesigServ.value += "$"
							}
						}	
					}	
				}
			}	
		}
	}catch (e){}	
	return true
}	

//Adicionar um node a um Xml
function AdicionarNode(objXML,strNomeNode,strValorNode)
{	
	var objElemento
	var objNodeFilho
	var intIndex
	var objNodeList
	
    if (objXML.xml == "")
    {
	   objXML.loadXML("<xDados></xDados>")
	}   
	//Verifica se já existe
	objNodeList = objXML.selectNodes("*/" + strNomeNode)
			
	if (objNodeList.length == 0)
	{
		//Cria
		if (strValorNode != ""){
			objNodeFilho = objXML.createNode("element", strNomeNode, "")
			objNodeFilho.text = strValorNode
			objXML.documentElement.appendChild (objNodeFilho)
		}	
	}	
	else
	{
		//Atualiza
		if (strValorNode != ""){
			objNodeList.item(0).text = strValorNode
		}else{//Se esta em branco remove o node para ir null para o SQL
			objNodeList.item(0).parentNode.removeChild(objNodeList.item(0)) 
		}	
	}	
}

//Resgata valor de um determminado node em um xml
function RequestNode(objXML,strNomeNode)
{
	var NodeList = objXML.selectNodes("*/" + strNomeNode)
	if (NodeList.length != 0)
	{
		return NodeList.item(0).text
		
	}	
	else
	{
		return ""		
	}
}

//Remove um node de um xml
function RemoverNode(objXml,strNomeNode,strValorNode)
{
	var objNodeCampos = objXml.selectNodes("//xDados["+strNomeNode+"='"+strValorNode+"']")
	if (parseInt("0" + objNodeCampos.length) > 0)
	{
		for (var intIndex=0;intIndex<objNodeCampos.length;intIndex++)
		{
			objNodeCampos.item(intIndex).parentNode.removeChild(objNodeCampos.item(intIndex)) 
		}	
	}
}

//Variável gerérica para controlar se algum checkbos/radio foi selecionada de uma lista grande
//isso não exige rodar toda a lista
var intSel = 0
//Incrementa a variável de controle da lista de checkbox/radio
function AddSelecaoChk(obj)
{
	if (obj.checked)
	{
		intSel += 1
	}
	else
	{
		intSel -= 1
	}
}

//Confimar a remoção de um item
function ConfirmarRemocao()
{
	if (intSel == 0)
	{
		alert("Selecione um item!")
		return false
	}
	else
	{
		if (window.confirm('Deseja excluir os registros selecionados ?') == false)
		{
			return false;
		}
		else
		{
			return true;
		}
	}
}

//Seleciona todos os checkbox para a exclusão nos cadastros básicos
function seleciona_tudo()
{
	for (var intIndex=0;intIndex<document.forms[0].elements.length;intIndex++)
	{
		var elemento = document.forms[0].elements[intIndex];
		if (elemento.name != 'excluirtudo' && !elemento.disabled)
			elemento.checked = document.forms[0].excluirtudo.checked;
	}
}
//Implemte o hint para objtos HTML normalmente utilizados no onmouseover
if (!document.layers&&!document.all)
event="test"

//Implemeta hint
function showtip(current,e,text){

if (document.all){
	thetitle=text.split('<br>')
	if (thetitle.length>1){
		thetitles=''
	for (i=0;i<thetitle.length;i++)
		thetitles+=thetitle[i]
		current.title=thetitles
}
else
	current.title=text
}

else if (document.layers){
	document.tooltip.document.write('<layer bgColor="white" style="border:1px solid black;font-size:12px;">'+text+'</layer>')
	document.tooltip.document.close()
	document.tooltip.left=e.pageX+5
	document.tooltip.top=e.pageY+5
	document.tooltip.visibility="show"
}

}
//Implementa hint
function hidetip(){
if (document.layers)
	document.tooltip.visibility="hidden"
}

//Implementa Right no JS
function Right(strIn, intlen)
{
	return (strIn.substr(strIn.length - intlen, intlen));
}

//Envia para a tela de impressão o dados são enviados por parâmetros no HTML
function TelaImpressao(intWidth,intHeight,strLabelCons)
{
	with (document.forms[0])
	{
		var objAryPram = new Array()
		objAryPram[0] = strLabelCons
		objAryPram[1] = hdnXls[0].value
		strRet = window.showModalDialog("Impressao.asp",objAryPram,"dialogHeight: "+intHeight+"px; dialogWidth: "+intWidth+"px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
	}	
}

//Função que Chama a página de cadastro de OS.
function AbrirCadastroOS(AcfID,PedID,ProID,EnvioEmail)
{
	with (document.forms[0])
	{
		strRet = window.showModalDialog("CadastraOS.asp?Acf=" + AcfID + "&Ped=" + PedID + "&Pro=" + ProID,"_blank","dialogHeight: 50px; dialogWidth: 500px; edge: Raised; center: Yes; help: No; resizable: No; status: No; scroll: No;");
		if(strRet == "true")
			parent.parent.window.ProcurarCadastroOS();
	}
}

//Função que Chama a página de alteração de Acesso a terceiros.
function AbrirCadastroSenha(AprovID,Utilizado)
{
	with (document.forms[0])
	{
			hdnAprovID.value = AprovID
			hdnAprov_Utilizado.value = Utilizado
			target = "_blank"
			action = "AlterarAutorizarAcesso.asp"
			submit()
	}
}



//Função que Chama a página de Inclusão de Acesso a terceiros.
function IncluirSenha()
{
	with (document.forms[0])
	{
			//var strNome = "Autorização de Acessos Terceiros"			
			//var objJanela = window.open()
			//objJanela.name =  strNome
			target = "_self"
			action = "IncluirAutorizarAcesso.asp"
			submit()
	}
}

//Envia para a planilha Excel
function AbrirXls()
{
	with (document.forms[0])
	{
		BreakItUp()
		target = "_blank"
		action = "ExcelExport.asp"
		submit()
	}
}


//Envia SQL para gerar a planilha Excel : PRSS - 12/01/2006
function AbrirXlsRecebe()
{
	with (document.forms[0])
	{
		BreakItUp()
		target = "_blank"
		action = "ExcelExportRecebe_CartasProvedor.asp"
		submit()
	}
}


function AbrirXlsAcesso()
{
		//alert(1)
	with (document.forms[0])
	{
		//	alert(3)
		//BreakItUp()
		target = "_blank"
		action = "ExcelExportAcesso.asp"
		submit()
	}
}

//Quebra campos que tem  seus valores maiores que 102399k pois é o máximo que o submit permite
function BreakItUp()
{
  //Set the limit for field size.
  var FormLimit = 202399

  //Get the value of the large input object.
  var TempVar = new String
  TempVar = document.forms[0].hdnXls[0].value

  if (TempVar.length > FormLimit)
  {
	with (document.forms[0]){
		for(var intIndex=0;intIndex<hdnXls.length;intIndex++){
			hdnXls[intIndex].value = ""
		}
		hdnXls[0].value = TempVar.substr(0, FormLimit)
		TempVar = TempVar.substr(FormLimit)
		//Limpa o array de Xls
		intIndex = 1
		while (TempVar.length > 0){
			hdnXls[intIndex].value = TempVar.substr(0, FormLimit)
			intIndex += 1
			TempVar = TempVar.substr(FormLimit) 
		}	
	}	
  }
  else{
	with (document.forms[0]){
		for(var intIndex=1;intIndex<hdnXls.length;intIndex++){
			hdnXls[intIndex].value = ""
		}
	}
  }
}
//Seta focus()
function setarFocus(strNomeCampo){
	eval("document.forms[0]."+strNomeCampo+".focus()")
}
//Valida um determinado domínio
function SearchDom(objSearch,strDom){

	var strSearch = new String(objSearch.value)
	if (strSearch == "") return true
	var objAryDom = strDom.toUpperCase().split(",")

	for (intIndex=0;intIndex<objAryDom.length;intIndex++){
	 	if (objAryDom[intIndex] == strSearch.toUpperCase()){
			return true
	 	}
	}
	alert("Valor fora do domínio ("+strDom+").")
	objSearch.value = ""
	objSearch.focus()
	return false
}

function ListarPedidosSolicitacao(dblSolId)
{
	with(document.forms[0])
	{
		var objAry = new Array(cboStatusSolic.value,txtMotivo.value)
		var intRet = window.showModalDialog("../Asp/ListaPedidoSolicitacao.asp?hdnSolId="+dblSolId,objAry,"dialogHeight: 200px; dialogWidth: 450px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
		if (intRet == 1){
			intRet = 0
		}
		IFrmLista.location.href = 'ProcessoMotivoPendencia.asp?strAcao=ResgatarLista&dblSolId=' + dblSolId
	}
}

function ListarPedidosHistoricoSolicitacao(dblSolId)
{
	with(document.forms[0])
	{
		var objAry = new Array(cboStatusSolic.value,txtMotivo.value)
		var intRet = window.showModalDialog("../Asp/ListaPedidoHistoricoSolicitacao.asp?hdnSolId="+dblSolId,objAry,"dialogHeight: 200px; dialogWidth: 350px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
		if (intRet == 1){
			intRet = 0
		}
		IFrmLista.location.href = 'ProcessoMotivoPendencia.asp?strAcao=ResgatarLista&dblSolId=' + dblSolId
	}
}

function AtualizarListaMotivo()
{
	with(document.forms[0])
	{
		try{
			IFrmLista.location.href = "ProcessoMotivoPendencia.asp?strAcao=ResgatarLista&dblSolId='" + hdnSolId.value + "'&dblPedId='" + hdnPedId.value + "'"
		}catch(e){}	
	}
}
//Valida formato do DM (DM-NNNNN/YYYY)
function ValidarDM(obj)
{
	if (obj == '[object]') var strDM = new String(obj.value)
	else var strDM = String(obj)

	if (strDM.length > 3)
	{
		if (strDM.length != 13){
			alert("O campo Pedido de Acesso não foi preenchido corretamente.");
			try{
				if (obj == '[object]') obj.focus();
			}catch(e){}	
			return false;
		}
		if (strDM.substr(2,1) != "-"){
			alert("O campo Pedido de Acesso não foi preenchido corretamente.");
			try{
				if (obj == '[object]') obj.focus();
			}catch(e){}	
			return false;
		}
			if (strDM.substr(8,1) != "/") {
			alert("O campo Pedido de Acesso não foi preenchido corretamente.");
			try{
				if (obj == '[object]') obj.focus();
			}catch(e){}	
			return false;
		}
		var objAry = strDM.split("/") 

		if (isNaN(objAry[1])){
			alert("O campo Pedido de Acesso não foi preenchido corretamente(Ano inválido).");
			try{
				if (obj == '[object]') obj.focus();
			}catch(e){}	
			return false;
		}
		var objAry2 = objAry[0].split("-") 

		if (isNaN(objAry2[1])){
			alert("O campo Pedido de Acesso não foi preenchido corretamente(Sequência inválido).");
			try{
				if (obj == '[object]') obj.focus();
			}catch(e){}	
			return false;
		}

	}
	return true
}

//Formata um Xml com os dados do form
function PopularXml()
{
	if (arguments.length > 0){
		with (document.forms[0]){
			hdnXmlReturn.value = ""
			for (var intIndex=0;intIndex<document.forms[0].elements.length;intIndex++)
			{
				var elemento = document.forms[0].elements[intIndex];
				if (elemento.type != 'button'){
					AdicionarNode(arguments[0],elemento.name,elemento.value)
				}	
			}
			try{//tenta adionar o padrão de designação
			for (var intIndex=0;intIndex<txtPadrao.length;intIndex++)
			{
				if (!IsEmpty(txtPadrao[intIndex].value))
				{
					AdicionarNode(arguments[0],"txtPadrao_"+intIndex,txtPadrao[intIndex].value)
				}
			}}catch(e){}
			hdnXmlReturn.value = arguments[0].xml
		}

	}
	else{

		with (document.forms[0]){
			hdnXmlReturn.value = ""
			for (var intIndex=0;intIndex<document.forms[0].elements.length;intIndex++)
			{
				var elemento = document.forms[0].elements[intIndex];
				if (elemento.type != 'button'){
					AdicionarNode(objXmlGeral,elemento.name,elemento.value)
				}	
			}
			try{//tenta adionar o padrão de designação
			for (var intIndex=0;intIndex<txtPadrao.length;intIndex++)
			{
				if (!IsEmpty(txtPadrao[intIndex].value))
				{
					AdicionarNode(objXmlGeral,"txtPadrao_"+intIndex,txtPadrao[intIndex].value)
				}
			}}catch(e){}
			hdnXmlReturn.value = objXmlGeral.xml
		}
	}
}
//Recoloca os dados do Xml no Form
function PopularForm(){
	if (arguments.length > 0){
		var	objNode = arguments[0].selectNodes("//xDados")
	}
	else{
		var	objNode = objXmlGeral.selectNodes("//xDados")
	}
	//Refaz a lista de Ids no IFRAME
	for (var intIndex=0;intIndex<objNode[0].childNodes.length;intIndex++)
	{
		try{
			var strNomdeName = new String(objNode[0].childNodes[intIndex].nodeName)
			var objChildForm = new Object(eval("document.forms[0]."+strNomdeName))
			if (strNomdeName.indexOf("rdo") != -1 || strNomdeName.indexOf("chk") != -1){
				var intIndexSelected = RequestNodeFac(objXmlGeral,strNomdeName+"Index",objNode[0].childNodes[intIndex].text)
				eval("document.forms[0]."+strNomdeName+"["+parseInt(intIndexSelected)+"].checked = true")
			}
			else{
				if (strNomdeName.indexOf("txtPadrao_") != -1)
				{
					var objAry = strNomdeName.split("_")
					eval("document.forms[0]."+objAry[0]+"["+parseInt(objAry[1])+"].value='"+objNode[0].childNodes[intIndex].text+"'")
				}else{
					objChildForm.value = objNode[0].childNodes[intIndex].text
				}	
			}	
		}catch(e){}
	}	
}
function AlterarSolicitacao(obj,acao)
{
	with(document.forms[0])
	{
		if (obj == '[object]'){	hdnSolId.value = obj.value}
		else{hdnSolId.value = obj}
		try{
			PopularXml()
		}catch(e){}	
		hdnAcao.value = "AlteracaoCad"
		target = self.name
		if (acao == '[object]')
		  {action = "AlteracaoCad.asp"}
		else
		  {action = "AlteracaoCad.asp?acao=" + acao}
		submit()
	}
}

function DesCanSolicitacao(obj,acao)
{
	with(document.forms[0])
	{
		if (obj == '[object]'){	hdnSolId.value = obj.value}
		else{hdnSolId.value = obj}
		try{
			PopularXml()
		}catch(e){}	
		hdnAcao.value = "AlteracaoCad"
		target = self.name
		if (acao == '[object]')
		  {action = "AlteracaoCadAprov.asp"}
		else
		  {action = "AlteracaoCadAprov.asp?acao=" + acao}
		submit()
	}
}


function AvaliarSolicitacao(obj)
{
	with(document.forms[0])
	{
		if (obj == '[object]'){	hdnSolId.value = obj.value}
		else{hdnSolId.value = obj}
		try{
			PopularXml()
		}catch(e){}	
		//alert('valida')
		hdnAcao.value = "AlteracaoCad"
		target = self.name
		action = "AvaliarAcesso.asp"
		submit()
	}
}




function CompararData(strData1,strData2,intOper,strMsg)
{
	if (strData1 == '[object]')	var objDat1 = new String(strData1.value)
	else var objDat1 = new String(strData1)
	
	if (strData2 == '[object]')	var objDat2 = new String(strData2.value)
	else var objDat2 = new String(strData2)
	
	switch (intOper)
	{
		case 1: //1 < 2
			if (InverterData(objDat1) > InverterData(objDat2))
			{
				alert(strMsg)
				return false
			}
			break
		case 2://1>2
			if (InverterData(objDat1) < InverterData(objDat2))
			{
				alert(strMsg)
				return false
			}
			break
		case 3://1=2
			if (InverterData(objDat1) != InverterData(objDat2))
			{
				alert(strMsg)
				return false
			}
			break
	}
	return true
}
function InverterData(datData)
{
	if (datData != "" && datData.length == 10)
	{
		return datData.substring(6,10) + datData.substring(3,5) + datData.substring(0,2)
	}
	else
	{
		return ""
	}	
}

//Função AdicionarCNL para Manobra
function AdicionarCNLManobra(index)
{
	with (document.forms[0])
	{
		var strLocal = new String(txtLocalConfiguracao.value);
		if(index == 1)
		{
			if (rdoUrbano[1].checked)
			{
				if (strLocal != "")
				{
					txtCNLPontaA.value = strLocal.split(" ")[0]
					txtCNLPontaB.value = strLocal.split(" ")[0]
				}	
			}
			else
			{
				txtCNLPontaA.value = strLocal.split(" ")[0]
				txtCNLPontaB.value = ''
			}
		}
		else
		{
			if (rdoUrbano2[1].checked)
			{
				if (strLocal != "")
				{
					txtCNLPontaA2.value = strLocal.split(" ")[0]
					txtCNLPontaB2.value = strLocal.split(" ")[0]
				}	
			}
			else
			{
				txtCNLPontaA2.value = strLocal.split(" ")[0]
				txtCNLPontaB2.value = ''
			}		
		}
	}	
}

function AddManobra(FacID)
{
	var Linha;
	with (document.forms[0])
	{
		var strLocal = new String(txtLocalConfiguracao.value);
		var strUrbando;
		var strFatura;
		
		switch (txtRede.value) 
		{
			case "DETERMINISTICO": //Det
				var strTimeSlot = new String(txtTimeslot.value);
				
				if (!ValidarCampos(txtFila,"Fila")) return false
				if (!ValidarCampos(cboCodProv,"Código do provedor")) return false
				if (!ValidarCampos(txtBastidor,"Bastidor")) return false
				if (!ValidarCampos(txtRegua,"Posição")) return false
				if (!ValidarCampos(txtPosicao,"Porta")) return false
				if (!ValidarCampos(txtTimeslot,"Timeslot")) return false
				if(strTimeSlot.substring(strTimeSlot.length-1) == "-")
				{
					alert("Campo Timeslot fora do padrão");
					return false;
				}
				if(parseFloat(strTimeSlot.substring(strTimeSlot.length-4)) > 31) 
				{
					alert("O maior Timeslot possível é 31");
					return false
				}
				if(strTimeSlot.substring(4,5) ==  "-")
				{
					if(parseFloat(strTimeSlot.substring(0,4)) > 31) 
					{
						alert("O maior Timeslot possível é 31");
						return false
					}
					if(parseFloat(strTimeSlot.substring(0,4)) >= parseFloat(strTimeSlot.substring(strTimeSlot.length-4)))
					{
						alert("O Timeslot inicial deve ser menor que o final");
						return false
					}
				}
				if (rdoUrbano[1].checked)
				{
					if (strLocal != "")
					{
						txtCNLPontaA.value = strLocal.split(" ")[0]
						txtCNLPontaB.value = strLocal.split(" ")[0]
					}	
				}
				else
				{
					if (IsEmpty(txtCNLPontaA.value))
						txtCNLPontaA.value = strLocal.split(" ")[0]
				}
				if (!ValidarCampos(cboPropModem,"Proprietário do Modem")) return false
				if (!ValidarCampos(txtQtdeModem,"Quantidade de Modens")) return false
				try{
					if (!rdoFatura[0].checked && !rdoFatura[1].checked)
					{
						alert("Fatura é um Campo Obrigatório.")
						return false
					}
				}
				catch(e){}		
				if (!IsEmpty(txtNumAcessoPtaEbt.value))
				{	
					if (!ValidarPadraoProvedorManobra(1)) return false
				}	
				if(rdoUrbano[0].checked == true)
					strUrbano = 'I'
				else
					strUrbano = 'U'	
				if(rdoFatura[0].checked == true)
					strFatura = 'S'
				else
					strFatura = 'N'
						
				Linha = AchaHdnGridManobra();
				if(Linha != -1)
					AlteraLinhaManobraDet(FacID,strUrbano,strFatura,Linha)
					//document.getElementById('hdnDados' + Linha).value = cboCodProv.value + '&&' + txtNumAcessoPtaEbt.value + '&&' + txtFila.value + '&&' + txtBastidor.value + '&&' + txtRegua.value + '&&' + txtPosicao.value + '&&' + txtTimeslot.value + '&&' + txtCCTOProvedor.value + '&&' + txtNumAcessoCLI.value + '&&' + txtCNLPontaA.value + '&&' + txtCNLPontaB.value	+ '&&' + txtQtdeModem.value + '&&' +	cboPropModem.value + '&&' + txtLink.value + '&&' + txtAreaObs.value + '&&' + strUrbano + '&&' + strFatura + '&&' + txtLink.value;
				else
					AddLinhaManobraDet(strUrbano,strFatura);
				
				break //FIM DET

			case "NAO DETERMINISTICO": //NDet
				//var strTimeSlot = new String(txtTimeslot.value); //Caso não necessário remover;
				
				if (!ValidarCampos(cboCodProv2,"Código do provedor")) return false
				if (!ValidarCampos(txtTronco,"Tronco")) return false
				if (!ValidarCampos(txtPar2,"Par")) return false
				if (rdoUrbano2[1].checked)
				{
					if (strLocal != "")
					{
						txtCNLPontaA2.value = strLocal.split(" ")[0]
						txtCNLPontaB2.value = strLocal.split(" ")[0]
					}	
				}
				else
				{
					if (IsEmpty(txtCNLPontaA2.value))
						txtCNLPontaA2.value = strLocal.split(" ")[0]
				}
				if (!ValidarCampos(cboPropModem2,"Proprietário do Modem")) return false
				if (!ValidarCampos(txtQtdeModem2,"Quantidade de Modens")) return false
				try{
					if (!rdoFatura2[0].checked && !rdoFatura2[1].checked)
					{
						alert("Fatura é um Campo Obrigatório.")
						return false
					}
				}
				catch(e){}	
								
				if (!IsEmpty(txtNumAcessoPtaEbt2.value))
				{
					if (!ValidarPadraoProvedorManobra(2)) return false
				}	
				if(rdoUrbano2[0].checked == true)
					strUrbano = 'I'
				else
					strUrbano = 'U'
				if(rdoFatura2[0].checked == true)
					strFatura = 'S'
				else
					strFatura = 'N'
						
				Linha = AchaHdnGridManobra();
				
				if(Linha != -1)
					AlteraLinhaManobraNDet(FacID,strUrbano,strFatura,Linha)
					//document.getElementById('hdnDados' + Linha).value = cboCodProv2.value + '&&' + txtNumAcessoPtaEbt2.value + '&&' + txtTronco.value + '&&' + txtPar2.value + '&&' + txtCCTOProvedor2.value + '&&' + txtNumAcessoCLI2.value + '&&' + txtCNLPontaA2.value + '&&' + txtCNLPontaB2.value + '&&' + txtQtdeModem2.value + '&&' + cboPropModem2.value + '&&' + txtAreaObs2.value + '&&' +  strUrbano + '&&' + strFatura;
				else
					AddLinhaManobraNDet(strUrbano,strFatura)
					
				break //FIM NAO DET

			case "ADE": //ADE
				if (!ValidarCampos(txtCabo3,"Cabo")) return false
				if (!ValidarCampos(txtPar3,"Número do Cabo de Acesso")) return false
				if (!IsEmpty(txtNumAcesso3.value))
				{
					if (!ValidarPadraoProvedorManobra(3)) return false
				}	
				if (!ValidarCampos(cboPropModem3,"Proprietário do Modem")) return false
				if (!ValidarCampos(txtQtdeModem3,"Quantidade de Modens")) return false
				
				Linha = AchaHdnGridManobra();
				if(Linha != -1)
					AlteraLinhaManobraADE(FacID,Linha)
					//document.getElementById('hdnDados' + Linha).value = txtNumAcesso3.value + '&&' + txtCabo3.value + '&&' + txtPar3.value + '&&' + txtPADE3.value + '&&' + txtDerivacao3.value + '&&' + cboPropModem3.value + '&&' + cboTCabo3.value + '&&' + txtQtdeModem3.value + '&&' + txtAreaObs3.value;
				else
					AddLinhaManobraADE()
				
				break //FIM ADE
		
			}//FIM SWITCH

			var intRet=alertbox('Deseja permanecer com os dados?','Sim','Não','Sair')
			if (intRet == 3) //Sair
			{	
				document.getElementById('hdnDados'+ Linha).value = '';
				return false;
			}
			else if (intRet == 2) //Nao
				LimparCamposManobra(FacID);
				
	}//End WITH
				
}

function AchaHdnGridManobra()
{
	var i;
	var y=0;
	with (document.forms[0])
	{
		for(i = 0; i < document.getElementById("hdnQtdLinha").value;i++)
		{
			if(document.getElementById('hdnFacID' + i) == '[object]')
			{
				if(document.getElementById('tblFac').rows.length == 2)
				{
					if(rdoSelFacilidade.checked == true)
					{
						return i;	
					}
				}
				else
				{
					if(rdoSelFacilidade[y].checked == true)
					{
						return i;	
					}
				}
				y++;
			}
		}
	}	
	return -1;
}

function AchaLinhaGridManobra()
{
	var i;
	var y;
	var k;
	y = 0;
	k = 0;
	with (document.forms[0])
	{
		for(i = 0; i < document.getElementById("hdnQtdLinha").value;i++)
		{
			if(document.getElementById('hdnFacID' + i) == '[object]')
			{
				y++;
				if(document.getElementById('tblFac').rows.length == 2)
				{
					if(rdoSelFacilidade.checked == true)
					{
						return y;	
					}
				}
				else
				{
					if(rdoSelFacilidade[k].checked == true)
					{
						return y;	
					}
				}
				k++;
			}
		}
	}	
	return -1;
}

function RemoveLinhaManobra(FacID)
{
	var i;
	var tam;
	var y;
	y = 0;
	with (document.forms[0])
	{
		y = AchaLinhaGridManobra();
		if (y != -1)
		{
			document.getElementById('tblFac').deleteRow(y);
			document.getElementById('btnRemover').disabled = true;
			hdnDelete.value = hdnDelete.value + FacID + ',';
			
			if (alertbox('Deseja permanecer com os dados?','Sim','Não')==2)
					LimparCamposManobra();
		}
		return;
	}
}

function AddLinhaManobraDet(strUrbano,strFatura)
{
	var tam;
	with (document.forms[0])
	{
		tam = document.getElementById("hdnQtdLinha").value;
		var row = document.getElementById('tblFac').insertRow(document.getElementById('tblFac').rows.length)
		var cell = row.insertCell(0);
		cell.innerHTML =  '<td><input type=radio name=rdoSelFacilidade index=' + tam + 'class=radio onclick=ExibeManobraDet()> </td>'  
		var cell = row.insertCell(1);
		cell.innerHTML =  '<td>' + cboCodProv[cboCodProv.selectedIndex].text + '</td>'
		var cell = row.insertCell(2);
		cell.innerHTML =  '<td>' + txtNumAcessoPtaEbt.value + '</td>'
		var cell = row.insertCell(3);
		cell.innerHTML =  '<td>' + txtFila.value + '</td>'
		var cell = row.insertCell(4);
		cell.innerHTML =  '<td>' + txtBastidor.value + '</td>'
		var cell = row.insertCell(5);
		cell.innerHTML =  '<td>' + txtRegua.value + '</td>'
		var cell = row.insertCell(6);
		cell.innerHTML =  '<td>' + txtPosicao.value + '</td>'
		var cell = row.insertCell(7);
		cell.innerHTML =  '<td>' + txtTimeslot.value + '</td>'		
		var cell = row.insertCell(8);
		cell.innerHTML =  '<td>' + txtNumAcessoCLI.value + '</td>'	
		var cell = row.insertCell(9);
		cell.innerHTML =  '<td><input type=hidden name=hdnFacID' + tam + '></td>'
		var cell = row.insertCell(10);
		cell.innerHTML =  '<td><input type=hidden name=hdnDados' + tam + '></td>'
		document.getElementById('hdnDados' + tam).value = cboCodProv.value + '&&' + txtNumAcessoPtaEbt.value + '&&' + txtFila.value + '&&' + txtBastidor.value + '&&' + txtRegua.value + '&&' + txtPosicao.value + '&&' + txtTimeslot.value + '&&' + txtCCTOProvedor.value + '&&' + txtNumAcessoCLI.value + '&&' + txtCNLPontaA.value + '&&' + txtCNLPontaB.value	+ '&&' + txtQtdeModem.value + '&&' +	cboPropModem.value + '&&' + txtLink.value + '&&' + txtAreaObs.value + '&&' + strUrbano + '&&' + strFatura + '&&' + txtLink.value;
		document.getElementById("hdnQtdLinha").value++;
	}
}

function AlteraLinhaManobraDet(FacID,strUrbano,strFatura,tam)
{
	var tam;
	var y;
	with (document.forms[0])
	{
		y = AchaLinhaGridManobra()
		var rows  = document.getElementById('tblFac').rows
		//rows[y].cells[0].innerHTML = '<td><input type=radio name=rdoSelFacilidade index=' + tam + 'class=radio onclick=ExibeManobraDet()> </td>';
		rows[y].cells[1].innerHTML = '<td>' + cboCodProv[cboCodProv.selectedIndex].text + '</td>'
		rows[y].cells[2].innerHTML = '<td>' + txtNumAcessoPtaEbt.value + '</td>'
		rows[y].cells[3].innerHTML = '<td>' + txtFila.value + '</td>'
		rows[y].cells[4].innerHTML = '<td>' + txtBastidor.value + '</td>'
		rows[y].cells[5].innerHTML = '<td>' + txtRegua.value + '</td>'
		rows[y].cells[6].innerHTML = '<td>' + txtPosicao.value + '</td>'
		rows[y].cells[7].innerHTML = '<td>' + txtTimeslot.value + '</td>'	
		rows[y].cells[8].innerHTML = '<td>' + txtNumAcessoCLI.value + '</td>'	
		rows[y].cells[9].innerHTML = '<td><input type=hidden name=hdnFacID' + tam + ' value=' + FacID + '></td>'
		rows[y].cells[10].innerHTML = '<td><input type=hidden name=hdnDados' + tam + '></td>'
		document.getElementById('hdnDados' + tam).value = cboCodProv.value + '&&' + txtNumAcessoPtaEbt.value + '&&' + txtFila.value + '&&' + txtBastidor.value + '&&' + txtRegua.value + '&&' + txtPosicao.value + '&&' + txtTimeslot.value + '&&' + txtCCTOProvedor.value + '&&' + txtNumAcessoCLI.value + '&&' + txtCNLPontaA.value + '&&' + txtCNLPontaB.value	+ '&&' + txtQtdeModem.value + '&&' +	cboPropModem.value + '&&' + txtLink.value + '&&' + txtAreaObs.value + '&&' + strUrbano + '&&' + strFatura + '&&' + txtLink.value;
	}
}

function AddLinhaManobraNDet(strUrbano,strFatura)
{
	var tam;
	with (document.forms[0])
	{
		tam = document.getElementById('tblFac').rows.length;
		tam--;
		var row = document.getElementById('tblFac').insertRow(-1)
		var cell = row.insertCell(0);
		cell.innerHTML =  '<td><input type=radio name=rdoSelFacilidade index=' + tam + 'class=radio onclick=ExibeManobraNaoDet()> </td>'  //\''+ hdnProvedor2.value + '\',\'' + txtNumAcessoPtaEbt2.value + '\' <%=objRSPag("Fac_Tronco")%>','<%=strPar%>','<%=objRSPag("Acf_NroAcessoPtaCli")%>','<%=objRSPag("Acf_NroAcessoCCTOProvedor")%>','<%=objRSPag("Acf_CCTOTipo")%>','<%=objRSPag("Acf_CnlPTA")%>','<%=objRSPag("Acf_CnlPTB")%>','<%=objRSPag("Acf_ProprietarioEquip")%>','<%=objRSPag("Acf_QtdEquip")%>','<%=objRSPag("Acf_CCTOFatura")%>','<%=objRSPag("Acf_Obs")%>','<%=objRSPag("Fac_ID")%>','<%=strQtd%>');ResgatarPadraoProvedor(<%=strPro%>)"></td>
		var cell = row.insertCell(1);
		cell.innerHTML =  '<td>' + cboCodProv2[cboCodProv2.selectedIndex].text + '</td>'
		var cell = row.insertCell(2);
		cell.innerHTML =  '<td>' + txtNumAcessoPtaEbt2.value + '</td>'
		var cell = row.insertCell(3);
		cell.innerHTML =  '<td>' + txtTronco.value + '</td>'
		var cell = row.insertCell(4);
		cell.innerHTML =  '<td>' + txtPar2.value + '</td>'
		var cell = row.insertCell(5);
		cell.innerHTML =  '<td>' + txtNumAcessoCLI2.value + '</td>'
		var cell = row.insertCell(6);
		cell.innerHTML =  '<td><input type=hidden name=hdnFacID' + tam + '></td>'
		var cell = row.insertCell(7);
		cell.innerHTML =  '<td><input type=hidden name=hdnDados' + tam + '></td>'
		document.getElementById('hdnDados' + tam).value = cboCodProv2.value + '&&' + txtNumAcessoPtaEbt2.value + '&&' + txtTronco.value + '&&' + txtPar2.value + '&&' + txtCCTOProvedor2.value + '&&' + txtNumAcessoCLI2.value + '&&' + txtCNLPontaA2.value + '&&' + txtCNLPontaB2.value + '&&' + txtQtdeModem2.value + '&&' + cboPropModem2.value + '&&' + txtAreaObs2.value + '&&' +  strUrbano + '&&' + strFatura;
		document.getElementById("hdnQtdLinha").value++;
	}
}

function AlteraLinhaManobraNDet(FacID,strUrbano,strFatura,tam)
{
	var tam;
	var y;
	with (document.forms[0])
	{
		y = AchaLinhaGridManobra()
		var rows  = document.getElementById('tblFac').rows
		rows[y].cells[1].innerHTML = '<td>' + cboCodProv2[cboCodProv2.selectedIndex].text + '</td>'
		rows[y].cells[2].innerHTML = '<td>' + txtNumAcessoPtaEbt2.value + '</td>'
		rows[y].cells[3].innerHTML = '<td>' + txtTronco.value + '</td>'
		rows[y].cells[4].innerHTML = '<td>' + txtPar2.value + '</td>'
		rows[y].cells[5].innerHTML = '<td>' + txtNumAcessoCLI2.value + '</td>'
		rows[y].cells[6].innerHTML = '<td><input type=hidden name=hdnFacID' + tam + ' value=' + FacID +'></td>'
		rows[y].cells[7].innerHTML = '<td><input type=hidden name=hdnDados' + tam + '></td>'
		document.getElementById('hdnDados' + tam).value = cboCodProv2.value + '&&' + txtNumAcessoPtaEbt2.value + '&&' + txtTronco.value + '&&' + txtPar2.value + '&&' + txtCCTOProvedor2.value + '&&' + txtNumAcessoCLI2.value + '&&' + txtCNLPontaA2.value + '&&' + txtCNLPontaB2.value + '&&' + txtQtdeModem2.value + '&&' + cboPropModem2.value + '&&' + txtAreaObs2.value + '&&' +  strUrbano + '&&' + strFatura;
	}
}

function AddLinhaManobraADE(strUrbano,strFatura)
{
	var tam;
	with (document.forms[0])
	{
		tam = document.getElementById('tblFac').rows.length;
		tam--;
		var row = document.getElementById('tblFac').insertRow(-1)
		var cell = row.insertCell(0);
		cell.innerHTML =  '<td><input type=radio name=rdoSelFacilidade index=' + tam + 'class=radio onclick=ExibeManobraADE()> </td>'  //\''+ hdnProvedor2.value + '\',\'' + txtNumAcessoPtaEbt2.value + '\' <%=objRSPag("Fac_Tronco")%>','<%=strPar%>','<%=objRSPag("Acf_NroAcessoPtaCli")%>','<%=objRSPag("Acf_NroAcessoCCTOProvedor")%>','<%=objRSPag("Acf_CCTOTipo")%>','<%=objRSPag("Acf_CnlPTA")%>','<%=objRSPag("Acf_CnlPTB")%>','<%=objRSPag("Acf_ProprietarioEquip")%>','<%=objRSPag("Acf_QtdEquip")%>','<%=objRSPag("Acf_CCTOFatura")%>','<%=objRSPag("Acf_Obs")%>','<%=objRSPag("Fac_ID")%>','<%=strQtd%>');ResgatarPadraoProvedor(<%=strPro%>)"></td>
		var cell = row.insertCell(1);
		cell.innerHTML =  '<td>' + txtNumAcesso3.value + '</td>'
		var cell = row.insertCell(2);
		cell.innerHTML =  '<td>' + txtCabo3.value + '</td>'
		var cell = row.insertCell(3);
		cell.innerHTML =  '<td>' + txtPar3.value + '</td>'
		var cell = row.insertCell(4);
		cell.innerHTML =  '<td>' + txtPADE3.value + '</td>'
		var cell = row.insertCell(5);
		cell.innerHTML =  '<td>' + txtDerivacao3.value + '</td>'
		var cell = row.insertCell(6);
		cell.innerHTML =  '<td>' + cboTCabo3.value + '</td>'
		var cell = row.insertCell(7);
		cell.innerHTML =  '<td><input type=hidden name=hdnFacID' + tam + '></td>'
		var cell = row.insertCell(8);
		cell.innerHTML =  '<td><input type=hidden name=hdnDados' + tam + '></td>'
		document.getElementById('hdnDados' + tam).value = txtNumAcesso3.value + '&&' + txtCabo3.value + '&&' + txtPar3.value + '&&' + txtPADE3.value + '&&' + txtDerivacao3.value + '&&' + cboPropModem3.value + '&&' + cboTCabo3.value + '&&' + txtQtdeModem3.value + '&&' + txtAreaObs3.value;
		document.getElementById("hdnQtdLinha").value++;
	}
}

function AlteraLinhaManobraADE(FacID,tam)
{
	var tam;
	var y;
	with (document.forms[0])
	{
		y = AchaLinhaGridManobra()
		var rows  = document.getElementById('tblFac').rows
		rows[y].cells[1].innerHTML = '<td>' + txtNumAcesso3.value + '</td>'
		rows[y].cells[2].innerHTML = '<td>' + txtCabo3.value + '</td>'
		rows[y].cells[3].innerHTML = '<td>' + txtPar3.value + '</td>'
		rows[y].cells[4].innerHTML = '<td>' + txtPADE3.value + '</td>'
		rows[y].cells[5].innerHTML = '<td>' + txtDerivacao3.value + '</td>'
		rows[y].cells[6].innerHTML = '<td>' + cboTCabo3.value + '</td>'
		rows[y].cells[7].innerHTML = '<td><input type=hidden name=hdnFacID' + tam + ' value=' + FacID +'></td>'
		rows[y].cells[8].innerHTML = '<td><input type=hidden name=hdnDados' + tam + '></td>'
		document.getElementById('hdnDados' + tam).value = txtNumAcesso3.value + '&&' + txtCabo3.value + '&&' + txtPar3.value + '&&' + txtPADE3.value + '&&' + txtDerivacao3.value + '&&' + cboPropModem3.value + '&&' + cboTCabo3.value + '&&' + txtQtdeModem3.value + '&&' + txtAreaObs3.value;
	}
}


function LimparCamposManobra(FacID)
{
	var Linha;
	with (document.forms[0]) 
	{
		if (txtRede.value != "")
		{
			switch (txtRede.value) 
			{
				case "DETERMINISTICO": //Det
					cboCodProv.value = ''
					txtNumAcessoPtaEbt.value = ''
					txtFila.value = ''
					txtBastidor.value = ''
					txtRegua.value = ''
					txtPosicao.value = ''
					txtTimeslot.value = ''
					txtCCTOProvedor.value = ''
					txtNumAcessoCLI.value = ''
					txtCNLPontaA.value = ''
					txtCNLPontaB.value = ''
					txtQtdeModem.value = ''
					cboPropModem.value = ''
					txtLink.value = ''
					txtAreaObs.value = ''	
					for (var intIndex=0;intIndex<rdoUrbano.length;intIndex++){
						rdoUrbano[intIndex].checked = false
					}	
					try{
					for (var intIndex=0;intIndex<rdoFatura.length;intIndex++){
						rdoFatura[intIndex].checked = false
					}}catch(e){}
					break

				case "NAO DETERMINISTICO": //NDet
					cboCodProv2.value = ''
					txtNumAcessoPtaEbt2.value = ''
					txtTronco.value = ''
					txtPar2.value = ''
					txtCCTOProvedor2.value = ''
					txtNumAcessoCLI2.value = ''
					txtCNLPontaA2.value = ''
					txtCNLPontaB2.value = ''
					txtQtdeModem2.value = ''
					cboPropModem2.value = ''
					txtAreaObs2.value = ''
					for (var intIndex=0;intIndex<rdoUrbano2.length;intIndex++)
						rdoUrbano2[intIndex].checked = false
					try
					{
						for (var intIndex=0;intIndex<rdoFatura2.length;intIndex++)
							rdoFatura2[intIndex].checked = false
					}
					catch(e){}
					break

				case "ADE": //ADE
					txtNumAcesso3.value = ''
					txtCabo3.value = ''
					txtPar3.value = ''
					txtPADE3.value = ''
					txtDerivacao3.value = ''
					cboPropModem3.value = ''
					cboTCabo3.value = ''
					txtQtdeModem3.value = ''
					txtAreaObs3.value = ''
					break
			}	
			//document.getElementById('btnAdicionarAlterar').disabled = true;
			document.getElementById('btnRemover').disabled = true;
			if(document.getElementById('tblFac').rows.length > 2)
			{
				rdoSelFacilidade[0].checked = true;
				rdoSelFacilidade[0].checked = false;
			}
			else if (document.getElementById('tblFac').rows.length == 2)
			{
				rdoSelFacilidade.checked = true;
				rdoSelFacilidade.checked = false;
			}
				
			document.getElementById('hdnFacIDAtual').value = '';
		}
	}	
}
function JanelaPosManobra()
{
	var blnDisabled 
	var TipoPlataforma
	var strPagina = "L"
	
	with (document.forms[0])
	{
		if(txtRede.value == "DETERMINISTICO")
		{
			PopularXml(objXmlReturn)
			objXmlReturn = window.showModalDialog('RedeDetManobra.asp',objXmlReturn,'dialogHeight: 380px; dialogWidth: 780px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
			if (RequestNode(objXmlReturn,"txtRedDetBastidor") != "")
			{
				//Facilidade
				txtBastidor.value	= RequestNode(objXmlReturn,"txtRedDetBastidor")
				txtRegua.value		= RequestNode(objXmlReturn,"txtRedDetRegua")
				txtPosicao.value	= RequestNode(objXmlReturn,"txtRedDetPosicao")
				txtTimeslot.value	= RequestNode(objXmlReturn,"txtRedDetTimeslot")
				txtFila.value		= RequestNode(objXmlReturn,"txtRedDetFila")
			}
		}
		else
		{
			if (!ValidarCampos(txtLocalEntrega,"Local de Entrega")) return
			if (!ValidarCampos(txtDistribuidor,"Distribuidor")) return
			if (!ValidarCampos(txtRede,"Rede")) return
			if (!ValidarCampos(txtProvedor,"Provedor")) return
			try{
				objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=L&cboLocalInstala='+hdnLocalInstala.value+'&cboDistLocalInstala='+hdnDisID.value+'&cboProvedor='+hdnProvedor2.value+'&strPagina='+strPagina+'&cboSistema='+hdnRedeID.value+'&cboPlataforma='+hdnPlaID.value ,objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: yes; resizable: yes; status: yes;");
			}
			catch(e){
				objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=L&cboLocalInstala='+hdnLocalInstala.value+'&cboDistLocalInstala='+hdnDisID.value+'&cboProvedor='+hdnProvedor2.value+'&strPagina='+strPagina+'&cboSistema='+hdnRedeID.value, objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: yes; resizable: yes; status: yes;");
			}
			if(txtRede.value == "NAO DETERMINISTICO")
			{
				if (objAryFac[0][0] != "")
				{
					txtTronco.value = objAryFac[0][0]
					txtPar2.value	= objAryFac[0][1]
				}					
			}
			else 
			{
				if (objAryFac[0][0] != "")
				{
					txtCabo3.value		= objAryFac[0][0]
					txtPar3.value		= objAryFac[0][1]
			    	txtDerivacao3.value	= objAryFac[0][2]
		        	cboTCabo3.value  	= objAryFac[0][3]
		        	txtPADE3.value      = objAryFac[0][4]
				}	
			}
		}
	}
}

function PosicoesLivresManobra(varLocalID,varDisID,varProID,varRedeID,varPlaID)
{
	var TipoPlataforma
	var strPagina = "L"
	with (document.forms[0])
	{
		if (!ValidarCampos(txtLocalEntrega,"Local de Entrega")) return
		if (!ValidarCampos(txtDistribuidor,"Distribuidor")) return
		if (!ValidarCampos(txtRede,"Rede")) return
		if (!ValidarCampos(txtProvedor,"Provedor")) return
		try{
			objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=L&cboLocalInstala='+varLocalID+'&cboDistLocalInstala='+varDisID+'&cboProvedor='+varProID+'&strPagina='+strPagina+'&cboSistema='+varRedeID+'&cboPlataforma='+varPlaID + '&abc='+varPlaID  ,objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: yes; resizable: yes; status: yes;");
		}
		catch(e){
			objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=L&cboLocalInstala='+varLocalID+'&cboDistLocalInstala='+varDisIDe+'&cboProvedor='+varProID+'&strPagina='+strPagina+'&cboSistema='+varRedeID, objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: yes; resizable: yes; status: yes;");
		}
		
		try{
			for (var intIndex=0;intIndex<objAryFacRet.length;intIndex++)
			{
				switch (txtRede.value) 
				{
					case "1": //Det
						if (objAryFac[intIndex][0] != "")
						{
							txtFila.value		= objAryFac[intIndex][0]
							txtBastidor.value	= objAryFac[intIndex][1]
							txtRegua.value		= objAryFac[intIndex][2]
							txtPosicao.value	= objAryFac[intIndex][3]
							txtTimeslot.value	= objAryFac[intIndex][4]
						}	
						break

					case "2": //NDet
						if (objAryFac[intIndex][0] != "")
						{
							txtTronco.value = objAryFac[intIndex][0]
							txtPar2.value	= objAryFac[intIndex][1]
						}						
						break

					case "3": //ADE
						if (objAryFac[intIndex][0] != "")
						{
							txtCabo3.value		= objAryFac[intIndex][0]
							txtPar3.value		= objAryFac[intIndex][1]
							txtDerivacao3.value	= objAryFac[intIndex][2]
							cboTCabo3.value	= objAryFac[intIndex][3]
							txtPADE3.value= objAryFac[intIndex][4]
						}	
						break
				}	
			}
		}
		catch(e){}	
	}	
}

function especialCharMask (especialChar){
        especialChar = especialChar.replace('/[áàãâä]/ui', 'a');
        especialChar = especialChar.replace('/[éèêë]/ui', 'e');
        especialChar = especialChar.replace('/[íìîï]/ui', 'i');
        especialChar = especialChar.replace('/[óòõôö]/ui', 'o');
        especialChar = especialChar.replace('/[úùûü]/ui', 'u');
        especialChar = especialChar.replace('/[ç]/ui', 'c');
        especialChar = especialChar.replace('/[^a-z0-9]/i', '_');
        especialChar = especialChar.replace('/_+/', '_'); //
        return especialChar;
}


//Completa os campos da tela de manobra com os campos passados para deterministico.
function ExibeManobraDet(var_Pro_ID,var_NroAcessoPtaEbt,var_Fac_Fila,var_Fac_Bastidor,var_Fac_Regua,var_Fac_Posicao,var_Time_slot,var_Acf_NroAcessoPtaCli,var_CCTOProvedor,var_CCTOTipo,var_Acf_CnlPTA,var_Acf_CnlPTB,var_Acf_ProprietarioEquip,var_Acf_QtdEquip,var_Acf_CCTOFatura,var_Acf_Obs,var_FacID,var_Link)
{
	var strDados
	var arrayCampos
	with (document.forms[0])
	{
		Linha = AchaHdnGridManobra()
		document.getElementById('btnAdicionarAlterar').disabled = false;
		document.getElementById('btnRemover').disabled = false;
		hdnFacIDAtual.value = var_FacID
		if(document.getElementById('hdnDados' + Linha).value == '')
		{
			cboCodProv.value = var_Pro_ID;
			txtNumAcessoPtaEbt.value = var_NroAcessoPtaEbt;
			txtFila.value  = var_Fac_Fila;
			txtBastidor.value = var_Fac_Bastidor;
			txtRegua.value = var_Fac_Regua;
			txtPosicao.value = var_Fac_Posicao;
			txtTimeslot.value = var_Time_slot
			txtNumAcessoCLI.value = var_Acf_NroAcessoPtaCli;
			txtCCTOProvedor.value = var_CCTOProvedor;
			if (var_CCTOTipo == "I")
				rdoUrbano[0].checked = true
			else
				rdoUrbano[1].checked = true
			txtCNLPontaA.value = var_Acf_CnlPTA
			txtCNLPontaB.value = var_Acf_CnlPTB
			cboPropModem.value = var_Acf_ProprietarioEquip
			txtQtdeModem.value = var_Acf_QtdEquip
			if (var_Acf_CCTOFatura == "S")
				rdoFatura[0].checked = true
			else
				rdoFatura[1].checked = true	
			txtAreaObs.value = var_Acf_Obs
			txtLink.value = var_Link
		}
		else
		{
			strDados = document.getElementById('hdnDados' + Linha).value
			arrayCampos = strDados.split("&&");
			cboCodProv.value = arrayCampos[0]
			txtNumAcessoPtaEbt.value = arrayCampos[1]
			txtFila.value = arrayCampos[2]
			txtBastidor.value = arrayCampos[3]
			txtRegua.value = arrayCampos[4]
			txtPosicao.value = arrayCampos[5]
			txtTimeslot.value = arrayCampos[6]
			txtCCTOProvedor.value = arrayCampos[7]
			txtNumAcessoCLI.value = arrayCampos[8]
			txtCNLPontaA.value = arrayCampos[9]
			txtCNLPontaB.value = arrayCampos[10]
			txtQtdeModem.value = arrayCampos[11]
			cboPropModem.value = arrayCampos[12]
			txtLink.value = arrayCampos[13]
			txtAreaObs.value = arrayCampos[14]
			if (arrayCampos[15] == "I")
				rdoUrbano[0].checked = true
			else
				rdoUrbano[1].checked = true
			if (arrayCampos[16] == "S")
				rdoFatura[0].checked = true
			else
				rdoFatura[1].checked = true	
			txtLink.value = arrayCampos[17]
		}
	}
}

function ValidarPadraoProvedorManobra(index)
{
	with (document.forms[0])
	{
		if(index ==1)
		{
			//Valida o tamanho do padrão mínimo/máximo permitido ao provedor
			if (txtNumAcessoPtaEbt.value.length != IFrmProcesso1.document.forms[0].hdnIntPadraoMin.value && txtNumAcessoPtaEbt.value.length != IFrmProcesso1.document.forms[0].hdnIntPadraoMax.value)
			{
				alert("Número do Acesso Pta Ebt fora do padrão.")
				txtNumAcessoPtaEbt.focus()
				return false
			}
			if (!IFrmProcesso1.ValidarPadraoManobra(document.forms[0].txtNumAcessoPtaEbt)) return false
		}
		else if(index ==2)
		{
			//Valida o tamanho do padrão mínimo/máximo permitido ao provedor
			if (txtNumAcessoPtaEbt2.value.length != IFrmProcesso1.document.forms[0].hdnIntPadraoMin.value && txtNumAcessoPtaEbt2.value.length != IFrmProcesso1.document.forms[0].hdnIntPadraoMax.value)
			{
				alert("Número do Acesso Pta Ebt fora do padrão.")
				txtNumAcessoPtaEbt2.focus()
				return false
			}
			if (!IFrmProcesso1.ValidarPadraoManobra(document.forms[0].txtNumAcessoPtaEbt2)) return false
		}
		else if(index ==3)
		{
			//Valida o tamanho do padrão mínimo/máximo permitido ao provedor
			if (txtNumAcesso3.value.length != IFrmProcesso1.document.forms[0].hdnIntPadraoMin.value && txtNumAcesso3.value.length != IFrmProcesso1.document.forms[0].hdnIntPadraoMax.value)
			{
				alert("Número do Acesso Pta Ebt fora do padrão.")
				txtNumAcesso3.focus()
				return false
			}
			if (!IFrmProcesso1.ValidarPadraoManobra(document.forms[0].txtNumAcesso3)) return false
		}
	}
	return true
}

//Completa os campos da tela de manobra com os campos passados para não deterministico.
function ExibeManobraNaoDet(var_Pro_ID,var_NroAcessoPtaEbt,var_Fac_Tronco,var_Fac_Par,var_Acf_NroAcessoPtaCli,var_CCTOProvedor,var_CCTOTipo,var_Acf_CnlPTA,var_Acf_CnlPTB,var_Acf_ProprietarioEquip,var_Acf_QtdEquip,var_Acf_CCTOFatura,var_Acf_Obs,var_FacID,Linha)
{
	var strDados
	var arrayCampos
	with (document.forms[0])
	{
		Linha = AchaHdnGridManobra();
		document.getElementById('btnAdicionarAlterar').disabled = false;
		document.getElementById('btnRemover').disabled = false;
		hdnFacIDAtual.value = var_FacID;
		if(document.getElementById('hdnDados' + Linha).value == '')
		{
			cboCodProv2.value = var_Pro_ID;
			txtNumAcessoPtaEbt2.value = var_NroAcessoPtaEbt;
			txtTronco.value  = var_Fac_Tronco;
			txtPar2.value = var_Fac_Par;
			txtNumAcessoCLI2.value = var_Acf_NroAcessoPtaCli;
			txtCCTOProvedor2.value = var_CCTOProvedor;
			if (var_CCTOTipo == "I")
				rdoUrbano2[0].checked = true
			else
				rdoUrbano2[1].checked = true
			txtCNLPontaA2.value = var_Acf_CnlPTA
			txtCNLPontaB2.value = var_Acf_CnlPTB
			cboPropModem2.value = var_Acf_ProprietarioEquip
			txtQtdeModem2.value = var_Acf_QtdEquip
			if (var_Acf_CCTOFatura == "S")
				rdoFatura2[0].checked = true
			else
				rdoFatura2[1].checked = true	
			txtAreaObs2.value = var_Acf_Obs
		}
		else
		{
			strDados = document.getElementById('hdnDados' + Linha).value
			arrayCampos = strDados.split("&&");
			cboCodProv2.value = arrayCampos[0]
			txtNumAcessoPtaEbt2.value = arrayCampos[1]
			txtTronco.value = arrayCampos[2]
			txtPar2.value = arrayCampos[3]
			txtCCTOProvedor2.value = arrayCampos[4]
			txtNumAcessoCLI2.value = arrayCampos[5]
			txtCNLPontaA2.value = arrayCampos[6]
			txtCNLPontaB2.value = arrayCampos[7]
			txtQtdeModem2.value = arrayCampos[8]
			cboPropModem2.value = arrayCampos[9]
			txtAreaObs2.value = arrayCampos[10]
			if (arrayCampos[11] == "I")
				rdoUrbano2[0].checked = true
			else
				rdoUrbano2[1].checked = true
			if (arrayCampos[12] == "S")
				rdoFatura2[0].checked = true
			else
				rdoFatura2[1].checked = true	
		}
	}
}

//Completa os campos da tela de manobra com os campos passados para ADE.
function ExibeManobraADE(var_NroAcessoPtaEbt,var_Fac_Tronco,var_Fac_Par,var_Fac_CxEmenda,var_Fac_Lateral,var_Fac_TipoCabo,var_Acf_ProprietarioEquip,var_Acf_QtdEquip,var_Acf_Obs,var_FacID)
{
	var arrayCampos
	var strDados = new String();
	with (document.forms[0])
	{
		Linha = AchaHdnGridManobra()
		document.getElementById('btnAdicionarAlterar').disabled = false;
		document.getElementById('btnRemover').disabled = false;
		hdnFacIDAtual.value = var_FacID
		if(document.getElementById('hdnDados' + Linha).value == '')
		{
			txtNumAcesso3.value = var_NroAcessoPtaEbt;
			txtCabo3.value  = var_Fac_Tronco;
			txtPar3.value = var_Fac_Par;
			txtPADE3.value = var_Fac_CxEmenda;
			txtDerivacao3.value = var_Fac_Lateral;
			cboTCabo3.value = var_Fac_TipoCabo;
			cboPropModem3.value = var_Acf_ProprietarioEquip;
			txtQtdeModem3.value = var_Acf_QtdEquip;
			txtAreaObs3.value = var_Acf_Obs;
		}
		else
		{
			strDados = document.getElementById('hdnDados' + Linha).value
			arrayCampos = strDados.split("&&");
			txtNumAcesso3.value = arrayCampos[0]
			txtCabo3.value = arrayCampos[1]
			txtPar3.value = arrayCampos[2]
			txtPADE3.value = arrayCampos[3]
			txtDerivacao3.value = arrayCampos[4]
			cboPropModem3.value = arrayCampos[5]
			cboTCabo3.value = arrayCampos[6]
			txtQtdeModem3.value = arrayCampos[7]
			txtAreaObs3.value = arrayCampos[8]
		}
	}
}



function upper(e)
{
	var campo;
	if(document.all)
		campo = window.event.srcElement;
	else
		campo = e.target;
		
	if (campo.type == 'text' && campo.alfatipo != 'min') 
	{
		campo.value = campo.value.toUpperCase()
	}	
	return true;
}
window.document.onfocusout = upper;
//PRSSILV
function MascaraMoeda(objTextBox, SeparadorMilesimo, SeparadorDecimal, e,TamanhoMaxCampo){
    var sep = 0;
    var key = '';
    var i = j = 0;
    var len = len2 = 0;
    var strCheck = '0123456789';
    var aux = aux2 = '';
    var whichCode = (window.Event) ? e.which : e.keyCode;
    if (whichCode == 13) return true;
    key = String.fromCharCode(whichCode); // Valor para o código da Chave
    if (strCheck.indexOf(key) == -1) return false; // Chave inválida
    len = objTextBox.value.length;
    for(i = 0; i < len; i++)
        if ((objTextBox.value.charAt(i) != '0') && (objTextBox.value.charAt(i) != SeparadorDecimal)) break;
    aux = '';
    for(; i < len; i++)
        if (strCheck.indexOf(objTextBox.value.charAt(i))!=-1) aux += objTextBox.value.charAt(i);
    aux += key;
    len = aux.length;

	if (len == TamanhoMaxCampo) {return};
    if (len == 0) objTextBox.value = '';
    if (len == 1) objTextBox.value = '0'+ SeparadorDecimal + '0' + aux;
    if (len == 2) objTextBox.value = '0'+ SeparadorDecimal + aux;
    if (len > 2) {
        aux2 = '';
        for (j = 0, i = len - 3; i >= 0; i--) {
            if (j == 3) {
                aux2 += SeparadorMilesimo;
                j = 0;
            }
            aux2 += aux.charAt(i);
            j++;
        }
        objTextBox.value = '';
        len2 = aux2.length;
        for (i = len2 - 1; i >= 0; i--)
        objTextBox.value += aux2.charAt(i);
        objTextBox.value += SeparadorDecimal + aux.substr(len - 2, len);
    }
    return false;
}
