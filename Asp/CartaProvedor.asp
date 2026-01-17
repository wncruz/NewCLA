<!--#include file="../inc/data.asp"-->
<HEAD>
<TITLE>Envia e-mail para o provedor</TITLE>
<LINK HREF = "..\css\CLA.CSS" REL ="stylesheet"/>
<SCRIPT LANGUAGE = "JavaScript">

	function  EnviaEmail()
	{
		with (document.forms[0])
		{
			//Valida numero do pedido
			if (txtPedID.value == "" )
			{
				alert("Preencha todos os campos para enviar o e-mail")
				txtSolID.setActive();
				return
			}
			ValidaNumPed(txtPedID.value)
			
			
			//Variáveis XML
			var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
			var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
			var repl , strXML , strRetorno
			
			//Busca XML padrão para envio de cartas a partir do numero do pedido
			strXML = "<root>"
			strXML = strXML + "<pedido>" +  txtPedID.value + "</pedido>"
			strXML = strXML + "</root>" 
						
			xmlDoc.loadXML(strXML);
			xmlhttp.Open("POST","emailprovedor_enviar.asp" , false);
			xmlhttp.Send(xmlDoc.xml);
			strRetorno = xmlhttp.responseText;
			
			//Verifica se o pedido existe 			
			var posFound = strRetorno.search("Pedido")

			if (posFound != -1){
					alert(strRetorno)
					return 
			}
			
			strXML = strRetorno
			
			// Busca modelo de carta a ser enviada 
			xmlDoc.loadXML(strXML);
			xmlhttp.Open("POST","RetornaModelo.asp" , false);
			xmlhttp.Send(xmlDoc.xml);
								
			// Executa a geração do Html da carta
			strRetorno = xmlhttp.responseText;
			xmlhttp.Open("POST",strRetorno , false);
			xmlhttp.Send(xmlDoc.xml);

							
			// Verifica o sucesso da geração da carta
			strXML  = xmlhttp.responseText;
			var posFound = strXML.search("http_404.htm")
					
			if (posFound != -1)
			{
				alert("Não foi possível enviar o e-mail, verifique o cadastro do provedor. - Formulário padrão.")
				return 
			}
			
			// verifica se a carta enviada utiliza o modelo padrão.		
			if (strRetorno != "ProcessoEmailProvedorPadrao.asp" && strRetorno != "CartaPadrao.asp"  && strRetorno != "" &&  strXML.substring(0,6) != "<table")
			{
				repl = /&/g	
				strXML = strXML.replace(repl,"&amp;");
				xmlDoc.loadXML(strXML);
				xmlhttp.Open("POST", "RetornaCarta.asp", false);
				xmlhttp.Send(xmlDoc.xml);
			}
					
			//Abri o modelo da carta para ser enviado
			strXML = xmlhttp.responseText;
			objWindow = window.open("About:blank", null, "status=no,toolbar=no,menubar=no,location=no,resizable=Yes,scrollbars = Yes");
			objWindow.document.write(strXML);		
			objWindow.document.close();
					
		}
	}

	function ValidaNumPed(StrPed)
	{
		var blnTraco = false 
		var blnBarra  = false
		var
		
		blnTraco = StrPed.search("-")
		blnBarra = StrPed.search("/")
		
		if (blnTraco == 0 || blnBarra == 0)
		{
			alert("Número de pedido inválido.")
			txtPedID.setActive();
			return
		}
		if (blnTraco != 2)
		{
			alert("Número de pedido inválido.")
			txtPedID.setActive();
			return
		}
		if (StrPed.substring(3,blnBarra).length > 6 )
		{
			alert("Número de pedido inválido.")
			txtPedID.setActive();
			return
		}
		if (StrPed.substring(blnBarra + 1 ,StrPed.length).length != 4 )
		{
			alert("Número de pedido inválido.")
			txtPedID.setActive();
			return
		}
		
	}
</SCRIPT>
</HEAD>
<BODY leftmargin="0" topmargin="0">
<!--#include file="../inc/header.asp"-->
<FORM  method=post >
<input type="hidden" name="hdnPedId">
<input type="hidden" name="hdnSolId">


<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align=center>Criação/Envio de Carta ao Provedor</p></th>
</tr>
<tr class=clsSilver >
	<td width=200px nowrap><font class="clsObrig">:: </font>Número do Pedido</td>
	<td  >
<input size="20" CLASS=TEXT name= "txtPedID"  maxlength=14 value="DM-">
	</td>
</tr>
<tr>
	<td align="center" colspan=2><br>
<input type="button" class="button" name="btnEmailPro" style="width:150px;text-align=center" value="Enviar e-mail para provedor" onclick ="javascript:EnviaEmail();">
	</td>
</td>
</tr>
</table>

</FORM>
</BODY>
