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
<FORM  method=post >
<input type="hidden" name="hdnPedId">
<input type="hidden" name="hdnSolId">
<TABLE border=0 cellPadding=0 cellSpacing=1 width="100%" >
<TR>
	<TD valign = "top" background=..\imagens\topo_embratel.jpg  height = 80 colspan = 2 ></TD>
</TR>
<TR>
	<th colspan=2><p align=center>Envio manual de e-mail</p></th>
	<!--<TD  colspan = 2 style = "font-size:10pt;COLOR:white;FONT-WEIGHT: bold" bgcolor=SteelBlue align  = center >Envio manual de e-mail</TD>-->
</TR>
<TR>
	<TD width = 15%> Número do Pedido: </TD>
	<TD> <input  style = " BORDER-RIGHT: 1px solid;BORDER-TOP: 1px solid;BORDER-LEFT: 1px solid;BORDER-BOTTOM: 1px solid;font-family:verdana" name= "txtPedID"  maxlength=14> </TD>
</TR>
</TABLE>
<P></P>
<TABLE WIDTH = 100%>
<TR>
	<td colspan = 2 align =CENTER ><input type="button" class="button" name="btnEmailPro" style="width:150px;text-align=center" value="Enviar e-mail " onclick ="javascript:EnviaEmail();"></td>
</TR>
</TABLE>
</FORM>
</BODY>
<%
	set objRSProv = nothing
%>

