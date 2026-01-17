<?xml version = "1.0" encoding ="ISO-8859-1" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="html" encoding ="ISO-8859-1"/>
	<xsl:template match="/">

	<HTML>
		<HEAD>
	
			<TITLE>Carta de Cancelamento/Desativação Para Brasil Telecom</TITLE>


			<STYLE>
				TD 
				{
					FONT-SIZE: 7pt; FONT-FAMILY: Tahoma
				}

				.TD_RESPOSTA 
				{
					BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; FONT-SIZE: 7pt; BORDER-LEFT: 0px solid; BORDER-BOTTOM: 0px solid; FONT-FAMILY: Tahoma
				}

				INPUT 
				{
					FONT-SIZE: 7pt; FONT-FAMILY: Tahoma
				}
		
				.TD_CAMPO 
				{
					FONT-WEIGHT: bold; FONT-SIZE: 7pt; FONT-FAMILY: Tahoma; BACKGROUND-COLOR: #c2c3d2
				}

				.td_btn 
				{
					BORDER-RIGHT: thin groove; BORDER-TOP: thin groove; FONT-WEIGHT: bold; FONT-SIZE: 7pt; BORDER-LEFT: thin groove; CURSOR: hand; COLOR: white; BORDER-BOTTOM: thin groove; FONT-FAMILY: Tahoma; BACKGROUND-COLOR: steelblue; TEXT-ALIGN: center
				}

				.TD_NEGRITO 
				{
					FONT-WEIGHT: bold; FONT-FAMILY: Arial
				}
		
				.TD_CAMPO_CLARO 
				{
					FONT-WEIGHT: bold; FONT-SIZE: 7pt; FONT-FAMILY: Tahoma; BACKGROUND-COLOR: #edebeb
				}
		
				.TD_SUBTIT 
				{
					FONT-WEIGHT: bold; FONT-SIZE: 7pt; FONT-FAMILY: Tahoma; BACKGROUND-COLOR: #b2e6f7; TEXT-DECORATION: underline
				}

				.TD_EMBRATEL 
				{
					BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; FONT-WEIGHT: normal; BORDER-LEFT: 1px solid; COLOR: darkblue; BORDER-BOTTOM: 1px solid; FONT-STYLE: oblique; FONT-FAMILY: 'Arial Black'; BACKGROUND-COLOR: #eeeeea
				}

				INPUT.button 
				{
					BORDER-RIGHT: #0f1f5f 1px solid; BORDER-TOP: #0f1f5f 1px solid; FONT-WEIGHT: normal; FONT-SIZE: 9px; BORDER-LEFT: #0f1f5f 1px solid; WIDTH: 100px; BORDER-BOTTOM: #0f1f5f 1px solid; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BACKGROUND-COLOR: #f1f1f1; TEXT-DECORATION: none
				}

				.table_interno 
				{
					BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 0px solid; BORDER-BOTTOM: 0px solid
				}

			</STYLE>

			<SCRIPT>
				function Imprimir()
				{
					window.print();
					//alert("teste");
				}

				function SendMail()
				{

					var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
					var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
					var strXML

					xmlDoc.loadXML(Envio.hdnstrXML.value);
					xmlhttp.Open("POST", "EnviaEmail.asp", false);
					xmlhttp.Send(xmlDoc.xml);

					strXML = xmlhttp.responseText;
					
					objWindow = window.open("About:blank", null, "status=no,toolbar=no,menubar=no,location=no,resizable=Yes,scrollbars = Yes");
					objWindow.document.write(strXML);		
					objWindow.document.close();
					document.close();
			
				}
			</SCRIPT>

		</HEAD>

<BODY>

<TABLE  border ="0" cellSpacing="0" cellPadding="0" width="100%" align="center">

    <TR>
		<TD width="100%" class="TD_EMBRATEL" style="FONT-SIZE: 14pt" vAlign="center" align="left" rowSpan="1">
			EMBRATEL
		</TD>
	</TR>

  	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			<BR> </BR>
		</TD>
	</TR>
  
    <TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			<B>CT. IMPLANTAÇÃO SUL - <xsl:value-of select="//ufcontratada"/>. Pedido: <xsl:value-of select="//numero"/> </B>
		</TD>
	</TR>
  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			<B><xsl:value-of select="//cidadecontratada"/>,  <xsl:value-of select="//data"/> </B>
		</TD>
	</TR>

	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			<BR> </BR>
		</TD>
	</TR>
  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			
		</TD>
	</TR>
	  
  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			<B>À</B>
		</TD>
	</TR>
  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			Brasil Telecom - <B> <xsl:value-of select="//ufcontratada"/> </B>
		</TD>
	</TR>
  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			Diretoria Adjunta de Negócios com Operadoras - Coordenação de Contratos com Operadoras - Núcleo Regional
		</TD>
	</TR>
  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			Endereço: <xsl:value-of select="//endercontratada"/>  
		</TD>
		
	</TR>
  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			Cidade: <xsl:value-of select="//cidadecontratada"/>
		</TD>
	</TR>

  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			A/C.: <xsl:value-of select="//contatocontratada"/> 
		</TD>
	</TR>

  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			<BR> </BR>
			<BR> </BR>
		</TD>
	</TR>
	
	<TR>
		<TD></TD>
	</TR>
  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			Ass.: Lista Consolidada das Solicitações de Retirada ref. <xsl:value-of select="//data"/>
		</TD>
	</TR>
  
    <TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			<BR> </BR>
			<BR> </BR>
		</TD>
	</TR>
	
	<TR>
		<TD></TD>
	</TR>
  
    <TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			Prezado Sr,
		</TD>
	</TR>

	<TR>
		<TD><BR> </BR></TD>
	</TR>

	<TR>
		<TD></TD>
	</TR>
  
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="left" width="40%" rowSpan="1">
			Solicitamos a desativação/cancelamento dos acessos/pedidos abaixo relacionados.
		</TD>
	</TR>
  
	<TR>
		<TD> 
			<BR> </BR>
			<BR> </BR>
			<BR> </BR>
			<BR> </BR>
			<BR> </BR>
		</TD>
	</TR>
	
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="center" width="40%" rowSpan="1">
		Atenciosamente,
		</TD>
	</TR>
	
	<TR>
		<TD> 
			<BR> </BR> 
			<BR> </BR>
		</TD>
	</TR>
	
		
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="center" width="40%" rowSpan="1">
			__________________________________________________________________________________________________ <BR></BR>
			<xsl:value-of select="//contatoebt"/> <BR></BR>
			Gerente Técnico de Redes do <xsl:value-of select="//ufcontratada"/> <BR></BR>
			Diretoria de Redes - SUL<BR></BR>
			<xsl:value-of select="//emailcontatoebt"/><BR></BR>
		</TD>
	</TR>
	
	<TR>
		<TD> 
			<BR> </BR> 
			<BR> </BR>
		</TD>
	</TR>
	
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="Left" width="40%" rowSpan="1">

			Recebido Por: <xsl:value-of select="//contatocontratada"/> <BR></BR> <BR></BR>
			Assinatura: ________________________________________________ <BR></BR> <BR></BR>
			Data: ____/____/______ <BR></BR> <BR></BR>
			
		</TD>
	</TR>


	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="Left" width="40%" rowSpan="1">
			EMBRATEL S/A - EMPRESA BRASILEIRA DE TELECOMUNICAÇÕES  <BR></BR>
			Endereço: <xsl:value-of select="//endercontratante"/> <BR></BR>
			Cidade:<xsl:value-of select="//cidadecontratante"/> - <xsl:value-of select="//ufcontratante"/> <BR></BR>
			Cep: <xsl:value-of select="//cepcontratante"/> <BR></BR>
			Telefone: <xsl:value-of select="//telefonecontratante"/> <BR></BR>
			Fax: <xsl:value-of select="//faxcontratante"/><BR></BR>
		</TD>
	</TR>
	
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="Left" width="40%" rowSpan="1">
			Relação de Acessos ref. CT  <xsl:value-of select="//numero"/> <BR></BR>
		</TD>
	</TR>
	
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="Left" width="40%" rowSpan="1">
		
		</TD>
	</TR>
	
	<TR>
		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="Left" width="40%" rowSpan="1">
			Acesso: <xsl:value-of select="//pontaebt"/> <BR></BR>
			Cliente: <xsl:value-of select="//clientenome"/> <BR></BR>
			Número do DM: <xsl:value-of select="//numero"/> <BR></BR>
			Velocidade: <xsl:value-of select="//clientevelocidade"/>  <BR></BR>
			Local: <xsl:value-of select="//clientecidade"/>  <BR></BR> <BR></BR> <BR></BR>
			
		</TD>
	</TR>
	
	<TR>

		<TD class="TD_RESPOSTA" style="FONT-SIZE: 10pt" vAlign="center" align="Left" width="40%" rowSpan="1">
			Recebido Por:_________________________________________________________ Data:____/____/_______
		</TD>

	</TR>
	
  </TABLE>

</BODY>

</HTML>

</xsl:template>
</xsl:stylesheet>
