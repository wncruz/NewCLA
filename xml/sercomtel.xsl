<?xml version = "1.0" encoding ="ISO-8859-1" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="html" encoding ="ISO-8859-1"/>
	<xsl:template match="/">
	
		<HTML>

		<HEAD>
			
			<TITLE>Pedido: <xsl:value-of select="//numero"/></TITLE>

			<STYLE>


				.TD_CAMPO 
				{
					FONT-SIZE: 9pt;
				    FONT-FAMILY: Arial
				}

				.TD_CAMPOFIM 
				{
					FONT-SIZE: 9pt; BORDER-BOTTOM: #000000 1px solid; FONT-FAMILY: Arial
				}
	
				.TD_PED 
				{
					FONT-WEIGHT: bold; FONT-SIZE: 14pt; COLOR: red; BORDER-BOTTOM: #000000 2px solid; FONT-FAMILY: Arial
				}
	
				.TD_EMBRATEL 
				{
					FONT-WEIGHT: bold; FONT-SIZE: 15pt; COLOR: blue; BORDER-BOTTOM: #000000 2px solid; FONT-FAMILY: Arial
				}

				.TD_NEGRITO 
				{
					FONT-WEIGHT: bold; FONT-SIZE: 9pt; FONT-FAMILY: Arial
				}
	
				.TD_SUBTIT 
				{
			
				    FONT-WEIGHT: bold;
				    FONT-SIZE: 11pt;
				    FONT-FAMILY: Tahoma;
				    COLOR: #800000; 
				    BORDER-BOTTOM: #000000 1px solid; 
				    
				    TEXT-DECORATION: underline
				
				}

				TR 
				{
					BACKGROUND-COLOR: #ffffff
				}

				INPUT.button 
				{
					BORDER-RIGHT: #0f1f5f 1px solid; BORDER-TOP: #0f1f5f 1px solid; FONT-WEIGHT: normal; FONT-SIZE: 9px; BORDER-LEFT: #0f1f5f 1px solid; WIDTH: 100px; BORDER-BOTTOM: #0f1f5f 1px solid; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BACKGROUND-COLOR: #f1f1f1; TEXT-DECORATION: none
				}

				.td_Div 
				{
					BORDER-BOTTOM: #000000 1px solid
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
					document.write(strXML);
								
				}
			</SCRIPT>

		</HEAD>

		<BODY leftMargin="0" topMargin="0" align="center" marginwidth="0" marginheight="0">
		
		<TABLE cellSpacing="3" cellPadding="0" width="95%" align="center" border="0">

				<TR>
					<TD class="TD_PED" align="left" width="35%"> Data: <FONT color="black"><B> <xsl:value-of select="//data"/> </B></FONT> </TD>
					<TD class="TD_PED" align="left" colSpan="2"> Nº Pedido: <FONT color="black"><B> <xsl:value-of select="//numero"/> </B></FONT> </TD>
				</TR>
  
				<TR>
					<TD class="TD_EMBRATEL" align="middle" colSpan="3">EMBRATEL S/A</TD>
				</TR>
	
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">
						<FONT color="black">SOLICITAÇÃO DE SERVIÇO EXPLORAÇÃO INDUSTRIAL DE LINHAS DEDICADAS-EILD</FONT>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">DADOS DO CLIENTE
					</TD>
				</TR>

				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">1. Nome Solicitante: <xsl:value-of select="//nomeGICL"/>
					</TD>
		
					<TD class="TD_CAMPO" align="left" width="35%">2. E-Mail: <xsl:value-of select="//proemail"/>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">3. Razão Social: <B> <xsl:value-of select="//nomecontratante"/> </B>
					</TD>
    
					<TD class="TD_CAMPO" align="left">4. CNPJ: <B> <xsl:value-of select="//cgc_cnpj_contratante"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">5. Telefone Solicitante: <xsl:value-of select="//TelefoneGICL"/>
					</TD>
		
					<TD class="TD_CAMPO" align="left" colSpan="2">6. Telefone p/ Cobrança: <xsl:value-of select="//telefonecontratante"/>
					</TD>
				</TR>
	
  
				<TR height="80">
					<TD class="TD_CAMPOFIM" vAlign="top" align="left" colSpan="3">7. Observações:   <B><xsl:value-of select="//observaçãos"/><br/><xsl:value-of select="//observaçãop"/></B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">DADOS DO SERVIÇO
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPOFIM" align="left" colSpan="2">8. Tipo de Serviço: <B> <xsl:value-of select="//acaooriginal"/> </B>
					</TD>
		
					<TD class="TD_CAMPOFIM" align="left">9. Nº Circuito: <B> <xsl:value-of select="//pontaebt"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">CARACTERÍSTICAS TÉCNICAS DO CIRCUITO
					</TD>
				</TR>
  
  
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="3">
						10. Âmbito do Circuito: <B> <xsl:value-of select="//tipocircuito"/> </B>  
					</TD>
				</TR>
  				
  				<TR>
					<TD class="TD_CAMPO" align="left" >11. Interface Física a ser Utilizada: <B> <xsl:value-of select="//clienteinterface"/> </B>  </TD>
				</TR>
  
  
  				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">12. Modo de Transmisão: <B>   </B>  </TD>
					<TD class="TD_CAMPO" align="left" >13. Portadora: <B> <xsl:value-of select="//portadora"/>  </B>  </TD>
				</TR>
  
  
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">14. Velocidade de Transmissão: 
						<B> <xsl:value-of select="//clientevelocidade"/></B>
					</TD>
					
					<TD class="TD_CAMPO" align="left" >15. Modo de Operação: 
						<B> Full Duplex</B>
					</TD>
					
				</TR>
				
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">16. Modalidade do Circuito:
					
					
						<xsl:if test="//chkind = -1">
							<span style = "font-family:'verdana'"><B> Indeterminado</B></span>
						</xsl:if>
				
						<xsl:if test="//chk12meses = -1">
							<span style = "font-family:'verdana'"><B> 12 Meses</B></span>
						</xsl:if>
					
						<xsl:if test="//chk24meses = -1">
							<span style = "font-family:'verdana'"><B> 24 Meses</B></span>
						</xsl:if>
					
						<xsl:if test="//chk36meses = -1">
							<span style = "font-family:'verdana'"><B> 36 Meses</B></span>
						</xsl:if>
					
						<xsl:if test="//chk48meses = -1">
							<span style = "font-family:'verdana'"><B> 48 Meses</B></span>
						</xsl:if>
					
						<xsl:if test="//chk60meses = -1">
							<span style = "font-family:'verdana'"><B> 60 Meses</B></span>
						</xsl:if>

						<xsl:if test="//chktemporario = -1">
							<span style = "font-family:'verdana'"><B>Temporário:</B></span>
						</xsl:if>
						
						<xsl:if test="//chktemporario = -1">
							de<xsl:value-of select="//temporariode"/>
							até <xsl:value-of select="//temporarioate"/>
						</xsl:if>
				
					
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">17. Modalidade de Instalação: <B> Com Modem </B>
					</TD>
					
					<TD class="TD_CAMPO" align="left" >18. Tipo Sinal: <B> Digital </B>
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPOFIM" align="left" colSpan="3">19. Dados do Acesso: <B> </B>
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">USUÁRIO - DADOS DA PONTA A
					</TD>
				</TR>

				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">20. Razão Social: <B> <xsl:value-of select="//clientenome"/> </B>
					</TD>
    
					<TD class="TD_CAMPO" align="left">21. CNPJ: <B> <xsl:value-of select="//clientecnpj"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">22. Rua: <B> <xsl:value-of select="//clienterua"/> </B>
					</TD>
					
					<TD class="TD_CAMPO" align="left">23. Nº: <B> <xsl:value-of select="//clientenumero"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">24. Complemento: <B> <xsl:value-of select="//clientecompl"/> </B>
					</TD>
					
					<TD class="TD_CAMPO" align="left" >25. Cep: <B> <xsl:value-of select="//clientecep"/> </B>
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">26. Cidade/Estado: <B><xsl:value-of select="//clientecidade"/> / <xsl:value-of select="//clienteuf"/> </B>
					</TD>
    				<TD class="TD_CAMPO" align="left" colSpan="2">27. Telefone Instalado no Local: <B><xsl:value-of select="//clientetelefone"/> </B>
					</TD>
    
				</TR>
	
				
				<TR>
					<TD class="TD_CAMPOFIM" align="left" colSpan="2">28. Bairro: <B><xsl:value-of select="//clientebairro"/></B>
					</TD>
		
					<TD class="TD_CAMPOFIM" align="left">29. Nome para Contato: <B> <xsl:value-of select="//clientecontato"/></B>
					</TD>
				</TR>
	
	
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">USUÁRIO - DADOS DA PONTA B
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">30. Razão Social: <B> <xsl:value-of select="//nomecontratante"/> </B>
					</TD>
    
					<TD class="TD_CAMPO" align="left">31. CNPJ: <B>  <xsl:value-of select="//cgc_cnpj_contratante"/> </B>
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">32. Rua: <B> <xsl:value-of select="//pontabend_rua"/> </B>
					</TD>

					<TD class="TD_CAMPO" align="left">33. Nº: <B><xsl:value-of select="//pontabend_numero"/></B>
					</TD>

				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">34. Complemento: <B> <xsl:value-of select="//pontabend_complemento"/> </B>
					</TD>
					<TD class="TD_CAMPO" align="left">35. Cep: <B> <xsl:value-of select="//pontabcep"/> </B>
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">36. Cidade/Estado: <B> <xsl:value-of select="//pontabcidade"/> / <xsl:value-of select="//pontabuf"/></B>
					</TD>
					<TD class="TD_CAMPO" align="left" colSpan="2">37. Teleforne Instalado no Local: <B> <xsl:value-of select="//telefonecontratante"/> </B>
					</TD>
				</TR>
  
				<TR>
		
					<TD class="TD_CAMPO" align="left" colSpan="2">38. Bairro: <B> <xsl:value-of select="//pontabend_bairro"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">39. Nome para Contato: <B> <xsl:value-of select="//contatocontratada"/> </B>
					</TD>
		
					<TD class="TD_CAMPO" align="left" colSpan="3">40. Setor: <B> <xsl:value-of select="//cargocontatocontratada"/> </B>
					</TD>
				</TR>
  
	</TABLE>

</BODY>
</HTML>
</xsl:template>
</xsl:stylesheet>
