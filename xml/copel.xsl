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
					FONT-SIZE: 9pt; FONT-FAMILY: Arial
				}

				.TD_CAMPOFIM 
				{
					FONT-SIZE: 9pt; BORDER-BOTTOM: #000000 2px solid; FONT-FAMILY: Arial
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
					FONT-WEIGHT: bold; FONT-SIZE: 11pt; COLOR: #800000; BORDER-BOTTOM: #000000 1px solid; FONT-FAMILY: Arial; TEXT-ALIGN: center
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
		
		<TABLE cellSpacing="3" cellPadding="0" width="85%" align="center" border="0">
		<TBODY>
				<TR>
					<TD class="TD_PED" align="left" width="35%"> Data: <FONT color="black"><B> <xsl:value-of select="//data"/> </B></FONT> </TD>
					<TD class="TD_PED" align="left" colSpan="2"> Nº Pedido: <FONT color="black"><B> <xsl:value-of select="//numero"/> </B></FONT> </TD>
				</TR>
  
				<TR>
					<TD class="TD_EMBRATEL" align="middle" colSpan="3">EMBRATEL S/A</TD>
				</TR>
	
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">
						<FONT color="black">SOLICITAÇÃO DE ACESSO FIBRA ÓPTICA</FONT>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">DADOS DO CLIENTE
					</TD>
				</TR>

				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">2. Nome Solicitante: <xsl:value-of select="//nomeGICL"/>
					</TD>
		
					<TD class="TD_CAMPO" align="left" width="35%">3. E-Mail: <xsl:value-of select="//proemail"/>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">4. Razão Social: <B> <xsl:value-of select="//nomecontratante"/> </B>
					</TD>
    
					<TD class="TD_CAMPO" align="left">5. CNPJ: <B> <xsl:value-of select="//cgc_cnpj_contratante"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left">6. Telefone Solicitante: <xsl:value-of select="//TelefoneGICL"/>
					</TD>
		
					<TD class="TD_CAMPO" align="left" colSpan="2">7. Telefone p/ Cobrança:
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPOFIM" align="left" colSpan="3">8. Aplicação do Serviço: <B> Entroncamento </B>
					</TD>
				</TR>
  <!-- LPEREZ 13/12/2005 -->
				<TR height="80">
					<TD class="TD_CAMPOFIM" vAlign="top" align="left" colSpan="3">9. Observações:   <B><xsl:value-of select="//observaçãos"/><br/><xsl:value-of select="//observaçãop"/></B>
					</TD>
				</TR>
	<!-- LP -->
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">DADOS DO SERVIÇO
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPOFIM" align="left" colSpan="2">10. Tipo de Serviço: <B> <xsl:value-of select="//servico_descricao"/> </B>
					</TD>
		
					<TD class="TD_CAMPOFIM" align="left">11. Nº Circuito: <B> <xsl:value-of select="//pontaebt"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">CARACTERÍSTICAS TÉCNICAS DO CIRCUITO
					</TD>
				</TR>
  
  
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="3">12. Interface A: <B> <xsl:value-of select="//clienteinterface"/>  </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="3">13. Interface B: <B> <xsl:value-of select="//pontabinterface"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="3">14. Velocidade de Operação: 
						<B> <xsl:value-of select="//clientevelocidade"/></B>
					</TD>
				</TR>
				
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="3">15. Prazo de Operação:
					
					
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
					<TD class="TD_CAMPOFIM" align="left" colSpan="3">16. Redundância: 
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">USUÁRIO - DADOS DA PONTA A
					</TD>
				</TR>

				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">17. Razão Social: <B> <xsl:value-of select="//clientenome"/> </B>
					</TD>
    
					<TD class="TD_CAMPO" align="left">18. CNPJ: <B> <xsl:value-of select="//clientecnpj"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="3">19. Rua: <B> <xsl:value-of select="//clienterua"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left">22. Nº: <B> <xsl:value-of select="//clientenumero"/> </B>
					</TD>
			
					<TD class="TD_CAMPO" align="left" colSpan="2">23. Complemento: <B> <xsl:value-of select="//clientecompl"/> </B>
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">24. Cidade/Estado: <B><xsl:value-of select="//clientecidade"/> / <xsl:value-of select="//clienteuf"/> </B>
					</TD>
    
				</TR>
	
				
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">27. Bairro: <B><xsl:value-of select="//clientebairro"/></B>
					</TD>
		
					<TD class="TD_CAMPO" align="left">28. Nome para Contato: <B> <xsl:value-of select="//clientecontato"/></B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="3">29. Telefone: <B> <xsl:value-of select="//clientetelefone"/></B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPOFIM" align="left" colSpan="3">30. Setor:
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_SUBTIT" align="middle" colSpan="3">USUÁRIO - DADOS DA PONTA B
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">31. Razão Social: <B> <xsl:value-of select="//nomecontratante"/> </B>
					</TD>
    
					<TD class="TD_CAMPO" align="left">32. CNPJ: <B>  <xsl:value-of select="//cgc_cnpj_contratante"/> </B>
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="3">33. Endereço: <B> <xsl:value-of select="//pontabend_rua"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left">34. Nº: <B><xsl:value-of select="//pontabend_numero"/></B>
					</TD>
					<TD class="TD_CAMPO" align="left" colSpan="2">35. Complemento: <B> <xsl:value-of select="//pontabend_complemento"/> </B>
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">36. Cidade/Estado: <B> <xsl:value-of select="//pontabcidade"/> / <xsl:value-of select="//pontabuf"/></B>
					</TD>
    
				</TR>
  
				<TR>
		
					<TD class="TD_CAMPO" align="left" colSpan="2">39. Bairro: <B> <xsl:value-of select="//pontabend_bairro"/> </B>
					</TD>
				</TR>
	
				<TR>
					<TD class="TD_CAMPO" align="left" colSpan="2">40. Nome para Contato: <B> <xsl:value-of select="//pontabcontato"/> </B>
					</TD>
		
					<TD class="TD_CAMPO" align="left" colSpan="3">41. Setor:
					</TD>
				</TR>
  
				<TR>
					<TD class="TD_CAMPOFIM" align="left" colSpan="3">42. Telefone: <B> <xsl:value-of select="//pontabtel"/> </B>
					</TD>
				</TR>
	</TBODY>
	</TABLE>

</BODY>
</HTML>
</xsl:template>
</xsl:stylesheet>
