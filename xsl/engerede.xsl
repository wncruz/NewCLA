<?xml version="1.0" encoding="ISO-8859-1" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="html" encoding ="ISO-8859-1"/>
	<xsl:template match="/">
	<html>
			<head>
			<title>ENGEREDES - Pedido: <xsl:value-of select="//numero"/></title>
			<style>
				.TD_CAMPO_BORDER
				{
				    FONT-WEIGHT: bold;
				    FONT-SIZE: 7pt;
				    FONT-FAMILY: Tahoma;
				    BACKGROUND-COLOR: #c2c3d2;
					BORDER-RIGHT: 1px solid;
					BORDER-TOP:  1px solid;
					BORDER-LEFT:  1px solid;
					BORDER-BOTTOM: 1px solid
				}
				TD
				{
				    FONT-SIZE: 7pt;
				    FONT-FAMILY: Tahoma;
				}
				.TD_RESPOSTA
				{
				    FONT-SIZE: 7pt;
				    FONT-FAMILY: Tahoma;
				    border-top: 1px solid ;
				    border-right: 1px solid ;
				    border-bottom: 1px solid ;
				    border-left: 1px solid;
				}
				INPUT
				{
				    FONT-SIZE: 7pt;
				    FONT-FAMILY: Tahoma;
				}
				.TD_CAMPO
				{
				    FONT-WEIGHT: bold;
				    FONT-SIZE: 7pt;
				    FONT-FAMILY: Tahoma;
				    BACKGROUND-COLOR: #c2c3d2
				}
				.td_btn
				{
				    FONT-WEIGHT: bold;
				    FONT-SIZE: 7pt;
				    CURSOR: hand;
				    COLOR: white;
				    BORDER-LEFT: thin groove;
				    BORDER-BOTTOM: thin groove;
				    BORDER-RIGHT: thin groove;
				    BORDER-TOP: thin groove;
				    FONT-FAMILY: Tahoma;
				    BACKGROUND-COLOR: steelblue;
				    TEXT-ALIGN: center
				}
				.TD_NEGRITO
				{
				    FONT-WEIGHT: bold;
				    FONT-FAMILY: Arial
				}
				.TD_CAMPO_CLARO
				{
				    FONT-WEIGHT: bold;
				    FONT-SIZE: 7pt;
				    FONT-FAMILY: Tahoma;
				    BACKGROUND-COLOR: #edebeb
				}
				.TD_SUBTIT
				{
				    FONT-WEIGHT: bold;
				    FONT-SIZE: 7pt;
				    FONT-FAMILY: Tahoma;
				    BACKGROUND-COLOR: #b2e6f7;
				    TEXT-DECORATION: underline
				}
				.TD_EMBRATEL
				{
				    FONT-WEIGHT: normal;
				    COLOR: darkblue;
				    FONT-STYLE: oblique;
				    FONT-FAMILY: 'Arial Black';
				    BACKGROUND-COLOR: #eeeeea;
				    border-top: 1px solid ;
				    border-right: 1px solid ;
				    border-bottom: 1px solid ;
				    border-left: 1px solid;
				}
				INPUT.button
				{
					font-family: Verdana, Arial, Helvetica, sans-serif;
					font-size: 9px;
					font-weight: normal;
					TEXT-ALIGN: center
					color: #000000;
					text-decoration: none;
					background-color: #f1f1f1;
					border-top: 1px solid #0F1F5F;
					border-right: 1px solid #0F1F5F;
					border-bottom: 1px solid #0F1F5F;
					border-left: 1px solid #0F1F5F;
					width:100px
				}
				.table_interno
				{
				    border-top: 1px solid ;
				    border-right: 1px solid ;
				    border-bottom: 1px solid ;
				    border-left: 1px solid;
				}
			</style>
			<script>
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
			</script>
			</head>
			<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0"  align = "center">
			<form name="form1">
			<TABLE align="center" class ="table_interno" cellPadding="0" cellSpacing="0" width = "80%">
			<TR>
				<TD>
			<TABLE align="center" border = "0" cellPadding="0" cellSpacing="1" width="100%">
				<TR>
					<TD ROWSPAN = "3" STYLE = "FONT-SIZE: 14pt"  CLASS  = "TD_EMBRATEL" ALIGN = "middle"  VALIGN="center" >Embratel</TD>
					<TD ROWSPAN = "3" VALIGN = "center" ALIGN = "middle" STYLE = "FONT-SIZE: 10pt" width  = "40%" CLASS="TD_RESPOSTA">
						<B>FORMULÁRIO DE PROVIMENTO<BR></BR>
						DE ACESSO DEDICADO <BR></BR>
						ANEXO II </B>
					</TD>
					<TD  STYLE = "FONT-SIZE: 10pt" CLASS="TD_RESPOSTA">
						<B>Nº: </B> <xsl:value-of select="//numero"/>
					</TD>
				</TR>
				<TR>
					<TD STYLE = "FONT-SIZE: 10pt" CLASS="TD_RESPOSTA">
						<B>Contrato: </B> <xsl:value-of select="//contrato"/>
					</TD>
				</TR>
				<TR>
					<TD STYLE = "FONT-SIZE: 10pt" CLASS="TD_RESPOSTA">
						<B>Data: </B> <xsl:value-of select="//data"/>
					</TD>
				</TR>
			</TABLE>
			<BR></BR>
			<BR></BR>

			<TABLE align="center" cellPadding="0" border = "0" cellSpacing="1" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO" width = "20%" > CONTRATANTE </TD>
					<TD  COLSPAN = "5" CLASS = "TD_RESPOSTA" ><font color = "white">.</font><xsl:value-of select="//nomecontratante"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" width = "15%"> ENDEREÇO </TD>
					<TD COLSPAN = "5" CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//endercontratante"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > CIDADE </TD>
					<TD width = "30%" CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//cidadecontratante"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "5%"> UF </TD>
					<TD  CLASS = "TD_RESPOSTA" width = "15%"> <font color = "white">.</font> <xsl:value-of select="//ufcontratante"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "5%"> CEP </TD>
					<TD  CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//cepcontratante"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > TELEFONE </TD>
					<TD width = "25%" CLASS = "TD_RESPOSTA"><font color = "white">.</font> <xsl:value-of select="//telefonecontratante"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "5%"> FAX </TD>
					<TD COLSPAN = "3" CLASS = "TD_RESPOSTA"><font color = "white">.</font> <xsl:value-of select="//faxcontratante"/></TD>
				</TR>
			</TABLE>
			
			<BR></BR>
			<BR></BR>
			
			
			<TABLE align="center"  cellPadding="0"  border = "0" cellSpacing="1" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO"  width = "20%" > CONTRATADA </TD>
					<TD  COLSPAN = "5" CLASS = "TD_RESPOSTA"> <font color = "white">.</font> <xsl:value-of select="//nomecontratada"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" width = "15%"> ENDEREÇO </TD>
					<TD COLSPAN = "5" CLASS = "TD_RESPOSTA"><font color = "white">.</font> <xsl:value-of select="//endercontratada"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > CIDADE </TD>
					<TD width = "30%" CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//cidadecontratada"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "5%"> UF </TD>
					<TD  CLASS = "TD_RESPOSTA" width = "15%"> <font color = "white">.</font> <xsl:value-of select="//ufcontratada"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "5%"> CEP </TD>
					<TD  CLASS = "TD_RESPOSTA"> <font color = "white">.</font> <xsl:value-of select="//cepcontratada"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > TELEFONE </TD>
					<TD width = "25%" CLASS = "TD_RESPOSTA"> <font color = "white">.</font> <xsl:value-of select="//telefonecontratada"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "5%"> FAX </TD>
					<TD COLSPAN = "3" CLASS = "TD_RESPOSTA"> <font color = "white">.</font> <xsl:value-of select="//faxcontratada"/></TD>
				</TR>
			</TABLE>
			<BR></BR>
			<BR></BR>


			<TABLE cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
			<TBODY>
				<TR>
					<TD width="30%">
						<B>TIPO DE SERVIÇO SOLICITADO:</B>
					</TD>
				</TR>
			</TBODY>
			</TABLE>



			<TABLE cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
			<TBODY>
				<TR>  
					<TD>
						( 
							<xsl:if test="//codacao = 1">
								<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
							</xsl:if>
						) Fornecimento 
					</TD>
					<TD>
						( 
							<xsl:if test="//codacao = 4">
								<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
							</xsl:if>
						) Cancelamento 
					</TD>
					<TD>
						( 
							<xsl:if test="//codacao = 2">
								<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
							</xsl:if>
						) Desativação para
					</TD>
				</TR>
				<TR>
						<TD>
						( 
							<xsl:if test="//codacao = 3">
								<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
							</xsl:if>
						) Mudança Velocidade 
					</TD>
					<TD>
						( 
							<xsl:if test="//codacao = 3">
								<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
							</xsl:if>
						) Mudança Endereço 
					</TD>
				</TR>
			</TBODY>
			</TABLE>
			<BR></BR>
			<BR></BR>
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD width = "30%"><B>PRAZO DE VIGÊNCIA DO ACESSO:</B></TD>
				</TR>
				<TR>
					<TD>
						<TABLE align="right" border="0" cellPadding="0" cellSpacing="0" width = "50%">
							<TR>
								<TD>				
									(	
									<xsl:if test="//chkind = 0">
										<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
									</xsl:if>
									)
									Indeterminado
								</TD>
								<TD >
									(	
									<xsl:if test="//chkind = 12">
										<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
									</xsl:if>
									)
									12 Meses
								</TD>
								<TD >
									(	
									<xsl:if test="//chkind = 24">
										<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
									</xsl:if>
									)
									24 Meses
								</TD>
							</TR>
							<TR>
								<TD >
									(	
									<xsl:if test="//chkind = 36">
										<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
									</xsl:if>
									)
									36 Meses
								</TD>
								<TD >
									(	
									<xsl:if test="//chkind = 48">
										<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
									</xsl:if>
									)
									48 Meses
								</TD>
								<TD >
									(	
									<xsl:if test="//chkind = 60">
										<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
									</xsl:if>
									)
									60 Meses
								</TD>
							</TR>
							<TR>
								<TD colspan = "3">
									(	
									<xsl:if test="//chkind = 1">
										<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
									</xsl:if>
									)
									Temporário de :
									<xsl:if test="//chktemporario = -1">
										de<xsl:value-of select="//temporariode"/>
										até <xsl:value-of select="//temporarioate"/>
									</xsl:if>
								</TD>
							</TR>
						</TABLE>
					
					</TD>
				</TR>
			</TABLE>
			
			
			<BR></BR>
			<BR></BR>
			
			
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO_BORDER"  width = "40%" ALIGN="middle"> Endereço Ponta A </TD>
					<TD CLASS = "TD_CAMPO_BORDER"  width = "35%" ALIGN="middle"> Responsável </TD>
					<TD CLASS = "TD_CAMPO_BORDER"  width = "25%" ALIGN="middle"> Telefone </TD>
				</TR>
				<TR>
					<TD  CLASS = "TD_RESPOSTA" >
						<font color = "white">.</font><xsl:value-of select="//pontaAend"/>
					</TD>
					<TD  CLASS = "TD_RESPOSTA" >
						<font color = "white">.</font><xsl:value-of select="//pontaAcontato"/>
					</TD>
					<TD  CLASS = "TD_RESPOSTA" >
						<font color = "white">.</font><xsl:value-of select="//pontaAtelefone"/>
					</TD>
				</TR>				
				<TR>
					<TD CLASS = "TD_CAMPO_BORDER"  width = "40%" ALIGN="middle"> Endereço Ponta B </TD>
					<TD CLASS = "TD_CAMPO_BORDER"  width = "35%" ALIGN="middle"> Responsável </TD>
					<TD CLASS = "TD_CAMPO_BORDER"  width = "25%" ALIGN="middle"> Telefone </TD>
				</TR>				
				<TR>
					<TD  CLASS = "TD_RESPOSTA" >
						<font color = "white">.</font><xsl:value-of select="//pontaBend"/>
					</TD>
					<TD  CLASS = "TD_RESPOSTA" >
						<font color = "white">.</font><xsl:value-of select="//pontaBcontato"/>
					</TD>
					<TD  CLASS = "TD_RESPOSTA" >
						<font color = "white">.</font><xsl:value-of select="//pontaBtelefone"/>
					</TD>
				</TR>			
			</TABLE>
			<BR></BR>
			
			
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO_BORDER" width="25%" ALIGN="middle">Nº Acesso</TD>
					<TD CLASS = "TD_CAMPO_BORDER" width="12%" ALIGN="middle">SEQ</TD>
					<TD CLASS = "TD_CAMPO_BORDER" width="12%" ALIGN="middle">DID</TD>
					<TD CLASS = "TD_CAMPO_BORDER" width="25%" ALIGN="middle">Interface</TD>
					<TD CLASS = "TD_CAMPO_BORDER" ALIGN="middle">Velocidade</TD>
				</TR>
				<TR>
					<TD CLASS = "TD_RESPOSTA"><font color = "white">.</font></TD>
					<TD CLASS = "TD_RESPOSTA"><font color = "white">.</font></TD>
					<TD CLASS = "TD_RESPOSTA"><font color = "white">.</font></TD>
					<TD CLASS = "TD_RESPOSTA"><font color = "white">.</font><xsl:value-of select="//clienteinterface"/></TD>
					<TD CLASS = "TD_RESPOSTA"><font color = "white">.</font><xsl:value-of select="//clientevelocidade"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_RESPOSTA"><font color = "white">.</font></TD>
					<TD CLASS = "TD_RESPOSTA"><font color = "white">.</font></TD>
					<TD CLASS = "TD_RESPOSTA"><font color = "white">.</font></TD>
					<TD CLASS = "TD_RESPOSTA"><font color = "white">.</font></TD>
					<TD CLASS = "TD_RESPOSTA"><font color = "white">.</font></TD>
				</TR>
			</TABLE>
			<BR></BR>
			<BR></BR>
			<B>VIABILIDADE</B>
			<BR></BR>
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD ROWSPAN ="4" align="center" CLASS="TD_CAMPO_BORDER" width="15%" style="WIDTH: 15%">
						Pedido<BR></BR>
						de<BR></BR>
						Viabilidade
					</TD>
					<TD CLASS="TD_CAMPO_CLARO" COLSPAN="2" align="middle">REPRESENTANTE DA EMBRATEL</TD>
					<TD CLASS="TD_CAMPO_CLARO" COLSPAN="2" align="middle">REPRESENTANTE DA CONTRATADA</TD>
				</TR>
				<TR>
					<TD class="TD_RESPOSTA" valign="middle">Assinatura</TD>
					<TD class="TD_RESPOSTA" valign="middle" height="25" WIDTH="30%"><font color = "white">.</font></TD>
					<TD class="TD_RESPOSTA" valign="middle">Assinatura</TD>
					<TD class="TD_RESPOSTA" valign="middle" height="25" WIDTH="30%"><font color = "white">.</font></TD>
				</TR>
				<TR>
					<TD class="TD_RESPOSTA">Nome</TD>
					<TD class="TD_RESPOSTA"><FONT color="white">.</FONT> <xsl:value-of select="//contatoebt"/> </TD>
					<TD class="TD_RESPOSTA">Nome</TD>
					<TD class="TD_RESPOSTA"><FONT color="white">.</FONT><xsl:value-of select="//contatocontratada"/></TD>
					
				</TR>
				<TR>
					<TD class="TD_RESPOSTA">Local/Data</TD>
					<TD class="TD_RESPOSTA"><FONT color="white">.</FONT> <xsl:value-of select="//localdataembratel"/></TD>
					<TD class="TD_RESPOSTA">Local/Data</TD>
					<TD class="TD_RESPOSTA"><FONT color="white">.</FONT> <xsl:value-of select="//localdatacontratada"/></TD>
					
				</TR>
				
				
			</TABLE>
			<BR></BR>
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD ROWSPAN = "5" align="middle" CLASS="TD_CAMPO_BORDER" style="WIDTH:15%" width="15%">
						Resultado<BR></BR>
						de<BR></BR>
						Viabilidade
					</TD>
					<TD CLASS="TD_CAMPO_CLARO" COLSPAN="2">
					<b>OK:</b>(s)(s/n)<br></br>
					Formulário de Provimento de Acesso Dedicado
					</TD>
					<TD CLASS="TD_CAMPO_CLARO" COLSPAN="2">Prazo de Atendimento: 15(dias) da Assinatura deste</TD>
				</TR>
				<TR>
					<TD CLASS="TD_CAMPO_CLARO" COLSPAN="2" align="middle">REPRESENTANTE DA EMBRATEL</TD>
					<TD CLASS="TD_CAMPO_CLARO" COLSPAN="2" align="middle">REPRESENTANTE DA CONTRATADA</TD>
				</TR>
				<TR>
					<TD class="TD_RESPOSTA" valign="middle">Assinatura</TD>
					<TD class="TD_RESPOSTA" valign="middle" height="25" WIDTH="30%"><font color = "white">.</font></TD>
					<TD class="TD_RESPOSTA" valign="middle">Assinatura</TD>
					<TD class="TD_RESPOSTA" valign="middle" height="25" WIDTH="30%"><font color = "white">.</font></TD>
				</TR>
				<TR>
					<TD class="TD_RESPOSTA">Nome</TD>
					<TD class="TD_RESPOSTA"><font color = "white">.</font></TD>
					<TD class="TD_RESPOSTA">Nome</TD>
					<TD class="TD_RESPOSTA"><font color = "white">.</font></TD>
					
				</TR>
				<TR>
					<TD class="TD_RESPOSTA">Local/Data</TD>
					<TD class="TD_RESPOSTA"><font color = "white">.</font></TD>
					<TD class="TD_RESPOSTA">Local/Data</TD>
					<TD class="TD_RESPOSTA"><font color = "white">.</font></TD>
					
				</TR>
				
				
			</TABLE>
			<BR></BR>
			<BR></BR>
			<B>Tipo de Acesso:</B><span>()Fibra</span><span>()Cabo Metálico</span><span>()Rádio</span>
			<BR></BR>
			<BR></BR>
			<B>AUTORIZAÇÃO DE SERVIÇO:</B>
			<BR></BR>
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD ROWSPAN = "4" align="middle" CLASS="TD_CAMPO_BORDER" width="15%" style="WIDTH: 15%">
						Data Prevista de<BR></BR>
						Entrega:<BR></BR>
						<FONT color="white">.</FONT> <xsl:value-of select="//previsaoentrega"/>
						
					</TD>
					<TD CLASS="TD_CAMPO_BORDER" COLSPAN="2" align="middle">REPRESENTANTE DA EMBRATEL</TD>
					<TD CLASS="TD_CAMPO_BORDER" COLSPAN="2" align="middle">REPRESENTANTE DA CONTRATADA</TD>
				</TR>
				<TR>
					<TD class="TD_RESPOSTA" valign="middle">Assinatura</TD>
					<TD class="TD_RESPOSTA" valign="middle" height="25" WIDTH="30%"><font color = "white">.</font></TD>
					<TD class="TD_RESPOSTA" valign="middle">Assinatura</TD>
					<TD class="TD_RESPOSTA" valign="middle" height="25" WIDTH="30%"><font color = "white">.</font></TD>
				</TR>
				<TR>
					<TD class="TD_RESPOSTA">Nome</TD>
					<TD class="TD_RESPOSTA"><FONT color="white">.</FONT> <xsl:value-of select="//contatoebt"/> </TD>
					<TD class="TD_RESPOSTA">Nome</TD>
					<TD class="TD_RESPOSTA"><FONT color="white">.</FONT> <xsl:value-of select="//contatocontratada"/></TD>
					
				</TR>
				<TR>
					<TD class="TD_RESPOSTA">Local/Data</TD>
					<TD class="TD_RESPOSTA"><FONT color="white">.</FONT> <xsl:value-of select="//localdataembratel"/></TD>
					<TD class="TD_RESPOSTA">Local/Data</TD>
					<TD class="TD_RESPOSTA"><FONT color="white">.</FONT> <xsl:value-of select="//localdatacontratada"/></TD>
					
				</TR>

				
				
			</TABLE>
			
			<BR></BR>
			<BR></BR>
			<TABLE align="center" border="1" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD CLASS="TD_CAMPO_CLARO" align="middle">Observação</TD>
					<TD CLASS="TD_CAMPO_CLARO" align="middle" width="70%"><font color = "#edebeb">.</font></TD>
				</TR>
				<TR>
					<TD CLASS="TD_CAMPO_CLARO" align="middle"><font color = "#edebeb">.</font></TD>
					<TD CLASS="TD_CAMPO_CLARO" align="middle" width="70%">
						<span style= "FONT-WEIGHT: normal">
							<font color = "white">.</font><xsl:value-of select="//observaçãos"/><BR/><xsl:value-of select="//observaçãop"/></span>
					</TD>
				</TR>
								
			</TABLE>
		</TD>
	</TR>
</TABLE>
			<P></P>
			</form>
			</BODY>
		</html>
  </xsl:template>
</xsl:stylesheet>