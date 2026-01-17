<?xml version = "1.0" encoding ="ISO-8859-1" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="html" encoding ="ISO-8859-1"/>
	<xsl:template match="/">
		<html>
			<head>
			<title>Pedido: <xsl:value-of select="//numero"/></title>
			<style>
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
			<TABLE align="center" class ="table_interno" cellPadding="0" cellSpacing="0" width = "80%">
			<TR><TD>
			<TABLE align="center" border = "0"  cellPadding="0" cellSpacing="1" width = "100%">
				<TR>
					<TD ROWSPAN = "3" STYLE = "FONT-SIZE:14PT" CLASS = "TD_EMBRATEL" ALIGN = "CENTER" VALIGN = "CENTER"  >Embratel</TD>
					<TD ROWSPAN = "3" VALIGN = "CENTER" ALIGN = "CENTER" STYLE = "FONT-SIZE:10PT" width ="40%" CLASS = "TD_RESPOSTA">
						<B>FORMULÁRIO DE PROVIMENTO DE SERVIÇO DE ACESSO DEDICADO</B>
					</TD>
					<TD  STYLE = "FONT-SIZE:10PT" CLASS = "TD_RESPOSTA">
						<B>Nº:</B><xsl:value-of select="//numero"/>
					</TD>
				</TR>
				<TR>
					<TD STYLE = "FONT-SIZE:10PT" CLASS = "TD_RESPOSTA" >
						<B>Data:</B> <xsl:value-of select="//data"/>
					</TD>
				</TR>
				<TR>
					<TD STYLE = "FONT-SIZE:10PT" CLASS = "TD_RESPOSTA">
						<B>Contrato:</B><xsl:value-of select="//contrato"/>
					</TD>
				</TR>
			</TABLE>
			<BR></BR>

			<TABLE align="center" cellPadding="0" border = "0" cellSpacing="1" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO" width = "20%" > CONTRATANTE </TD>
					<TD  COLSPAN = "5" CLASS = "TD_RESPOSTA" ><font color = "white">.</font> <xsl:value-of select="//nomecontratante"/></TD>
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
			<P></P>
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
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD width = "30%"><B>PRAZO DE VIGÊNCIA DA CONTRATAÇÃO DO ACESSO:</B></TD>
				</TR>
			</TABLE>
			<BR></BR>
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD >
										
						(	
						<xsl:if test="//chkind = -1">
							<span style = "font-family:'verdana'"><B>x</B></span>
						</xsl:if>
						)
						Indeterminado
					</TD>
					<TD >
					
						(	
						<xsl:if test="//chk12meses = -1">
							<span style = "font-family:'verdana'"><B>x</B></span>
						</xsl:if>
						)
						12 Meses
					</TD>
					<TD >
						(	
						<xsl:if test="//chk24meses = -1">
							<span style = "font-family:'verdana'"><B>x</B></span>
						</xsl:if>
						)
						24 Meses
					</TD>
					<TD >
						(	
						<xsl:if test="//chk36meses = -1">
							<span style = "font-family:'verdana'"><B>x</B></span>
						</xsl:if>
						)
						36 Meses
					</TD>
					<TD >
						(	
						<xsl:if test="//chk48meses = -1">
							<span style = "font-family:'verdana'"><B>x</B></span>
						</xsl:if>
						)
						48 Meses
					</TD>
					<TD >
						(	
						<xsl:if test="//chk60meses = -1">
							<span style = "font-family:'verdana'"><B>x</B></span>
						</xsl:if>
						)
						60 Meses
					</TD>
				</TR>
				<TR>
					<TD colspan = "3">
						(	
						<xsl:if test="//chktemporario = -1">
							<span style = "font-family:'verdana'"><B>x</B></span>
						</xsl:if>
						)
						<b>Temporário:</b>
						<xsl:if test="//chktemporario = -1">
							de<xsl:value-of select="//temporariode"/>
							até <xsl:value-of select="//temporarioate"/>
						</xsl:if>
					</TD>
				</TR>
			</TABLE>
			<BR></BR>
			<BR></BR>
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD><B>DADOS DE INSTALAÇÃO DO ACESSO(PONTA A = CLIENTE)</B></TD>
				</TR>
			</TABLE>
			<BR></BR>
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="1" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO"  width = "20%"> NOME DO CLIENTE </TD>
					<TD  COLSPAN = "5" CLASS = "TD_RESPOSTA" ><font color = "white">.</font><xsl:value-of select="//clientenome"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO"  > ENDEREÇO </TD>
					<TD  COLSPAN = "5" CLASS = "TD_RESPOSTA"><font color = "white">.</font><xsl:value-of select="//clienteend"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO"  > CIDADE </TD>
					<TD CLASS = "TD_RESPOSTA" width = "30%"> <font color = "white">.</font><xsl:value-of select="//clientecidade"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "5%"> UF </TD>
					<TD CLASS = "TD_RESPOSTA" width = "15%"> <font color = "white">.</font><xsl:value-of select="//clienteuf"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "5%"> CEP </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clientecep"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO"  > CONTA CORRENTE </TD>
					<TD  COLSPAN = "5" CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clienteccta"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > CGC ou CNPJ </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clientecnpj"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" COLSPAN = "3"> INSCRIÇÃO MUNICIPAL </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clienteiem"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > INSCRIÇÃO ESTADUAL </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clienteiee"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO"  COLSPAN = "3"> PESSOA DE CONTATO </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clientecontato"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > TELEFONE CONTATO</TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clientetelefone"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" COLSPAN = "3"> INTERFACE </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clienteinterface"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > VELOCIDADE</TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clientevelocidade"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" COLSPAN = "3"> DATA BACKLOG </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clienteback"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > NÚMERO DA SEV</TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clientesev"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" COLSPAN = "3"> PRAZO DE ENTREGA SEV </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//clienteprazo"/></TD>
				</TR>
			</TABLE>
			<BR></BR>
			<BR></BR>
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="0" width = "100%">
				<TR>
					<TD><B>DADOS DE INSTALAÇÃO DO ACESSO(PONTA B = EMBRATEL)</B></TD>
				</TR>
			</TABLE>
			<BR></BR>	
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="1" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO"  width = "20%" > NOME  </TD>
					<TD  CLASS = "TD_RESPOSTA" COLSPAN = "5"> <font color = "white">.</font><xsl:value-of select="//pontabnome"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO"  > ENDEREÇO </TD>
					<TD  CLASS = "TD_RESPOSTA" COLSPAN = "5"> <font color = "white">.</font><xsl:value-of select="//pontabend"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO"  > CIDADE </TD>
					<TD CLASS = "TD_RESPOSTA" width = "30%"> <font color = "white">.</font><xsl:value-of select="//pontabcidade"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "5%"> UF </TD>
					<TD CLASS = "TD_RESPOSTA" width = "15%"> <font color = "white">.</font><xsl:value-of select="//pontabuf"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "5%"> CEP </TD>
					<TD CLASS = "TD_RESPOSTA" > <font color = "white">.</font><xsl:value-of select="//pontabcep"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > CGC ou CNPJ </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//pontabcnpj"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" COLSPAN = "3" > INSCRIÇÃO MUNICIPAL </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" > INSCRIÇÃO ESTADUAL </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font></TD>
					<TD CLASS = "TD_CAMPO_CLARO" COLSPAN = "3" > PESSOA DE CONTATO </TD>
					<TD CLASS = "TD_RESPOSTA"> <font color = "white">.</font><xsl:value-of select="//pontabcontato"/></TD>
				</TR>
			</TABLE>
			<BR></BR>
			<BR></BR>
			
			
			<TABLE align="center" cellPadding="0" border = "0" cellSpacing="1" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" width = "100%" > <B> PACOTE MIGRAÇÃO: </B></TD>
				</TR>
			</TABLE>
			<BR></BR>
			
			<TABLE align="center" cellPadding="0" border = "0" cellSpacing="1" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" width = "20%"> (x) Acesso Novo </TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "20%"> (<font color = "white">.</font>) Migração de Custos </TD>
					<TD CLASS = "TD_CAMPO_CLARO" width = "20%"> (<font color = "white">.</font>) Projeto de Otimização </TD>
				</TR>
			</TABLE>
			<BR></BR>
			<BR></BR>
			
			<TABLE align="center" cellPadding="0" border = "0" cellSpacing="1" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" width = "100%" > <B> INFORMAÇÕES COMPLEMENTARES: </B></TD>
				</TR>
			</TABLE>
			<BR></BR>
			
			<TABLE align="center" cellPadding="0" border = "0" cellSpacing="1" width = "100%">
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" width = "20%" > CONTRATO </TD>
					<TD CLASS = "TD_RESPOSTA" width = "30%"> <font color = "white">.</font> <xsl:value-of select="//servicocontrato"/></TD>
					<TD CLASS = "TD_CAMPO_CLARO" COLSPAN = "3" width = "20%"> SERVIÇO </TD>
					<TD CLASS = "TD_RESPOSTA" width = "30%"> <font color = "white">.</font> <xsl:value-of select="//servicodescricao"/></TD>
				</TR>
				<TR>
					<TD CLASS = "TD_CAMPO_CLARO" width = "20%" > DESIGNAÇÃO DO CCTO </TD>
					<TD CLASS = "TD_RESPOSTA" width = "30%"> <font color = "white">.</font> <xsl:value-of select="//designacaoservico"/> </TD>
					<TD CLASS = "TD_CAMPO_CLARO" COLSPAN = "3" width = "20%" > E-MAIL GI SOLICITANTE </TD>
					<TD CLASS = "TD_RESPOSTA" width = "30%"> <font color = "white">.</font> <xsl:value-of select="//emailgicl"/></TD>
				</TR>
			</TABLE>
			<BR></BR>


<!-- LPEREZ 13/12/2005-->			
			<TABLE align="center" border="0" cellPadding="0" cellSpacing="1" width = "100%">
				<TR>
					<TD class= "TD_CAMPO_CLARO" width = "20%"  height="50px"  ALIGN = "LEFT" VALIGN = "TOP">
						<B>OBSERVAÇÕES:</B>
					</TD>
					<TD CLASS = "TD_RESPOSTA"  width= "80%" height="50px"  STYLE = "height:80px" valign = "top">
						<span style= "FONT-WEIGHT:normal"><font color = "white">.</font><xsl:value-of select="//observaçãos"/><br/><font color = "white">.</font><xsl:value-of select="//observaçãop"/></span>
					</TD>
				</TR>
			</TABLE>
<!-- LP -->			
			<p></p>
			<p></p>
			<TABLE align="left" border="0" cellPadding="0" cellSpacing="1" width = "100%">
				<TR>
					<TD ALIGN = "LEFT" VALIGN = "TOP">
						<B>Elaborado por: </B><xsl:value-of select="//contatogla"/>
						<br></br>
						<B>Responsável: </B><xsl:value-of select="//contatoebt"/>
						<br></br><xsl:value-of select="//cargoebt"/>
					</TD>
					
				</TR>
			</TABLE>
			</TD></TR>
			</TABLE>
			<P></P>
			</BODY>
		</html>
  </xsl:template>
</xsl:stylesheet>
