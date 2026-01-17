<?xml version = "1.0" encoding ="ISO-8859-1" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="html" encoding ="ISO-8859-1" />
	<xsl:template match="/">
		<html>
			<head>
			<title>Pedido: <xsl:value-of select="//numero"/></title>
			<style>
				
				.td_negrito
				{
				    FONT-WEIGHT: bold;
				    FONT-SIZE: 8pt;
				    FONT-FAMILY: Tahoma;
				    TEXT-ALIGN: left
				}
				TD
				{
				    FONT-SIZE: 8pt;
				    FONT-FAMILY: Tahoma;
				    TEXT-ALIGN: left
				}
				.td_principal
				{
				    FONT-SIZE: 8pt;
				    FONT-WEIGHT: bold;
				    COLOR: white;
				    FONT-FAMILY: Tahoma;
				    BACKGROUND-COLOR: steelblue;
				    TEXT-ALIGN: center
				}
				.td_titulo
				{
				    border-top: 1px solid ;
				    border-right: 1px solid ;
				    border-bottom: 1px solid ;
				    border-left: 1px solid ;
				    FONT-SIZE: 8pt;
				    FONT-WEIGHT: bold;
				    FONT-FAMILY: Tahoma;
				    BACKGROUND-COLOR: lightgrey;
				    TEXT-ALIGN: center
				}
				INPUT
				{
				    BORDER-RIGHT: thin groove;
				    BORDER-TOP: thin groove;
				    FONT-SIZE: 8pt;
				    BORDER-LEFT: thin groove;
				    BORDER-BOTTOM: thin groove;
				    FONT-FAMILY: Tahoma;
				}
				.td_btn
				{
				    BORDER-RIGHT: thin groove;
				    BORDER-TOP: thin groove;
				    FONT-WEIGHT: bold;
				    FONT-SIZE: 8pt;
				    BORDER-LEFT: thin groove;
				    CURSOR: hand;
				    COLOR: white;
				    BORDER-BOTTOM: thin groove;
				    FONT-FAMILY: Tahoma;
				    BACKGROUND-COLOR: steelblue;
				    TEXT-ALIGN: center
				}
				.td_resposta
				{
				    border-top: 1px solid ;
				    border-right: 1px solid ;
				    border-bottom: 1px solid ;
				    border-left: 1px solid ;
				    FONT-SIZE: 8pt;
				    WIDTH: 100px;
				    FONT-FAMILY: tahoma
				}
				.TD_TELEMAR
				{
				    FONT-WEIGHT: normal;
				    COLOR: darkblue;
				    FONT-STYLE: oblique;
				    FONT-FAMILY: 'Arial Black';
				    FONT-SIZE: 9pt
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
			</style>
			<script>
				function Imprimir()
				{
					window.print();
					//alert("teste");
				}
				function VoltarOrigem()
				{
					try
					{
						with (document.forms[0])
						{	
							//target = self.name 
							action = "facilidade.asp"
							submit()
						}
					}
					catch(e)
					{
						alert(e.description)
					}
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
				<TABLE align="center" cellPadding="0" cellSpacing="0" width="80%" style = "border-top: 1px solid ;border-right: 1px solid ;border-bottom: 1px solid ;border-left: 1px solid ;">
					<TR>
						<TD class="td_principal" align="middle">FORMULÁRIO DE SOLICITAÇÃO DE SERVIÇOS</TD>
					</TR>
					<TR>
						<TD>
							<TABLE BORDER = "0">
								<TR>
									<TD ROWSPAN = "2" ALIGN = "MIDDLE" CLASS = "TD_TELEMAR"> GPON  </TD>
									<TD class="td_negrito" noWrap = "true">Pedido Nº</TD>
									<TD class="td_resposta"><xsl:value-of select="//numero"/></TD>
								</TR>
								<TR>
									<TD class = "td_negrito">Data</TD>
									<TD class="td_resposta"><xsl:value-of select="//familia/data"/></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="td_titulo" align="middle">DADOS DO CLIENTE</TD>
					</TR>
					<TR>
						<TD>
							<TABLE>
								<TR>
									<TD noWrap = "true">Razão Social</TD>
									<TD class="td_resposta" style = "width=450px"><font color = "white">.</font><xsl:value-of select="//empresa/cliente"/></TD>
								</TR>
								<TR>
									<TD>Endereço</TD>
									<TD class="td_resposta" style = "width=450px"><font color = "white">.</font><xsl:value-of select="//empresa/endereco"/></TD>
								</TR>
								<TR>
									<TD>Nº da OE</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//empresa/OE"/></TD>
									<TD>Designação do Serviço</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//empresa/designacao"/></TD>
									<TD noWrap = "true" >Velocidade do Serviço</TD>
									<TD class="td_resposta" ><font color = "white">.</font><xsl:value-of select="//empresa/velocidade"/></TD>
								</TR>
							</TABLE>							
						</TD>
					</TR>
					<TR>
						<TD ALIGN = "MIDLLE" CLASS="TD_TITULO">DADOS TÉCNICOS</TD>
					</TR>
					<TR>
						<TD>		
							<TABLE> 
								<TR>		
									<TD> Fabricante ONT </TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/FabricanteONT"/></TD>
									<TD> PE </TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/PE"/></TD>
								</TR>
								<TR>		
									<TD> Modelo ONT </TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/ModeloONT"/></TD>
									<TD> Porta PE (e também SVLAN) </TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/PortaPE"/></TD>
								</TR>
								<TR>
									<TD nowrap = "true">Designação ONT</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/DesignacaoONT"/></TD>
									<TD nowrap = "true">VLAN (CVLAN)</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/VLAN"/></TD>
								</TR>
								<TR>
									<TD nowrap = "true">Porta ONT</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/PortaONT"/></TD>
									<TD nowrap = "true">Estação</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/Estacao"/></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</form>
				
			</BODY>
		</html>

  </xsl:template>
</xsl:stylesheet>