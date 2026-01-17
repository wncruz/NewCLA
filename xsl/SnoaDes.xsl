<?xml version = "1.0" encoding ="ISO-8859-1" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="html" encoding ="ISO-8859-1" />
	<xsl:template match="/">
		<html>
			<head>
			<title>Número do Bilhete Snoa: <xsl:value-of select="//numero"/></title>
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
					xmlhttp.Open("POST", "EnviaEmailSnoa.asp", false);
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
						<TD class="td_principal" align="middle">FORMULÁRIO DE SOLICITAÇÃO DE RETIRADA / CANCELAMENTO</TD>
					</TR>
					<TR>
						<TD>
							<TABLE BORDER = "0">
								<TR>
									<TD class="td_negrito" noWrap = "true">Nº bilhete SNOA</TD>
									<TD class="td_resposta"><xsl:value-of select="//numero"/></TD>
									<TD class="td_negrito" noWrap = "true" >Nº do contrato correspondente</TD>
									<TD class="td_resposta" STYLE="width=220px"><xsl:value-of select="//familia/numcontrato"/></TD>
								</TR>
								<TR>
									<TD class = "td_negrito">Data</TD>
									<TD class="td_resposta"><xsl:value-of select="//familia/data"/></TD>
									<TD class = "td_negrito">Tipo de Circuito</TD>
									<TD class="td_resposta" STYLE="width=220px"><xsl:value-of select="//familia/circuito"/></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="td_titulo" align="middle">EMPRESA CONTRATANTE</TD>
					</TR>
					<TR>
						<TD>
							<TABLE BORDER = "0">
							<TR>
								<TD class="td_negrito" noWrap = "true">Razão Social   </TD>
								<TD class="td_resposta" style = "width:540px;height:10px;TEXT-ALIGN:left"><xsl:value-of select="//empresa/cliente"/></TD>
							</TR>
							<TR>
								<TD class="td_negrito">Endereço   </TD>
								<TD class="td_resposta" style = "width:540px;height:10px;TEXT-ALIGN:left" ><xsl:value-of select="//empresa/endereço"/></TD>
							</TR>
						</TABLE>
							<TABLE >
								<TR>
									<TD class="td_negrito" >Cidade</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//empresa/cidade"/></TD>
									<TD class="td_negrito" >CEP</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//empresa/cep"/></TD>
									<TD class="td_negrito" noWrap = "true" >Inscr. Municipal</TD>
									<TD class="td_resposta" ><font color = "white">.</font><xsl:value-of select="//empresa/inscricao_municipal"/></TD>
								</TR>
								<TR>
									<TD class="td_negrito" >CNPJ</TD>
									<TD class="td_resposta" ><font color = "white">.</font><xsl:value-of select="//empresa/cnpj"/></TD>
									<TD class="td_negrito" >UF</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//empresa/uf"/></TD>
									<TD class="td_negrito" noWrap = "true" >Inscr. Estadual</TD>
									<TD class="td_resposta" ><font color = "white">.</font><xsl:value-of select="//empresa/inscrição"/></TD>
								</TR>
								<TR>
									<TD class="td_negrito" >Telefone</TD>
									<TD class="td_resposta" ><font color = "white">.</font><xsl:value-of select="//empresa/telefone"/></TD>
									<TD class="td_negrito" >EMAIL</TD>
									<TD class="td_resposta"  ><font color = "white">.</font><xsl:value-of select="//empresa/email"/></TD>
								</TR>	
														
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="td_titulo" ALIGN = "MIDDLE">DADOS DA SOLICITAÇÃO</TD>
					</TR>
					<TR>
						<TD>
							<TABLE >
							    <TR>
									<TD class="td_negrito" >Cliente</TD>
									<TD class="td_resposta" style = "width:540px;height:10px;TEXT-ALIGN:left" ><font color = "white">.</font><xsl:value-of select="//solicitacao/cliente"/> </TD>
								</TR>
								<TR>
									<TD class="td_negrito" >Serviço </TD>
									<TD class="td_resposta" style = "width:540px;height:10px;TEXT-ALIGN:left"  ><font color = "white">.</font><xsl:value-of select="//solicitacao/serviço"/></TD>
									<TD class="td_negrito" >Data </TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//solicitacao/data"/> </TD>
								</TR>
								<TR>
									<TD class="td_negrito" >Taxa de Transmissão</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/velocidade"/></TD>
									<TD class="td_negrito" nowrap = "true">Quantidade de linhas</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/linhas"/></TD>
									
								</TR>
								<TR>
									<TD class="td_negrito" nowrap="true">Motivo do Cancelamento</TD>
									<TD class="td_resposta" style = "width:540px;height:10px;TEXT-ALIGN:left" ><font color = "white">.</font><xsl:value-of select="//solicitacao/motivo"/></TD>
									
								</TR>
								<TR>
									<TD class="td_negrito" nowrap="true">DESIGNAÇÃO</TD>
									<TD class="td_resposta" style = "width:540px;height:10px;TEXT-ALIGN:left" ><font color = "white">.</font><xsl:value-of select="//tecnico/designacao"/></TD>
									
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="td_titulo" ALIGN = "MIDDLE" >Ponta A</TD>
					</TR>
					<TR>
						<TD>
							<TABLE >
							
								<TR>
									<TD class="td_negrito" >Contato p/ agendamento</TD>
									<TD class = "td_resposta" style = "width:540px;height:10px;TEXT-ALIGN:left" ><xsl:value-of select="//pontaa/contato"/></TD>
									<TD class="td_negrito"  >Telefone</TD>
									<TD class = "td_resposta" COLSPAN = "3"><xsl:value-of select="//pontaa/telefone"/></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
						
					<TR>
						<TD class="td_titulo" ALIGN = "MIDDLE">Ponta B</TD>
					</TR>
					<TR>
						<TD>
							<TABLE >	
								<TR>
									<TD class="td_negrito" >Contato p/ agendamento</TD>
									<TD class = "td_resposta" style = "width:540px;height:10px;TEXT-ALIGN:left" ><xsl:value-of select="//pontab/contato"/></TD>
									<TD class="td_negrito" >Telefone</TD>
									<TD class = "td_resposta" COLSPAN = "3"><xsl:value-of select="//pontab/telefone"/></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD STYLE = "FONT-WEIGHT:bold;COLOR:steelblue; FONT-FAMILY: Tahoma;TEXT-ALIGN: center;border-top: 1px solid ;border-right: 1px solid ;border-bottom: 1px solid ;border-left: 1px solid;"></TD>
					</TR>
					<TR>
						<TD class = "td_titulo">OBSERVAÇÕES OU COMENTÁRIOS COMPLEMENTARES</TD>
					</TR>
					
   					<TR STYLE = "BORDER-RIGHT: thin solid; BORDER-TOP: thin solid;BORDER-LEFT: thin solid;BORDER-BOTTOM: thin solid;">
						<TD>
							<TABLE align="center" cellPadding="0" cellSpacing="0" width="100%">
								<TR>
									<TD VALIGN = "TOP" STYLE = "BORDER-RIGHT: thin solid; BORDER-TOP: thin solid;BORDER-LEFT: thin solid;BORDER-BOTTOM: thin solid;">
										<TABLE width="750">											
											<TR>
												<TD class="td_resposta" style = "width:940px;height:25px;TEXT-ALIGN:center" valign= "center"><font color = "white">.</font> <xsl:value-of select="//observacao"/></TD>
											</TR>											
										</TABLE>	
									</TD>
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