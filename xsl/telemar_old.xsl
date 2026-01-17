<?xml version = "1.0" encoding ="ISO-8859-1" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="html" encoding ="ISO-8859-1" />
	<xsl:template match="/">
		<html>
			<head>
			<title>Pedido: <xsl:value-of select="//familia/numpedido"/></title>
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

				<TABLE align="center" cellPadding="0" cellSpacing="0" width="80%" style = "border-top: 1px solid ;border-right: 1px solid ;border-bottom: 1px solid ;border-left: 1px solid ;">
					<TR>
						<TD class="td_principal" align="middle">
							FORMULÁRIO DE SOLICITAÇÃO DE SERVIÇOS - ACESSO INDUSTRIAL
						</TD>
					</TR>
					<TR>
						<TD>
							<TABLE BORDER = "0">
								<TR>
									<TD ROWSPAN = "2" ALIGN = "MIDDLE" CLASS = "TD_TELEMAR">TELEMAR</TD>
									<TD class="td_negrito" noWrap = "true">Pedido Nº</TD>
									<TD class="td_resposta"><xsl:value-of select="//familia/numpedido"/></TD>
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
						<TD class="td_titulo" align="middle">
							EMPRESA CONTRATANTE
						</TD>
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
								<TD class="td_resposta" style = "width=450px"><font color = "white">.</font><xsl:value-of select="//empresa/endereço"/></TD>
							</TR>
						</TABLE>
						<TABLE border="0">
							<TR>
								<TD>Cidade</TD>
								<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//empresa/cidade"/></TD>
								<TD>CEP</TD>
								<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//empresa/cep"/></TD>
							</TR>
							<TR>
								<TD>CGC</TD>
								<TD class="td_resposta" style = "width=150px"><font color = "white">.</font><xsl:value-of select="//empresa/cgc"/></TD>
								<TD>UF</TD>
								<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//empresa/uf"/></TD>
								<TD noWrap = "true" >Inscr. Estadual</TD>
								<TD class="td_resposta" ><font color = "white">.</font><xsl:value-of select="//empresa/inscrição"/></TD>
							</TR>
							<TR>
								<TD>Telefone</TD>
								<TD class="td_resposta" style = "width=150px" ><font color = "white">.</font><xsl:value-of select="//empresa/telefone"/></TD>
								<TD colSpan="3" rowSpan="2">
									A contratada está indicada no Anexo único do 
									Convênio ICMS 126/98 (CONFAZ)? 
								</TD>
								<TD class="td_resposta" colSpan="2" rowSpan="2"><font color = "white">.</font><xsl:value-of select="//empresa/icms"/></TD>
								<TD></TD>
							</TR>
							<TR>
								<TD>Fax</TD>
								<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//empresa/fax"/></TD>
								<TD></TD>
								<TD></TD>
								<TD></TD>
								<TD></TD>
							</TR>
						</TABLE>
					</TD>
					</TR>
					<TR>
						<TD class="td_titulo" ALIGN = "MIDDLE">DADOS DA SOLICITAÇÃO</TD>
					</TR>
					<TR>
						<TD>
							<TABLE>
							        <TR>
									<TD class = "td_negrito">Cliente</TD>
									<TD class="td_resposta" style = "width=500px"><font color = "white">.</font><xsl:value-of select="//solicitacao/cliente"/> </TD>
								</TR>
							</TABLE>
							<TABLE>
								<TR>
									<TD class = "td_negrito">Serviço </TD>
									<TD class="td_resposta"  style = "width=250px" ><font color = "white">.</font><xsl:value-of select="//solicitacao/serviço"/></TD>
									<TD class = "td_negrito">Data </TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//solicitacao/data"/> </TD>
									<TD class = "td_negrito">Circuito</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//solicitacao/circuito"/></TD>
								</TR>
							</TABLE>
							<TABLE>
								<TR>
									<TD ><B>DE</B>(caso 3)</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//solicitacao/de"/></TD>
									<TD nowrap = "true"><B>PARA</B>(caso 3)</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//solicitacao/para"/></TD>
									<TD class = "td_negrito" ROWSPAN = "2" VALIGN = "center" align = "center" > Período de contratação</TD>
									<TD>De</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//solicitacao/periodode"/></TD>
								</TR>
								<TR>
									<TD class = "td_negrito" nowrap="true">Tempo de Contratação</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//solicitacao/tempo"/></TD>
									<TD></TD>
									<TD></TD>
									<TD>Até</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//solicitacao/periodoate"/></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TR>
							<TD class="td_titulo" ALIGN = "MIDDLE">
								ENDEREÇO DAS PONTAS
							</TD>
						</TR>
						
						<TD>
							<TABLE>
								<TR>
									<TD class = "td_negrito" NOWRAP = "true">Endereço da PONTA A</TD> 
									<TD class="td_resposta" style = "width=400px" ><font color = "white">.</font><xsl:value-of select="//pontaa/endereço"/></TD>
								</TR>
							</TABLE>
							<TABLE>
								<TR>
									<TD>Cidade</TD>
									<TD class="td_resposta"  style = "width=300px"><font color = "white">.</font><xsl:value-of select="//pontaa/cidade"/></TD>
									<TD>CEP</TD>
									<TD colspan = "3" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/cep"/></TD>
									<TD>UF</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/uf"/></TD>
								</TR>
								<TR>
									<TD width = "10%">Código do Logradouro</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/logradouro"/></TD>
									<TD>Sigla</TD>
									<TD colspan = "3" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/sigla"/></TD>
									<TD width = "10%">Código Site</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/site"/></TD>
								</TR>
								<TR>
									<TD width = "10%">Latitude</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/latitude"/></TD>
									<TD>longitude</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/longitude"/></TD>
									<TD>Estação</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/estação"/></TD>
									<TD width = "10%">Desig E1 Canalizado</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/e1"/></TD>
								</TR>
								<TR>
									<TD>Dentro da ATB</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/atb"/></TD>
									<TD>Ponto de referência</TD>
									<TD colspan = "3" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/referencia"/></TD>
									<TD class = "td_negrito">Time Slot</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/slot"/></TD>
								</TR>
							</TABLE>
							<TABLE width = "100%">
								<TR>
									<TD width = "25%">Nome de Contato</TD>
									<TD width = "25%" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/contato"/></TD>
									<TD width = "25%" >Telefone</TD>
									<TD width = "25%" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontaa/telefone"/></TD>
								</TR>
								<TR>
									<TD class = "td_negrito" >Interface Elétrica</TD>
									<TD width = "25%" class="td_resposta" style = "width=200px"><font color = "white">.</font><xsl:value-of select="//pontaa/eletrica"/></TD>
									<TD class = "td_negrito" >Interface Física</TD>
									<TD class="td_resposta" style = "width=450px" ><font color = "white">.</font><xsl:value-of select="//pontaa/fisica"/>
									</TD>		
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD>
							<TABLE>
								<TR>
									<TD class = "td_negrito" NOWRAP = "true">Endereço da PONTA B</TD> 
									<TD class="td_resposta" style = "width=400px" ><font color = "white">.</font><xsl:value-of select="//pontab/endereço"/></TD>
								</TR>
							</TABLE>
							<TABLE>
								<TR>
									<TD>Cidade</TD>
									<TD class="td_resposta"  style = "width=300px"><font color = "white">.</font><xsl:value-of select="//pontab/cidade"/></TD>
									<TD>CEP</TD>
									<TD colspan = "3" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/cep"/></TD>
									<TD>UF</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/uf"/></TD>
								</TR>
								<TR>
									<TD width = "10%">Código do Logradouro</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/logradouro"/></TD>
									<TD>Sigla</TD>
									<TD colspan = "3" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/sigla"/></TD>
									<TD width = "10%">Código Site</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/site"/></TD>
								</TR>
								<TR>
									<TD width = "10%">Latitude</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/latitude"/></TD>
									<TD>longitude</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/longitude"/></TD>
									<TD>Estação</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/estação"/></TD>
									<TD width = "10%">Desig E1 Canalizado</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/e1"/></TD>
								</TR>
								<TR>
									<TD>Dentro da ATB</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/atb"/></TD>
									<TD>Ponto de referência</TD>
									<TD colspan = "3" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/referencia"/></TD>
									<TD class = "td_negrito">Time Slot</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/slot"/></TD>
								</TR>
							</TABLE>
							<TABLE width = "100%">
								<TR>
									<TD width = "25%">Nome de Contato</TD>
									<TD width = "25%" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/contato"/></TD>
									<TD width = "25%" >Telefone</TD>
									<TD width = "25%" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//pontab/telefone"/></TD>
								</TR>
								<TR>
									<TD class = "td_negrito" >Interface Elétrica</TD>
									<TD width = "25%" class="td_resposta" style = "width=200px"><font color = "white">.</font><xsl:value-of select="//pontab/eletrica"/></TD>
									<TD class = "td_negrito" >Interface Física</TD>
									<TD class="td_resposta" style = "width=450px" ><font color = "white">.</font><xsl:value-of select="//pontab/fisica"/>
									</TD>		
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
									<TD CLASS = "TD_NEGRITO" NOWRAP = "true">Quantidade de linhas</TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/linhas"/></TD>
									<TD CLASS = "TD_NEGRITO"> Modem </TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/modem"/></TD>
								</TR>
									<TD CLASS = "TD_NEGRITO">Velocidade</TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/velocidade"/></TD>
									<TD CLASS = "TD_NEGRITO">Aplicação</TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/aplicação"/></TD>
									<TD CLASS = "TD_NEGRITO">Tecnologia</TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/tecnologia"/></TD>
									<TR>
									<TD nowrap = "true" class="td_negrito">Classe de operação</TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/operacao"/></TD>
									<TD nowrap = "true" class= "td_negrito" >Padrão de qualidade</TD>
									<TD colspan = "3"  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//tecnico/qualidade"/>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD_TITULO">DADOS PARA FATURAMENTO / COBRANÇA</TD>
					</TR>
					<TR>
						<TD>
							<TABLE>
								<TR>
									<TD NOWRAP = "true" >Razão Social</TD>
									<TD class="td_resposta" style = "width = 450px"><font color = "white">.</font><xsl:value-of select="//faturamento/razao"/></TD>
								</TR>
								<TR>
									<TD NOWRAP = "true">Endereço (Rua, nº, bairro)</TD>
									<TD class="td_resposta" style = "width = 450px"><font color = "white">.</font><xsl:value-of select="//faturamento/endereço"/></TD>
								</TR>
							</TABLE>
							<TABLE>
								<TR>
									<TD>Cidade</TD>
									<TD style = "width = 250px" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//faturamento/cidade"/></TD>
									<TD>Cep</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//faturamento/cep"/></TD>
									<TD>UF</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//faturamento/uf"/></TD>		
								</TR>
								<TR>
									<TD>Nome de Contato</TD>
									<TD width = "40%" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//faturamento/contato"/></TD>
									<TD>Telefone</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//faturamento/telefone"/></TD>
									<TD>Fax</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//faturamento/fax"/></TD>
								</TR>
							</TABLE>
							<TABLE>
								<TR>
									<TD class ="td_negrito">Data de vencimento da fatura</TD>
									<TD  width = "30%" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//faturamento/vencimento"/></TD>
									<TD class ="td_negrito">Prazo para pagamento  </TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//faturamento/prazo"/></TD>
									<TD colspan = "2">dias</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD CLASS = "TD_TITULO">DADOS COMPLEMENTARES</TD>
					</TR>
					<TR>
						<TD>
							<TABLE width = "100%">
								<TR>
									<TD>Representante legal da contratante</TD>
									<TD  class="td_resposta" style ="width=250px"><font color = "white">.</font><xsl:value-of select="//complemento/representante"/></TD>
									<TD>E-mail</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//complemento/e-mail"/></TD>
								</TR>
							</TABLE>
							<TABLE width = "100%">
								<TR>
									<TD>Telefone fixo</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//complemento/telefone"/></TD>
									<TD>Fax</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//complemento/fax"/></TD>
									<TD ROWSPAN = "2" COLSPAN = "2" STYLE = "BORDER-RIGHT: thin solid; BORDER-TOP: thin solid;BORDER-LEFT: thin solid;BORDER-BOTTOM: thin solid;width:150px">
									Assinatura:
									</TD>
								</TR>
								<TR>
									<TD>Endereço</TD>
									<TD colspan = "3" class="td_resposta" style ="width=350px"><font color = "white">.</font><xsl:value-of select="//complemento/endereço"/></TD>
									<TD></TD>
									<TD></TD>
								</TR>
								<TR>
									<TD>Cidade</TD>
									<TD COLSPAN = "3" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//complemento/cidade"/></TD>
									<TD>UF</TD>
									<TD class="td_resposta"><xsl:value-of select="//complemento/uf"/>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD STYLE = "FONT-WEIGHT:bold;COLOR:steelblue; FONT-FAMILY: Tahoma;TEXT-ALIGN: center;border-top: 1px solid ;border-right: 1px solid ;border-bottom: 1px solid ;border-left: 1px solid;">
							Preenchimento exclusivo da Telemar
						</TD>
					</TR>
					<TR>
						<TD class = "td_titulo">
							SUPORTE A CANAIS
						</TD>
					</TR>
					<TR>
						<TD>
							<TABLE> 
								<TR>
									<TD nowrap = "true">Elemento PEP nº</TD>
									<TD width = "10%" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//suporte/pep"/></TD>
									<TD nowrap = "true">Nº da LP original (casos 2 ou 3 da solicitação do cliente)</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//suporte/lp"/></TD>		
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class = "td_titulo">
							DADOS PARA FATURAMENTO DA COBRANÇA
						</TD>
					</TR>
					<TR>
						<TD>
							<TABLE>
								<TR>
									<TD CLASS = "TD_NEGRITO">Portfólio</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//cobranca/portifolio"/></TD>
									<TD CLASS = "TD_NEGRITO">Projeto de cobrança</TD>
									<TD colspan = "3"  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//cobranca/projeto"/></TD>
								</TR>
								<TR>
									<TD CLASS = "TD_NEGRITO">Manual</TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//cobranca/manual"/></TD>
									<TD CLASS = "TD_NEGRITO">Conta</TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//cobranca/conta"/></TD>
									<TD CLASS = "TD_NEGRITO">Nº da Conta</TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//cobranca/numconta"/></TD>
								</TR>
							</TABLE>
							<TABLE>
								<TR>
									<TD nowrap = "true"> <b>Taxa da Instalação</b>(Valor, em R$)</TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//cobranca/taxa"/> </TD>
									<TD nowrap = "true"><b>Mensalidade</b>(Valor em R$)</TD>
									<TD  class="td_resposta"><font color = "white">.</font><xsl:value-of select="//cobranca/mensalidade"/></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class = "td_titulo">
							DADOS COMPLEMENTARES
						</TD>
					</TR>
					<TR>
						<TD>
							<TABLE>
								<TR>
									<TD NOWRAP = "true" >Responsável pela conta do cliente </TD>
									<TD class="td_resposta" style = "width=400px"><font color = "white">.</font><xsl:value-of select="//telemar/responsavel"/></TD>
									<TD NOWRAP = "true" ><B>Anexo I - FSS_AGO/02 REV 01</B></TD>
								</TR>
							</TABLE>
							<TABLE>	
								<TR>
									<TD>Telefone Fixo</TD>
									<TD class="td_resposta"><font color = "white">.</font><xsl:value-of select="//telemar/telefone"/></TD>
									<TD>Fax</TD>
									<TD WIDTH = "20%" class="td_resposta"><font color = "white">.</font><xsl:value-of select="//telemar/fax"/></TD>
									<TD>E-mail</TD>
									<TD class="td_resposta" style ="width=200px"><font color = "white">.</font><xsl:value-of select="//telemar/email"/></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
   					<TR STYLE = "BORDER-RIGHT: thin solid; BORDER-TOP: thin solid;BORDER-LEFT: thin solid;BORDER-BOTTOM: thin solid;">
						<TD>
							<TABLE>
								<TR>
									<TD STYLE = "BORDER-RIGHT: thin solid; BORDER-TOP: thin solid;BORDER-LEFT: thin solid;BORDER-BOTTOM: thin solid;">
										<TABLE >
											<TR>
												<TD CLASS = "TD_NEGRITO">Circuito para interconexão?</TD>
												<TD style = "font-size:9px">(sim ou Não?)</TD>
												<TD class="td_resposta" style ="width=40px" ><font color = "white">.</font><xsl:value-of select="//telemar/circuito"/></TD>
											</TR>
											<TR>
												<TD colspan = "3" CLASS = "TD_NEGRITO">Preechimento exclusivo da Telemar</TD>
											</TR>
											<TR>
												<TD>Parecer da área de Interconexão:</TD>
												<TD style = "font-size:9px">Previsto no PTI</TD>
												<TD class="td_resposta" style ="width=40px" ><font color = "white">.</font><xsl:value-of select="//telemar/interconexao"/></TD>
											</TR>	
											<TR>
												<TD colspan = "2">Data de ativação prevista no PTI</TD>
												<TD class="td_resposta" style ="width=40px" ><font color = "white">.</font><xsl:value-of select="//telemar/data"/></TD>
											</TR>
										</TABLE>	
									</TD>
									<TD VALIGN = "TOP" STYLE = "BORDER-RIGHT: thin solid; BORDER-TOP: thin solid;BORDER-LEFT: thin solid;BORDER-BOTTOM: thin solid;">
										<TABLE>
											<TR>
												<TD ><b>Observações ou comentários do cliente ou da Telemar:</b></TD>
											</TR>
											<TR>
												<TD style = "font-size:10px" ><xsl:value-of select="//telemar/fraseobs"/></TD>
											</TR> 
											<TR>
												<TD class="td_resposta" style = "width:310px;height:70px;TEXT-ALIGN:center" valign= "center"><font color = "white">.</font><xsl:value-of select="//telemar/obs"/></TD>
											</TR>
											
										</TABLE>	
									</TD>
								</TR>
							</TABLE>		
						</TD>
					</TR>
				</TABLE>
			</BODY>
		</html>
  </xsl:template>
</xsl:stylesheet>
