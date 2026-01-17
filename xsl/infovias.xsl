<?xml version="1.0" encoding="ISO-8859-1" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="html" encoding ="ISO-8859-1"/>
	<xsl:template match="/">
	<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=ISO-8859-1"/>
<TITLE>Pedido: </TITLE>
<STYLE>
				
				.observacao
				{
					FONT-SIZE: 10pt;
				    FONT-FAMILY: arial;
				    FONT-WEIGHT: bold;
				    color: darkred;
				    align:left;
				    OVERFLOW: visible;
				    TEXT-ALIGN: left;
				}
				td
				{
					FONT-SIZE: 8pt;
				    FONT-FAMILY: arial
				
				}
				.comun
				{
					font-size:8pt;
					color: black;
					font-family: arial
				}
				.campo
				{
					FONT-SIZE: 8pt;
				    FONT-FAMILY: arial;
				    FONT-WEIGHT: bold
				    
				}
				.valor{
					FONT-SIZE: 8pt;
				    FONT-FAMILY: arial
				    
				    
				}
				.SEPARADOR
				{
					BORDER-BOTTOM:DARKBLUE SOLID 4PX
				}
				.TD_CAMPO 
				{
					FONT-SIZE: 8pt;
				    FONT-FAMILY: Arial
				}

				.TD_CAMPOFIM 
				{
					FONT-SIZE: 8pt; BORDER-BOTTOM: #000000 1px solid; FONT-FAMILY: Arial
				}
	
				.TD_PED 
				{
					FONT-WEIGHT: bold; FONT-SIZE: 12pt; COLOR: red; BORDER-BOTTOM: #000000 2px solid; FONT-FAMILY: Arial
				}
	
				.TD_EMBRATEL 
				{
					FONT-WEIGHT: bold; FONT-SIZE: 14pt; COLOR: blue; BORDER-BOTTOM: #000000 2px solid; FONT-FAMILY: Arial
				}

				.TD_NEGRITO 
				{
					FONT-WEIGHT: bold; FONT-SIZE: 8pt; FONT-FAMILY: Arial
				}
	
				.TD_SUBTIT 
				{
			
				    FONT-WEIGHT: bold;
				    FONT-SIZE: 10pt;
				    FONT-FAMILY: arial;
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
<BODY leftMargin="0" topMargin="0"  marginwidth="0" marginheight="0">
<form name="form1">
<TABLE cellSpacing="3" cellPadding="0" width="650PX" align="CENTER" border="0" class="comun" >
	<tr>
		<TD class="SEPARADOR" >
			<img src="../imagens/infovias.gif" width="650px"></img>
		</TD>
		
	</tr>
	<tr>
		<TD class="separador">
			<TABLE width="100%" style="font-size=12px"> 
				<TR>
					<TD COLSPAN="2" class="campo" style="font-size:10pt">
						Cliente:EMBRATEL
					</TD>
					<TD class="campo" style="font-size:10pt">Data da OS:</TD>
					<TD  style="font-size:10pt"><xsl:value-of select="//data"/></TD>
					<TD class="campo" style="font-size:10pt">Nº OS:</TD>
					<TD  style="font-size:10pt"><xsl:value-of select="//numero"/></TD>
				</TR>
				<TR>
					<TD class="campo" style="font-size:10pt">Ref.Viabilidade Nº:</TD>
					<TD  style="font-size:10pt"></TD>
					<TD class="campo" style="font-size:10pt">Data pra entrega do circuito:</TD>
					<TD  style="font-size:10pt"><xsl:value-of select="//dataprevista"/>	</TD>
					<TD class="campo" style="font-size:10pt">Contrato Nº:</TD>
					<TD  style="font-size:10pt"><xsl:value-of select="//contrato"/>	</TD>
				</TR>
			</TABLE>
		</TD>
	</tr>
	<tr>
		<td>
			<table width="98%" align="center">
				<tr>
					<td class="campo" width="135px">Empresa Contratante:</td>
					<td  colspan="3"><xsl:value-of select="//nomecontratante"/></td>
					
				</tr>
				<tr>
					<td class="campo" >Natureza do Pedido:</td>
					<td  colspan="3">
						<TABLE cellSpacing="0" cellPadding="0" width="100%" align="center" border="0" bgcolor="#c0c0c0">
							<TBODY >
								<TR >  
									<TD  bgcolor="#c0c0c0">
										( 
											<xsl:if test="//codacao = 1">
												<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
											</xsl:if>
										) Ativação
									</TD>
									<TD bgcolor="#c0c0c0">
										( 
											<xsl:if test="//codacao = 4">
												<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
											</xsl:if>
										) Cancelamento
									</TD>
									<TD bgcolor="#c0c0c0">
										( 
											<xsl:if test="//codacao = 3">
												<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
											</xsl:if>
										) Mudança de Velocidade
									</TD>
									
								</TR>
								<TR >
									
									<TD bgcolor="#c0c0c0">
										( 
											<xsl:if test="//codacao = 3">
												<span style = "FONT-FAMILY: 'verdana'"><B>x</B></span>
											</xsl:if>
										) Mudança de Endereço
									</TD>
									<TD bgcolor="#c0c0c0">
										( ) Mudança Interna
									</TD>
								</TR>
							</TBODY>
						</TABLE>
					
					</td>
					
				</tr>
				<tr>
					<td class="campo">Contato Comercial:</td>
					<td width="150px"><xsl:value-of select="//contatoebt"/></td>
					<td class="campo">Tel:</td>
					<td ><xsl:value-of select="//telefoneebt"/></td>
					
				</tr>
				<tr>
					<td class="campo" >E-Mail para Contato:</td>
					<td colspan="3"><xsl:value-of select="//mailfrom"/></td>
					
				</tr>
				<tr>
					<td class="campo">Contato Técnico:</td>
					<td ><xsl:value-of select="//contatogla"/></td>
					<td class="campo">Tel:</td>
					<td ><xsl:value-of select="//telefonegla"/></td>
				</tr>
				<tr>
					<td class="campo" >E-Mail para Contato:</td>
					<td  colspan="3">
					<xsl:value-of select="//emailgla"/></td>
					
				</tr>
				<tr>
					<td class="campo"></td>
					<td ></td>
					<td class="campo"></td>
					<td ></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td alIgn="center" STYLE="font-size:16pt; color:darkred">
		SDH/PDH
		</td>
	</tr>
	<tr>
		<td style="font-weight:bold">
		Site A
		</td>
	</tr>
	<tr>
		<td>
			
			<table style="border:solid 2px black" width="90%" align="center" >
				<tr>
					<td COLSPAN="5">
						<SPAN class="campo">Operadora:</SPAN>
						<font color = "white">.</font><xsl:value-of select="//pontaAnome"/>
					</td>
					<td><span class="campo">CNPJ:</span><font color = "white">.</font><xsl:value-of select="//pontaAcnpj"/></td>
					
				</tr>
				<tr>
					<td COLSPAN="6">
						<SPAN class="campo">Endereço:</SPAN>
						<font color = "white">.</font><xsl:value-of select="//pontaAend"/>
					</td>
					
				</tr>
				<tr>
					<td COLSPAN="3">
						<SPAN class="campo">Cidade:</SPAN>
						<font color = "white">.</font><xsl:value-of select="//pontaAcid"/>
					</td>
					
					<td  colspan="3">
						<span class="campo">Interface no Site A:</span>
						<font color = "white">.</font><xsl:value-of select="//clienteinterface"/>
					</td>
					
				</tr>
				<tr>
					<td COLSPAN="3" WIDTH="200PX">
						<SPAN class="campo" >Contato:</SPAN>
						<font color = "white">.</font><xsl:value-of select="//pontaAcontato"/>
					
						<span class="campo" WIDTH="150PX"> Tel:</span>
						<font color = "white">.</font><xsl:value-of select="//pontaAtelefone"/>
					
					</td>
					<td colspan="3" >
						<span class="campo">E-Mail:</span>
						<font color = "white">.</font><xsl:value-of select="//pontaAemail"/>
					</td>
					
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="separador"></td>
	</tr>
	<tr>
		<td></td>
	</tr>
	<tr>
		<td style="font-weight:bold">
		Site B
		</td>
	</tr>
	<tr>
		<td>
			
			<table style="border:solid 2px black" width="90%" align="center">
				<tr>
					<td COLSPAN="5">
						<SPAN class="campo">Cliente Final:</SPAN>
						<font color = "white">.</font><xsl:value-of select="//pontabnome"/>
					</td>
					
					<td>
						<span class="campo">CNPJ:</span>
						<font color = "white">.</font><xsl:value-of select="//pontabcnpj"/>
					</td>
				</tr>
				<tr>
					<td COLSPAN="6">
						<SPAN class="campo">Endereço:</SPAN>
						<font color = "white">.</font><xsl:value-of select="//pontaBend"/>
					</td>
				</tr>
				<tr>
					<td COLSPAN="3">
						<SPAN class="campo">Cidade:</SPAN>
						<font color = "white">.</font><xsl:value-of select="//pontabcid"/>
					</td>
					<td  colspan="3">
						<span class="campo">Degrau:</span>
					</td>
					
				</tr>
				<tr>
					<td COLSPAN="3" >
						<SPAN class="campo" >Contato Cliente:</SPAN>
						<xsl:value-of select="//pontabcontato"/>
						<span class="campo" WIDTH="50PX"> Tel:</span>
						<xsl:value-of select="//pontaBtelefone"/>
					</td>
					<td colspan="3" >
						<span class="campo">E-Mail:</span>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td></td>
	</tr>
	<tr>
		<td>
			<TABLE width="100%">
				<TR>
					<TD COLSPAN="10" class="campo">CARACTERÍSTICAS DO SERVIÇO</TD>
					
				</TR>
				<TR>
					<TD></TD>
					<TD></TD>
					<TD class="campo" >Rede:</TD>
					<TD></TD>
					<TD  colspan="2">
						<span class="campo">Interface Site B:</span>
						<font color = "white">.</font><xsl:value-of select="//clienteinterface"/>
					</TD>
					<TD  colspan="2">
						<span class="campo">Quantidade:</span>
						<font color = "white">.</font><xsl:value-of select="//quantidade"/>
					</TD>
					<TD colspan="2">
						<span class="campo">Velocidade:</span>
						<font color = "white">.</font><xsl:value-of select="//velacessofisico"/>
					</TD>
				</TR>
				<TR>
					<TD class="campo">Custos:</TD>
					<TD colspan= "2">Valor Mensal s/Impostos</TD>
					<TD colspan="3"></TD>
					<TD class="campo">Prazo de Contratação:</TD>
					<TD colspan="3"><font color = "white">.</font><xsl:value-of select="//chkind"/></TD>
					
					
					
				</TR>
				<TR>
					<TD class="campo"></TD>
					<TD colspan= "2">Taxa de Ativação s/Impostos</TD>
					<TD colspan="3"></TD>
					<TD class="campo"></TD>
					<TD colspan="3"></TD>
					
				</TR>

				
			</TABLE>
		</td>
	</tr>
	<TR>
		<TD></TD>
	</TR>
	<tr>
		<td style="font-weight:bold">
		Observações:
		</td>
	</tr>
	<tr>
		<td>
			
			<table style="border:solid 2px black" width="80%" align="center">
				<tr>
					<td class="observacao" HEIGHT="50PX"><font color = "white">.</font><xsl:value-of select="//observaçãos"/><BR/><xsl:value-of select="//observaçãop"/></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="separador"></td>
	</tr>
	<tr>
		<td></td>
	</tr>
	<tr>
		<td>
		
		<p>O Cliente abaixo assinado, por meio desta Ordem de Serviço, solicita os serviços aqui descritos no prazo informado acima, mediante as condições comerciais e técnicas definidas no Contrato que ampara esta solicitação de serviços.</p>
		<p>Todos os preços definidos neste instrumento são líquidos de Impostos (PIS, COFINS, ISS e ICMS), os quais serão adicionados na época do faturamento nas alíquotas então vigentes.</p>
		</td>
	</tr>
	
	<tr>
		<td></td>
	</tr>
	<tr>
		<td style="FONT-STYLE: italic">Assinaturas Autorizadas</td>
	</tr>
	<tr>
		<td></td>
	</tr>
	<tr>
		<td>
			<table WIDTH="80%" align="center">
				<tr>
					<td STYLE="BORDER-BOTTOM:BLACK SOLID 1PX"></td>
					<td width="20%"></td>
					<td STYLE="BORDER-BOTTOM:BLACK SOLID 1PX"></td>
				</tr>
				<tr>
					<td colspan="3"></td>
				</tr>
				<tr>
					<td colspan="3"></td>
				</tr>
				<tr>
					<td align="center">________________________</td>
					<td></td>
					<td align="center">________________________</td>
				</tr>
			
				<tr>
					<td align="center">Empresa de Infovias</td>
					<td></td>
					<td align="center">EMBRATEL</td>
				</tr>
			</table>
		</td>
	</tr>
	
	
	<tr>
		<td class="separador"></td>
	</tr>
	
	
	
</TABLE>
<br></br>
</form>
</BODY>
</HTML>



  </xsl:template>
</xsl:stylesheet>