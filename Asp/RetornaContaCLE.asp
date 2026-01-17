<!--#include file="../inc/data.asp"-->
<%
	dim objXmlDoc 
	dim strHTML, strConta15  
	dim mdCodRetorno
	dim msRazaoSocial
	dim msNomeFantasia
	
	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	
	'Atribuição de valores para as variáveis 	
	objXmlDoc.load(Request)
	set ndConta15	= objXmlDoc.selectSingleNode("//conta15")
	strConta15 = ndConta15.Text
	
	'''
	'''Set obj035 = CreateObject("EOL.CLE035")
  '''obj035.ServerAddress = "10.4.2.1:5026@RPC/SRVPCLE/CALLNAT" '10.4.2.1:5026@RPC/SRVPCLE/CALLNAT' //Prod  //Desenv 10.2.4.83:5024@RPC/SRVCLE/CALLNAT
 ''' obj035.Logon
  '**********ContaSubConta = strConta15'"000001655200000"
 ''' obj035.CLESP035 CDbl(strConta15), msRazaoSocial, msNomeFantasia, _
  '''                    dTipoDocumento, msNumeroDocumentoCliente, _
 '''                     msInscricaoEstadual, msInscricaoMunicipal, _
  '''                    mdCodigoCNL, mdCodigoLogradouro, _
  '''                    mdCodigoSite, mdCodigoLocal, _
  '''                    msCodigoLocalidade, msTipoEndereco, _
  ''''''                    msEndereco, msNumeroEndereco, _
  '''                    msComplementoEndereco, msBairro, _
  '''                    msCidade, msUF, mdCep, msDDDTelefone, _
  '''                    msNumeroTelefone, msRamalTelefone, msDDDFax, _
  '''                    msNumeroFax, msMatriculaCS, msEmail, _
 '''                     msNomeContatoCom, msDDDTelefoneCom, _
 '''                     msNumeroTelefoneCom, msRamalTelefoneCom, _
 '''                     msDDDFaxCom, msNumeroFaxCom, msEmailCom, _
 '''                     msCargoCom, msNomeContatoFin, msDDDTelefoneFin, _
  '''                    msNumeroTelefoneFin, msRamalTelefoneFin, _
   '''                   msDDDFaxFin, msNumeroFaxFin, msEmailFin, _
  '''                    msCargoFin, msNomeContatoTec, msDDDTelefoneTec, _
   '''                   msNumeroTelefoneTec, msRamalTelefoneTec, msDDDFaxTec, _
  '''                    msNumeroFaxTec, msEmailTec, msCargoTec, mdCodCNLCps, _
  '''                    mdCodLogrCps, mdCodSiteCps, mdCodLocalCps, _
  '''                    msCodigolocalidadeCps, msTipoEnderecoCps, msEnderecoCps, _
   '''                   msNumeroEnderecoCps, msComplementoEnderecoCps, msBairroCps, msCidadeCps, _
  '''                    msUFCps, mdCEPCps, msCodCgcSacado, msInscricaoEstadualSacado, _
  '''                    msInscricaoMunicipalSacado, mdCodRetorno, sDiretoriaPrincipal, sDiretoriaAssociada
  
  	
			'Set Tronco = CreateObject("EOL.CLESP22O")	
			
			'Set Tronco = CreateObject("EOL.CLESP22O")	
  			'Tronco.ServerAddress = "ETBHMGBA.NT.EMBRATEL.COM.BR:1971@RPC/SRVHCLE/CALLNAT" '10.4.2.1:5026@RPC/SRVPCLE/CALLNAT' //Prod  //Desenv 10.2.4.83:5024@RPC/SRVCLE/CALLNAT
  			
			
			Set Tronco = CreateObject("EOL.CLESP22O")

  			'Alterado por Fabio Pinho em 26/04/2016 - Inicio - ver 1.0'
  			'Produção'
  			'10.4.2.1:5026@RPC/SRVPCLE/CALLNAT' //Prod  //Desenv 10.2.4.83:5024@RPC/SRVCLE/CALLNAT
  			Tronco.ServerAddress = "ETBPRDBA.NT.EMBRATEL.COM.BR:1971@RPC/SRVPCLE/CALLNAT" 
  			
  			'Homologação'
  			'Tronco.ServerAddress = "ETBHMGBA.NT.EMBRATEL.COM.BR:1971@RPC/SRVHCLE/CALLNAT"
			'Alterado por Fabio Pinho em 26/04/2016 - Fim - ver 1.0''

			Tronco.Logon
			Tronco.CLESP220 strConta15, msRazaoSocial , msNomeFantasia , SEGMENTO	, PORTE , mdCodRetorno , TXT-MSG
			
			'response.write "<script>alert('"&strRazaoSocial&"')</script>"
			'response.write "<script>alert('"&strNomeFantasia&"')</script>"
			'response.write "<script>alert('"&SEGMENTO&"')</script>"
			'response.write "<script>alert('"&PORTE&"')</script>"
			'response.write "<script>alert('"&COD-RETORNO&"')</script>"
			'response.write "<script>alert('"&TXT-MSG&"')</script>"
	
	if isnull(msRazaoSocial) or trim(msRazaoSocial)="" then
		msRazaoSocial="*"
	end if
	if isnull(msNomeFantasia) or trim(msNomeFantasia)="" then
		msNomeFantasia="*"
	end if
	if isnull(SEGMENTO) or trim(SEGMENTO)="" then
		SEGMENTO="*"
	end if		
	if isnull(PORTE) or trim(PORTE)="" then
		PORTE="*"
	end if	
	if isnull(mdCodRetorno) or trim(mdCodRetorno)="" then
		mdCodRetorno="*"
	end if	
	if isnull(TXT-MSG) or trim(TXT-MSG)="" then
		TXT-MSG="*"
	end if		
	
	'PORTE= "*"

	'strXML = " <?xml version=1.0 encoding=UTF-8 ?> "
	strXML =  "<root>"
	strXML = strXML + "<razaoSocial><![CDATA["  +  TRIM(CSTR(msRazaoSocial))      + "]]></razaoSocial>"
	strXML = strXML + "<nomeFantasia><![CDATA[" +  TRIM(CSTR(msNomeFantasia))     + "]]></nomeFantasia>"
	strXML = strXML + "<segmento><![CDATA[" +  TRIM(CSTR(SEGMENTO))     + "]]></segmento>"
	'strXML = strXML + "<porte><![CDATA[" +  PORTE     + "]]></porte>"	
	
	
	strXML = strXML + "<porte>"   +  TRIM(CSTR(PORTE)) + "</porte>"	
	strXML = strXML + "<codRetorno>"   +  TRIM(CSTR(mdCodRetorno)) + "</codRetorno>"
	strXML = strXML + "<msg>"   +  TRIM(CSTR(TXT-MSG)) + "</msg>"
	strXML = strXML + "</root>"	
	
	
	'doc1.save(Server.MapPath("1123-out-completo.xml"))
	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strXML)
%>