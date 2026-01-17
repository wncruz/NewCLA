<!--#include file="../inc/data.asp"-->

<% 

dim dblSolId 
dim dblPedId 
dim dblProId
dim dblEscEntrega 
dim intTipoProcesso 
dim objRSped1
dim objRSPontoB
dim strData
dim objRSPro 
dim strProEmail, strProNome, strParmProc
dim strServico, strAssunto
dim varUF, intCont
dim strTemporario

dim ndPed 
dim ndSol 
dim ndProv
dim ndEsc 
dim ndTipo

varUF = array("AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO")


set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	
	'Atribuição de valores para as variáveis 	
	
objXmlDoc.load(Request)
strCaminho = server.MapPath("..\")


set ndPed =  objXmlDoc.selectSingleNode("//ped")
set ndSol  =  objXmlDoc.selectSingleNode("//sol")
set ndProv =  objXmlDoc.selectSingleNode("//Prov")
'set ndEsc =  objXmlDoc.selectSingleNode("//Esc")
set ndTipo =  objXmlDoc.selectSingleNode("//ndTipo")

dblPedId  = ndPed.Text
dblSolId  = ndSol.Text
dblProId = ndProv.Text
'dblEscEntrega = ndEsc.Text
intTipoProcesso = ndTipo.Text
DblNdTipo = cint(ndTipo.text)

	
'set objRSped1 = db.execute("CLA_sp_view_cartaprovedor " & dblPedId & " , " & dblProId)
set objRSped1 = db.execute("CLA_sp_view_cartaprovedor " & dblPedId )
set objRSPontoB = db.execute("CLA_sp_view_ponto null, " & dblPedId )

Set objRSPro = db.execute("CLA_sp_sel_provedoremail " & dblProId & ",null,'" & objRSped1("Est_Sigla") &"','"& objRSped1("Cid_Sigla") &"'") 
if Not objRSPro.Eof and Not objRSPro.bof then
	strProEmail = Trim(objRSPro("Cpro_Contratadaemail"))
	strProNome	= Trim(objRSPro("Pro_Nome"))
	if not isnull(Trim(objRSPro("Cpro_Contratadaemail"))) then 
		strFromEmail = Trim(objRSPro("Cpro_Contratanteemail"))
	else
		strFromEmail = "acessos@embratel.com.br"
	end if 
End if

strParmProc = dblProId + "|" + dblPedId  + "|" + intTipoProcesso + "|" + strUserName 


		'--> PSOUTO 10/05/2006 
	if DblNdTipo <> 2 then
		DblNdTipo = objRSped1("tprc_id")  
	END IF 

	'PSOUTO 10/05/2006
	'select case objRSped1("tprc_id") '--> 
	SELECT CASE DblNdTipo
 
case 1 
	strServico = "Ativação"
case 2 
	strServico = "Retirada"
case 3
	dim strVelAntiga
	dim ObjMudanca
	set ObjMudanca = db.execute ("CLA_sp_sel_CartaMudanca " & dblSolId & " , " & objRSped1("Acf_id"))
								
	if not ObjMudanca.eof then  						
		  strVelAntiga = ObjMudanca("Log_Valor")
	end if
	set	 ObjMudanca =  nothing 
	strServico = "Alteração"
case 4 
	blnCancelamento = true
	strServico = "Cancelamento"
end select 

							
IF  isnull(objRSped1("Acl_DtIniAcessoTemp")) THEN
	strTemporario = "Permanente"
ELSE
	strTemporario = "Temporário"
END IF 


strAssunto = strServico & "  -  " & trim(objRSped1("Cli_nome")) & "  -  " & ucase(objRSped1("Ped_Prefixo")) & "-" & right("00000" & objRSped1("Ped_Numero"),5) & "/" & objRSped1("Ped_Ano")

DIM   strXML
	
	//Msgusuario.innerText = "Aguarde..."
	strXML = "<?xml version = ""1.0"" encoding =""ISO-8859-1""?><root>"
	strXML = strXML & "<arquivo>Telemar</arquivo>"
	strXML = strXML & "<parmproc>" & strParmProc & "</parmproc>"
	strXML = strXML & "<assunto>" & strAssunto & "</assunto>"
	strXML = strXML & "<proemail>" & strProEmail & "</proemail>"
	strXML = strXML & "<mailfrom>" & strFromEmail &  " </mailfrom>"
	strXML = strXML & "<pronome>"  & strProNome & "</pronome>"
	strXML = strXML & "<familia><numpedido>" & objRSped1("ped_Prefixo") & "-" & objRSped1("ped_Numero") & "/" & objRSped1("ped_Ano") & "</numpedido>"
	strXML = strXML & "<numcontrato>CDCT 0002/02</numcontrato>"
	strXML = strXML & "<data>" & FormatDateTime(objRSped1("ped_data"), vbshortdate) & "</data>"
	strXML = strXML & "<circuito>TC Data(EILD) s/ICMS</circuito>"
	strXML = strXML & "</familia>"
	strXML = strXML & "<empresa>"
	strXML = strXML & "<cliente>EMPRESA BRASILEIRA DE TELECOMUNICAÇÕES - EMBRATEL S/A</cliente>" 
	strXML = strXML & "<endereço>RUA DOS INGLESES, 600 - 3ºANDAR</endereço>" 
	strXML = strXML & "<cidade>SÃO PAULO</cidade>" 
	strXML = strXML & "<cep>01329-904</cep>" 
	strXML = strXML & "<cgc>33.530.486/0125-69</cgc>" 
	strXML = strXML & "<uf>SP</uf>"
	strXML = strXML & "<inscrição>108240571-119</inscrição>" 
	strXML = strXML & "<telefone>(11) 2121-2690</telefone>" 
	strXML = strXML & "<fax>(11) 2121-2154</fax>" 
	strXML = strXML & "<icms></icms>" 
	strXML = strXML & "</empresa>"
	strXML = strXML & "<solicitacao>"
	strXML = strXML & "<cliente>" &  objRSped1("Cli_nome") & "</cliente>"
	strXML = strXML & "<serviço>" & strServico & "</serviço>"
	strXML = strXML & "<data>" & strData & "</data>"
	strXML = strXML & "<circuito>" & strTemporario & "</circuito>"	
	strXML = strXML & "<de>"& strVelAntiga & "</de>"
	strXML = strXML & "<para>" & objRSped1("VEL_DESC") & "</para>"
	
			dim TpoCircuito
			strData = ""
			if not  isnull(objRSped1("Acl_DtIniAcessoTemp")) then 
				strData =  FormatDateTime(objRSped1("Acl_DtIniAcessoTemp")) 
			end if 
			
	strXML = strXML & "<periodode>" & strData & "</periodode>"
		
			strData = ""
			if not isnull(objRSped1("Acl_DtFimAcessoTemp")) then 
				strData =  FormatDateTime(objRSped1("Acl_DtFimAcessoTemp")) 
			end if 

	strXML = strXML & "<periodoate>" & strData & "</periodoate>"
	
			if isnull(objRSped1("Tct_Meses"))  then 
				strPrazo = ""
			else 
				if objRSped1("Tct_Meses") = "0" then 
					strPrazo = "Indeterminado"
				else
					strPrazo = objRSped1("Tct_Meses") & " Meses"
				end if 
			end if 
	 
	strXML = strXML & "<tempo>" & strPrazo & "</tempo>"
	strXML = strXML & "</solicitacao>"
	strXML = strXML & "<pontaa>"
	strXML = strXML & "<endereço>" & trim(objRSped1("TPL_SIGLA")) & " " & trim(objRSped1("esc_nomelogr")) & " Nº" & objRSped1("esc_nrologr") & " "  & objRSped1("esc_bairro") & " " & objRSped1("esc_complemento")  & "</endereço>"  
	strXML = strXML & "<cidade>" & objRSped1("cid_desc") & "</cidade>" 
	strXML = strXML & "<cep>" & objRSped1("esc_cod_cep") & "</cep>" 
	strXML = strXML & "<uf>" & objRSped1("est_sigla") & "</uf>" 
	strXML = strXML & "<logradouro>" & objRSped1("cid_desc")& "</logradouro>"
	strXML = strXML & "<sigla>" & objRSped1("CID_SIGLA") & "</sigla>" 
	strXML = strXML & "<site></site>"  
	strXML = strXML & "<latitude></latitude>" 
	strXML = strXML & "<longitude></longitude>"  
	strXML = strXML & "<estação>" & objRSped1("cid_sigla") & "-" & objRSped1("esc_sigla") & "</estação>"  
	strXML = strXML & "<e1>" & objRSped1("FAC_LINK") & "</e1>" 
	strXML = strXML & "<atb></atb>"
	strXML = strXML & "<referencia></referencia>"
	
	dim objRSFac
	dim strAnt
	dim strXMLaux  
	dim strRepresentacao
	dim strRepresentacaoTot
	
	Set objRSFac= db.Execute("CLA_SP_Sel_Facilidade " & dblPedId)
	intCont =0 
	strRepresentacaoTot = ""
	strRepresentacao = ""
	if not objRSFac.eof then 
		do While not objRSFac.Eof 
			if not isNull(objRSFac("Fac_Representacao")) then
				strRepresentacao = objRSFac("Fac_Representacao")
			else
				Select Case objRSped1("SIS_ID")
					Case 1
						strRepresentacao = objRSFac("Fac_TimeSlot")
					Case Else	
						strRepresentacao = objRSFac("Fac_Par")
				End Select	
			End if
			if strRepresentacao <>  strAnt then 
				if strRepresentacaoTot = "" then 
					strRepresentacaoTot = strRepresentacaoTot &  strRepresentacao
				else
					strRepresentacaoTot = strRepresentacaoTot & "-" &  strRepresentacao
				end if 
			end if 
			strAnt =  strRepresentacao
			intCont =  intCont + 1 
			objRSFac.movenext
		loop
	END IF 
	set objRSFac= nothing 
	
	strXML = strXML & "<slot>"& strRepresentacaoTot &"</slot>"
	strXML = strXML & "<contato>" & objRSped1("esc_contato") & "</contato>" 
	strXML = strXML & "<telefone>" & objRSped1("esc_telefone") & "</telefone>"
	strXML = strXML & "<eletrica></eletrica>"
	strXML = strXML & "<fisica>" & objRSped1("Acl_InterfaceEst") & "</fisica>"
	strXML = strXML & "</pontaa>"
	strXML = strXML & "<pontab>"
	strXML = strXML & "<endereço>" & trim(objRSPontoB("TPL_SIGLA")) & " " & trim(objRSPontoB("End_NomeLogr")) & " Nº" & trim(objRSPontoB("End_NroLogr")) & " " & trim(objRSPontoB("End_Bairro")) & " " & trim(objRSPontoB("Aec_Complemento")) & "</endereço>"  
	strXML = strXML & "<cidade>" & objRSPontoB("CID_DESC")   & "</cidade>" 
	strXML = strXML & "<cep>" & objRSPontoB("end_Cep") & "</cep>" 
	strXML = strXML & "<uf>" & objRSPontoB("est_sigla") & "</uf>"
	strXML = strXML & "<logradouro>" & objRSPontoB("CID_DESC") & "</logradouro>"
	strXML = strXML & "<sigla>" & objRSPontoB("CID_SIGLA") & "</sigla>" 
	strXML = strXML & "<site></site>" 
	strXML = strXML & "<latitude></latitude>" 
	strXML = strXML & "<longitude></longitude>"  
	strXML = strXML & "<estação></estação>"  
	strXML = strXML & "<e1></e1>" 
	strXML = strXML & "<atb></atb>"
	strXML = strXML & "<referencia></referencia>"
	strXML = strXML & "<slot></slot>"
	strXML = strXML & "<contato>" & objRSPontoB("Aec_Contato") & "</contato>" 
	strXML = strXML & "<telefone>" & objRSPontoB("Aec_Telefone") & "</telefone>"
	strXML = strXML & "<eletrica></eletrica>"
	strXML = strXML & "<fisica>" & objRSPontoB("AcF_Interface") & "</fisica>"
	strXML = strXML & "</pontab>"
	strXML = strXML & "<tecnico>"
	strXML = strXML & "<linhas></linhas>" 
	strXML = strXML & "<modem>4 fios</modem>"
	strXML = strXML & "<velocidade>" & objRSped1("VEL_DESC") & "</velocidade>"
	strXML = strXML & "<aplicação>Acesso</aplicação>" 
	'strXML = strXML & "<tecnologia>" & objRSped1("TEC_SIGLA") & "</tecnologia>"
	'Campo tecnologia de acordo com a definição deve possuir valor fixo 
	strXML = strXML & "<tecnologia>Par Metalico</tecnologia>"
	
	if isnull(objRSped1("Acf_Tipovel")) = 1  then 
		strOperacao = "Assincrono"
	else
		strOperacao = "Sincrono"
	end if 
	'strXML = strXML & "<operacao>" & strOperacao & "</operacao>"
	'Campo com valor fixo segundo a definição
	strXML = strXML & "<operacao></operacao>" 
	strXML = strXML & "<qualidade></qualidade>"
	strXML = strXML & "</tecnico>"
	strXML = strXML & "<faturamento>" 
	strXML = strXML & "<razao>EMPRESA BRASILEIRA DE TELECOMUNICAÇÕES EMBRATEL S/A</razao>"
	'Campos alterados de acordo com a definição recebida.
	'strXML = strXML & "<endereço>" & trim(objRSped1("TPL_SIGLA")) & " " & trim(objRSped1("esc_nomelogr")) & " " & objRSped1("esc_nrologr") & " " & objRSped1("esc_complemento") & "</endereço>"   
	'strXML = strXML & "<cidade>" & objRSped1("cid_desc") & "</cidade>"
	'strXML = strXML & "<cep>" & objRSped1("esc_cod_cep") & "</cep>" 
	'strXML = strXML & "<uf>" & objRSped1("est_sigla") & "</uf>"  
	'strXML = strXML & "<contato>" & objRSped1("esc_contato") & "</contato>"
	'strXML = strXML & "<telefone>" & objRSped1("esc_telefone") &"</telefone>"  
	strXML = strXML & "<endereço>RUA DOS INGLESES, 600 - 7º ANDAR </endereço>"   
	strXML = strXML & "<cidade>SÃO PAULO</cidade>"
	strXML = strXML & "<cep>01329-904</cep>" 
	strXML = strXML & "<uf>SP</uf>"  
	strXML = strXML & "<contato>JORGE MATSUDA</contato>"
	strXML = strXML & "<telefone>(11) 2121-2380</telefone>"  
	strXML = strXML & "<fax></fax>" 
	strXML = strXML & "<vencimento></vencimento>"
	strXML = strXML & "<prazo></prazo>"
	strXML = strXML & "</faturamento>"
	strXML = strXML & "<complemento>" 
			dim objRepresGLA 
			set objRepresGLA = db.execute("CLA_sp_view_agentesolicitacao " & dblSolId )
			dim strNomeRepresent , strEmailRepresent , strTelefone
						
			do while not objRepresGLA.eof
				if UCASE(objRepresGLA("AGE_DESC")) = "GLA" then
					  strNomeRepresent   = objRepresGLA("USU_NOME") 
					  strEmailRepresent	 = objRepresGLA("USU_EMAIL") 
					  strTelefone = objRepresGLA("USU_RAMAL") 
				end if 
				objRepresGLA.MoveNext
			loop
						
			set objRepresGLA = nothing 
	'strXML = strXML & "<representante>" & strNomeRepresent & "</representante>"  
	'strXML = strXML & "<e-mail>" & strEmailRepresent & "</e-mail>"  
	'strXML = strXML & "<uf>" & objRSped1("est_sigla") & "</uf>" 
	strXML = strXML & "<representante>RUBENS NALLIN FILHO</representante>"  
	strXML = strXML & "<e-mail>RUBENSN@EMBRATEL.COM.BR</e-mail>"  
	strXML = strXML & "<telefone>(11) 2121-2169</telefone>" 
	strXML = strXML & "<fax>(11) 2121-2154</fax>"
	strXML = strXML & "<endereço>RUA DOS INGLESES, 600 - 3º ANDAR</endereço>" 
	strXML = strXML & "<cidade>SÃO PAULO</cidade>" 
	strXML = strXML & "<uf>SP</uf>" 
	strXML = strXML & "</complemento>" 
	strXML = strXML & "<suporte>" 
	strXML = strXML & "<pep></pep>"
	'psouto 10/05/2006
	'if objRSped1("tprc_id")  <>  "1" then 
	if DblNdTipo  <>  "1" then 
	
		strXML = strXML & "<lp>" & objRSped1("Acf_NroAcessoPtaEbt") & "</lp>"
	else
		strXML = strXML & "<lp></lp>"
	end if 
	strXML = strXML & "</suporte>" 
	strXML = strXML & "<cobranca>" 
	strXML = strXML & "<portifolio></portifolio>"
	strXML = strXML & "<projeto></projeto>"
	strXML = strXML & "<manual></manual>"
	strXML = strXML & "<conta>Conta Customizada </conta>"
	strXML = strXML & "<numconta>18500/3106417</numconta>"
	strXML = strXML & "<taxa></taxa>" 
	strXML = strXML & "<mensalidade></mensalidade>" 
	strXML = strXML & "</cobranca>" 
	strXML = strXML & "<telemar>" 
	'@@
	strXML = strXML & "<responsavel></responsavel>" 
	strXML = strXML & "<telefone></telefone>" 
	strXML = strXML & "<fax></fax>" 
	strXML = strXML & "<email></email>"
	strXML = strXML & "<circuito></circuito>"
	strXML = strXML & "<interconexao></interconexao>"
	strXML = strXML & "<data></data>" 
	strXML = strXML & "<obs>TRATA-SE DE DO CONTRATO DE SERVICO DE EILD UARJ.21 - 022/98 DE 25/03/1998 </obs>"
	strXML = strXML & "</telemar>" 
	strXML = strXML & "</root>"

	set objRSped1 = nothing 
	set objRSPontoB = nothing 
	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strXML)

%>
