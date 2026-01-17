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
DIM DblNdTipo

dim strVlan
dim strPE
dim strPorta
dim strLink

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
DblNdTipo = cint(ndTipo.text)
'dblEscEntrega = ndEsc.Text
intTipoProcesso = ndTipo.Text

	
'set objRSped1 = db.execute("CLA_sp_view_cartaprovedor " & dblPedId & " , " & dblProId)
set objRSped1 = db.execute("CLA_sp_view_cartaprovedor " & dblPedId )

if objRSped1.eof then 
	Response.Write ("<table width=100% ><tr><td style=""text-align:center""><font color = red>Pedido não encontrado.</font></td></tr></table>")
	Response.end()
end if 

set objRSPontoB = db.execute("CLA_sp_view_ponto null, " & dblPedId & ",null," & objRSped1("Sol_ID"))

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
	strXML = strXML & "<arquivo>BSODLIGHT</arquivo>"
	strXML = strXML & "<parmproc>" & strParmProc & "</parmproc>"
	strXML = strXML & "<assunto>" & strAssunto & "</assunto>"
	strXML = strXML & "<proemail>" & strProEmail & "</proemail>"
	strXML = strXML & "<mailfrom>" & strFromEmail &  " </mailfrom>"
	strXML = strXML & "<pronome>"  & strProNome & "</pronome>"
	'alterado por PSOUTO 18/05/2006
	strXML = strXML &	"<numero>"
	IF intTipoProcesso <> 2 THEN ' NÃO É DESATIVACAO(LIBERACAO DE ESTOQUE)
	strXML = strXML &	objRSped1("ped_Prefixo") & "-" & objRSped1("ped_Numero") & "/" & objRSped1("ped_Ano") 
	END IF 
	strXML = strXML & "</numero>"
	'strXML = strXML &	"<numero>"& objRSped1("ped_Prefixo") & "-" & objRSped1("ped_Numero") & "/" & objRSped1("ped_Ano") & "</numero>"
	' /PSOUTO
	strXML = strXML & "<familia><numcontrato>" & objRSPro("CPro_ContratadaContrato") &  "</numcontrato>"
	strXML = strXML & "<data>" & FormatDateTime(objRSped1("ped_data"), vbshortdate) & "</data>"
	strXML = strXML & "<circuito>TC Data(EILD) s/ICMS</circuito>"
	strXML = strXML & "</familia>"
	strXML = strXML & "<empresa>"


	strXML = strXML & "<cliente> " & objRSPro("Cpro_ContratanteRazao") & "</cliente>" 
	strXML = strXML & "<endereço> " & objRSPro("Cpro_ContratanteEnd") & "</endereço>" 
	strXML = strXML & "<cidade> " & objRSPro("ContratanteCid") & "</cidade>" 
	strXML = strXML & "<cep> "& objRSPro("Cpro_ContratanteCep") & "</cep>" 
	strXML = strXML & "<cgc> "& objRSPro("Cpro_ContratanteCGC_CNPJ") & "</cgc>" 
	strXML = strXML & "<uf>"& objRSPro("Cpro_ContratanteEstSigla")& "</uf>"
	strXML = strXML & "<inscrição>81617341</inscrição>" 
	strXML = strXML & "<telefone> "& objRSPro("CPro_ContratanteTelefone") & "</telefone>" 
	strXML = strXML & "<fax> " & objRSPro("CPro_ContratanteFax")& "</fax>" 


''	strXML = strXML & "<cliente>EMPRESA BRASILEIRA DE TELECOMUNICAÇÕES - EMBRATEL S/A</cliente>" 
''  	strXML = strXML & "<endereço>RUA CAMERINO, 96 - SALA 206 - CENTRO</endereço>" 
''	strXML = strXML & "<cidade>RIO DE JANEIRO</cidade>" 
''	strXML = strXML & "<cep>20080-010</cep>" 
''	strXML = strXML & "<cgc>33.350.496/0001-29</cgc>" 
''	strXML = strXML & "<uf>RJ</uf>"
''	strXML = strXML & "<inscrição>81617341</inscrição>" 
''	strXML = strXML & "<telefone>2121-9655 / 21216040</telefone>" 
''	strXML = strXML & "<fax>2121-7950</fax>" 
''	strXML = strXML & "<icms></icms>" 
	strXML = strXML & "</empresa>"
	strXML = strXML & "<solicitacao>"
	strXML = strXML & "<DesignacaoServico>" & objRSped1("Acl_DesignacaoServico") & "</DesignacaoServico>"
	strXML = strXML & "<DescVelAcessoLog>" & objRSped1("DescVelAcessoLog") & "</DescVelAcessoLog>"
	strXML = strXML & "<cliente>" &  objRSped1("Cli_nome") & "</cliente>"
	strXML = strXML & "<serviço>" & strServico & "</serviço>"
	strXML = strXML & "<data>" & date() & "</data>"
	strXML = strXML & "<circuito>" & strTemporario & "</circuito>"	
	strXML = strXML & "<de>"& strVelAntiga & "</de>"
	strXML = strXML & "<para>" & objRSped1("VEL_DESC") & "</para>"
	strXML = strXML & "<cnpj> " & objRSped1("Aec_CNPJ") & "</cnpj>" 
	
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
	strXML = strXML & "<endereço>" & trim(objRSped1("TPL_SIGLA")) & " " & trim(objRSped1("esc_nomelogr")) & " Nº" & objRSped1("esc_nrologr") & " " & objRSped1("esc_complemento") & " "  & objRSped1("esc_bairro")  & "</endereço>"  
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
	
	IF DblNdTipo = 2 OR DblNdTipo = 4 THEN
		Set objRSFac= db.Execute("CLA_SP_Sel_Facilidade " & dblPedId & ",NULL,NULL,NULL,NULL,'E'")
	else
		Set objRSFac= db.Execute("CLA_SP_Sel_Facilidade " & dblPedId)
	END IF
	
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
			
			'strVlan = objRSFac("Fac_Vlan")
			'strPE = objRSFac("Fac_PE")
			'strPorta = objRSFac("Fac_Porta")
			'strLink = objRSFac("Fac_Link")
			
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
	strXML = strXML & "<endereço>" & objRSPontoB("TPL_SIGLA") & " " 
	strXML = strXML & objRSPontoB("End_NomeLogr") & " N" 
	strXML = strXML & objRSPontoB("End_NroLogr") & " " 
	strXML = strXML & objRSPontoB("Aec_Complemento") & " " 
	strXML = strXML & objRSPontoB("End_Bairro") & "</endereço>"  
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
	'strXML = strXML & "<vlan>"& strVlan &"</vlan>"	
	'strXML = strXML & "<pe>"& strPe &"</pe>"
	'strXML = strXML & "<porta>"& strPorta &"</porta>"
	'strXML = strXML & "<link>"& strLink &"</link>"
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
	strXML = strXML & "<designacao>" & objRSped1("Acf_NroAcessoPtaEbt") & "</designacao>"
	
	if cLng(objRSped1("Vel_Conversao")) >= 2048 then
	  vel_contrato = "SUPERLINK"
	else
	  vel_contrato = "EILD PADRÃO"
	end if
	strXML = strXML & "<veltpcontrato>" & vel_contrato & "</veltpcontrato>"
	
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
	strXML = strXML & "<endereço>RUA CAMERINO, 96 - SALA 310 - CENTRO</endereço>"   
	strXML = strXML & "<cidade>RIO DE JANEIRO</cidade>"
	strXML = strXML & "<cep>20080-010</cep>" 
	strXML = strXML & "<uf>RJ</uf>"  
	strXML = strXML & "<contato>HELENO ABREU</contato>"
	strXML = strXML & "<telefone>21-21219642</telefone>"  
	strXML = strXML & "<fax></fax>" 
	strXML = strXML & "<vencimento></vencimento>"
	strXML = strXML & "<prazo></prazo>"
	
	select case objRSPro("Cpro_ContratanteEstSigla")
	case "AL"
	  contacustomizada = "4500429"
	case "AM"
	  contacustomizada = "2000233"
	case "AP"
	  contacustomizada = "2000239"
	case "BA"
	  contacustomizada = "4500453"
	case "CE"
	  contacustomizada = "2000447"
	case "ES"
	  contacustomizada = "0000859"
	case "MA"
	  contacustomizada = "2000474"
	case "MG"
	  contacustomizada = "0000877"
	case "MG"
	  contacustomizada = "6000326"
	case "PA"
	  contacustomizada = "2000437"
	case "PB"
	  contacustomizada = "2000223"
	case "PE"
	  contacustomizada = "2000285"
	case "PI"
	  contacustomizada = "2000457"
	case "RJ"
	  contacustomizada = "3500371"
	case "RN"
	  contacustomizada = "2000267"
	case "RR"
	  contacustomizada = "2000395"
	case "SE"
	  contacustomizada = "4500436"
	End select
	
	strXML = strXML & "<conta_ptaa>" & contacustomizada  & "</conta_ptaa>"
	
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
	strXML = strXML & "<representante>" & objRSPro("Cpro_ContratanteContato") & "</representante>"  
	strXML = strXML & "<e-mail> " & objRSPro("Cpro_ContratanteEmail")  & "</e-mail>"  
	strXML = strXML & "<telefone> " & objRSPro("Cpro_ContratanteTelefone")  & "</telefone>" 
	strXML = strXML & "<fax> " & objRSPro("Cpro_ContratanteFax")  & "</fax>"
	strXML = strXML & "<endereço>" & objRSPro("Cpro_ContratanteEnd")  & "</endereço>" 
	strXML = strXML & "<cidade>" & objRSPro("ContratanteCid")  & "</cidade>" 
	strXML = strXML & "<uf>" & objRSPro("Cpro_ContratanteEstSigla")  & "</uf>" 	
	strXML = strXML & "</complemento>" 
	strXML = strXML & "<suporte>" 
	strXML = strXML & "<pep></pep>"
	'PSOUTO 10/05/2006
	'if objRSped1("tprc_id")  <>  "1" then
	if DblNdTipo  <>  "1" then  
	    if objRSped1("Acf_NroAcessoPtaEbt") = "" or isnull(objRSped1("Acf_NroAcessoPtaEbt")) or isempty(objRSped1("Acf_NroAcessoPtaEbt")) then
		  strXML = strXML & "<lp>" & session("ss_Acf_NroAcessoPtaEbt") & " </lp>" 'facilidade.asp
		else
		  strXML = strXML & "<lp>" & objRSped1("Acf_NroAcessoPtaEbt") & " </lp>"
		end if
	else
		strXML = strXML & "<lp></lp>"
	end if 
	strXML = strXML & "</suporte>" 
	strXML = strXML & "<cobranca>" 
	strXML = strXML & "<portifolio></portifolio>"
	strXML = strXML & "<projeto></projeto>"
	strXML = strXML & "<manual></manual>"
	strXML = strXML & "<conta>Conta Customizada </conta>"
	strXML = strXML & "<numconta>18500/3500371</numconta>"
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
	'strXML = strXML & "<obs>TRATA-SE DE DO CONTRATO DE SERVICO DE EILD UARJ.21 - 022/98 DE 25/03/1998 </obs>"
	'strXML = strXML & "<fraseobs>Para propósito de substituição conforme item 9.31 do contrato</fraseobs>"
	'PARA PROPÓSITO DE SUBSTITUIÇÃO CONFORME ITEM 9.31 DO CONTRATO
'	strXML = strXML & "<obs>" & Trim(objRSped1("pED_obs")) & "</obs>"
'' LPEREZ 13/12/2005

	strXML = strXML & "<fraseobs>Para propósito de substituição conforme item 9.2.1 do contrato</fraseobs>"

	'' PRSSILV - 05/03/2009 - Correção causa raiz OBS.
		if Trim(objRSped1("SOL_Obs")) <> Trim(objRSped1("Ped_Obs")) then
		  var_obs = Trim(objRSped1("SOL_Obs")) & "<p></p>" & Trim(objRSped1("Ped_Obs"))
		else
		  var_obs = Trim(objRSped1("Sol_Obs"))
		end if
		
		if trim(var_obs) = "" and DblNdTipo <> 2 then
		  var_obs = Trim(objRSped1("Ped_Obs"))
		end if
		
		strXML = strXML &	"<observacao>" & var_obs & "</observacao>"
'' PRSSILV
'' LP
	strXML = strXML & "</telemar>" 
	strXML = strXML & "</root>"

	set objRSped1 = nothing 
	set objRSPontoB = nothing 
	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strXML)

%>
