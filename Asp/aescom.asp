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
dim strServico , strAcao , strXML
dim ndPed 
dim ndSol 
dim ndProv
dim ndEsc 
dim ndTipo

Response.ContentType = "text/HTML;charset=ISO-8859-1"

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

if objRSped1.eof then 
	Response.Write ("<table width=100% ><tr><td style=""text-align:center""><font color = red>Pedido não encontrado.</font></td></tr></table>")
	Response.end()
end if 

set objRSPontoB = db.execute("CLA_sp_view_ponto null, " & dblPedId & ",null," & objRSped1("Sol_ID"))

Set objRSPro = db.execute("CLA_sp_sel_ProvedorEmail  " & dblProId & ",null,'" & objRSped1("Est_Sigla") & "','" & objRSped1("Cid_Sigla") &"'") 
if Not objRSPro.Eof and Not objRSPro.bof then
	strProEmail = Trim(objRSPro("cpro_contratadaemail"))
	strProNome	= Trim(objRSPro("Pro_Nome"))
End if
strParmProc = dblProId + "|" + dblPedId  + "|" + intTipoProcesso + "|" + strUserName 

	dim blnAtiva , blnDestaivacao, blnCancelamento, blnAltera 			
	dim strEnderEst, strCidadeEst
	blnAtiva  = 0
	blnDestaivacao = 0 
	blnCancelamento = 0 
	blnAltera =0 
		
			
			'--> PSOUTO 10/05/2006 
	if DblNdTipo <> 2 then
		DblNdTipo = objRSped1("tprc_id")  
	END IF 

	'PSOUTO 10/05/2006
	'select case objRSped1("tprc_id") '--> 
	SELECT CASE DblNdTipo

		case 1 
				blnAtiva = -1
				strServico = "Ativação"
				strAcao = "Instalar Acesso"
		case 2
				blnDestaivacao = -1 
				strServico = "Desativação"
				strAcao = "Retirar Acesso"
		case 3 
				blnAltera = -1 
				strServico = "Alteração"
				strAcao = "Alterar Acesso"
		case 4 
				blnCancelamento = -1 
				strServico = "Cancelamento"
				strAcao = "Cancelar Acesso"
	end select 

	if not isnull(objRSped1("esc_nomelogr")) then 
			strEnderEst = replace(trim(objRSped1("TPL_SIGLA")) & " " & trim(objRSped1("esc_nomelogr")) & " " & objRSped1("esc_nrologr") & " " & objRSped1("esc_complemento"),"&","&amp;") & " " & objRSped1("esc_bairro")
		else
			strEnderEst = ""
		end if
		
		if not isnull(objRSped1("cid_DESC")) then 
			strCidadeEst = replace(objRSped1("cid_DESC"),"&","&amp")
		else
			strCidadeEst = ""
	end if
			


		strXML = "<?xml version = ""1.0"" encoding =""ISO-8859-1""?><ROOT>"
		strXML = strXML &	"<contrato>" & objRSPro("CPro_ContratadaContrato") & "</contrato>"
		strXML = strXML &	"<arquivo>aescom</arquivo>"
		strXML = strXML &	"<parmproc>" & strParmProc & "</parmproc>"
		strXML = strXML &	"<assunto>" & strServico & "  -  " & trim(replace(objRSped1("Cli_nome"),"&","&amp;")) & "  -  " & ucase(objRSped1("Ped_Prefixo")) & "-" & right("00000" & objRSped1("Ped_Numero"),5) & "/" & objRSped1("Ped_Ano") &  "</assunto>"
		strXML = strXML &	"<mailfrom>" & objRSPro("CPro_ContratanteEmail") &  " </mailfrom>"
		strXML = strXML &	"<proemail>" & strProEmail &  " </proemail>"
		strXML = strXML &	"<pronome>" & replace(strProNome,"&","&amp;") &  "</pronome>"
		'alterado por PSOUTO 18/05/2006
		strXML = strXML &	"<numero>"
		IF intTipoProcesso <> 2 THEN ' NÃO É DESATIVACAO(LIBERACAO DE ESTOQUE)
		strXML = strXML &	objRSped1("ped_Prefixo") & "-" & objRSped1("ped_Numero") & "/" & objRSped1("ped_Ano") 
		END IF 
		strXML = strXML & "</numero>"
		'strXML = strXML &	"<numero>"& objRSped1("ped_Prefixo") & "-" & objRSped1("ped_Numero") & "/" & objRSped1("ped_Ano") & "</numero>"
		' /PSOUTO
		strXML = strXML &	"<data>" & formatdatetime(objRSped1("ped_data"),2) &  "</data>"
		strXML = strXML &	"<pontaebt>" & objRSped1("acf_NroAcessoPtaEbt") &  "</pontaebt>"
		strXML = strXML &	"<acao>" & strAcao &  "</acao>"
		strXML = strXML &	"<nomecontratada> " & objRSPro("CPro_ContratadaRazao")  &  "</nomecontratada>"
		strXML = strXML &	"<endercontratada>" & objRSPro("Cpro_ContratadaEnd")  & "</endercontratada>"
		strXML = strXML &	"<cidadecontratada>" & objRSPro("ContratadaCid")  & "</cidadecontratada>"
		strXML = strXML &	"<telefonecontratada>" & objRSPro("CPro_ContratadaTelefone")  & "</telefonecontratada>"
		strXML = strXML &	"<cepcontratada>" & objRSPro("Cpro_ContratadaCep")  & "</cepcontratada>"
		strXML = strXML &	"<faxcontratada>" & objRSPro("Cpro_ContratadaFax")  & "</faxcontratada>"
		strXML = strXML &	"<ufcontratada>" & objRSPro("Cpro_ContratadaEstSigla")  & "</ufcontratada>"
		strXML = strXML &	"<nomecontratante>"& objRSPro("CPro_ContratanteRazao")  &"</nomecontratante>"
		strXML = strXML &	"<endercontratante>"& objRSPro("Cpro_ContratanteEnd") & "</endercontratante>"
		strXML = strXML &	"<cidadecontratante>" &   objRSPro("ContratanteCid") & "</cidadecontratante>"
		strXML = strXML &	"<telefonecontratante>" & objRSPro("CPro_ContratanteTelefone")  & "</telefonecontratante>"
		strXML = strXML &	"<cepcontratante>"& objRSPro("Cpro_ContratanteCep")   &"</cepcontratante>"
		strXML = strXML &	"<faxcontratante>" &  objRSPro("Cpro_ContratanteFax") & "</faxcontratante>"
		strXML = strXML &	"<ufcontratante>" & objRSPro("Cpro_ContratanteEstSigla") & "</ufcontratante>"
		strXML = strXML &	"<chkcancelamento>" & blnCancelamento &  "</chkcancelamento>"
		strXML = strXML &	"<chkativacao>" & blnAtiva &  "</chkativacao>"
		strXML = strXML &	"<chkdesativacao>" & blnDestaivacao &  "</chkdesativacao>"
		strXML = strXML &	"<chkvelocidade>" & blnAltera &  "</chkvelocidade>"
		strXML = strXML &	"<chkendereço></chkendereço>"
		
			dim bln12, bln24 , bln36, bln48 , bln60 , blnInd , blnIndT
			bln12 = 0
			bln24 = 0 
			bln36 = 0 
			bln48 = 0
			bln60 = 0
			blnInd =0 
			blnIndT= 0
		
			select case trim(ucase(objRSped1("Tct_Meses")))
				case "12"
						bln12 =  -1 
				case "24" 
						bln24 = -1
				case "36" 
						bln36  = -1
				case "48" 
						bln48 = -1
				case "60"
						bln60 = -1
				case "0"
						blnIndT = - 1
				case else 
						blnind = -1
			end select 
		
		strXML = strXML &	"<chkind>" & blnIndT &  "</chkind>"
		strXML = strXML &	"<chktemporario>" & blnind &  "</chktemporario>"
		strXML = strXML &	"<chk12meses>" & bln12 &  "</chk12meses>"
		strXML = strXML &	"<chk24meses>" & bln24 &  "</chk24meses>"
		strXML = strXML &	"<chk36meses>" & bln36 &  "</chk36meses>"
		strXML = strXML &	"<chk48meses>" & bln48 &  "</chk48meses>"
		strXML = strXML &	"<chk60meses>" & bln60 &  "</chk60meses>"
		
		strXML = strXML &	"<temporariode>" & objRSped1("Acl_DtIniAcessoTemp") &  "</temporariode>"
		strXML = strXML &	"<temporarioate>" & objRSped1("Acl_DtFimAcessoTemp") &  "</temporarioate>"
		
		'@@ Clinte 
		'strXML = strXML & "<cidade>" & objRSPontoB("CID_DESC")   & "</cidade>" 
		'strXML = strXML & "<cep>" & objRSPontoB("end_Cep") & "</cep>" 
		'strXML = strXML & "<uf>" & objRSPontoB("est_sigla") & "</uf>"
		'strXML = strXML & "<logradouro>" & objRSPontoB("CID_DESC") & "</logradouro>"
		'strXML = strXML & "<sigla>" & objRSPontoB("CID_SIGLA") & "</sigla>" 
		
	
		strXML = strXML &	"<clientenome>" & trim(objRSped1("Cli_nome")) & "</clientenome>"
		strXML = strXML &	"<clienteend>" &   trim(objRSPontoB("TPL_SIGLA")) & " " & trim(objRSPontoB("End_NomeLogr")) & " " & trim(objRSPontoB("End_NroLogr")) & " " & trim(objRSPontoB("Aec_Complemento")) & " " & trim(objRSPontoB("End_Bairro")) & "</clienteend>"
		strXML = strXML &	"<clientecidade>" & objRSPontoB("CID_DESC")   &  "</clientecidade>"
		strXML = strXML &	"<clienteuf>"  & objRSPontoB("est_sigla") & "</clienteuf>"
		strXML = strXML &	"<clientecep>" & objRSPontoB("end_Cep") &  "</clientecep>"
		strXML = strXML &	"<clienteccta>" & objRSPontoB("Cli_CC") & "</clienteccta>"
		strXML = strXML &	"<clientecnpj>" & objRSPontoB("Aec_CNPJ") & "</clientecnpj>"
		strXML = strXML &	"<clienteiem>" & objRSPontoB("Aec_IM") & "</clienteiem>"
		strXML = strXML &	"<clienteiee>" & objRSPontoB("Aec_IE") & "</clienteiee>"
		strXML = strXML &	"<clientecontato>" & objRSPontoB("Aec_Contato") & "</clientecontato>"
		strXML = strXML &	"<clientetelefone>( " & mid(objRSPontoB("Aec_Telefone"),1,2) & " ) " & mid(objRSPontoB("Aec_Telefone"),3,len(objRSPontoB("Aec_Telefone")) - 2) & "</clientetelefone>"
		strXML = strXML &	"<clienteinterface>" & objRSPontoB("Acf_Interface") & "</clienteinterface>"
		strXML = strXML &	"<clientevelocidade>" & objRSped1("Vel_Desc") &  "</clientevelocidade>"
		strXML = strXML &	"<clienteback></clienteback>"
		strXML = strXML &	"<clientesev>" & trim(objRSped1("Sol_SevSeq")) & "</clientesev>"
		strXML = strXML &	"<clienteprazo></clienteprazo>"
		
		
		'@@ PontaB
		strXML = strXML &	"<pontabnome>EMPRESA BRASILEIRA DE TELECOMUNICAÇÔES S/A - EMBRATEL</pontabnome>"
		strXML = strXML &	"<pontabend>" & trim(objRSped1("TPL_SIGLA")) & " " & trim(objRSped1("esc_nomelogr")) & " " & objRSped1("esc_nrologr") & " " & objRSped1("esc_complemento") & " " & objRSped1("esc_bairro") & "</pontabend>"
		strXML = strXML &	"<pontabcidade>" & objRSped1("cid_desc") & "</pontabcidade>"
		strXML = strXML &	"<pontabuf>" & objRSped1("Est_sigla") & "</pontabuf>"
		strXML = strXML &	"<pontabcep>" & objRSped1("esc_cod_cep")  & "</pontabcep>"
		strXML = strXML &	"<pontabcnpj>" & objRSped1("esc_CNPJ")  & "</pontabcnpj>"
		strXML = strXML &	"<pontabie></pontabie>"
		strXML = strXML &	"<pontabtel>" & objRSped1("esc_telefone") &  "</pontabtel>"
		strXML = strXML &	"<pontabcontato>" & objRSped1("esc_contato")  & "</pontabcontato>"
		
		
		
		strXML = strXML &	"<nomeembratel></nomeembratel>"
		strXML = strXML &	"<nomerepresentante></nomerepresentante>"
		strXML = strXML &	"<localdataembratel></localdataembratel>"
		strXML = strXML &	"<localdatarepresentante></localdatarepresentante>"
'' LPEREZ 13/12/2005		
		strXML = strXML &	"<observaçãop>" & objRSped1("pED_obs") & "</observaçãop>"
		strXML = strXML &	"<observaçãos>" & objRSped1("SOL_obs") & "</observaçãos>"
'' LP		
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
		
		strXML = strXML &	"<contatogla>" & strNomeRepresent &  "</contatogla>"
		strXML = strXML &	"<contatoebt>" & objRSPro("Cpro_ContratanteContato") & "</contatoebt>"
		strXML = strXML &	"<cargoebt>" & objRSPro("Cpro_ContratanteDepto") & "</cargoebt>"
		strXML = strXML &	"</ROOT>"

	
	set objRSPro = nothing 
	set objRSped1 = nothing
	set objRSPontoB = nothing
	
	Response.Write (strXML)
%>
