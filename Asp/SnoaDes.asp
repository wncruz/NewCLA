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

	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
		
		
	objXmlDoc.load(Request)
	strCaminho = server.MapPath("..\")


	'set ndPed =  objXmlDoc.selectSingleNode("//ped")
	set ndSol  =  objXmlDoc.selectSingleNode("//sol")
	set ndUser =  objXmlDoc.selectSingleNode("//user")
	'set ndEsc =  objXmlDoc.selectSingleNode("//Esc")
	'set ndTipo =  objXmlDoc.selectSingleNode("//ndTipo")

	set ndAcfId =  objXmlDoc.selectSingleNode("//acfid")
	'dblPedId  = ndPed.Text
	dblSolId  = ndSol.Text
	strUserName = ndUser.Text
	
	DblAcf_ID = ndAcfId.Text
	
	set ndstrVelocidade  			=  objXmlDoc.selectSingleNode("//Velocidade")
	set ndstrContatoEnd_A  			=  objXmlDoc.selectSingleNode("//ContatoEnd_A")
	set ndstrTelEnd_A  				=  objXmlDoc.selectSingleNode("//TelEnd_A")
	set ndstrDesignacaoFornecedora  =  objXmlDoc.selectSingleNode("//DesignacaoFornecedora")
	set nddblSnoa  					=  objXmlDoc.selectSingleNode("//Snoa")
	set ndstrQtdLinhas  			=  objXmlDoc.selectSingleNode("//QtdLinhas")
	
	
	strVelocidade					= ndstrVelocidade.Text		
	strContatoEnd_A					= ndstrContatoEnd_A.Text
	strTelEnd_A						= ndstrTelEnd_A.Text	
	strDesignacaoFornecedora		= ndstrDesignacaoFornecedora.Text	
	dblSnoa							= nddblSnoa.Text	
	strQtdLinhas					= ndstrQtdLinhas.Text
	
	'response.write "<script>alert('"&strVelocidade&"')</script>"
	'response.write "<script>alert('"&strContatoEnd_A&"')</script>"
	'response.write "<script>alert('"&strTelEnd_A&"')</script>"
	'response.write "<script>alert('"&dblSolId&"')</script>"
	'response.write "<script>alert('"&strDesignacaoFornecedora&"')</script>"
	'response.write "<script>alert('"&dblSnoa&"')</script>"
	'response.write "<script>alert('"&dblAcfId&"')</script>"
	'response.write "<script>alert('"&strQtdLinhas&"')</script>"
	
	Vetor_Campos(1)="adInteger,4,adParamInput," & dblSnoa
	Vetor_Campos(2)="adWChar,100,adParamInput," & strContatoEnd_A
	
	Vetor_Campos(3)="adInteger,4,adParamInput," & dblSolId
	Vetor_Campos(4)="adWChar,100,adParamInput," & strDesignacaoFornecedora
	
	Vetor_Campos(5)="adWChar,20,adParamInput," & strTelEnd_A
	
	Vetor_Campos(6)="adWChar,4,adParamInput," & strQtdLinhas
	Vetor_Campos(7)="adInteger,4,adParamInput," & strVelocidade
	
	Vetor_Campos(8)="adInteger,4,adParamInput," & DblAcf_ID
	

	call APENDA_PARAM("cla_sp_upd_PedidoSnoa",8,Vetor_Campos)
	ObjCmd.Execute



	'Retorna Endereco da Ponta A e Ponta B'
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
	strSqlRet = APENDA_PARAMSTR("CLA_sp_view_SNOA_DES",1,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)

	

	
	if not objRS.Eof then

		str_ped_id						= objRS("ped_id")
		str_tprc_id 					= objRS("tprc_id")                    
		str_sol_id 						= objRS("sol_id")
		str_Sol_SevSeq 					= objRS("Sol_SevSeq")
		str_ped_data 					= objRS("ped_data")
		str_DM 							= objRS("DM")
		str_Acl_IDAcessoLogico 			= objRS("Acl_IDAcessoLogico")
		str_esc_identrega 				= objRS("esc_identrega")
		str_esc_idconfiguracao 			= objRS("esc_idconfiguracao")
		
		str_Est_sigla_PTA 				= objRS("Est_sigla_PTA")
		str_Cid_Desc_PTA 				= objRS("Cid_Desc_PTA")
		str_Cid_Sigla_PTA 				= objRS("Cid_Sigla_PTA")
		str_Tpl_Sigla_PTA 				= objRS("Tpl_Sigla_PTA")
		str_End_NomeLogr_PTA 			= objRS("End_NomeLogr_PTA")
		str_End_NroLogr_PTA 			= objRS("End_NroLogr_PTA")
		str_End_Bairro_PTA 				= objRS("End_Bairro_PTA")
		str_End_CEP_PTA 				= objRS("End_CEP_PTA")
		str_Aec_Complemento_PTA 		= objRS("Aec_Complemento_PTA")
		str_Aec_Contato_PTA 			= objRS("Aec_Contato_PTA")
		str_Aec_Telefone_PTA 			= objRS("Aec_Telefone_PTA")
		str_Aec_CNPJ_PTA 				= objRS("Aec_CNPJ_PTA")
		str_Email_PTA 					= objRS("Email_PTA")
		
		str_Est_Sigla_PTB               = objRS("Est_Sigla_PTB")
		str_TPL_Sigla_PTB               = objRS("TPL_Sigla_PTB")
		str_Cid_Desc_PTB                = objRS("Cid_Desc_PTB")
		str_ESC_CID_SIGLA_PTB 			= objRS("ESC_CID_SIGLA_PTB")
		str_Esc_NomeLogr_PTB 			= objRS("Esc_NomeLogr_PTB")
		str_Esc_NroLogr_PTB 			= objRS("Esc_NroLogr_PTB")
		str_Esc_Cod_Cep_PTB 			= objRS("Esc_Cod_Cep_PTB")
		str_Esc_Bairro_PTB 				= objRS("Esc_Bairro_PTB")
		str_Esc_Complemento_PTB 		= objRS("Esc_Complemento_PTB")
		str_Esc_Contato_PTB 			= objRS("Esc_Contato_PTB")
		str_Esc_Telefone_PTB 			= objRS("Esc_Telefone_PTB")
		str_Esc_CNPJ_PTB 				= objRS("Esc_CNPJ_PTB")
		str_Esc_Email_PTB 				= objRS("Esc_Email_PTB")

		str_ACAO 						= objRS("ACAO")
		str_Status						= objRS("Status")
		str_Cliente						= objRS("Cliente")
		str_NContratoServico			= objRS("NContratoServico")
		str_DesignacaoServico			= objRS("DesignacaoServico")
		str_Servico 					= objRS("Servico")
		str_DataSolicitacao				= objRS("DataSolicitacao")
		str_VelocidadeServico 			= objRS("VelocidadeServico")

		strProId						= objRS("PRO_ID")
		
		'DblAcf_ID						= objRS("acf_id")
		
		'strSnoaNum						= trim(objRS("snoa_num"))
	end if
	
	'<!-- CH-83646VWR - Inicio -->
	Vetor_Campos(1)="adInteger,4,adParamInput," & dblSolId
	Vetor_Campos(2)="adInteger,4,adParamOutput,0"
	Call APENDA_PARAM("CLA_sp_sel_SnoaAnterior",2,Vetor_Campos)

	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	'response.write "<script>alert('"&Trim(DBAction)&"')</script>"

	if DBAction <> 0 then
		strSnoaNum		= Trim(DBAction)
	Else
		strSnoaNum		= dblSnoa
	end if
	
	'@Pro_ID int = null , 
   	'@UF_PTA varchar(2) = null , 
   	'@Cid_Sigla_PTA varchar(4) = null
	Vetor_Campos(1)="adInteger,2,adParamInput," & strProId
	Vetor_Campos(2)="adWChar,2,adParamInput," & str_Est_sigla_PTA
	Vetor_Campos(3)="adWChar,4,adParamInput," & str_Cid_Sigla_PTA
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_RepresentanteContratoProvedor",3,Vetor_Campos)
	Set objRS3 = db.Execute(strSqlRet)
	if not objRS3.Eof then
		'Representante Comercial'
		str_Nome_ReprComer 					= objRS3("CPro_RepresentanteNome")
		str_End_ReprComer 					= objRS3("CPro_RepresentanteEnd")
		str_Est_sigla_ReprComer				= objRS3("CPro_RepresentanteEstSigla")
		str_End_CEP_ReprComer				= objRS3("CPro_RepresentanteCEP")
		str_Cid_Sigla_ReprComer				= objRS3("CPro_RepresentanteCidSigla")
		str_Cid_Desc_ReprComer				= objRS3("RepresentanteCid")
		str_Email_ReprComer 				= objRS3("CPro_RepresentanteEmail")
		str_Telefone_ReprComer				= objRS3("CPro_RepresentanteTelefone")
		str_Fax_ReprComer					= objRS3("CPro_RepresentanteFax")

		'Dados de Faturamento e Cobrança'
		str_Cpro_ContratanteRazao			= objRS3("Cpro_ContratanteRazao")
		str_Cpro_ContratanteCGC_CNPJ		= objRS3("Cpro_ContratanteCGC_CNPJ")
		str_Cpro_inscricao_estadual			= objRS3("inscricao_estadual")
		str_Cpro_ContratanteEnd				= objRS3("Cpro_ContratanteEnd")
		str_Cpro_ContratanteCEP				= objRS3("Cpro_ContratanteCEP")
		str_Cpro_Cid_Desc					= objRS3("Cid_Desc")
		str_Cpro_ContratanteEstSigla		= objRS3("Cpro_ContratanteEstSigla")
		str_Cpro_conta_customizada_ponta_a	= objRS3("conta_customizada_ponta_a")
		str_Cpro_conta_customizada_ponta_b	= objRS3("conta_customizada_ponta_b")
		str_Cpro_ContratanteContato			= objRS3("Cpro_ContratanteContato")
		str_Cpro_ContratanteEmail			= objRS3("Cpro_ContratanteEmail")
		str_Cpro_ContratanteTelefone		= objRS3("Cpro_ContratanteTelefone")
		str_Cpro_ContratanteFax				= objRS3("Cpro_ContratanteFax")
	end if

	'NOVAS PROCEDURES - SNOA - INICIO'
	Vetor_Campos(1)="adInteger,2,adParamInput," & DblAcf_ID
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_AssocFisicoTecnico",1,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)

	if not objRS.Eof then

		str_aft_id = objRS("aft_id")
		str_acf_id = objRS("acf_id")
		str_Transmissao = objRS("vel_desc")
		'str_aplicacao_id = objRS("aplicacao_id")
		'str_meioPreferencial_id = objRS("meioPreferencial_id")
		'str_caracteristicaTecnica_id = objRS("caracteristicaTecnica_id")
		str_QtdLinha = objRS("QtdLinha")
		'str_QinQ = objRS("QinQ")
		'str_Prazo_contratacao_id = objRS("Prazo_contratacao_id")
		'str_Produto_id = objRS("Produto_id")
		'str_Finalidade_id = objRS("Finalidade_id")
		'str_Descricao_estudo = objRS("Descricao_estudo")
		'str_motivo_id = objRS("motivo_id")
		'str_justificativa_des = objRS("justificativa_des")
		'str_EILD_SEQ_SUGERIDO = objRS("EILD_SEQ_SUGERIDO")

		'str_Prestadora_id = objRS("Fornecedora_id")
		'str_Contrato_id = objRS("Contrato_id")
		

	end if

	'str_Transmissao = DblAcf_ID
	
	'NOVAS PROCEDURES - SNOA - INICIO'
	Vetor_Campos(1)="adInteger,2,adParamInput,null"
	Vetor_Campos(2)="adInteger,2,adParamInput," & strSnoaNum
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_designacaoFornecedora",2,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)

	if not objRS.Eof then

		str_Designacao = objRS("designacao_recurso_fornecedora")
	
	end if
	
		
	
	'strUserName = "EDAR"
	'strParmProc = strProId + "|" + str_ped_id  + "|" + str_tprc_id + "|" + strUserName 
	
	strAssunto = " Retirada do número do Bilhete  -  " & strSnoaNum 

	DIM   strXML
	
	//Msgusuario.innerText = "Aguarde..."
	strXML = "<?xml version = ""1.0"" encoding =""ISO-8859-1""?><root>"
	strXML = strXML & "<arquivo>SnoaDes</arquivo>"
	strXML = strXML & "<parmproc>" & strParmProc & "</parmproc>"
	strXML = strXML & "<assunto>" & strAssunto & "</assunto>"
	strXML = strXML & "<proemail>" & strProEmail & "</proemail>"
	strXML = strXML & "<mailfrom>" & strFromEmail &  " </mailfrom>"
	strXML = strXML & "<pronome>"  & strProNome & "</pronome>"
	
	strXML = strXML & "<proid>"  & strProId & "</proid>"
	strXML = strXML & "<pedid>"  & str_ped_id & "</pedid>"
	strXML = strXML & "<tprcid>"  & str_tprc_id & "</tprcid>"
	strXML = strXML & "<username>"  & strUserName & "</username>"
	'alterado por PSOUTO 18/05/2006
	strXML = strXML &	"<numero>" & strSnoaNum & "</numero>"
	
	strXML = strXML &	"<familia> "
	strXML = strXML &	"	<numcontrato></numcontrato> "
	strXML = strXML &	"	<data>" & str_DataSolicitacao &  "</data> "
	strXML = strXML &	"	<circuito>SUPERLINK</circuito> "
	strXML = strXML &	"</familia> "
	strXML = strXML &	"<empresa> "
	strXML = strXML &	"	<cliente>" & str_Cpro_ContratanteRazao &  "</cliente> "
	strXML = strXML &	"	<endereço>" & str_Cpro_ContratanteEnd &  "</endereço> "
	strXML = strXML &	"	<cidade>" & str_Cpro_Cid_Desc &  "</cidade> "
	strXML = strXML &	"	<cep>" & str_Cpro_ContratanteCEP &  "</cep> "
	strXML = strXML &	"	<inscricao_municipal>" & str_Cpro_inscricao_estadual &  "</inscricao_municipal> "
	strXML = strXML &	"	<cnpj>" & str_Cpro_ContratanteCGC_CNPJ &  "</cnpj> "
	strXML = strXML &	"	<uf>" & str_Cpro_ContratanteEstSigla &  "</uf> "
	strXML = strXML &	"	<inscrição>" & str_Cpro_inscricao_estadual &  "</inscrição> "
	strXML = strXML &	"	<telefone>" & str_Cpro_ContratanteTelefone &  "</telefone> "
	strXML = strXML &	"	<email>" & str_Cpro_ContratanteEmail  &  "</email> "
	strXML = strXML &	"</empresa> "
	strXML = strXML &	"<solicitacao> "
	strXML = strXML &	"	<cliente>" & str_Cliente &  "</cliente> "
	strXML = strXML &	"	<serviço> 1 -  Retirada  </serviço> "
	strXML = strXML &	"	<data>" & str_DataSolicitacao &  "</data> "
	strXML = strXML &	"	<motivo>005 - Desistência do serviço pelo cliente final da operadora contratante</motivo> "
	strXML = strXML &	"</solicitacao> "
	strXML = strXML &	"<tecnico> "
	strXML = strXML &	"	<velocidade>" & str_Transmissao &  "</velocidade> "
	strXML = strXML &	"	<linhas>" & str_QtdLinha &  "</linhas> "
	strXML = strXML &	"	<designacao>" & str_Designacao &  "</designacao> "
		
	strXML = strXML &	"</tecnico> "
	strXML = strXML &	"<pontaa> "
	strXML = strXML &	"	<contato>" & str_Aec_Contato_PTA &  "</contato> "
	strXML = strXML &	"	<telefone>" & str_Aec_Telefone_PTA &  "</telefone> "
	strXML = strXML &	"</pontaa> "
	strXML = strXML &	"<pontab> "
	strXML = strXML &	"	<contato>" & str_Esc_Contato_PTB &  "</contato> "
	strXML = strXML &	"	<telefone>" & str_Esc_Telefone_PTB &  "</telefone> "
	strXML = strXML &	"</pontab> "
	strXML = strXML &	"<observacao>" & str_justificativa_des &  "</observacao> "

	
	strXML = strXML &	" </root> "
	
	
	
	

	set objRSped1 = nothing 
	set objRSPontoB = nothing 
	
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strXML)

%>
