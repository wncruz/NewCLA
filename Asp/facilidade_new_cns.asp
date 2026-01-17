<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Facilidade.asp
'	- Descrição			: Alocação de facilidade
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<!--#include file="../inc/AlocacaoFac.asp"-->
<%

'response.write "<script>alert('"&Trim(Request.QueryString("SolId"))&"')</script>"

dblSolId = Trim(Request.Form("hdnSolId"))
if dblSolId = "" then dblSolId = Trim(Request.QueryString("SolId"))

dblPedId = Request.Form("hdnPedId")
if dblPedId = "" then dblPedId = Trim(Request.QueryString("PedId"))

'strJanOrig = Request.Form("hdnPaginaOrig")

'if Trim(Request.Form("hdnSolId")) <> "" then

if Trim(dblSolId) <> "" then


	set ObjRsetOrisol = db.execute("select orisol_id , Acl_IDAcessoLogico from cla_solicitacao where sol_id = " & dblSolId ) 'Request.Form("hdnSolId"))
	if not ObjRsetOrisol.eof then
		orisolid = ObjRsetOrisol("orisol_id")
		strIdLogico		= Trim(ObjRsetOrisol("Acl_IDAcessoLogico"))
	end if
	
	'GPON
	set ObjRsetTecID = db.execute("select top 1 cla_acessofisico.Pro_ID,cla_tecnologia.Tec_Nome,cla_acessofisico.Acf_ID, cla_tecnologia.Tec_ID from cla_solicitacao inner join cla_acessologico 		on cla_solicitacao.Acl_IDAcessoLogico = cla_acessologico.Acl_IDAcessoLogico inner join cla_acessologicofisico 	on cla_acessologico.Acl_IDAcessoLogico = cla_acessologicofisico.Acl_IDAcessoLogico inner join cla_acessofisico 		on cla_acessologicofisico.acf_id = cla_acessofisico.acf_id left join cla_tecnologia 		on cla_acessofisico.tec_id = cla_tecnologia.tec_id where ( ( cla_acessologicofisico.alf_flagAux is null  ) or (cla_acessologicofisico.alf_flagAux is not null and pro_id = 145) ) and sol_id = " & dblSolId ) 'Request.Form("hdnSolId") )
	if not ObjRsetTecID.eof then
		strTecnologia = ObjRsetTecID("Tec_Nome")
		strTecId = ObjRsetTecID("Tec_ID")
		DblAcf_ID = ObjRsetTecID("Acf_ID")
		strProId = ObjRsetTecID("Pro_ID")
	end if
	
	
	'dblSolId = Request.Form("hdnSolId")
	
	'dblPedId = Request.Form("hdnPedId")
	
	strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

	
	if dblPedId = "" and TRIM(strTecnologia) <> "GPON" then 'jcartus
		Response.Write "<script language=javascript>alert('Pedido indisponível.');window.location.replace('facilidade_main.asp');</script>"
		Response.End
	End if

	strProId = Request.Form("cboProvedor")
	if strProId = "" then  strProId	= Trim(objRS("Pro_id"))

		
	strPabx_Virtual     	= Trim(objRS("PABX_VIRTUAL"))	
	strExigePortadora 		= Trim(objRS("Pro_ExigePortadora"))
	strNroSolic				= dblSolId 'Trim(objRS("Sol_id"))
	strPropAcesso			= Trim(objRS("Acf_Proprietario"))
	
	strDataSolic			= Formatar_Data(Trim(objRS("Sol_Data")))
	strDM					= UCASE(objRS("Ped_Prefixo") & "-" & right("00000" & objRS("Ped_Numero"),5) & "/" & objRS("Ped_Ano"))
	strDataPedido			= Formatar_Data(objRS("Ped_Data"))
	strSatus				= Trim(objRS("Sts_Desc"))
	strCliente				= Trim(objRS("Cli_Nome"))
	intTipoProcesso 		= objRS("Tprc_idLog")
	strObs					= Trim(objRS("Sol_Obs"))
	PedObs					= Trim(objRS("Ped_Obs"))
	strAcao					= AcaoPedido(objRS("Tprc_id"))
	intTipoAcao				= objRS("Tprc_id")
	Id_Acesso				= objRS("Id_Acesso")
	dblNroSev				= objRS("Sol_SevSeq")

	strCidSigla				= Trim(objRS("Cid_Sigla"))
	strUfSigla				= Trim(objRS("Est_Sigla"))
	strTplSigla 			= Trim(objRS("Tpl_Sigla"))
	strNomeLogr				= Trim(objRS("End_NomeLogr"))
	strNroEnd				= Trim(objRS("End_NroLogr"))
	strCep					= Trim(objRS("End_Cep"))

	strNroServico			= Trim(objRS("Acl_NContratoServico"))
	strDesigServico 		= Trim(objRS("Acl_DesignacaoServico"))
	strServico				= Trim(objRS("Ser_Desc"))
	strVelServico			= Trim(objRS("DescVelAcessoLog"))
	strVelAcessoFis			= Trim(objRS("DescVelAcessoFis"))

	strPrmId				= Trim(objRS("Prm_id"))
	strRegId				= Trim(objRS("Reg_id"))

	strLocalInstala = Request.Form("estacao")
	if strLocalInstala = "" then
		strLocalInstala		= Trim(objRS("Esc_IdEntrega"))
	end if
	
	strLocalConfig			= Trim(objRS("Esc_IdConfiguracao"))
	strRecurso				= Trim(objRS("Rec_IDEntrega"))

	strDistrib = Request.Form("distribuidor")
	if strDistrib = "" then
		strDistrib			= Trim(objRS("Dst_Id"))
	end if

	strRede = objRS("Sis_ID")
		
	If (strRede = 4 or strRede = 5) And strProId = 154 Then
		strRede 			= 13
	End If

	'<!-- CH-83646VWR - Inicio -->
	Vetor_Campos(1)="adInteger,4,adParamInput," & dblSolId
	Vetor_Campos(2)="adInteger,4,adParamOutput,0"
	Call APENDA_PARAM("CLA_sp_sel_SnoaAnterior",2,Vetor_Campos)

	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	'response.write "<script>alert('"&Trim(DBAction)&"')</script>"

	if DBAction <> 0 then
		strNumPedSnoaAnt		= Trim(DBAction)
	Else
		strNumPedSnoaAnt		= ""
	end if
	'<!-- CH-83646VWR - Fim -->


	set objRSFis = db.execute("CLA_sp_view_acessofisico null," & dblPedId)
	if Not objRSFis.Eof and not objRSFis.Bof then
		strTecnologia		= objRSFis("Tec_Sigla")
		strVelFis			= objRSFis("Vel_Desc")
		intTipoVel			= objRSFis("Acf_TipoVel")
	End if

	'GPON
	if TRIM(strTecnologia) <> "GPON" or isnull(strTecnologia) then
		DblAcf_ID 			= objRSFis("Acf_ID")
	end if
	
	'<!-- OI-8064742-16 - Inicio -->
	'set ObjRssnoa = db.execute("select top 1 cla_acessofisico.Pro_ID,cla_tecnologia.Tec_Nome,cla_acessofisico.Acf_ID, cla_tecnologia.Tec_ID from cla_solicitacao inner join cla_acessologico 	On cla_solicitacao.Acl_IDAcessoLogico = cla_acessologico.Acl_IDAcessoLogico inner join cla_acessologicofisico 	on cla_acessologico.Acl_IDAcessoLogico = cla_acessologicofisico.Acl_IDAcessoLogico 	inner join cla_acessofisico 		on cla_acessologicofisico.acf_id = cla_acessofisico.acf_id 	left join cla_tecnologia 		on cla_acessofisico.tec_id = cla_tecnologia.tec_id 	where sol_id = " & dblSolId ) 'Request.Form("hdnSolId") )

	set ObjRssnoa = db.execute("select top 1 cla_acessofisico.Pro_ID,cla_tecnologia.Tec_Nome,cla_acessofisico.Acf_ID, cla_tecnologia.Tec_ID from cla_solicitacao inner join cla_acessologico 	On cla_solicitacao.Acl_IDAcessoLogico = cla_acessologico.Acl_IDAcessoLogico inner join cla_acessologicofisico 	on cla_acessologico.Acl_IDAcessoLogico = cla_acessologicofisico.Acl_IDAcessoLogico 	inner join cla_acessofisico 		on cla_acessologicofisico.acf_id = cla_acessofisico.acf_id 	left join cla_tecnologia 		on cla_acessofisico.tec_id = cla_tecnologia.tec_id 	where alf_flag is null and sol_id = " & dblSolId ) 'Request.Form("hdnSolId") )
	'<!-- OI-8064742-16 - Fim -->
	
	if not ObjRssnoa.eof then
		strTecnologia = ObjRssnoa("Tec_Nome")
		strTecId = ObjRssnoa("Tec_ID")
		DblAcf_ID = ObjRssnoa("Acf_ID")
		strProId = ObjRssnoa("Pro_ID")
	end if
	SET ObjRssnoa = Nothing


	dim hdnNroAcessoEbt
	hdnNroAcessoEbt 		= ""
	set ObjTemp = db.execute ("select Acf_NroAcessoPtaEbt, pro_id from cla_acessofisico with (nolock) where acf_ID=" & DblAcf_ID)
	if not ObjTemp.eof and not ObjTemp.Bof then	 
		if ObjTemp("pro_id") = 145 and NOT isnull(ObjTemp("Acf_NroAcessoPtaEbt")) then 'FO ETHERNET
			hdnNroAcessoEbt = ObjTemp("Acf_NroAcessoPtaEbt")		  		  	
		end if
	end if

	If trim(strLoginRede) = "T3FRRP"  Then
		response.write "<script>alert('"&dblSolId&"')</script>"
		response.write "<script>alert('"&dblPedId&"')</script>"
		response.write "<script>alert('"&DblAcf_ID&"')</script>"
	End IF

	Set objRSFis = Nothing

	'NOVAS PROCEDURES - SNOA - INICIO'
	Vetor_Campos(1)="adInteger,2,adParamInput," & DblAcf_ID
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_AssocFisicoTecnico",1,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)

	if not objRS.Eof then

		str_aft_id = objRS("aft_id")
		str_acf_id = objRS("acf_id")
		str_vel_idTransmissao = objRS("vel_idTransmissao")
		str_aplicacao_id = objRS("aplicacao_id")
		str_meioPreferencial_id = objRS("meioPreferencial_id")
		str_caracteristicaTecnica_id = objRS("caracteristicaTecnica_id")
		str_QtdLinha = objRS("QtdLinha")
		str_QinQ = objRS("QinQ")
		str_Prazo_contratacao_id = objRS("Prazo_contratacao_id")
		str_Produto_id = objRS("Produto_id")
		str_Finalidade_id = objRS("Finalidade_id")
		str_Descricao_estudo = objRS("Descricao_estudo")
		str_motivo_id = objRS("motivo_id")
		str_justificativa_des = objRS("justificativa_des")
		str_EILD_SEQ_SUGERIDO = objRS("EILD_SEQ_SUGERIDO")

		str_Prestadora_id = objRS("Fornecedora_id")
		str_Contrato_id = objRS("Contrato_id")

	end if


	'Ponta A'
	Vetor_Campos(1)="adInteger,2,adParamInput," & DblAcf_ID
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_AssocFisicoPontaInstalacao",1,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)

	if not objRS.Eof then

		str_afp_id_A 					= objRS("afp_id")
		str_acf_id_A 					= objRS("acf_id")
		str_Latitude_A 					= objRS("Latitude")
		str_Longitude_A 				= objRS("Longitude")
		str_Equipamento_A 				= objRS("Equipamento")
		str_CentroFios_A 				= objRS("CentroFios")
		str_Ponto_Referencia_A 			= objRS("Ponto_Referencia")
		str_Canais_desc_A 				= objRS("Canais_desc")
		str_Entrega_Canalizado_A 		= objRS("Entrega_Canalizado")
		str_E1_Canalizado_A 			= objRS("E1_Canalizado")
		str_Time_Slot_A 				= objRS("Time_Slot")
		str_InterfaceEletrica_id_A 		= objRS("InterfaceEletrica_id")
		str_InterfaceFisica_id_A 		= objRS("InterfaceFisica_id")

		str_afp_id_B 					= objRS("afp_id_B")
		str_acf_id_B 					= objRS("acf_id_B")
		str_Latitude_B 					= objRS("Latitude_B")
		str_Longitude_B 				= objRS("Longitude_B")
		str_Equipamento_B 				= objRS("Equipamento_B")
		str_CentroFios_B 				= objRS("CentroFios_B")
		str_Ponto_Referencia_B 			= objRS("Ponto_Referencia_B")
		str_Canais_desc_B 				= objRS("Canais_desc_B")
		str_Entrega_Canalizado_B 		= objRS("Entrega_Canalizado_B")
		str_E1_Canalizado_B 			= objRS("E1_Canalizado_B")
		str_Time_Slot_B 				= objRS("Time_Slot_B")
		str_InterfaceEletrica_id_B 		= objRS("InterfaceEletrica_id_B")
		str_InterfaceFisica_id_B 		= objRS("InterfaceFisica_id_B")


	end if

	'Ponta B'
	'Vetor_Campos(1)="adInteger,2,adParamInput," & DblAcf_ID
	'strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_AssocFisicoPontaInstalacao",1,Vetor_Campos)
	'Set objRS = db.Execute(strSqlRet)

	'if not objRS.Eof then

		'str_afp_id_B 					= objRS("afp_id")
		'str_acf_id_B 					= objRS("acf_id")
		'str_Latitude_B 					= objRS("Latitude_B")
		'str_Longitude_B 				= objRS("Longitude_B")
		'str_Equipamento_B 				= objRS("Equipamento")
		'str_CentroFios_B 				= objRS("CentroFios")
		'str_Ponto_Referencia_B 			= objRS("Ponto_Referencia")
		'str_Canais_desc_B 				= objRS("Canais_desc")
		'str_Entrega_Canalizado_B 		= objRS("Entrega_Canalizado")
		'str_E1_Canalizado_B 			= objRS("E1_Canalizado")
		'str_Time_Slot_B 				= objRS("Time_Slot")
		'str_InterfaceEletrica_id_B 		= objRS("InterfaceEletrica_id")
		'str_InterfaceFisica_id_B 		= objRS("InterfaceFisica_id")

	'end if

	'Retorna Endereco da Ponta A e Ponta B'
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblPedId
	strSqlRet = APENDA_PARAMSTR("CLA_sp_view_SNOA",1,Vetor_Campos)
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

	end if

	If trim(strLoginRede) = "T3FRRP" then 'or  trim(strLoginRede) = "EDAR" Then
		response.write "<script>alert('CLA_sp_sel_RepresentanteContratoProvedor')</script>"
		response.write "<script>alert('"&strProId&"')</script>"
		response.write "<script>alert('"&str_Est_sigla_PTA&"')</script>"
		response.write "<script>alert('"&str_Cid_Sigla_PTA&"')</script>"
	End IF

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
	'NOVAS PROCEDURES - SNOA - FIM'


Else
	Response.Write "<script language=javascript>window.location.replace('facilidade_main.asp')</script>"
	Response.End
End if

strUserGicL		= strUserName

'Deterministico
blnGravado = false

intCount = 0

Dim strStatus
strStatus = ""

%>

<script language='javascript' src="../javascript/solicitacao.js"></script>
<script language='javascript' src="../javascript/xmlObjects.js"></script>
<script language='javascript' src="../javascript/help.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>

<SCRIPT LANGUAGE=javascript>

	function GravarPortasSwitch ()
	{
		//alert("1");
		with (document.Form2)
		{
			//alert("2");
			campo = hdnIDrdoVlanSwitch.value;
			//alert(campo);
			//alert(txtDesigONT.value);
			
			if ( IsEmpty(campo) ){
				alert("Favor selecionar uma porta");
				return;
			}
			
			hdnvlanSwitch.value = document.getElementById("vlanSwitch" + campo).value;
			hdnportaoltSwitch.value = document.getElementById("portaPE" + campo).value;
			hdnpeSwitch.value = document.getElementById("pe" + campo).value;
			hdnSvlanSwitch.value = document.getElementById("svlan" + campo).value;
			hdndesigRadioIP.value = document.getElementById("desigRadioIP" + campo).value;
			 
			if (hdndisabledSvlan.value == "False"){	
						
				if (hdnSvlanSwitch.value==null || Trim(hdnSvlanSwitch.value)==""){
					alert("Favor informar a SVLAN");
					return;
				}
			}
			if (hdndisabledCvlan.value == "False" ){
				
				if (hdnvlanSwitch.value==null || Trim(hdnvlanSwitch.value)==""){
					alert("Favor informar a VLAN");
					return;
				}
			}
			
			if (hdnportaoltSwitch.value==null || Trim(hdnportaoltSwitch.value)==""){
				alert("Favor informar a Porta PE");
				return;
			}
			if (hdnpeSwitch.value==null || Trim(hdnpeSwitch.value)==""){
				alert("Favor informar a PE");
				return;
			}
			if (hdndesigRadioIP.value==null || Trim(hdndesigRadioIP.value)==""){
				alert("Favor informar a Designação do Radio IP");
				return;
			}
			
			
			for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
			 { 
					var elemento = document.Form2.elements[intIndex]; 
					if (elemento.name == "rdoPortaSwitchID" ){ 
							if (elemento.checked){
								hdnrdoPortaSwitchID.value = elemento.value ;
								//alert(elemento.value)
							}
					} 
			} 	
			
						
			target = "IFrmProcesso"
			action = "AlocarPortaSwitch.asp"
			submit()
			
		}
	}
	

function VoltarOrigem()
{
	var strPagina = new String("<%=Request.Form("hdnPaginaOrig")%>")

	//alert(strPagina)

	with (document.forms[0])
	{
		target = self.name
		if (strPagina == "" || strPagina.toUpperCase().indexOf("FACILIDADE.ASP") != -1)
		{
			action = "facilidade_main.asp"

		}else
		{
			//action = strPagina
			window.close();
		}
		//submit()
		window.close();
	}
}

</script>

<script language="VBScript">
function AlterarPorta()
	returnvalue=MsgBox ("Confirma a alteração do modelo do equipamento?",36,"Confirmação de alteração de equipamento.")
                
    If returnvalue=6 Then
		document.Form2.action = "facilidade.asp"
		document.Form2.hdnAlocar.value = "Alocar"
		document.Form2.submit()
	Else
         
    End If
End function
</SCRIPT>

<table border=0 cellspacing="1" cellpadding="1" width="760">
	
	<Form name="Form2" method="Post">
	
		<input type=hidden name=hdnUsuID value="<%=dblUsuID %>">
		<input type=hidden name=hdnAlocar>
		<input type=hidden name=hdnUsuario>
		<input type=hidden name=hdnOrigemEst>
		<input type=hidden name=hdnCNLEstUsu>
		<input type=hidden name=hdnComplEstUsu>
		
		<input type="hidden" name="hdnProvedor" value="<%=strProId%>">
		<input type="hidden" name="hdnEstacaoAtual" >
		<input type="hidden" name="hdnNomeLocal" >
		
		
		
		<input type=hidden	 name="hdnPlataforma" value="<%=strPla%>" >
		<input type="hidden" name="hdnSubAcao">
		<input type="hidden" name="hdnRede" value="<%=strRede%>">
		<input type="hidden" name="hdnPedId"		value="<%=Request.Form("hdnPedId")%>">
		<input type="hidden" name="hdnSolId"		value="<%=Request.Form("hdnSolId")%>">
		<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
		
	

		<input type="hidden" name="txtRazaoSocial"	value="<%=strCliente%>">
		<input type="hidden" name="hdnTipoProcesso"	value="<%=intTipoProcesso%>">
		<input type="hidden" name="hdnTipoAcao"	    value="<%=intTipoAcao%>">
		<input type="hidden" name="hdnFacDetid">
		<input type="hidden" name="hdnOEOrigem"     value="<%=orisolid%>">
		<input type="hidden" name="hdnDstId">


		<!-- Dados da Solicitação -->
		<input type="hidden" name="hdnSol_Id"	    value="<%=str_sol_id%>">
		<input type="hidden" name="hdnPed_Id"		value="<%=str_ped_id%>">
		<input type="hidden" name="hdnAcao"			value="<%=str_ACAO%>">
		

		<!-- Ação -->

		<!-- Dados Técnicos -->


		<!-- Dados do Acesso - Ponta A -->
		<input type="hidden" name="txtEndCid"		value="<%=strCidSigla%>">
		<input type="hidden" name="cboUFEnd"		value="<%=strUfSigla%>">
		<input type="hidden" name="cboLogrEnd"		value="<%=strTplSigla%>">
		<input type="hidden" name="txtEnd"			value="<%=strNomeLogr%>">
		<input type="hidden" name="txtNroEnd"		value="<%=strNroEnd%>">
		<input type="hidden" name="txtCepEnd"		value="<%=strCep%>">

		<!-- Validar Latitude e Longitude - Inicio-->
		<input type="hidden" name="hdntipoPonta_A"					value="A">
		<input type="hidden" name="hdnSgl_tipo_lograd_A"			value="<%=str_Tpl_Sigla_PTA%>">
		<input type="hidden" name="hdnDes_titulo_nome_lograd_A"		value="<%=str_End_NomeLogr_PTA%>">
		<input type="hidden" name="hdnDes_bairro_A"					value="<%=str_End_Bairro_PTA%>">
		<input type="hidden" name="hdnDes_localid_A"				value="<%=str_Cid_Desc_PTA%>">
		<input type="hidden" name="hdnDes_uf_A"						value="<%=str_Est_sigla_PTA%>">
		<input type="hidden" name="hdnNum_CEP_A"					value="<%=str_End_CEP_PTA%>">
		<input type="hidden" name="hdnTxtNum_A"						value="<%=str_End_NroLogr_PTA%>">
		<input type="hidden" name="hdnTxtComple_A"					value="<%=str_Aec_Complemento_PTA%>">
		
		<input type="hidden" name="hdntipoPonta_B"					value="B">
		<input type="hidden" name="hdnSgl_tipo_lograd_B"			value="<%=str_TPL_Sigla_PTB%>">
		<input type="hidden" name="hdnDes_titulo_nome_lograd_B"		value="<%=str_Esc_NomeLogr_PTB%>">
		<input type="hidden" name="hdnDes_bairro_B"					value="<%=str_Esc_Bairro_PTB%>">
		<input type="hidden" name="hdnDes_localid_B"				value="<%=str_Cid_Desc_PTB%>">
		<input type="hidden" name="hdnDes_uf_B"						value="<%=str_Est_Sigla_PTB%>">
		<input type="hidden" name="hdnNum_CEP_B"					value="<%=str_Esc_Cod_Cep_PTB%>">
		<input type="hidden" name="hdnTxtNum_B"						value="<%=str_Esc_NroLogr_PTB%>">
		<input type="hidden" name="hdnTxtComple_B"					value="<%=str_Esc_Complemento_PTB%>">

		<input type="hidden" name="hdntipoPonta">
		<input type="hidden" name="hdnSgl_tipo_lograd">
		<input type="hidden" name="hdnDes_titulo_nome_lograd">
		<input type="hidden" name="hdnDes_bairro">
		<input type="hidden" name="hdnDes_localid">
		<input type="hidden" name="hdnDes_uf">
		<input type="hidden" name="hdnNum_CEP">
		<input type="hidden" name="hdnTxtNum">
		<input type="hidden" name="hdnTxtComple">
		<!-- Validar Latitude e Longitude - Fim-->


		<input type="hidden" name="hdnEild"	value="<%=strEild%>">

		<input type=hidden name="hdnIdAcessoFisico" value="0">
		<input type=hidden name="hdnIdAcessoFisico1" value="0">
		<input type=hidden name="hdnCompartilhamento" value="0">
		<input type=hidden name="hdnCompartilhamento1" value="0">
		<input type=hidden name="hdnAlteracao" value="0">
		<input type=hidden name="hdnPodeAlterar" value="S">
		<input type=hidden name="hdnPropIdFisico" >
		<input type=hidden name="hdnPropIdFisico1" >
		<input type=hidden name="hdnXmlReturn" value='<%=FormatarStrXml(Request.Form("hdnXmlReturn"))%>'>

		<input type=hidden name="hdnIdLog" value="<%=strIdLogico%>">
		<input type=hidden name="hdnAcfIdRadio" value="<%=DblAcf_ID%>">

		<input type=hidden name="hdnPropAcesso" value="<%=strPropAcesso%>">
		<input type=hidden name="hdnTecId" value="<%=intTec%>">
		<input type=hidden name="hdnTecSigla" value="<%=strTecSigla%>">
		<input type=hidden name="hdnEnviarEmail">
		<input type=hidden name="hdnControleNovoPed" value="NAO">
		<input type=hidden name="hdnTecnologia" value="<%=strTecnologia%>">

		<input type=Hidden name=hdnIDrdoVlanSwitch > 
		<input type=Hidden name=hdnvlanSwitch> 
		<input type=Hidden name=hdnSvlanSwitch> 
		<input type=hidden name="hdndisabledSvlan" value="<%=disabledSvlan%>">
		<input type=hidden name="hdndisabledCvlan" value="<%=disabledCvlan%>">

		<input type=Hidden name=hdnportaoltSwitch>  
		<input type=Hidden name=hdnpeSwitch> 
		<input type=Hidden name=hdnQtdPorta >
		<input type=Hidden name=hdndesigRadioIP >
		<input type=Hidden name=hdnSwitchIP >
		<input type=Hidden name=hdnrdoPortaSwitchID >

		<%
		For Each Perfil in objDicCef
			if Perfil = "GAT" then dblCtfcIdGAT = objDicCef(Perfil)
			if Perfil = "GAE" then dblCtfcIdGAE = objDicCef(Perfil)
		Next
		%>
		<input type=hidden name="hdnCtfcIdGLA" value="<%=dblCtfcIdGAT%>">
		<input type=hidden name="hdnCtfcIdGLAE" value="<%=dblCtfcIdGAE%>">

		<input type="hidden" name="hdnIntIndice">
		<input type="hidden" name="hdnXml" >
		<input type="hidden" name="hdnXmlFacLibera">

		<!--Novo - Inicio-->
		<input type=hidden name=hdnQinQ>
		<input type=hidden name=hdnContratoFornec>
		<!--Novo - Fim-->

		<%
		dim idFornecedora
		idFornecedora = hdnContratoFornec
		%>

		<!--<tr><th colspan=6><p align=center>Alocação de Facilidade</p></th></tr>-->
		<tr><th colspan=6><p style="FONT-SIZE: 14px" align=center>Consulta de Pedido SNOA</p></th></tr>

		<tr>
			<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Dados do Pedido</th>
		</tr>
		<tr class=clsSilver>
			<td nowrap width=170><font class="clsObrig">:: </font>Fornecedora</td>
			<td>&nbsp;
				<span id="solicPedSnoaCboFornecedora">
					<select name="CboFornecedora" style="width:200px" onChange="ResgatarContrato(this.value);" disabled>
					<option value=""></option>
					<%	
						set objRS = db.execute("CLA_sp_sel_PrestadoraSNOA")
						While not objRS.Eof
							strItemSel = ""
							if Trim(str_Prestadora_id) = Trim(objRS("Prestadora_id")) then strItemSel = " Selected " End if
							Response.Write "<Option value=" & objRS("Prestadora_id") & strItemSel & ">" & objRS("Razao_social") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					%>
					</select>
				</span>
			</td>

			<td align=right><font class="clsObrig">:: </font>Contrato &nbsp;</td>
			<td colspan="3">&nbsp;
				<span id="solicPedSnoaCboContrFornec">
					<select name="CboContrFornec" style="width:200px" disabled>
					<option value=""></option>
					<%	
						set objRS = db.execute("CLA_sp_sel_ContratoPrestadoraSNOA " & str_Prestadora_id) 
						While not objRS.Eof
							strItemSel = ""
							if Trim(str_Contrato_id) = Trim(objRS("Bilhete_ContratoID")) then
								strItemSel = " Selected "
							End if
							Response.Write "<Option value=" & objRS("Bilhete_ContratoID") & strItemSel & ">" & objRS("Contrato_solicitante") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					%>
					</select>
				</span>
			</td>

		</tr>

		<tr>
			<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Dados da Solicitação</th>
			<!--<th align=rigth><a href="javascript:DetalharFac()"><font color=white>Mais...</font></a></th>-->
		</tr>

		<tr class=clsSilver>
			<td nowrap width=170>Solicitação de Acesso Nº</td>
			<td>&nbsp;<%=str_sol_id%></td>
			
			<td align=right>Id Lógico</td>
			<td>&nbsp;<%=str_Acl_IDAcessoLogico%></td>
			
			<td align=right >Data da Solicitação</td>
			<td>&nbsp;<%=str_DataSolicitacao%></td>
		</tr>

		<tr class=clsSilver>
			<td width=170>Designação da Solicitante</td>
			<td>&nbsp;<%=str_DM%></td>
			<td align=right>Data do Pedido</td>
			<td colspan=3 nowrap>&nbsp;<%=str_ped_data%></td>
		</tr>

		<tr class=clsSilver>
			<td width=170>Ação</td>
			<td colspan=5>&nbsp;<%=str_Acao%></td>
		</tr>

		<tr class=clsSilver>
			<td width=170>Status</td>
			<td colspan=5>&nbsp;<%=str_Status%></td>
		</tr>

		<tr class=clsSilver>
			<td width=170>Cliente</td>
			<td colspan="5">&nbsp;<%=str_Cliente%></td>
		</tr>

		<!--
		<tr class=clsSilver>
			<td width=170>Endereço</td>
			<td colspan="5">&nbsp;<%=strEndereco%></td>
		</tr>
		-->

		<tr class=clsSilver>
			<td width=170>Nº Contrato Serviço</td>
			<td nowrap>&nbsp;<%=str_NContratoServico%></td>
			
			<td align=right>Designação do Serviço</td>
			<td colspan=3>&nbsp;<%=str_DesignacaoServico%></td>
		</tr>

		<tr class=clsSilver>
			<td width=170>Serviço</td>
			<td >&nbsp;<%=str_Servico%></td>
			<td align=right>Velocidade do Serviço</td>
			<td colspan=3>&nbsp;<%=str_VelocidadeServico%></td>
		</tr>

		<tr class=clsSilver>
			<td width=170>Velocidade do Acesso Físico</td>
			<td colspan=5>&nbsp;<%=strVelFis%>&nbsp;<%=TipoVel(intTipoVel)%></td>
		</tr>
		
		<!-- Retirar
		<%
		if (intTipoProcesso = 4 and intTipoAcao = 4) or (intTipoProcesso = 3 and intTipoAcao = 3)then %>
			<tr class="clsSilver">
				<td width=170>Observações</td>        
				<td colspan="6"><textarea name="txtObsProvedor" cols="90" rows="2"><%=PedObs%></textarea></td>
			</tr>
		<%Else
 			if intTipoProcesso = 1 and intTipoAcao = 1 then%>	
				<tr class="clsSilver">
					<td width=170>Observações</td>        
					<td colspan="6"><textarea name="txtObsProvedor" cols="90" rows="2"><%=strObs%></textarea></td>
				</tr>
			<%End if%>
		<%End if%>
		-->

		<tr class=clsSilver>
			<td width=170>Finalidade</td>
			<td>&nbsp;
				<span id="solicPedSnoaCboFinalidade">
					<select name="cboFinalidade" style="width:200px" disabled>
					<option value=""></option>
					<%	
						set objRS = db.execute("CLA_sp_sel_Finalidade")
						While not objRS.Eof
							strItemSel = ""
							if Trim(str_Finalidade_id) = Trim(objRS("Finalidade")) then strItemSel = " Selected " End if
							Response.Write "<Option value=" & objRS("Finalidade") & strItemSel & ">" & objRS("Finalidade_desc") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					%>
					</select>
				</span>
			</td>

			<td align=right>Prazo de Contratação</td>
			<td colspan=3 nowrap>&nbsp;
				<span id="solicPedSnoaCboPrazContr">
					<select name="cboPrazContr" style="width:200px" disabled>
					<option value=""></option>
					<%	
						set objRS = db.execute("CLA_sp_sel_PrazoContratacao")
						While not objRS.Eof
							strItemSel = ""
							if Trim(str_Prazo_contratacao_id) = Trim(objRS("Prazo_contratacao")) then
								strItemSel = " Selected "
							End if
							Response.Write "<Option value=" & objRS("Prazo_contratacao") & strItemSel & ">" & objRS("Prazo_contratacao_desc") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					%>
					</select>
				</span>
			</td>

		</tr>

		<tr class=clsSilver>
			<td nowrap width=170>&nbsp;Estudo de Viabilidade</td>
			<td colspan=5>&nbsp;&nbsp;<textarea readonly name='txtDescrEstudo' cols='89' rows='2'><%=strObs%></textarea></td>
		</tr>

		<!-- CH-83646VWR - Inicio -->
		<tr class=clsSilver>
			<td nowrap width=170>&nbsp;Nº Pedido SNOA Anterior</td>
			<td colspan=5>&nbsp;&nbsp;<input type="text" class="text" name="txtNumPedSnoaAnt" value="<%=trim(strNumPedSnoaAnt)%>" size="89" readonly>
			</td>
		</tr>
		<!-- CH-83646VWR - Fim -->


</table>

<%
'Localiza informações da solução SSA
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblNroSev
	Vetor_Campos(2)="adInteger,2,adParamOutput,0"
	Call APENDA_PARAM("CLA_sp_sel_historico_solucao_ssa",2,Vetor_Campos)
	
	Set objRSCli = ObjCmd.Execute
	DBAction = ObjCmd.Parameters("RET").value

	if DBAction = 0 then
	
		If Not objRSCli.eof and  Not objRSCli.bof then%>
			
			<table cellspacing=1 cellpadding=0 border=0 width=760> 
				<tr class=clsSilver>
					<td width=321>Status do Pedido:  	 <b> <%= TratarAspasJS(Trim(objRSCli("STATUS_PEDIDO"))) %>  </b></td> 
				  	<td width=241>Status da SEV:  <b> <%=  TratarAspasJS(Trim(objRSCli("STATUS_SEV"))) %>  </b></td> 
				</tr> 
				
				<tr class=clsSilver>
					<td width=321>Nº do pedido de Venda do SNOA:  	 <b> <%=  TratarAspasJS(Trim(objRSCli("Pedido_Venda_SNOA"))) %>  </b></td> 
					<td width=241>Caracteristíca do EILD :  <b> <%=  TratarAspasJS(Trim(objRSCli("C_TIPO_EILD"))) %>  </b></td> 
				</tr> 
			
				<tr class=clsSilver>
					<td width=321>Ponta:  	 <b> <%=  TratarAspasJS(Trim(objRSCli("Ponta"))) %>  </b></td> 
					<td width=241>Nº SEV Ponta A :  <b> <%=  TratarAspasJS(Trim(objRSCli("sevPontaA_seq"))) %>  </b></td> 
				</tr> 
			
				<tr class=clsSilver>
					<td width=321>Operadora Solicitante:  	 <b><%= TratarAspasJS(Trim(objRSCli("OPR_DES"))) %> </b></td> 
					<td width=321>Pabx Virtual:  	 <b><%= TratarAspasJS(Trim(objRSCli("pabx_des"))) %> </b></td> 
				</tr> 
			
				<table cellpadding=0 cellspacing=1 width=760 > 
					<tr> 
						<th nowrap>&nbsp;Tecnologia</th> 
						<th>&nbsp;Provedor</th> 
						<th>&nbsp;Opc Indicada</th>
						<th>&nbsp;Atendimento</th> 
						<th>&nbsp;EILD</th> 
						<th nowrap >&nbsp;Estação</th> 
						<th>&nbsp;Usuário</th> 
						<th>&nbsp;Data</th> 
						<th>&nbsp;Análise</th> 
						<th>&nbsp;OBS</th> 
					</tr>
			
				<%While Not objRSCli.eof%>
				
					<tr class=clsSilver > 
						<td width=15% nowrap >&nbsp; <%=  TratarAspasJS(Trim(objRSCli("FAC_DES"))) %> </td>	
						<td >&nbsp; <%=  TratarAspasJS(Trim(objRSCli("FOR_DES"))) %> </td> 
						<td >&nbsp;<%= TratarAspasJS(Trim(objRSCli("SOL_SELECIONADA"))) %></td> 
						<td >&nbsp; <%=  TratarAspasJS(Trim(objRSCli("ATE_DESC"))) %> </td> 
						<td >&nbsp; <%=  TratarAspasJS(Trim(objRSCli("TIPO_EILD"))) %> </td> 
						<td nowrap>&nbsp; <%=  TratarAspasJS(Trim(objRSCli("ESTACAO"))) %> </td>
						<td nowrap >&nbsp; <%=  TratarAspasJS(Trim(objRSCli("SEV_USRNAMESOLUC"))) %> </td> 
						<td nowrap >&nbsp; <%=  TratarAspasJS(Trim(objRSCli("SEV_DTULTRESPSOLUC"))) %> </td> 
						<td >&nbsp; <%=  TratarAspasJS(Trim(objRSCli("ANALISE"))) %> </td> 
						<td width=20% >&nbsp; <%=  TratarAspasJS(Trim(objRSCli("SOL_OBS"))) %>  </td> 
					</tr> 
					<%						
					if ( ( Trim(objRSCli("Pro_ID")) = strProId ) and ( Trim(objRSCli("TIPO_EILD")) = "EILD PMS" ) ) then
						strPonta = Trim(objRSCli("Ponta"))
						strEild  = "S"
					end if
					objRSCli.MoveNext
				Wend%>
			
				<%IF strPabx_Virtual <> "" then %>
					<tr>
						<th colspan=9 class="clsSilver">&nbsp;PABX VIRTUAL BROADSOFT: <%= strPabx_Virtual %></th>
					</tr>
				<%END IF%>	

				</table> 
					
			</table> 
			
		<%End if
	
	End if%>

<%

	intTipoProcesso = 3 
	intTipoAcao = 3

	if intTipoProcesso = 3 and intTipoAcao = 3 then

		Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
		Vetor_Campos(2)="adInteger,2,adParamInput," & dblPedId
		strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_logEvento",2,Vetor_Campos)
		Set objRS = db.Execute(strSqlRet)

		if not objRS.Eof then
		
			if objRS("Retorno") = "Verdadeiro" then%>
				<table cellpadding=0 cellspacing=1 width=760>
					<tr>
						<tr>
							<th colspan=4 style="FONT-SIZE: 14px">&nbsp;•&nbsp;Informações de Origem</th>
						</tr>
					<tr>
						<th nowrap width=100 >&nbsp;Data</th>
						<th width=200>&nbsp;Campo</th>
						<th width=360>&nbsp;Valor</th>
						<th width=100>&nbsp;Usuário</th>
					</tr>
				<%
				blnCor = true
				While Not objRS.Eof
					
					if blnCor then
						strHtml = strHtml &  "<tr class=clsSilver >"
						blnCor = false
					Else
						strHtml = strHtml &  "<tr class=clsSilver2>"
						blnCor = true
					End if
					
					strHtml = strHtml &  "<td nowrap >&nbsp;"& Formatar_Data(objRS("Log_Data")) &"</td>"
					strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Log_Campo") & "</td>"
					strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Log_Valor") & "</td>"
					strHtml = strHtml &  "<td nowrap>&nbsp;"& objRS("Usu_UserName") & "</td>"
					strHtml = strHtml &  "</tr>"
					objRS.MoveNext
				Wend
				Response.Write strHtml
				%>
				</table>
			<%End if
		End if
	End if
%>

<%'GPON
	if isnull(strTecnologia) then
		strTecnologia = ""
	end if
%>

<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr>
		<td>
			<iframe	id			= "IFrmMotivoPend"
				    name        = "IFrmMotivoPend"
				    width       = "100%"
				    height      = "220px"
				    src			= "../inc/MotivoPendencia.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>"
				    frameborder = "0"
				    scrolling   = "no"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>

<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr>
		<td>
			<iframe	id			= "IFrmListaStatusSNOA"
				    name        = "IFrmListaStatusSNOA"
				    width       = "100%"
				    height      = "180px"
				    src			= "../inc/ListaStatusSNOA.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>&telaaceitar=0"
				    frameborder = "0"
				    scrolling   = "no"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>

<!-- CH-56418KBG - Inicio -->
<%

	//dblSolId = 1535099
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_designacaoFornecedora ",1,Vetor_Campos)
	Set objRS_DSF = db.Execute(strSqlRet)

	if not objRS_DSF.eof then %>

		<table border=0 cellspacing="1" cellpadding="0" width="760">
			<tr>
				<td>
					<iframe	id			= "IFrmListaDesigFornecSNOA"
						    name        = "IFrmListaDesigFornecSNOA"
						    width       = "100%"
						    height      = "60px"
						    src			= "../inc/ListaDesigFornecSNOA.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>&telaaceitar=<%=telaaceitar%>"
						    frameborder = "0"
						    scrolling   = "no"
						    align       = "left">
					</iFrame>
				</td>
			</tr>
		</table>

	<% end if %>	
<!-- CH-56418KBG - Fim -->

<%
	'fabio pinho'
	'response.write "<script>alert('"&Trim(intTipoProcesso)&"')</script>"
	'response.write "<script>alert('"&Trim(intTipoAcao)&"')</script>"

	if intTipoProcesso = 3 and intTipoAcao = 3 then


			'if objRS("Retorno") = "Verdadeiro" then%>
				<table cellpadding=0 cellspacing=1 width=760>
				
					<!--<tr><th colspan=6><p align=center>Solicitação de Pedido SNOA</p></th></tr>-->

					<tr>
						<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Ação</th>
					</tr>

					<tr class=clsSilver>
						<td nowrap width=170>&nbsp;Tipo da Ação</td>
						<td colspan=5>&nbsp;
							<span id="solicPedSnoaTipoAcao">
								<select name="cboTipoAcao" style="width:200px" disabled>
								<!--<option value=""></option>-->
								<%	
									set objRS = db.execute("CLA_sp_sel_AcaoSnoa " & str_tprc_id)
									While not objRS.Eof
										strItemSel = ""
										'if Trim(strDistrib) = Trim(objRS("Dst_ID")) then strItemSel = " Selected " End if
										Response.Write "<Option value=" & objRS("acaoSnoa_id") & strItemSel & ">" & objRS("acaoSnoa_desc") & "</Option>"
										objRS.MoveNext
									Wend
									strItemSel = ""
								
								%>
								</select>
							</span>
						</td>
						
					</tr>
					
					<% 	if str_tprc_id = "4" then%>
					<tr class=clsSilver>
						<td nowrap width=170>&nbsp;Motivo</td>
						<td colspan=5>&nbsp; 
								<span id="spncboMotivoSNOA">
									<select name="cboMotivoSNOA" style="width:200px" disabled>
									<option value=""></option>
									<%	
										set objRS = db.execute("CLA_sp_sel_MotivoSNOA")
										While not objRS.Eof
											strItemSel = ""
											'if Trim(str_vel_idTransmissao) = Trim(objRS("StsSnoa")) then strItemSel = " Selected " End if
											Response.Write "<Option value=" & objRS("motSnoa") & strItemSel & ">" & objRS("motSnoa_Desc") & "</Option>"
											objRS.MoveNext
										Wend
										strItemSel = ""
									%>
									</select>
								</span>
			
								
			
							</td>
							
						</tr>
						<tr class=clsSilver>
							<td nowrap width=170>&nbsp;Justificativa</td>
							<td colspan=5>				
									&nbsp;
									<textarea readonly name="txtMotivo" cols="80" rows="2" onKeyPress="MaxLength(this,1000)"></textarea>
							</td>
						</tr>
			
					<% end if%>

					<tr></tr>
					<!-- Dados de Faturamento e Cobrança - Inicio-->
					<tr>
						<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Dados de Faturamento e Cobrança</th>
					</tr>

					<tr class=clsSilver>
						<td width=170>&nbsp;Razão Social</td>
						<td colspan="5">&nbsp;
							<input type="text" class="text" name="txtCproContratanteRazao" value="<%=trim(str_Cpro_ContratanteRazao)%>" size="88" maxlength="50" readonly>
						</td>
					</tr>

					<tr class="clsSilver">
						<td width=170px>CNPJ</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtCproContratanteCNPJ"  maxlength="14"
							readonly
							size="25"  value="<%=trim(str_Cpro_ContratanteCGC_CNPJ)%>" >
						</td>
						<td width=170px>&nbsp;</font>Inscrição Estadual</td>
	    				<td>
	      					<input type="text" class="text" name="txtCproInscricaoEstadual"  value="<%=trim(str_Cpro_inscricao_estadual)%>" maxlength="50" size="27" readonly >&nbsp;
	    				</td>
					</tr>

					<tr class="clsSilver">
						<td width=170px>Endereço</td>
						<td colspan=3>&nbsp;
							<input type="text" class="text" name="txtCproContratanteEnd" value="<%=trim(str_Cpro_ContratanteEnd)%>" size="88" maxlength="50" readonly >
						</td>
					</tr>
					
					<tr class="clsSilver">
						<td width=170px>UF / Municipio</td>
						<td>&nbsp;
                        	<input type=text size=2 maxlength=4 class=text name="txtCproContratanteEstSigla" value="<%=trim(str_Cpro_ContratanteEstSigla)%>" readonly>
                        	&nbsp;
							<input type=text size=25 readonly  class=text
							name="txtCproCidDesc" readonly value="<%=trim(str_Cpro_Cid_Desc)%>">
                		</td>
	
	
						<td nowrap right>CEP&nbsp;</td>
						<td>
							<input type="text" class="text" name="txtCproContratanteCEP" value="<%=trim(str_Cpro_ContratanteCEP)%>" maxlength="9" size="12" readonly>
						</td>
					</tr>

					<tr class="clsSilver">
						<td width=170px>Contato</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtCproContratanteContato" value="<%=trim(str_Cpro_ContratanteContato)%>" maxlength="20" size="30" readonly>
						</td>
						<td>Email</td>
						<td>
							<input type="text" class="text" name="txtCproContratanteEmail" value="<%=trim(str_Cpro_ContratanteEmail)%>"	maxlength="50" size="27" readonly> <!--onkeyUp="ValidarTipo(this,0)">-->
						</td>
					</tr>

					<tr class="clsSilver">
						<td width=170px>Telefone</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtCproContratanteTelefone" value="<%=trim(str_Cpro_ContratanteTelefone)%>" maxlength="20" size="30" onkeyUp="ValidarTipo(this,0)" readonly>
						</td>
						<td>Fax</td>
						<td>
							<input type="text" class="text" name="txtCproContratanteFax" value="<%=trim(str_Cpro_ContratanteFax)%>"	maxlength="20" size="27" onkeyUp="ValidarTipo(this,0)" readonly>
						</td>
					</tr>


					<tr class="clsSilver">
						<td width=170px>Conta Customizada Ponta A</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtCproContaCustomizadaPonta_a" value="<%=trim(str_Cpro_conta_customizada_ponta_a)%>" maxlength="50" size="30" readonly>
						</td>
						<td>Conta Customizada Ponta B</td>
						<td>
							<input type="text" class="text" name="txtCproContaCustomizadaPonta_b" value="<%=trim(str_Cpro_conta_customizada_ponta_b)%>" maxlength="50" size="27" readonly>
						</td>
					</tr>
					<!-- Dados de Faturamento e Cobrança - Fim-->


					<tr></tr>
					<!-- Representante Comercial - Inicio-->
					<tr>
						<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Representante Comercial</th>
					</tr>

					<tr class=clsSilver>
						<td width=170>&nbsp;Nome</td>
						<td colspan="5">&nbsp;
							<input type="text" class="text" name="txtNome_ReprComer" value="<%=trim(str_Nome_ReprComer)%>" size="88" maxlength="50" readonly>
						</td>
					</tr>

					<!-- Inicio Endereco-->
					<tr class="clsSilver">
						<td width=170px>Endereço</td>
						<td colspan=3>&nbsp;
							<input type="text" class="text" name="txtEnd_ReprComer" value="<%=trim(str_End_ReprComer)%>" size="88" maxlength="50" readonly>
						</td>
					</tr>

					
					<tr class="clsSilver">
						<td width=170px>UF</td>
						<td>&nbsp;
							<select name="cboUFEnd_ReprComer"  disabled>
								<Option value=""></Option>
			            		<%
			                 		set objRS = db.execute("CLA_sp_sel_estado ''")
			                 		While not objRS.Eof
				             			strItemSel = ""
				              			if Trim(str_Est_sigla_ReprComer) = Trim(objRS("Est_Sigla")) then strItemSel = " Selected " End if
				              			Response.Write "<Option value=" & objRS("Est_Sigla") & strItemSel & ">" & objRS("Est_Sigla") & "</Option>"
				              			objRS.MoveNext
			                  		Wend
			                  		strItemSel = ""
		                   		%>
			       			</select>
		
						</td>

						<td nowrap right>CEP&nbsp;</td>
						<td>
							<input type="text" class="text" name="txtCepEnd_ReprComer" value="<%=trim(str_End_CEP_ReprComer)%>" maxlength="9" size="12" readonly>
						</td>
					</tr>

					<tr class="clsSilver">
						<td nowrap>Município (CNL - ANATEL)</td>
						<td nowrap colspan=3 >&nbsp;
                        	<input type=text size=5 maxlength=4 class=text name="txtEndCid_ReprComer" value="<%=str_Cid_Sigla_ReprComer%>"
							onBlur="if (ValidarTipo(this,1)){ResgatarCidadeSnoa(document.forms[0].cboUFEnd_ReprComer,1,this)}" readonly>&nbsp;
							<input type=text size=27 readonly style="BACKGROUND-COLOR:#eeeeee" class=text
							name="txtEndDescCid_ReprComer" readonly value="<%=str_Cid_Desc_ReprComer%>" tabIndex=-1 readonly>
                		</td>
					</tr>
	
	
					<tr class="clsSilver">
						<td>Email</td>
						<td nowrap colspan=3 >&nbsp;
							<input type="text" class="text" name="txtEmailTec_ReprComer" value="<%=strim(str_Email_ReprComer)%>" maxlength="50" size="88" readonly>&nbsp;
						</td>
					</tr>


					<tr class="clsSilver">
						<td colspan=4 align=right><span id=spnCEPSInstala_ReprComer></span></td>
					</tr>
	
					<tr class="clsSilver">
						<td width=170px>Telefone</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtTelefone_ReprComer" value="<%=trim(str_Telefone_ReprComer)%>" maxlength="20" size="30" readonly>
						</td>
						<td>Fax</td>
						<td>
							<input type="text" class="text" name="txtFax_ReprComer" value="<%=trim(str_Fax_ReprComer)%>"	maxlength="20" size="30" onkeyUp="ValidarTipo(this,0)" readonly>
						</td>
					</tr>
					<!-- Fim Endereco-->

					<!-- Representante Comercial - Fim-->
					<tr></tr>


					<tr>
						<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Dados Técnicos</th>
					</tr>

					<tr class=clsSilver>
						<td nowrap width=170><font class="clsObrig">:: </font>Taxa de Transmição</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboVelocidade">
								<select name="cboVelocidade" style="width:200px" disabled>
								<option value=""></option>
								<%	
									set objRS = db.execute("CLA_sp_sel_VelocidadeTransmissao ")
									While not objRS.Eof
										strItemSel = ""
										if Trim(str_vel_idTransmissao) = Trim(objRS("Vel_Transmissao")) then strItemSel = " Selected " End if
										Response.Write "<Option value=" & objRS("Vel_Transmissao") & strItemSel & ">" & objRS("Vel_desc") & "</Option>"
										objRS.MoveNext
									Wend
									strItemSel = ""
								%>
								</select>
							</span>
						</td>

						<td align=right >Característica Técnica &nbsp;</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboCaracTec">
								<select name="cboCaracTec" style="width:200px" disabled>
								<option value=""></option>
								<%	
									set objRS = db.execute("CLA_sp_sel_CaracteristicaTecnica")
									While not objRS.Eof
										strItemSel = ""
										if Trim(str_caracteristicaTecnica_id) = Trim(objRS("CaracteristicaTecnica")) then strItemSel = " Selected " End if
										Response.Write "<Option value=" & objRS("CaracteristicaTecnica") & strItemSel & ">" & objRS("CaracteristicaTecnica_desc") & "</Option>"
										objRS.MoveNext
									Wend
									strItemSel = ""
								%>
								</select>
							</span>
						</td>

					</tr>

					<tr class=clsSilver>
						<td nowrap width=170>&nbsp;Q in Q</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboQinQ" >
								<input type="radio" disabled="disabled" name="rdoQinQ" onClick="javascript:document.Form1.hdnQinQ.value = 'S';" value="S"
								<%if str_QinQ = "S" then%> checked <%end if%>>&nbsp; Sim
			
								<input type="radio" disabled="disabled" name="rdoQinQ" onClick="javascript:document.Form1.hdnQinQ.value = 'N';" value="N"
								<%if str_QinQ <> "S" then%> checked <%end if%>>&nbsp; Não
							</span>
						</td>

						<td align=right ><font class="clsObrig">:: </font>Quantidade de Linhas</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtQtdLinhas" value="<%=str_QtdLinha%>" size="30" maxlength="20" 
							onBlur="ValidarTipo(this,0)" readonly>
						</td>
					</tr>

					<tr class=clsSilver>
						<td nowrap width=170>&nbsp;Meio Preferencial</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboMeioPref">
								<select name="cboMeioPref" style="width:200px" disabled>
								<option value=""></option>
								<%	
									set objRS = db.execute("CLA_sp_sel_MeioPreferencial")
									While not objRS.Eof
										strItemSel = ""
										if Trim(str_meioPreferencial_id) = Trim(objRS("MeioPreferencial")) then strItemSel = " Selected " End if
										Response.Write "<Option value=" & objRS("MeioPreferencial") & strItemSel & ">" & objRS("MeioPreferencial_desc") & "</Option>"
										objRS.MoveNext
									Wend
									strItemSel = ""
								%>
								</select>
							</span>
						</td>

						<td align=right >Aplicação &nbsp;</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboAplicacao">
								<select name="cboAplicacao" style="width:200px" disabled>
								<option value=""></option>
								<%	
									set objRS = db.execute("CLA_sp_sel_Aplicacao")
									While not objRS.Eof
										strItemSel = ""
										if Trim(str_aplicacao_id) = Trim(objRS("Aplicacao")) then strItemSel = " Selected " End if
										Response.Write "<Option value=" & objRS("Aplicacao") & strItemSel & ">" & objRS("Aplicacao_desc") & "</Option>"
										objRS.MoveNext
									Wend
									strItemSel = ""
								%>
								</select>
							</span>
						</td>

					</tr>

					<tr class=clsSilver>
						<td colspan=6>
							<P style="FONT-SIZE: 9pt">
								A Característica Técnica definida pela fornecedora conforme disponibilidade constante no contrato, e não necessariamente, a que foi requisitada neste pedido.
								<br>
								A Característica Técnica deve seguir conforme disponibilidade da Oferta de Referência.
							</p>
						</td>
					</tr>

				</table>
<%
	End if
%>


<%
	'fabio pinho - segunda parte'
	'response.write "<script>alert('"&Trim(intTipoProcesso)&"')</script>"
	'response.write "<script>alert('"&Trim(intTipoAcao)&"')</script>"

	if intTipoProcesso = 3 and intTipoAcao = 3 then%>


			<!--
			'str_afp_id_A 					= objRS("afp_id")
			'str_acf_id_A 					= objRS("acf_id")
			'str_Latitude_A 					= objRS("Latitude")
			'str_Longitude_A 				= objRS("Longitude")
			'str_Equipamento_A 				= objRS("Equipamento")
			'str_CentroFios_A 				= objRS("CentroFios")
			'str_Ponto_Referencia_A 			= objRS("Ponto_Referencia")
			'str_Canais_desc_A 				= objRS("Canais_desc")
			'str_Entrega_Canalizado_A 		= objRS("Entrega_Canalizado")
			'str_E1_Canalizado_A 			= objRS("E1_Canalizado")
			'str_Time_Slot_A 				= objRS("Time_Slot")
			'str_InterfaceEletrica_id_A 		= objRS("InterfaceEletrica_id")
			'str_InterfaceFisica_id_A 		= objRS("InterfaceFisica_id")
			-->


				<!-- Dados do Acesso - Ponta A - Inicio -->
				<table cellpadding=0 cellspacing=1 width=760>
				
					<tr>
						<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Dados do Acesso - Ponta A</th>
					</tr>

					<tr class=clsSilver>
						<td width=170>&nbsp;Usuário</td>
						<td colspan="5">&nbsp;
							<input type="text" class="text" name="txtUsuario_A" value="<%=str_Cliente%>" size="80" maxlength="20" readonly>
						</td>
					</tr>


					<tr class=clsSilver>
						<td width=170>&nbsp;Latitude</td>
						<td colspan="4">
							&nbsp;
							<input type="text" class="text" name="txtLatEnd_A" value="<%=str_Latitude_A%>" size="20" maxlength="20" readonly>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							Longitude&nbsp;
							&nbsp;&nbsp;
							<input type="text" class="text" name="txtLongEnd_A" value="<%=str_Longitude_A%>" size="20" maxlength="20" readonly>
							&nbsp;&nbsp;<input type=button name=btnProcurarCoord_A value="Localizar Coordenadas" 	
							class="button" onclick="Validar_LatLong('A')" tabindex=-1 
							onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros dígitos (Alt+D)');" accesskey="D" disabled="disabled">
						</td>

					</tr>

					<!-- Inicio Endereco-->
					<tr class="clsSilver">
						<td width=170px><font class="clsObrig">:: </font>UF</td>
						<td>&nbsp;
							<select name="cboUFEnd_A" disabled>
								<Option value=""></Option>
			            		<%
			                 		set objRS = db.execute("CLA_sp_sel_estado ''")
			                 		While not objRS.Eof
				             			strItemSel = ""
				              			if Trim(str_Est_sigla_PTA) = Trim(objRS("Est_Sigla")) then strItemSel = " Selected " End if
				              			Response.Write "<Option value=" & objRS("Est_Sigla") & strItemSel & ">" & objRS("Est_Sigla") & "</Option>"
				              			objRS.MoveNext
			                  		Wend
			                  		strItemSel = ""
		                   		%>
			       			</select>
		
						</td>

						<td nowrap right><font class="clsObrig">:: </font>CEP&nbsp;</td>
						<td>
							<input type="text" class="text" name="txtCepEnd_A" readonly value="<%=str_End_CEP_PTA%>" maxlength="9" size="12">
						</td>
					</tr>

					<tr class="clsSilver">
						<td nowrap><font class="clsObrig">:: </font>Município (CNL - ANATEL)</td>
						<td nowrap colspan=3 >&nbsp;
                        	<input type=text size=5 maxlength=4 class=text name="txtEndCid_A" readonly value="<%=str_Cid_Sigla_PTA%>"
							onBlur="if (ValidarTipo(this,1)){ResgatarCidade(document.forms[1].cboUFEnd,1,this)}">&nbsp;
							<input type=text size=27 readonly style="BACKGROUND-COLOR:#eeeeee" class=text
							name="txtEndCidDesc" readonly value="<%=str_Cid_Desc_PTA%>" tabIndex=-1>
                		</td>
					</tr>
	
					<tr class="clsSilver">
						<td width=170px><font class="clsObrig">:: </font>Tipo do Logradouro</td>
						<td colspan=3>&nbsp;
							<select name="cboLogrEnd_A"  disabled>
							<option value=""></option>
							<% set objRS = db.execute("CLA_sp_sel_tplogradouro")
								While not objRS.Eof
									strItemSel = ""
									if Trim(str_Tpl_Sigla_PTA) = Trim(objRS("Tpl_Sigla")) then strItemSel = " Selected " End if
									Response.Write "<Option value=""" & Trim(objRS("Tpl_Sigla")) &""" " & strItemSel & ">" & Trim(objRS("Tpl_Sigla")) & "</Option>"
									objRS.MoveNext
								Wend
								strItemSel = ""
							%>
							</select>
						</td>
					</tr>
	
					<tr class="clsSilver">
						<td><font class="clsObrig">:: </font>Nome Logr</td>
						<td nowrap>&nbsp;
							<input type="text" class="text" name="txtEnd_A" readonly="TRUE" value="<%=str_End_NomeLogr_PTA%>" maxlength="60" size="40">
						</td>
						<td width=170px><font class="clsObrig">:: </font> Número</td>
						<td>
							<input type="text" class="text" name="txtNroEnd_A" value="<%=str_End_NroLogr_PTA%>" readonly="TRUE"	maxlength="10" size="10">
						</td>
					</tr>
	
					<tr class="clsSilver">
						<td>&nbsp;Complemento</td>
						<td colspan=3>&nbsp;
							<input type="text" class="text" name="txtComplEnd_A"  value="<%=str_Aec_Complemento_PTA%>" maxlength="120" size="80" readonly>
						</td>
					</tr>

					<tr class="clsSilver">
						<td>&nbsp;Ponto de Referência</td>
						<td colspan=3>&nbsp;
							<input type="text" class="text" name="txtPontoRefencia_A"  value="<%=str_Ponto_Referencia_A%>" maxlength="80" size="80" readonly>
						</td>
					</tr>

	
					<tr class="clsSilver">
						<td width=170px><font class="clsObrig">:: </font>Bairro</td>
						<td colspan = 3 >&nbsp;
							<input type="text" class="text" name="txtBairroEnd_A"
							value="<%=str_End_Bairro_PTA%>" maxlength="120" size="40" readonly="TRUE">&nbsp;
						</td>			
					</tr>

					<tr class="clsSilver">
						<td colspan=4 align=right><span id=spnCEPSInstala_A></span></td>
					</tr>
	
					<tr class="clsSilver">
						<td width=170px><font class="clsObrig">:: </font>Contato</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtContatoEnd_A" value="<%=str_Aec_Contato_PTA%>" maxlength="30" size="30">
						</td>
						<td><font class="clsObrig">:: </font>Telefone</td>
						<td>
							<input type="text" class="text" name="txtTelEnd_A" value="<%=str_Aec_Telefone_PTA%>"	maxlength="15" size="11" onkeyUp="ValidarTipo(this,0)" readonly>
						</td>
					</tr>
	
					<tr class="clsSilver">
						<td width=170px>CNPJ para fins de transporte (sem valor fiscal) </td>
						<td>&nbsp;
							<input type="text" class="text" name="txtCNPJ_A"  maxlength="14"
							<%=bbloqueia%>
							size="16" onKeyUp="ValidarTipo(this,0)" value="<%=str_Aec_CNPJ_PTA%>" >&nbsp;(99999999999999)
						</td>
						<td width=170px>&nbsp;</font>Email Contato Técnico</td>
	    				<td>
	      					<input type="text" class="text" name="txtEmailTec_A"  value="<%=str_Email_PTA%>" maxlength="30" size="30">&nbsp;
	    				</td>
					</tr>
					<!-- Fim Endereco-->

					<tr class=clsSilver>
						<td nowrap width=170><font class="clsObrig">:: </font>Incluir Equipamento</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboEquip">
								
								<input type="radio" disabled="disabled" name="rdoEquip_A" onClick="javascript:document.Form1.hdnQinQ.value = 'S';" value="S"
								<%if str_Equipamento_A = "S" then%> checked <%end if%>>&nbsp; Sim
			
								<input type="radio" disabled="disabled" name="rdoEquip_A" onClick="javascript:document.Form1.hdnQinQ.value = 'N';" value="N"
								<%if str_Equipamento_A <> "S" then%> checked <%end if%>>&nbsp; Não

							</span>
						</td>

						<td align=left><font class="clsObrig">:: </font>é o mesmo do Centro de Fios</td>
						<td>&nbsp;
							<span id="solicPedSnoaCentroFios">
								<input type="radio" disabled="disabled" name="rdoCentroFios_A" onClick="javascript:document.Form1.hdnQinQ.value = 'S';" value="S"
								<%if str_CentroFios_A = "S" then%> checked <%end if%>>&nbsp; Sim
			
								<input type="radio" disabled="disabled" name="rdoCentroFios_A" onClick="javascript:document.Form1.hdnQinQ.value = 'N';" value="N"
								<%if str_CentroFios_A <> "S" then%> checked <%end if%>>&nbsp; Não
							</span>
						</td>

					</tr>

					<tr class=clsSilver>
						<td nowrap width=170><font class="clsObrig">:: </font>Deseja entrega via Canalizado</td>
						<td colspan = 3 >&nbsp;
							<span id="solicPedSnoaEntrCanalizada_A">
								<input type="radio" name="rdoEntrCanalizada_A" disabled="disabled"
								onClick="javascript:document.forms[0].txtTimeSlot_A.disabled=false,document.forms[0].txtTimeSlot_A.value='', document.forms[0].txtE1Canalizado_A.disabled=false,document.forms[0].txtE1Canalizado_A.value='';" value="S"
								<%if str_Entrega_Canalizado_A = "S" then%> checked <%end if%>>&nbsp; Sim
			
								<input type="radio" name="rdoEntrCanalizada_A" disabled="disabled"
								onClick="javascript:document.forms[0].txtTimeSlot_A.disabled=true,document.forms[0].txtTimeSlot_A.value='', document.forms[0].txtE1Canalizado_A.disabled=true,document.forms[0].txtE1Canalizado_A.value='';" value="N"
								<%if str_Entrega_Canalizado_A <> "S" then%> checked <%end if%>>&nbsp; Não
							</span>
						</td>
					</tr>

					<tr class=clsSilver>

						<td nowrap width=170>&nbsp;Time Slot</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtTimeSlot_A"  disabled="disabled" value="<%=str_Time_Slot_A%>" size="28" maxlength="20" >
						</td>

						
						<td align=left >E1 Canalizado &nbsp;</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtE1Canalizado_A"  disabled="disabled" value="<%=str_E1_Canalizado_A%>" size="28" maxlength="20" >
						</td>
					</tr>

					<tr class=clsSilver>
						<td nowrap width=170>&nbsp;Interface Elétrica</td>
						<td>&nbsp;
							<span id="solicPedSnoaInterfEletr_A">
								<select name="cboInterfEletr_A" style="width:190px" disabled>
								<option value=""></option>
								<%	
									set objRS = db.execute("CLA_sp_sel_InterfaceEletrica")
									While not objRS.Eof
										strItemSel = ""
										if Trim(str_InterfaceEletrica_id_A) = Trim(objRS("InterfaceEletrica")) then 
											strItemSel = " Selected "
										End if
										Response.Write "<Option value=" & objRS("InterfaceEletrica") & strItemSel & ">" & objRS("InterfaceEletrica_desc") & "</Option>"
										objRS.MoveNext
									Wend
									strItemSel = ""

								%>
								</select>
							</span>
						</td>

						
						<td align=left >Interface Física &nbsp;</td>
						<td>&nbsp;
							<span id="solicPedSnoaInterfFisica_A">
								<select name="cboInterfFisica_A" style="width:190px" disabled>
								<option value=""></option>
								<%	
									set objRS = db.execute("CLA_sp_sel_InterfaceFisica")
									While not objRS.Eof
										strItemSel = ""
										if Trim(str_InterfaceFisica_id_A) = Trim(objRS("InterfaceFisica")) then
											strItemSel = " Selected "
										End if
										Response.Write "<Option value=" & objRS("InterfaceFisica") & strItemSel & ">" & objRS("InterfaceFisica_desc") & "</Option>"
										objRS.MoveNext
									Wend
									strItemSel = ""

								%>
								</select>
							</span>
						</td>
					</tr>
				</table>
				<!-- Dados do Acesso - Ponta A - Fim -->

				<!-- Dados do Acesso - Ponta B - Inicio -->
				<table cellpadding=0 cellspacing=1 width=760>
				
					<tr>
						<th style="FONT-SIZE: 14px" colspan=6>&nbsp;•&nbsp;Dados do Acesso - Ponta B</th>
					</tr>

					<tr class=clsSilver>
						<td width=170>&nbsp;Usuário</td>
						<td colspan="5">&nbsp;
							<input type="text" class="text" name="txtUsuario_B" value="EMBRATEL SA" size="80" maxlength="20" readonly>
						</td>
					</tr>


					<tr class=clsSilver>
						<td width=170>&nbsp;Latitude</td>
						<td colspan="4">
							&nbsp;
							<input type="text" class="text" name="txtLatEnd_B" value="<%=str_Latitude_B%>" size="20" maxlength="20" readonly>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							Longitude&nbsp;
							&nbsp;&nbsp;
							<input type="text" class="text" name="txtLongEnd_B" value="<%=str_Longitude_B%>" size="20" maxlength="20" readonly>
							&nbsp;&nbsp;<input type=button name=btnProcurarCoord_B value="Localizar Coordenadas " 	
							class="button" onclick="Validar_LatLong('B')" tabindex=-1 
							onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros dígitos (Alt+D)');" accesskey="D" disabled="disabled">
						</td>

					</tr>

					<!-- Inicio Endereco-->
					<tr class="clsSilver">
						<td width=170px><font class="clsObrig">:: </font>UF</td>
						<td>&nbsp;
							<select name="cboUFEnd_B" disabled>
								<Option value=""></Option>
			            		<%
			                 		set objRS = db.execute("CLA_sp_sel_estado ''")
			                 		While not objRS.Eof
				             			strItemSel = ""
				              			if Trim(str_Est_Sigla_PTB) = Trim(objRS("Est_Sigla")) then strItemSel = " Selected " End if
				              			Response.Write "<Option value=" & objRS("Est_Sigla") & strItemSel & ">" & objRS("Est_Sigla") & "</Option>"
				              			objRS.MoveNext
			                  		Wend
			                  		strItemSel = ""
		                   		%>
			       			</select>
		
						</td>

						<td nowrap right><font class="clsObrig">:: </font>CEP&nbsp;</td>
						<td>
							<input type="text" class="text" name="txtCepEnd_B"  readonly value="<%=str_Esc_Cod_Cep_PTB%>" maxlength="9" size="12">
						</td>
					</tr>

					<tr class="clsSilver">
						<td nowrap><font class="clsObrig">:: </font>Município (CNL - ANATEL)</td>
						<td nowrap colspan=3 >&nbsp;
                        	<input type=text size=5 maxlength=4 class=text name="txtEndCid_B" readonly value="<%=str_ESC_CID_SIGLA_PTB%>"
							onBlur="if (ValidarTipo(this,1)){ResgatarCidade(document.forms[1].cboUFEnd,1,this)}">&nbsp;
							<input type=text size=27 readonly style="BACKGROUND-COLOR:#eeeeee" class=text
							name="txtEndCidDesc" readonly value="<%=str_Cid_Desc_PTB%>" tabIndex=-1>
                		</td>
					</tr>

	
					<tr class="clsSilver">
						<td width=170px><font class="clsObrig">:: </font>Tipo do Logradouro</td>
						<td colspan=3>&nbsp;
							<select name="cboLogrEnd_B"  disabled>
							<option value=""></option>
							<% set objRS = db.execute("CLA_sp_sel_tplogradouro")
								While not objRS.Eof
									strItemSel = ""
									if Trim(str_Tpl_Sigla_PTB) = Trim(objRS("Tpl_Sigla")) then strItemSel = " Selected " End if
									Response.Write "<Option value=""" & Trim(objRS("Tpl_Sigla")) &""" " & strItemSel & ">" & Trim(objRS("Tpl_Sigla")) & "</Option>"
									objRS.MoveNext
								Wend
								strItemSel = ""
							%>
							</select>
						</td>
					</tr>
	
					<tr class="clsSilver">
						<td><font class="clsObrig">:: </font>Nome Logr</td>
						<td nowrap>&nbsp;
							<input type="text" class="text" name="txtEnd_B" readonly="TRUE" value="<%=str_Esc_NomeLogr_PTB%>" maxlength="60" size="40">
						</td>
						<td width=170px><font class="clsObrig">:: </font> Número</td>
						<td>
							<input type="text" class="text" name="txtNroEnd_B" value="<%=str_Esc_NroLogr_PTB%>" readonly="TRUE"	maxlength="10" size="10">
						</td>
					</tr>
	
					<tr class="clsSilver">
						<td>&nbsp;Complemento</td>
						<td colspan=3>&nbsp;
							<input type="text" class="text" name="txtComplEnd_B"  value="<%=str_Esc_Complemento_PTB%>" maxlength="120" size="80" readonly>
						</td>
					</tr>

					<tr class="clsSilver">
						<td>&nbsp;Ponto de Referência</td>
						<td colspan=3>&nbsp;
							<input type="text" class="text" name="txtPontoRefencia_B"  value="<%=str_Ponto_Referencia_B%>" maxlength="80" size="80" readonly>
						</td>
					</tr>

	
					<tr class="clsSilver">
						<td width=170px><font class="clsObrig">:: </font>Bairro</td>
						<td colspan = 3 >&nbsp;
							<input type="text" class="text" name="txtBairroEnd_B"
							value="<%=str_Esc_Bairro_PTB%>" maxlength="120" size="40" readonly="TRUE">&nbsp;
						</td>			
					</tr>

					<tr class="clsSilver">
						<td colspan=4 align=right><span id=spnCEPSInstala_B></span></td>
					</tr>
	
					<tr class="clsSilver">
						<td width=170px><font class="clsObrig">:: </font>Contato</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtContatoEnd_B" value="<%=str_Esc_Contato_PTB%>" maxlength="30" size="30">
						</td>
						<td><font class="clsObrig">:: </font>Telefone</td>
						<td>
							<!--
							<input type="text" class="text" name="txtTelEndArea_A" value="<%=str_Aec_Telefone_PTA%>"	maxlength="2" size="2" onkeyUp="ValidarTipo(this,0)">&nbsp;
							-->
							<input type="text" class="text" name="txtTelEnd_B" value="<%=str_Esc_Telefone_PTB%>"	maxlength="15" size="11" onkeyUp="ValidarTipo(this,0)" readonly>
						</td>
					</tr>
	
					<tr class="clsSilver">
						<td width=170px>CNPJ para fins de transporte (sem valor fiscal) </td>
						<td>&nbsp;
							<input type="text" class="text" name="txtCNPJ_B"  maxlength="14"
							<%=bbloqueia%>
							size="16" onKeyUp="ValidarTipo(this,0)" value="<%=str_Esc_CNPJ_PTB%>" readonly>&nbsp;(99999999999999)
						</td>
						<td width=170px>&nbsp;</font>Email Contato Técnico</td>
	    				<td>
	      					<input type="text" class="text" name="txtEmailTec_B"  value="<%=str_Esc_Email_PTB%>" maxlength="30" size="30" readonly>&nbsp;
	    				</td>
					</tr>
					<!-- Fim Endereco-->

					<tr class=clsSilver>
						<td nowrap width=170><font class="clsObrig">:: </font>Incluir Equipamento</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboEquip_B">
								
								<input type="radio" disabled="disabled" name="rdoEquip_B" onClick="javascript:document.Form1.hdnQinQ.value = 'S';" value="S"
								<%if str_Equipamento_B = "S" then%> checked <%end if%>>&nbsp; Sim
			
								<input type="radio" disabled="disabled" name="rdoEquip_B" onClick="javascript:document.Form1.hdnQinQ.value = 'N';" value="N"
								<%if str_Equipamento_B <> "S" then%> checked <%end if%>>&nbsp; Não

							</span>
						</td>

						<td align=left><font class="clsObrig">:: </font>é o mesmo do Centro de Fios</td>
						<td>&nbsp;
							<span id="solicPedSnoaCentroFios_B">
								<input type="radio" disabled="disabled" name="rdoCentroFios_B" onClick="javascript:document.Form1.hdnQinQ.value = 'S';" value="S"
								<%if str_CentroFios_B = "S" then%> checked <%end if%>>&nbsp; Sim
			
								<input type="radio" disabled="disabled" name="rdoCentroFios_B" onClick="javascript:document.Form1.hdnQinQ.value = 'N';" value="N"
								<%if str_CentroFios_B <> "S" then%> checked <%end if%>>&nbsp; Não
							</span>
						</td>

					</tr>

					<tr class=clsSilver>
						<td nowrap width=170><font class="clsObrig">:: </font>Deseja entrega via Canalizado</td>
						<td colspan = 3 >&nbsp;
							<span id="solicPedSnoaEntrCanalizada_B">
								<input type="radio" name="rdoEntrCanalizada_B" disabled="disabled"
								onClick="javascript:document.forms[0].txtTimeSlot_A.disabled=false,document.forms[0].txtTimeSlot_A.value='', document.forms[0].txtE1Canalizado_A.disabled=false,document.forms[0].txtE1Canalizado_A.value='';" value="S"
								<%if str_Entrega_Canalizado_B = "S" then%> checked <%end if%>>&nbsp; Sim
			
								<input type="radio" name="rdoEntrCanalizada_B" disabled="disabled"
								onClick="javascript:document.forms[0].txtTimeSlot_A.disabled=true,document.forms[0].txtTimeSlot_A.value='', document.forms[0].txtE1Canalizado_A.disabled=true,document.forms[0].txtE1Canalizado_A.value='';" value="N"
								<%if str_Entrega_Canalizado_B <> "S" then%> checked <%end if%>>&nbsp; Não
							</span>
						</td>
					</tr>

					<tr class=clsSilver>

						<td nowrap width=170>&nbsp;Time Slot</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtTimeSlot_B"  disabled="disabled" value="<%=str_Time_Slot_B%>" size="28" maxlength="20" >
						</td>

						
						<td align=left >E1 Canalizado &nbsp;</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtE1Canalizado_B"  disabled="disabled" value="<%=str_E1_Canalizado_B%>" size="28" maxlength="20" >
						</td>
					</tr>

					<tr class=clsSilver>
						<td nowrap width=170>&nbsp;Interface Elétrica</td>
						<td>&nbsp;
							<span id="solicPedSnoaInterfEletr_B">
								<select name="cboInterfEletr_B" style="width:190px" disabled>
								<option value=""></option>
								<%	
									set objRS = db.execute("CLA_sp_sel_InterfaceEletrica")
									While not objRS.Eof
										strItemSel = ""
										if Trim(str_InterfaceEletrica_id_B) = Trim(objRS("InterfaceEletrica")) then 
											strItemSel = " Selected "
										End if
										Response.Write "<Option value=" & objRS("InterfaceEletrica") & strItemSel & ">" & objRS("InterfaceEletrica_desc") & "</Option>"
										objRS.MoveNext
									Wend
									strItemSel = ""

								%>
								</select>
							</span>
						</td>

						
						<td align=left >Interface Física &nbsp;</td>
						<td>&nbsp;
							<span id="solicPedSnoaInterfFisica_B">
								<select name="cboInterfFisica_B" style="width:190px" disabled>
								<option value=""></option>
								<%	
									set objRS = db.execute("CLA_sp_sel_InterfaceFisica")
									While not objRS.Eof
										strItemSel = ""
										if Trim(str_InterfaceFisica_id_B) = Trim(objRS("InterfaceFisica")) then
											strItemSel = " Selected "
										End if
										Response.Write "<Option value=" & objRS("InterfaceFisica") & strItemSel & ">" & objRS("InterfaceFisica_desc") & "</Option>"
										objRS.MoveNext
									Wend
									strItemSel = ""

								%>
								</select>
							</span>
						</td>
					</tr>
				</table>
				<!-- Dados do Acesso - Ponta B - Fim -->

	<%
	End if
%>


<table width="760" border=0>
	<tr>
		<td>
			<table width=100% border=0 align=center cellspacing=1 cellpadding=1>
				<tr class=clsSilver2>
					<td colspan=4 align=center>
						<input type=button	class="button" name=btnVoltar value=Fechar onclick="VoltarOrigem()" accesskey="B" onmouseover="showtip(this,event,'Fechar (Alt+B)');">&nbsp;
						<!--
						<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
						-->
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>

<table width="760">
	<tr>
		<td>
			<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
		</td>
	</tr>
	<tr>
		<td>
			&nbsp;&nbsp;&nbsp;&nbsp;Legenda: A - Alfanumérico;  N - Numérico;  L - Letra
		</td>
	</tr>

</table>

<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso"
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no"
	    align       = "left">
</iFrame>
<iframe	id			= "IFrmProcesso2"
	    name        = "IFrmProcesso2"
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no"
	    align       = "left">
</iFrame>
<iframe	id			= "IFrmProcesso3"
	    name        = "IFrmProcesso3"
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no"
	    align       = "left">
</iFrame>
</body>
</html>

<SCRIPT LANGUAGE=javascript>
function sair()
{
	window.returnValue=0;
	if(window.history.length < 1)
	{
		window.close()
	}else
	{
		window.navigate(window.document.referrer)
	}
	
}
</script>
<%'GPON

if strTecnologia <> "GPON" AND strTecnologia <> "FO EDD" then%>
<SCRIPT LANGUAGE=javascript>
<!--
//var objXmlFacOrigem = new ActiveXObject("Microsoft.XMLDOM")
//objXmlFacOrigem.loadXML("<%=strXmlFacOrigem%>")
//RetornaCboPlataforma(document.forms[0].cboRede.value,<%=strPla%>)
//if (document.forms[0].cboRede.disabled) document.forms[0].cboRede.disabled = true
//-->
</SCRIPT>
<%end if 'GPON%>
