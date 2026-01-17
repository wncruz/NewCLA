<!--#include file="../inc/FacilidadesPedido.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsultaProcesso.asp
'	- Descrição			: Detalhe de uma solicitação

dblSolId = Trim(Request.Form("hdnSolId"))
dblPedId = Trim(Request.Form("hdnPedId"))

If Trim(dblSolId) = "" then
	Response.Write "<script language=javascript>window.location.replace('main.asp')</script>"
	Response.End
End if

Set objRSSolic =	db.execute("CLA_sp_view_solicitacaomin " & dblSolId)

if objRSSolic.Eof and objRSSolic.Bof then 'Não existe registros
	Response.Write "<script language=javascript>alert('Pedido indisponível.');window.location.replace('main.asp');</script>"
	Response.End
End if

if Not objRSSolic.Eof and Not objRSSolic.Bof then
	
	strPabx_Virtual     = Trim(objRSSolic("PABX_VIRTUAL"))

	IF ISNULL(strPabx_Virtual) THEN
		strPabx_Virtual = ""
	END IF
	
	IF strPabx_Virtual = "NULL" THEN
		strPabx_Virtual = ""
	END IF
	
    strAntAcesso = objRSSolic("Acl_AntAcesso")
	strNroSev = Trim(objRSSolic("Sol_SevSeq"))
	strRazaoSocial = Trim(objRSSolic("Cli_Nome"))
	strNomeFantasia = Trim(objRSSolic("Cli_NomeFantasia"))
	strContaSev = Trim(objRSSolic("Cli_CC"))
	strSubContaSev = Trim(objRSSolic("Cli_SubCC"))
	
	segmento = Trim(objRSSolic("segmento"))
	porte = Trim(objRSSolic("porte"))
	
	intTipoProc	= Trim(objRSSolic("Tprc_ID"))
''@@ LPEREZ - 24/10/2005
	strGrupoCli	= Trim(objRSSolic("GCli_ID"))
	
	strIA = Trim(objRSSolic("IA"))
	strPONTA = Trim(objRSSolic("ID_ENDERECO"))
	'response.write "<script>alert('"&strIA&"')</script>"
''@@LP


	if Trim(objRSSolic("Sol_OrderEntry")) <> "" then
		strOrderEntry = Trim(objRSSolic("Sol_OrderEntry"))
		intTamSis = len(strOrder)-12
	End if

	'Solicitação
''@@ LPEREZ - 24/10/2005
	strOrigemSol	= Trim(objRSSolic("OriSol_ID"))
	strProjEspecial = Trim(objRSSolic("Sol_IndProjEspecial"))
''@@LP

	strDtPedido = Formatar_Data(Trim(objRSSolic("Sol_Data")))
	strTipoContratoServico = Trim(objRSSolic("Acl_TipoContratoServico"))
	strNroContrServico = Trim(objRSSolic("Acl_NContratoServico"))
	strIdLogico = Trim(objRSSolic("Acl_IDAcessoLogico"))

	strDesigAcessoPri	= Trim(objRSSolic("Acl_IdAcessoLogicoPrincipal"))
	'''strDtDesEntrAcesServ= Formatar_Data(objRSSolic("Acl_DtDesejadaEntregaAcessoServico"))
	'''strDtEntrAcesServ	= Formatar_Data(Trim(objRSSolic("Acl_DtEntregaServico")))
	'''strDtConcluSolic	= Formatar_Data(Trim(objRSSolic("Sol_DtConclusao")))

	'''strDtConstrAcesFisTer 	= Formatar_Data(Trim(objRSSolic("Sol_DtConclusao"))) 		'Data Construção Acesso Físico Terceiro
	'''strDtDesaAcesFisTer 	= Formatar_Data(Trim(objRSSolic("Acl_DtDesaticavaoServico"))) 	'Data Desativação do Acesso Físico Terceiro
	'strDtCanPedAcesFisTer	= 'Data Cancelamento de Pedido de Acesso Físico Terceiro
	'''strDtAtivServ		= Formatar_Data(Trim(objRSSolic("Acl_DtAtivacaoServico")))
	'''strDtAtivServ		= Formatar_Data(Trim(objRSSolic("Acl_DtAtivacaoServico"))) 'Data da Ativação do Serviço
	'''strDtCanSolic		= Formatar_Data(Trim(objRSSolic("Sol_DtCancelamento"))) 'Data de Cancelamento da Solicitação

	'strDtPrevEntrAcesProv = Formatar_Data(Trim(objRSSolic("Ped_DtPrevistaAtendProv")))
	strDtIniTemp = Formatar_Data(Trim(objRSSolic("Acl_DtIniAcessoTemp")))
	strDtFimTemp = Formatar_Data(Trim(objRSSolic("Acl_DtFimAcessoTemp")))
	strDtDevolucao = Formatar_Data(Trim(objRSSolic("Acl_DtDevolAcessoTemp")))
	strServicoPedido = Trim(objRSSolic("Ser_Desc"))
	strVelServico = Trim(objRSSolic("DescVelAcessoLog"))
	strDesigncaoServico = Trim(objRSSolic("Acl_DesignacaoServico"))

	'''StrDtInicioAvaliacao  = Formatar_Data(Trim(objRSSolic("Sol_DtInicioAvaliacao")))
	'''StrDtFimAvaliacao  = Formatar_Data(Trim(objRSSolic("Sol_DtFimAvaliacao")))


    'PSOUTO
    'REMOVIDO - NÃO EXISTE USO PARA NADA
	'if strRegimeCntr <> "" then
	'	Set objRS = db.execute("CLA_sp_sel_regimecontrato 0," & strProvedor)
	'	if Not objRS.Eof and Not objRS.Bof then
	'		strRegimeCntr =  Trim(objRS("Pro_Nome")) & " - " & Trim(objRS("Tct_Desc"))
	'	End if
	'End if
	

	'if strProvedor <> "" then
	'	Set objRS = db.execute("CLA_sp_sel_provedor " & strProvedor)
	'	if Not objRS.Eof and Not objRS.Bof then
	'		strProvedor = objRS("Pro_Nome")
	'	End if
	'
	'End if
	
	

	'strPromocao = Trim(objRSSolic("Prm_ID"))
	'if strPromocao <> "" then
	'	Set objRS = db.execute("CLA_sp_sel_promocaoprovedor " & strPromocao)
	'	if Not objRS.Eof and Not objRS.Bof then
	'		strPromocao = objRS("Prm_Desc")
	'	End if
'
	'End if
	
	' /PSOUTO

	
	
	
	'Eduardo Araujo RN-88361 item 3.1
	
	'acesso = trim (objRSSolic("Acf_Proprietario"))
	'response.write "<script>alert('"&acesso&"')</script>"
	if acesso = "EBT" THEN
		' Data Construção do Acesso Físico EBT
		strDtConstrAcesFisEBT   = Formatar_Data(objRSSolic("acf_dtConstrAcessoFis"))
		' Data Desativação do Acesso Físico EBT
		strDtDesaAcesFisEBT 	= Formatar_Data(objRSSolic("acf_dtDesatAcessoFis"))
		' Data Cancelamento do Pedido de Acesso Físico EBT
		strDtCanPedAcesFisEBT	= Formatar_Data(objRSSolic("Acf_DtCancAcessoFis"))
		' Data Construção Acesso Físico Terceiro
		strDtConstrAcesFisTer 	= NULL
		' Data Desativação do Acesso Físico Terceiro
		strDtDesaAcesFisTer	= NULL
		' Data Cancelamento de Pedido de Acesso Físico Terceiro
		strDtCanPedAcesFisTer	= NULL
	
	ELSE 
		' Data Construção Acesso Físico Terceiro
		strDtConstrAcesFisTer 	= Formatar_Data(objRSSolic("acf_dtConstrAcessoFis"))
		' Data Desativação do Acesso Físico Terceiro
		strDtDesaAcesFisTer	= Formatar_Data(objRSSolic("acf_dtDesatAcessoFis"))
		' Data Cancelamento de Pedido de Acesso Físico Terceiro
		strDtCanPedAcesFisTer	= Formatar_Data(objRSSolic("Acf_DtCancAcessoFis"))
		' Data Construção do Acesso Físico EBT
		strDtConstrAcesFisEBT   = NULL
		' Data Desativação do Acesso Físico EBT
		strDtDesaAcesFisEBT 	= NULL
		' Data Cancelamento do Pedido de Acesso Físico EBT
		strDtCanPedAcesFisEBT	= NULL
	
	END IF 
	
	
	' Data de Entrega do Acesso ao Serviço
	strDtEntrAcesServ	= Formatar_Data(objRSSolic("Acl_DtEntregaServico"))
	
	' Data da Ativação do Serviço
	strDtAtivServ		= Formatar_Data(objRSSolic("Sol_DtConclusao"))
	
	' Data de Conclusão da Solicitação
	strDtConcluSolic	= Formatar_Data(objRSSolic("Sol_DtConclusao"))
	
	' Data de Cancelamento da Solicitação
	strDtCanSolic		= Formatar_Data(objRSSolic("sol_dtCancelamento"))
	
	' Data Inicio de Avaliação
	StrDtInicioAvaliacao	= Formatar_Data(objRSSolic("sol_dtInicioAvaliacao"))
	
	' Data Fim de Avaliação
	StrDtFimAvaliacao	= Formatar_Data(objRSSolic("sol_dtFimAvaliacao"))	
	
	''' Eduardo Araujo RN-88361 item 3.1 Témino
	


	strObs = Trim(objRSSolic("Sol_Obs"))

	'Soluções indicadas pelo SSA
	''if Trim(strNroSev) <> "" then

	'''	Vetor_Campos(1)="adInteger,2,adParamInput," & strNroSev
	'''	Vetor_Campos(2)="adInteger,2,adParamOutput,0"
	'''	Call APENDA_PARAM("CLA_sp_sel_solucao_ssa",2,Vetor_Campos)
	'''	ObjCmd.Execute'pega dbbtnGravar
	'''	DBAction = ObjCmd.Parameters("RET").value
	'''	if DBAction = 0 then 'Sucesso

	'''		Set objRSCli = db.execute("CLA_sp_sel_solucao_ssa " & strNroSev & ",0")
	'''		if Not objRSCli.eof and Not objRSCli.bof then
	'''			strSolSel = "<table border=0 cellspacing=1 cellpadding=1><tr class=clsSilver2><td>Provedor</td><td>Facilidade</td><td>Prazo</td></tr>"
				'Soluções indicadas pelo SSA
	'''			While Not objRSCli.eof
	'''				if Trim(objRSCli("Sol_Selecionada")) = 1 then
	'''					strSolSel = strSolSel & "<tr><td>" & Trim(objRSCli("For_Des")) & "</td><td bgcolor=#f2f2f2>" & Trim(objRSCli("Fac_Des")) & "</td><td bgcolor=#f2f2f2>" & Trim(objRSCli("Sol_PrazoCompleto")) & "</td></tr>"
	'''				End if
	'''				objRSCli.MoveNext
	'''			Wend
	'''			strSolSel = strSolSel & "</table>"
	'''			strRespostaSSA = strSolSel
	'''		End if
	'''	End if
	'''End if
	
	
	
	
	
	''%%''
	
	if Trim(strNroSev) <> "" then
			
			'Localiza informações da solução SSA
			Vetor_Campos(1)="adInteger,2,adParamInput," & strNroSev
			Vetor_Campos(2)="adInteger,2,adParamOutput,0"
			'Call APENDA_PARAM("CLA_sp_sel_solucao_ssa",2,Vetor_Campos)
			Call APENDA_PARAM("CLA_sp_sel_historico_solucao_ssa",2,Vetor_Campos)
			
			Set objRSCli = ObjCmd.Execute
			DBAction = ObjCmd.Parameters("RET").value
		
			if DBAction = "0" then
				
				If Not objRSCli.eof and  Not objRSCli.bof then
		
					strSolSel = " <table cellspacing=1 cellpadding=0 border=0 width=760> "
					strSolSel = strSolSel & "	<tr class=clsSilver>"
					strSolSel = strSolSel & "	  <td width=321>Status do Pedido:  	 <b>" & TratarAspasJS(Trim(objRSCli("STATUS_PEDIDO"))) & " </b></td> "
					strSolSel = strSolSel & "	  <td width=241>Status da SEV:  <b>" & TratarAspasJS(Trim(objRSCli("STATUS_SEV"))) & " </b></td> "
					strSolSel = strSolSel & "	</tr> "
					strSolSel = strSolSel & "	<tr class=clsSilver>"
					strSolSel = strSolSel & "	  <td width=321>Nº do pedido de Venda do SNOA:  	 <b>" & TratarAspasJS(Trim(objRSCli("Pedido_Venda_SNOA"))) & " </b></td> "
					strSolSel = strSolSel & "	  <td width=241>Caracteristíca do EILD :  <b>" & TratarAspasJS(Trim(objRSCli("C_TIPO_EILD"))) & " </b></td> "
					strSolSel = strSolSel & "	</tr> "
					
					strSolSel = strSolSel & "	<tr class=clsSilver>"
					strSolSel = strSolSel & "	  <td width=321>Ponta:  	 <b>" & TratarAspasJS(Trim(objRSCli("Ponta"))) & " </b></td> "
					strSolSel = strSolSel & "	  <td width=241>Nº SEV Ponta A :  <b>" & TratarAspasJS(Trim(objRSCli("sevPontaA_seq"))) & " </b></td> "
					strSolSel = strSolSel & "	</tr> "
					
					strSolSel = strSolSel & "	<tr class=clsSilver>"
					strSolSel = strSolSel & "	  <td width=321>Operadora Solicitante:  	 <b>" & TratarAspasJS(Trim(objRSCli("OPR_DES"))) & " </b></td> "
					'strSolSel = strSolSel & "	  <td ></td> "
					strSolSel = strSolSel & "	  <td width=321>Pabx Virtual:  	 <b>" & TratarAspasJS(Trim(objRSCli("pabx_des"))) & " </b></td> "
					''strSolSel = strSolSel & "	  <td ></td> "
					strSolSel = strSolSel & "	</tr> "
					
					strSolSel = strSolSel & "	  		<table cellpadding=0 cellspacing=1 width=760 > "
					strSolSel = strSolSel & "				<tr> "
					strSolSel = strSolSel & "					<th nowrap>&nbsp;Tecnologia</th> "
					strSolSel = strSolSel & "					<th>&nbsp;Provedor</th> "
					'strSolSel = strSolSel & "					<th>&nbsp;Opc Indicada</th> "							
					'strSolSel = strSolSel & "					<th>&nbsp;Atendimento</th> "
					'strSolSel = strSolSel & "					<th>&nbsp;EILD</th> "
					strSolSel = strSolSel & "					<th >&nbsp;Estação</th> "
					strSolSel = strSolSel & "					<th>&nbsp;Usuário</th> "
					strSolSel = strSolSel & "					<th>&nbsp;Data</th> "
					'strSolSel = strSolSel & "					<th>&nbsp;Análise</th> "
					strSolSel = strSolSel & "					<th>&nbsp;OBS</th> "
					strSolSel = strSolSel & "				</tr> "

					
					
						
					While Not objRSCli.eof
						
						
						strSolSel = strSolSel & "				<tr class=clsSilver > "
						strSolSel = strSolSel & "					<td width=15% nowrap >&nbsp;" & TratarAspasJS(Trim(objRSCli("FAC_DES"))) & "</td>	"
						strSolSel = strSolSel & "					<td >&nbsp;" & TratarAspasJS(Trim(objRSCli("FOR_DES"))) & "</td> "
						'strSolSel = strSolSel & "					<td >&nbsp;" & TratarAspasJS(Trim(objRSCli("SOL_SELECIONADA"))) & "</td> "
						'strSolSel = strSolSel & "					<td >&nbsp;" & TratarAspasJS(Trim(objRSCli("ATE_DESC"))) & "</td> "
						'strSolSel = strSolSel & "					<td >&nbsp;" & TratarAspasJS(Trim(objRSCli("TIPO_EILD"))) & "</td> "
						strSolSel = strSolSel & "					<td nowrap>&nbsp;" & TratarAspasJS(Trim(objRSCli("ESTACAO"))) & "</td> "
						strSolSel = strSolSel & "					<td nowrap >&nbsp;" & TratarAspasJS(Trim(objRSCli("SEV_USRNAMESOLUC"))) & "</td> "
						strSolSel = strSolSel & "					<td nowrap >&nbsp;" & TratarAspasJS(Trim(objRSCli("SEV_DTULTRESPSOLUC"))) & "</td> "
						'strSolSel = strSolSel & "					<td >&nbsp;" & TratarAspasJS(Trim(objRSCli("ANALISE"))) & "</td> "
						strSolSel = strSolSel & "					<td width=20% >&nbsp;" & TratarAspasJS(Trim(objRSCli("SOL_OBS"))) & " </td> "
						strSolSel = strSolSel & "				</tr> "
						
						objRSCli.MoveNext
					Wend
					
					strSolSel = strSolSel & "			</table> "
							
					strSolSel = strSolSel & " </table> "
					
					strRespostaSSA = strSolSel
					'Response.Write "<script language=javascript>parent.strProvedorSelSev.innerHTML = '" & strSolSel & "';</script>"
					
		
				End if
			End if
			
	End if

	
	''%%''
	

	'Configuração
	dblEscIdEntrega = Trim(objRSSolic("Esc_IDEntrega"))
	dblEscIdConfig = Trim(objRSSolic("Esc_IDConfiguracao"))

	'Endereço do local de instalação
	if Trim(dblEscIdEntrega) <> "" then
		Set objRS = db.execute("CLA_sp_sel_estacao " & dblEscIdEntrega)
		if Not objRS.Eof And Not objRS.Bof then
			strLocalInstala = objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla")
			strEndereco = Trim(Cstr("" & objRS("Tpl_Sigla"))) & " " &  Trim(Cstr("" & objRS("Esc_NomeLogr"))) & " nº " & Trim(Cstr("" & objRS("Esc_NroLogr"))) & " - " & Trim(Cstr("" & objRS("Est_Sigla"))) & " - " & Trim(Cstr("" & objRS("Cid_Sigla")))
			strContEndLocalInstala	=	objRS("Esc_Contato")
			strTelEndLocalInstala	=	objRS("Esc_Telefone")
		End if
	End if

	'Endereço do local de instalação
	if Trim(dblEscIdConfig) <> "" then
		Set objRS = db.execute("CLA_sp_sel_estacao " & dblEscIdConfig)
		if Not objRS.Eof And Not objRS.Bof then
			strLocalConfig = objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla")
			strEndereco = Trim(Cstr("" & objRS("Tpl_Sigla"))) & " " &  Trim(Cstr("" & objRS("Esc_NomeLogr"))) & " nº " & Trim(Cstr("" & objRS("Esc_NroLogr"))) & " - " & Trim(Cstr("" & objRS("Est_Sigla"))) & " - " & Trim(Cstr("" & objRS("Cid_Sigla")))
			strContEndLocalConfig	= objRS("Esc_Contato")
			strTelEndLocalConfig	= objRS("Esc_Telefone")
		End if
	End if



	'Usuario de coordenação embratel
	Set objRS = db.execute("CLA_sp_view_agentesolicitacao " & dblSolId)

	if Not objRS.Eof  and  Not objRS.bof then
		While Not objRS.Eof
			Select Case Trim(Ucase(objRS("Age_Desc")))
				Case "GLA"
					strGLA = Trim(objRS("Usu_Username"))
					strNomeGLA = Trim(objRS("Usu_Nome"))
					strRamalGLA = Trim(objRS("Usu_Ramal"))
				Case "GICN"
					strGICN = Trim(objRS("Usu_Username"))
					strNomeGICN = Trim(objRS("Usu_Nome"))
					strRamalGICN = Trim(objRS("Usu_Ramal"))
				Case "GICL"
					strGICL = Trim(objRS("Usu_Username"))
					strNomeGICL = Trim(objRS("Usu_Nome"))
					strRamalGICL = Trim(objRS("Usu_Ramal"))
					if Trim(objRS("Agp_Origem")) = "P" then
						strUserGICL = strGICL
					End if
				Case "GLAE"
					strGLAE = Trim(objRS("Usu_Username"))
					strNomeGLAE = Trim(objRS("Usu_Nome"))
					strRamalGLAE = Trim(objRS("Usu_Ramal"))

				Case "AVALIADOR"
					strAVL = Trim(objRS("Usu_Username"))
					strNomeAVL = Trim(objRS("Usu_Nome"))
					strRamalAVL = Trim(objRS("Usu_Ramal"))


			End Select
			objRS.MoveNext
		Wend
	End if

	'Descrição da cidade para o endereço de instalação
	Set objRS = db.execute("CLA_sp_sel_cidade '" & strUFEndInstala & "','" & strEndCidInstala & "'")
	if Not objRS.Eof  and  Not objRS.bof then
		 strEndCidInstala = LimparStr(Trim(objRS("Cid_Desc")))
	End if

	'Descrição da cidade para o endereço do ponto intemediário
	Set objRS = db.execute("CLA_sp_sel_cidade '" & strUFEndInterme & "','" & strEndCidInterme & "'")
	if Not objRS.Eof then
		 strEndCidinterme = LimparStr(Trim(objRS("Cid_Desc")))
	End if

	dblOrgId = Trim(objRSSolic("Org_Id"))
	strStatus = Trim(objRSSolic("Sts_id"))
	if Trim(strStatus) <> "" then
		Set objRS = db.execute("CLA_sp_sel_Status " & strStatus)
		if Not objRS.Eof And Not objRS.Bof then
			strStatus = objRS("Sts_Desc")
		End if
	End if

	strHistorico = Trim(objRSSolic("StsSol_Historico"))

End if
%>
<!--INÍCIO CAMPOS PROCESSO -->
<table cellspacing="1" cellpadding="1" width="760" border="0">
    <tr>
        <th width="25%">&nbsp;•&nbsp;Solicitação de Acesso</th>
        <th width="25%">&nbsp;Nº&nbsp;:&nbsp;<span id="spnSolId"><%=dblSolId%></span></th>
        <th width="25%">&nbsp;Tipo&nbsp;:&nbsp;
		<%
		Select Case Cint("0" & intTipoProc)
			Case 1
				Response.Write "Ativação"
			Case 2
				Response.Write "Desativação"
			Case 3
				Response.Write "Alteração"
			Case 4
				Response.Write "Cancelamento"
		End Select
        %>
        </th>
        <th width="25%">&nbsp;Data&nbsp;:&nbsp;<%=strDtPedido%></th>
    </tr>
    <tr>
        <th colspan="4">
        &nbsp;•&nbsp;Informações do Cliente</td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Nro SEV</td>
        <td colspan="3">&nbsp;<%=strNroSev%></td>
    </tr>

    <!-- LPEREZ - 24/10/2005 -->
    <tr class="clsSilver">
        <td width="170">Projeto Especial&nbsp;</td>
        <td colspan="3">
            <table border="0" cellspacing="0" width="100%">
                <tr width="100%">
                    <td width="100px">&nbsp;<%=strProjEspecial%></td>
                    <td width="100px">Grupo&nbsp;</td>
                    <td width="100px">&nbsp;
						<%
						set gr = db.execute("CLA_sp_sel_GrupoCliente 0")
							do while not gr.eof
								if trim(strGrupoCli) = trim(gr("GCli_ID")) then
									Response.Write ucase(gr("GCli_Descricao"))
								else
									Response.Write ""
								end if
							gr.movenext
							loop
                        %>
                    </td>
                    <td>Origem Solicitação</td>
                    <% 
						if len(trim(strOrigemSol)) > 0 then 
						sql = "select orisol_descricao  from cla_origemsolicitacao where orisol_id =  " & strOrigemSol
						set Ori = db.execute(sql)
						
                    %>
                    <td><% Response.Write ucase(Ori("orisol_descricao")) %></td>
                    <%
						else
                    %>
                    <td>&nbsp;</td>
                    <%
						end if 
						
                    %>
                </tr>
            </table>
        </td>

    </tr>

    <tr class="clsSilver">
        <td width="170">Razão Social</td>
        <td colspan="3">&nbsp;<%=strRazaoSocial%>
        </td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Nome Fantasia</td>
        <td colspan="3">&nbsp;<%=strNomeFantasia%></td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Conta Corrente</td>
        <td>&nbsp;<%=strContaSev%></td>
        <td align="right">Sub Conta&nbsp;</td>
        <td>&nbsp;<%=strSubContaSev%></td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Segmento</td>
        <td>&nbsp;<%=segmento%></td>
        <td align="right">Porte&nbsp;</td>
        <td>&nbsp;<%=porte%></td>
    </tr>
    <tr>
        <th colspan="4">&nbsp;•&nbsp;Informações do Serviço&nbsp;</th>
    </tr>
    <tr class="clsSilver">
        <td width="170">Id Acesso Lógico</td>
        <td colspan="3">&nbsp;<%=strIdLogico%></td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Ação</td>
        <td colspan="3">&nbsp;<%=AcaoPedido(intTipoProc)%></td>
    </tr>
    <% IF strOrigemSol <> "10" THEN %>
    <tr class="clsSilver">
        <td width="170">Order Entry</td>
        <td colspan="3">&nbsp;<%=strOrderEntry%> </td>
    </tr>
    <%END IF %>
    <% IF strOrigemSol = "10" THEN %>
    <tr class="clsSilver">
        <td width="170">SOLICITACAO - CFD</td>
        <td>&nbsp;<%=strIA%> </td>
        <td align="right">Ponta</td>
        <td>&nbsp;<%=strPONTA%> </td>
    </tr>
    <%END IF %>
    <tr class="clsSilver">
        <td width="170">Serviço</td>
        <td>&nbsp;<%=strServicoPedido%> <%IF strPabx_Virtual <> "" then %><br>
            <br>
            PABX VIRTUAL BROADSOFT: <%= strPabx_Virtual %><br>
            <%END IF%></td>
        <td align="right">Velocidade&nbsp;</td>
        <td>&nbsp;<%=strVelServico%></td>
    </tr>
    <tr class="clsSilver">
        <td width="170px">Antecipação de Acesso</td>
        <td colspan="3">&nbsp;<%if strAntAcesso = "S" then%>Sim<%else%>Não<%end if%>
        </td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Designação do Serviço</td>
        <td colspan="3">&nbsp;<%=strDesigncaoServico%></td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Nro. Contrato Serviço</td>
        <td colspan="3">&nbsp;
			<%	Select Case strTipoContratoServico
					Case 1
						Response.Write "Contrato de Serviço - "
					Case 2
						Response.Write "Contrato de Referência - "
					Case 3
						Response.Write "Carta de Compromisso - "
				End Select
            %>
            <%=strNroContrServico%>
        </td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Data Desejada de Entrega do Acesso ao Serviço</td>
        <td>&nbsp;<%=strDtDesEntrAcesServ%></td>
        <td align="right">Data Prevista de Entrega do Acesso pelo Provedor</td>
        <td>&nbsp;<%=strDtPrevEntrAcesProv%></td>
    </tr>
    <tr class="clsSilver">
        <td rowspan="2">Acesso Temporário<br>
            (dd/mm/aaaa)</td>
        <td>&nbsp;Início&nbsp;</td>
        <td>&nbsp;Fim&nbsp;</td>
        <td>&nbsp;Devolução&nbsp;</td>
    </tr>
    <tr class="clsSilver">
        <td>&nbsp;<%=strDtIniTemp%></td>
        <td>&nbsp;<%=strDtFimTemp%></td>
        <td>&nbsp;<%=strDtDevolucao%></td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Designação do Acesso Principal (678)</td>
        <td colspan="3">&nbsp;<%=strDesigAcessoPri%></td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Observações p/ Provedor</td>
        <td colspan="3">&nbsp;<%=strObs%></td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Resposta
            <br>
            Indicada no SSA</td>
        <td colspan="3" valign="top">&nbsp;</td>
    </tr>
    <tr>
        <th colspan="4">
        &nbsp;•&nbsp;Datas</td>
    </tr>
    <!--
	<tr class=clsSilver>
		<td width=170>Data Construção do Acesso Físico EBT</td>
		<td>&nbsp;<%=strDtConstrAcesFisEBT%></td>
		<td align=right width=170>Data Construção Acesso Físico Terceiro</td>
		<td>&nbsp;<%=strDtConstrAcesFisTer%></td>
	</tr>
	<tr class=clsSilver>
		<td width=170>Data Desativação do Acesso Físico EBT</td>
		<td>&nbsp;<%=strDtDesaAcesFisEBT%></td>
		<td align=right width=170>Data Desativação do Acesso Físico Terceiro</td>
		<td>&nbsp;<%=strDtDesaAcesFisTer%></td>
	</tr>
	<tr class=clsSilver>
		<td width=170>Data Cancelamento do Pedido de Acesso Físico EBT</td>
		<td>&nbsp;<%=strDtCanPedAcesFisEBT%></td>
		<td align=right width=170>Data Cancelamento de Pedido de Acesso Físico Terceiro</td>
		<td>&nbsp;<%=strDtCanPedAcesFisTer%></td>
	</tr>
	-->
    <tr class="clsSilver">
        <td width="170">Data de Entrega do Acesso ao Serviço</td>
        <td>&nbsp;<%=strDtEntrAcesServ%></td>
        <td align="right" width="170">Data da Ativação do Serviço</td>
        <td>&nbsp;<%=strDtAtivServ%></td>
    </tr>
    <tr class="clsSilver">
        <td width="170">Data de Conclusão da Solicitação</td>
        <td>&nbsp;<%=strDtConcluSolic%></td>
        <td align="right" width="170">Data de Cancelamento da Solicitação</td>
        <td>&nbsp;<%=strDtCanSolic%></td>
    </tr>


    <tr class="clsSilver">
        <td width="170">Data Inicio de Avaliação</td>
        <td>&nbsp;<%=StrDtInicioAvaliacao%></td>
        <td align="right" width="170">Data Fim de Avaliação</td>
        <td>&nbsp;<%=StrDtFimAvaliacao%></td>
    </tr>


    <table border="0" cellspacing="1" cellpadding="0" width="760">
        <tr>
            <th colspan="4">&nbsp;•&nbsp;Análise do estudo de viabilidade (SEV   SSA) &nbsp;
            </th>
        </tr>
        <tr>
            <td>


                <%
				if strNroSev <> "" then
					
						
						'Localiza informações da solução SSA
						Vetor_Campos(1)="adInteger,2,adParamInput," & strNroSev ' dblNroSev
						Vetor_Campos(2)="adInteger,2,adParamOutput,0"
						'Call APENDA_PARAM("CLA_sp_sel_solucao_ssa",2,Vetor_Campos)
						Call APENDA_PARAM("CLA_sp_sel_historico_solucao_ssa",2,Vetor_Campos)
						
						
						Set objRSCli = ObjCmd.Execute
						if DBAction = 0 then
							If Not objRSCli.eof and  Not objRSCli.bof then
                %>
                <table cellspacing="1" cellpadding="0" border="0" width="760">
                    <tr class="clsSilver">
                        <td>SEV:  	 <b><%=strNroSev %>   </b></td>
                        <td></td>
                    </tr>
                    <tr class="clsSilver">

                        <td>Status do Pedido:  	 <b><%=objRSCli("STATUS_PEDIDO")%>   </b></td>
                        <td>Status da SEV:  <b><%=objRSCli("STATUS_SEV")%>   </b></td>
                    </tr>
                    <tr class="clsSilver">
                        <td>Nº do pedido de Venda do SNOA:  	 <b><%= TratarAspasJS(Trim(objRSCli("Pedido_Venda_SNOA"))) %>  </b></td>
                        <td>Caracteristíca do EILD :  <b><%= TratarAspasJS(Trim(objRSCli("C_TIPO_EILD"))) %>  </b></td>
                    </tr>

                    <tr class="clsSilver">
                        <td>Ponta:  	 <b><%= TratarAspasJS(Trim(objRSCli("Ponta"))) %>  </b></td>
                        <td>Nº SEV Ponta A :  <b><%= TratarAspasJS(Trim(objRSCli("sevPontaA_seq"))) %>  </b></td>
                    </tr>

                    <tr class="clsSilver">
                        <td>Operadora Solicitante:  	 <b><%= TratarAspasJS(Trim(objRSCli("OPR_DES"))) %>  </b></td>
                        <td>Pabx Virtual:  	 <b><%= TratarAspasJS(Trim(objRSCli("pabx_des"))) %>  </b>| Migração do Pabx Virtual:  	 <b><%= TratarAspasJS(Trim(objRSCli("SEV_MIGRA_PABX"))) %>  </td>
                    </tr>

                    <table cellpadding="0" cellspacing="1" width="760">
                        <tr>
                            <th nowrap>&nbsp;Facilidade</th>
                            <th nowrap>&nbsp;Tecnologia</th>
                            <th>&nbsp;Provedor</th>
                            <!--<th>&nbsp;Opc Indicada</th>
												<th>&nbsp;Atendimento</th> 
												<th>&nbsp;EILD</th> 
												<th>&nbsp;SIP</th>
												-->
                            <th>&nbsp;Estação</th>
                            <th>&nbsp;Usuário</th>
                            <th>&nbsp;Data vencimento</th>
                            <!--<th>&nbsp;Análise</th> -->
                            <th>&nbsp;OBS</th>
                        </tr>
                        <%
									 While Not objRSCli.eof
                        %>
                        <tr class="clsSilver">
                            <td nowrap>&nbsp; <%=objRSCli("FAC_DES")%>  </td>
                            <td nowrap>&nbsp; <%=objRSCli("TEC_DES")%>  </td>
                            <td>&nbsp;  <%=objRSCli("FOR_DES")%> </td>
                            <!--<td >&nbsp;<%= TratarAspasJS(Trim(objRSCli("SOL_SELECIONADA"))) %></td> 
													<td >&nbsp;  <%=objRSCli("ATE_DESC")%> </td> 
													<td >&nbsp;  <%=objRSCli("TIPO_EILD")%> </td> 
													<td >&nbsp;  <%=objRSCli("SOL_SINALIZACAO_SIP")%> </td>-->
                            <td>&nbsp;  <%=objRSCli("ESTACAO")%> </td>
                            <td>&nbsp;  <%=objRSCli("SEV_USRNAMESOLUC")%>  </td>
                            <td>&nbsp; <%=objRSCli("data_vencimento")%> </td>
                            <!--<td >&nbsp;  <%=objRSCli("ANALISE")%></td> -->
                            <td>&nbsp;  <%=objRSCli("SOL_OBS")%>  </td>
                        </tr>
                        <%
									objRSCli.MoveNext
									Wend
							End If
						 End If
					
				End if
                        %>
                    </table>

                </table>


            </td>
        </tr>


    </table>


    <table border="0" cellspacing="1" cellpadding="0" width="760">
        <tr>
            <th colspan="4">&nbsp;•&nbsp;Acessos Físicos</th>
        </tr>
        <%
Dim objDicAcf
Set objDicAcf = Server.CreateObject("Scripting.Dictionary")

intIndex = 1

Vetor_Campos(1)="adInteger,4,adParamInput,"
Vetor_Campos(2)="adInteger,4,adParamInput,"
Vetor_Campos(3)="adInteger,4,adParamInput," & strIdLogico
Vetor_Campos(4)="adInteger,4,adParamInput," & dblSolId
Vetor_Campos(5)="adInteger,4,adParamInput,A"

'strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",5,Vetor_Campos)
strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_view_ponto2",5,Vetor_Campos)

Set objRSFis = db.Execute(strSqlRet)
intAecLog = ""

Dim ArrAcfIsFis(2,2)
intIndice = 0
if Not objRSFis.EOF and not objRSFis.BOF then
	While not objRSFis.EOF
		'if objRSFis("Alf_FlagAux") <> "R" or isnull(objRSFis("Alf_FlagAux")) then
			
			ArrAcfIsFis(intIndice,0) = objRSFis("Acf_ID")
			ArrAcfIsFis(intIndice,1) = objRSFis("Acf_IDAcessoFisico")
			
			intIndice = intIndice + 1

			strIDAcessoFisico	= objRSFis("Acf_IDAcessoFisico")
			
			if objRSFis("DesignacaoContrato") <> "" then
				strNroAcesso		= objRSFis("DesignacaoContrato")
			else
			strNroAcesso		= objRSFis("Acf_NroAcessoPtaEbt")
			end if 
			
			'strNroAcesso		= objRSFis("Acf_NroAcessoPtaEbt")
			dblAcfId			= objRSFis("Acf_Id")

			'if intAecLog <> objRSFis("Aec_ID") then
				intIndexIdFis = 1
				intAecLog = objRSFis("Aec_ID")
				intTec = objRSFis("Tec_ID")
				strPropAcessoFis = objRSFis("Acf_Proprietario")
				if objRSFis("Alf_TipoCCTo") = "I" then
					strTipoPonto    = "INSTALAÇÃO"
				Else
					strTipoPonto    = "INTERMEDIÁRIO"
				End if
				dblProId			= objRSFis("Pro_Id")
				strTecnologia = ""
				if objRSFis("Tec_Id") <> "" then
					Set objRSAux = db.Execute("CLA_Sp_Sel_Tecnologia " & objRSFis("Tec_Id"))
					if not objRSAux.Eof and Not objRSAux.Bof then
						strTecnologia	= objRSAux("Tec_Sigla")
					End if
				End if
				if objRSFis("Vel_IDAcessoFis") <> "" then
					Set objRSAux = db.Execute("CLA_sp_sel_velocidade " & objRSFis("Vel_IDAcessoFis"))
					if not objRSAux.Eof and Not objRSAux.Bof then
						strVelAcessoFis = objRSAux("Vel_Desc")
					End if
				End if

				dblRegId			= objRSFis("Reg_ID")
				'strPrmDesc			= objRSFis("Prm_Desc")
				strNroPI			= objRSFis("Acf_NroPI")
				strInterFaceEndInstala = objRSFis("Acf_Interface")
				strInterfaceEbt		   = Trim(objRSFis("Acf_InterfaceEstEntregaFisico"))
				intTipoVel			= objRSFis("Acf_TipoVel")
				strUFEndInstala		= objRSFis("Est_Desc")
				strEndCidInstala	= objRSFis("Cid_Desc")
				strLogrEndInstala	= objRSFis("Tpl_Sigla")
				strEndInstala		= objRSFis("End_NomeLogr")
				strNroEndInstala	= objRSFis("End_NroLogr")
				strComplEndInstala	= objRSFis("Aec_Complemento")
				strBairroEndInstala	= objRSFis("End_Bairro")
				strCepEndInstala	= objRSFis("End_CEP")
				strContatoEndInstala= objRSFis("Aec_Contato")
				strTelEndInstala	= objRSFis("Aec_Telefone")
				strCNPJInstala		= objRSFis("Aec_CNPJ")
				strSiglaCentroCli	= objRSFis("Aec_SiglaCentroCliente")
				strIEInstala		= objRSFis("Aec_IE")
				strIMInstala		= objRSFis("Aec_IM")
				strPropEndInstala	= objRSFis("Aec_PropEnd")
				strCodSAP			= objRSFis("Aec_CodSap")
				
				''@@Davif Inclusão Campos Acomp. Provedor
				strDtPrevEntregaProvedor = Formatar_Data(Trim(objRSFis("Acf_DtPrevistaEntregaAcesso")))
				strDtReaprPrevEntregaProvedor = Formatar_Data(Trim(objRSFis("Acf_DtReaprPrevistaEntregaAcesso")))
				StrProjetoEspecial = Trim(objRSFis("Acf_ProjetoEspecial"))
				StrProjEsp_NumeroProposta = Trim(objRSFis("Acf_ProjEsp_NumeroProposta"))
				StrProjEsp_Prazo = objRSFis("Acf_ProjEsp_Prazo")
				'MSCAPRI - CH-97014NBU
				strDtAceite		  =	Formatar_Data(Trim(objRSFis("Acf_DtAceite")))
				strDtConstrucao	  =	Formatar_Data(Trim(objRSFis("Acf_DtConstrAcessoFis")))
				strDtDesativacao  =	Formatar_Data(Trim(objRSFis("Acf_DtDesatAcessoFis")))
				strDtCancelamento = Formatar_Data(Trim(objRSFis("Acf_DtCancAcessoFis")))

				STSSIGLAESTENTREGAFISICO      = objRSFis("ACF_SIGLAESTENTREGAFISICO")
				STSCOMPLSIGLAESTENTREGAFISICO = objRSFis("ACF_COMPLSIGLAESTENTREGAFISICO")
				

				
        %>
        <tr class="clsSilver2">
            <td colspan="4">&nbsp;•&nbsp;Acesso <%=intIndex%>&nbsp;<%'=strDM%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Prop do Acesso Físico</td>
            <td width="170">&nbsp;<%=strPropAcessoFis%></td>
            <td width="170">Velocidade</td>
            <td width="170">&nbsp;<%=strVelAcessoFis%>
                <%
					if intTipoVel = 1 then Response.Write " - ESTRUTURADA"
					if intTipoVel = 0 then Response.Write " - NÃO ESTRUTURADA"
                %>
            </td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Tipo do Ponto</td>
            <td colspan="3">&nbsp;<%=strTipoPonto%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Proprietário do Modem</td>
            <td>&nbsp;<%=strProModem%></td>
            <td>Quantidade de Modens</td>
            <td>&nbsp;<%=intQtdeEquip%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Provedor</td>
            <td colspan="3">&nbsp;<%set objRS = db.execute("CLA_sp_sel_provedor 0")
					While not objRS.Eof
						if Trim(dblProId) = Trim(objRS("Pro_ID")) then
							Response.Write UCase(objRS("Pro_Nome"))
						End if
						objRS.MoveNext
					Wend
            %>
            </td>





        </tr>
        <tr class="clsSilver">

            <td width="170">Facilidade</td>
            <td>&nbsp;<%set objRS = db.execute("CLA_sp_sel_newconsultaTecnologiaFacilidade " & dblSolId & ", null, " & dblAcfId  )
					While not objRS.Eof
						
						Response.Write UCase(objRS("newfac_Nome"))
						
						objRS.MoveNext
					Wend
            %>
			
					
					
            </td>
            <td width="170">Tecnologia</td>
            <td>&nbsp;<%set objRS = db.execute("CLA_sp_sel_newconsultaTecnologiaFacilidade " & dblSolId  & ", null, " & dblAcfId  )
					While not objRS.Eof
						
						Response.Write UCase(objRS("newtec_Nome"))
						
						objRS.MoveNext
					Wend
            %>
            </td>



        </tr>
        <tr class="clsSilver">
            <td width="170">Prazos de Contratação de Acesso</td>
            <td>&nbsp;<%
					if Trim(dblProId) <> "" then
						set objRS = db.execute("CLA_sp_sel_regimecontrato 0," & dblProId)
						While not objRS.Eof
							if Trim(dblRegId) = Trim(objRS("Reg_ID")) then
								Response.Write LimparStr(Trim(objRS("Pro_Nome"))) & " - " & LimparStr(Trim(objRS("Tct_Desc")))
							End if
							objRS.MoveNext
						Wend
					End if
            %>
            </td>
            <td>Promoção</td>
            <td>&nbsp;<%=strPrmDesc%>
            </td>
        </tr>
        <tr class="clsSilver">
            <td width="170">UF</td>
            <td>&nbsp;<%=strUFEndInstala%></td>
            <td>Cidade</td>
            <td>&nbsp;<%=strEndCidInstala%></td>
        </tr>
        <tr class="clsSilver">
            <td>Tipo Logr</td>
            <td>&nbsp;<%=strLogrEndInstala%></td>
            <td>Nome Logr</td>
            <td>&nbsp;<%=strEndInstala%></td>
        </tr>
        <tr class="clsSilver">
            <td>Número</td>
            <td>&nbsp;<%=strNroEndInstala%></td>
            <td>Complemento</td>
            <td>&nbsp;<%=strComplEndInstala%></td>
        </tr>
        <tr class="clsSilver">
            <td>Bairro</td>
            <td>&nbsp;<%=strBairroEndInstala%></td>
            <td>CEP</td>
            <td>&nbsp;<%=strCepEndInstala%></td>
        </tr>
        <tr class="clsSilver">
            <td>Contato</td>
            <td>&nbsp;<%=strContatoEndInstala%></td>
            <td>Telefone</td>
            <td>&nbsp;<%=strTelEndInstala%></td>
        </tr>
        <tr class="clsSilver">
            <td>CNPJ</td>
            <td>&nbsp;<%=strCNPJInstala%></td>
            <td>Sigla do Centro do Cliente</td>
            <td>&nbsp;<%=strSiglaCentroCli%></td>
        </tr>
        <tr class="clsSilver">
            <td>I.E.</td>
            <td>&nbsp;<%=strIEInstala%></td>
            <td>I.M</td>
            <td>&nbsp;<%=strIMInstala%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Proprietário do Endereço de Instalação</td>
            <td colspan="3">&nbsp;<%=strPropEndInstala%></td>
        </tr>
        <tr class="clsSilver">
            <td>Código SAP</td>
            <td>&nbsp;<%=strCodSAP%></td>
            <td>Número PI</td>
            <td>&nbsp;<%=strNroPI%></td>
        </tr>

        <tr class="clsSilver">
            <td>Data prevista de entrega pelo provedor</td>
            <td>&nbsp;<%=strDtPrevEntregaProvedor%></td>
            <td>Data Reaprazada prevista de entrega pelo provedor</td>
            <td>&nbsp;<%=strDtReaprPrevEntregaProvedor%></td>
        </tr>

        <tr class="clsSilver">
            <td>Projeto Especial</td>
            <td>&nbsp;<%=strProjetoEspecial%></td>
            <td>Número da Proposta</td>
            <td>&nbsp;<%=StrProjEsp_NumeroProposta%></td>
        </tr>

        <tr class="clsSilver">
            <td>Prazo  </td>
            <td colspan="3">&nbsp;<%=StrProjEsp_Prazo%> &nbsp;(Dias) </td>
        </tr>

        <tr class="clsSilver">
            <td><font color="#FF0000">Interface Est. Cliente</font></td>
            <td colspan="3">&nbsp;<%=strInterFaceEndInstala%></td>
        </tr>

        <tr class="clsSilver">
            <td><font color="red">Interface Est. EBT</font></td>
            <td colspan="3">&nbsp;<%= strInterfaceEbt%></td>
        </tr>
        <tr class="clsSilver">
            <td>Sigla Estação Entrega do Acesso Físico</td>
            <td colspan="3">&nbsp;<%= STSSIGLAESTENTREGAFISICO%> <%= STSCOMPLSIGLAESTENTREGAFISICO %></td>
        </tr>

        <%
			
			'End if 'Separa por Aec_ID (Divisão de ponto)
			if not objDicAcf.Exists(Cstr(dblAcfId)) then

				Call objDicAcf.Add(Cstr(dblAcfId),Cstr(dblAcfId))

				Set objRSPed =	db.execute("CLA_sp_view_pedido " & dblSolId & ",null,null,null,null,null,null,null,null,'T',"& dblAcfId)
				''Response.Write ( "db.execute(CLA_sp_view_pedido " & dblSolId & ",null,null,null,null,null,null,null,null,'T',"& dblAcfId )

				if not objRSPed.Eof and not objRSPed.Bof then
					While not objRSPed.Eof

						'if Trim(objRSPed("Tprc_id")) = Trim(intTipoProc) then
							strDM	= Ucase(Trim(objRSPed("Ped_Prefixo"))) & "-" & Trim(right("00000" & objRSPed("Ped_Numero"),5)) & "/" & Trim(objRSPed("Ped_Ano"))
							'response.write "<script>alert('"&objRSPed("Ped_ID")&"')</script>"
							dblPedId = objRSPed("Ped_ID")
						'End if
						objRSPed.MoveNext
					Wend
				Else
					if not isNull(intTec) and intTec <> 3 then
						'Localiza codigo do pedido no CRMSF
						Set objRSCRM = db.execute("CLA_SP_SEL_CRMSPROCESSO " & strIdLogico & "," & dblSolId)
						if not objRSCRM.Eof and not objRSCRM.Bof then
							strDM = objRSCRM("CodInstalar")
						End If
					End if
				End if
        %>
        <tr class="clsSilver">
            <!--<td rowspan=2>ID Acesso Físico (<%=intIndexIdFis%>)</td>-->
            <td rowspan="2">ID Acesso Físico</td>
            <td class="clsSilver2">&nbsp;Pedido&nbsp;</td>
            <td class="clsSilver2">&nbsp;Id Físico&nbsp;</td>
            <td class="clsSilver2">&nbsp;Nº do Acesso&nbsp;</td>
        </tr>
        <tr class="clsSilver">
            <td>&nbsp;<%=strDM%></td>
            <td>&nbsp;<%=strIDAcessoFisico%></td>
            <td>&nbsp;<%=strNroAcesso%></td>
        </tr>
        <!--MSCAPRI - CH-97014NBU -->
        <tr class="clsSilver">
            <td class="clsSilver2" colspan="4">&nbsp;•&nbsp;Datas do Acesso Físico - Acesso <%=intIndex%>
            </th>
        </tr>
        <tr class="clsSilver">
            <td width="170">Data de Aceite</td>
            <td>&nbsp;<%=strDtAceite%></td>
            <td width="170">Data de Construção</td>
            <td>&nbsp;<%=strDtConstrucao%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Data de Desativação</td>
            <td>&nbsp;<%=strDtDesativacao%></td>
            <td width="170">Data de Cancelamento</td>
            <td>&nbsp;<%=strDtCancelamento%></td>
        </tr>

        <%
                                intIndex = intIndex + 1
				intIndexIdFis = intIndexIdFis + 1
			End if
		'End if
		 objRSFis.MoveNext



	Wend
End if
        %>
    </table>
    <table border="0" cellspacing="1" cellpadding="0" width="760">

        <tr>
            <th colspan="4">&nbsp;•&nbsp;Informações da Embratel</th>
        </tr>
        <tr class="clsSilver">
            <td width="170">Local de Entrega</td>
            <td colspan="3">&nbsp;<%=strLocalInstala%></td>
        </tr>

        <tr class="clsSilver">
            <td width="170">Local de Configuração</td>
            <td colspan="3">&nbsp;<%=strLocalConfig%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Contato</td>
            <td>&nbsp;<%=strContEndLocalInstala%></td>
            <td>Telefone</td>
            <td>&nbsp;<%=strTelEndLocalInstala%></td>
        </tr>

        <tr>
            <th colspan="4">
            &nbsp;•&nbsp;Coordenação Embratel
		</td>
        </tr>
        <!--
	<tr class=clsSilver>
		<td width=170>Órgão de Venda</td>
		<td colspan="3">&nbsp;<%
		 Set objRS = db.Execute("CLA_sp_sel_Orgao " & dblOrgId)
		 if not objRS.Eof and not objRS.Bof then
			Response.Write objRS("Org_Nome")
		 End if
		%></td>
	</tr>
-->
        <tr class="clsSilver">
            <td width="170">UserName GIC-N</td>
            <td colspan="3">&nbsp;<%=strGICN%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Nome GIC-N</td>
            <td>&nbsp;<%=strNomeGICN%></td>
            <td>Ramal</td>
            <td>&nbsp;<%=strRamalGICN%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">UserName GIC-L</td>
            <td colspan="3">&nbsp;<%=strGICL%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Nome GIC-L</td>
            <td>&nbsp;<%=strNomeGICL%></td>
            <td>Ramal</td>
            <td>&nbsp;<%=strRamalGICL%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">UserName GLA</td>
            <td colspan="3">&nbsp;<%=strGLA%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Nome GLA</td>
            <td>&nbsp;<%=strNomeGLA%></td>
            <td>Ramal</td>
            <td>&nbsp;<%=strRamalGLA%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">UserName GLA-E</td>
            <td colspan="3">&nbsp;<%=strGLAE%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Nome GLA-E</td>
            <td>&nbsp;<%=strNomeGLAE%></td>
            <td>Ramal</td>
            <td>&nbsp;<%=strRamalGLAE%></td>
        </tr>


        <tr class="clsSilver">
            <td width="170">UserName AVL</td>
            <td colspan="3">&nbsp;<%=strAVL%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Nome AVL</td>
            <td>&nbsp;<%=strNomeAVL%></td>
            <td>Ramal</td>
            <td>&nbsp;<%=strRamalAVL%></td>
        </tr>


        <tr>
            <th colspan="4">&nbsp;•&nbsp;Status
            </th>
        </tr>
        <tr class="clsSilver">
            <td width="170">Status</td>
            <td colspan="3">&nbsp;<%=strStatus%></td>
        </tr>
        <tr class="clsSilver">
            <td width="170">Histórico</td>
            <td colspan="3">&nbsp;<%=strHistorico%></td>
        </tr>
    </table>
    <%=strRespostaSSA%>
    <table cellspacing="1" cellpadding="1" width="760" border="0">
        <tr>
            <th colspan="4">&nbsp;•&nbsp;Histórico</th>
        </tr>
        <%
		Vetor_Campos(1)="adInteger,2,adParamInput,"
		Vetor_Campos(2)="adInteger,2,adParamInput," & dblSolId
		Vetor_Campos(3)="adInteger,2,adParamInput,1"
		Vetor_Campos(4)="adInteger,2,adParamInput,"

		strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_StatusSolicitacao",4,Vetor_Campos)
		Set objRS = db.Execute(strSqlRet)

		blnCor = true
		strHtml = strHtml &  ""
		While Not objRS.Eof
			if blnCor then
				strHtml = strHtml &  "<tr class=clsSilver >"
				blnCor = false
			Else
				strHtml = strHtml &  "<tr class=clsSilver2>"
				blnCor = true
			End if
			strHtml = strHtml &  "<td width=15% nowrap >"& Formatar_Data(objRS("StsSol_Data")) &"</td>"
			strHtml = strHtml &  "<td >&nbsp;" & objRS("stssol_cont") & "</td>"
			strHtml = strHtml &  "<td >&nbsp;" & objRS("stssol_freq") & "</td>"
			strHtml = strHtml &  "<td nowrap >&nbsp;" & objRS("Pedido") & "</td>"
			strHtml = strHtml &  "<td nowrap >&nbsp;" & objRS("Tecnologia") & "</td>"
			strHtml = strHtml &  "<td width=20% >" & objRS("Usu_UserName") & "</td>"
			strHtml = strHtml &  "<td width=30% >"& objRS("Sts_Desc") & "</td>"
			strHtml = strHtml &  "<td width=35% >"& objRS("StsSol_Historico") &"</td>"
			strHtml = strHtml &  "</tr>"
			objRS.MoveNext
		Wend
		Response.Write strHtml%>
    </table>












    <%

Dim intIndex
Dim strSql
Dim intCount
Dim strClass

			Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
			'Vetor_Campos(1)="adInteger,2,adParamInput," & dblAcfId

			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_newConsultaFacilidadeServico",1,Vetor_Campos)
			'Response.Write strSqlRet
			'Response.end
			Set objRS = db.Execute(strSqlRet)
			objRS.Close
			objRS.CursorLocation = adUseClient
			objRS.Open
intCount=1



	if objRS.Bof then
    %>
    <table border="0" cellspacing="1" cellpadding="0" width="760">
        <tr>
            <th colspan="4">&nbsp;Facilidades do Serviço - <%=strIDAcessoFisico %></th>
        </tr>
    </table>
    <table border="0" cellspacing="1" cellpadding="0" width="760">
        <tr>
            <th colspan="4">&nbsp;•&nbsp;Facilidades</th>
        </tr>
    </table>
    <table cellspacing="1" cellpadding="0" width="760" border="0">
        <tr class="clsSilver">
            <td>
                <%
		Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
		Vetor_Campos(2)="adInteger,2,adParamInput," & dblAcfId
		strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_newAlocacaoAcesso",2,Vetor_Campos)
		'Response.Write strSqlRet
		Set objRSFacilidades = db.Execute(strSqlRet)
		'objRSFacilidades.CursorLocation = adUseClient
		'objRSFacilidades.Open
intCount=1
if not objRSFacilidades.Eof and not objRSFacilidades.Bof then
	'For intIndex = 1 to objRS.PageSize
	While Not objRSFacilidades.Eof
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
                %>
                <tr class="<%=strClass%>">
                    <td width="170">&nbsp;<%=TratarAspasHtml(objRSFacilidades("label"))%></td>

                    <td colspan="3"><%=trim(objRSFacilidades("conteudo"))%></td>
                </tr>
                <%
		intCount = intCount+1
		objRSFacilidades.MoveNext
		'if objRS.EOF then Exit For
	'Next
	Wend
                %>
            </td>
        </tr>
    </table>
    <%
End if

	end if

	While Not objRS.Eof

		
		dblAcfId = objRS("acf_id")
		
		if Trim(objRS("orisol_id"))= "10" then 

			'strVlan =  "VLAN CFD"
			'strSVlan = "SVLAN CFD"
			'strPorta = "Porta CFD"
			'strPe =    "PE CFD"

			strVlan =  "VLAN "
			strSVlan = "SVLAN "
			strPorta = "Porta "
			strPe =    "Eqpto Agregador "


		else

			strVlan =  "VLAN"
			strSVlan = "SVLAN"
			strPorta = "Porta"
			strPe =    "PE"


		end if
	
	For t = 0 to 1
		
		if ArrAcfIsFis(t,0) = dblAcfId then strIDFisico = ArrAcfIsFis(t,1)
	
	Next

	'For intIndex = 1 to objRS.PageSize
	'	if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
if not objRS.Eof and not objRS.Bof then
    %>
    <table border="0" cellspacing="1" cellpadding="0" width="760">
        <tr>
            <th colspan="4">&nbsp;Facilidades do Serviço - <%=strIDFisico %></th>
        </tr>
    </table>

    <table cellspacing="1" cellpadding="0" width="760" border="0">
        <tr class="clsSilver">
            <td>
                <tr class="clsSilver">
                    <td width="170"><%=strPE %></td>

                    <td colspan="3"><%=trim(objRS("newfacServico_PE"))%></td>
                </tr>
                <tr class="clsSilver2">
                    <td width="170"><%=strPorta %></td>

                    <td colspan="3"><%=trim(objRS("newfacServico_Porta"))%></td>
                </tr>
                <tr class="clsSilver">
                    <td width="170"><%=strVlan %></td>

                    <td colspan="3"><%=trim(objRS("newfacServico_Vlan"))%></td>
                </tr>
                <tr class="clsSilver2">
                    <td width="170"><%=strSVlan %></td>

                    <td colspan="3"><%=trim(objRS("newfacServico_SVlan"))%></td>
                </tr>



                <%	
end if 		

                %>

            </td>
        </tr>
    </table>


    <table border="0" cellspacing="1" cellpadding="0" width="760">
        <tr>
            <th colspan="4">&nbsp;•&nbsp;Facilidades</th>
        </tr>
    </table>

    <table cellspacing="1" cellpadding="0" width="760" border="0">
        <tr class="clsSilver">
            <td>
                <%
'Dim intIndex
'Dim strSql
'Dim intCount
'Dim strClass

'strSql = "CLA_sp_sel_AssocTecnologiaFacilidade"

'Call PaginarRS(0,strSql)
			
			Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
			Vetor_Campos(2)="adInteger,2,adParamInput," & dblAcfId
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_newAlocacaoAcesso",2,Vetor_Campos)
			'Response.Write strSqlRet
			Set objRSFacilidades = db.Execute(strSqlRet)
			'objRSFacilidades.CursorLocation = adUseClient
			'objRSFacilidades.Open
	intCount=1
	if not objRSFacilidades.Eof and not objRSFacilidades.Bof then
		'For intIndex = 1 to objRS.PageSize
		While Not objRSFacilidades.Eof
			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
                %>
                <tr class="<%=strClass%>">
                    <td width="170">&nbsp;<%=TratarAspasHtml(objRSFacilidades("label"))%></td>

                    <td colspan="3"><%=trim(objRSFacilidades("conteudo"))%></td>
                </tr>
                <%
			intCount = intCount+1
			objRSFacilidades.MoveNext
			'if objRS.EOF then Exit For
		'Next
		Wend
	End if
	objRS.MoveNext
Wend
                %>
            </td>
        </tr>
    </table>


    <!--FIM CAMPOS PROCESSO -->
