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
'response.write "<script>alert('"&Trim(Request.Form("hdnSolId"))&"')</script>"

if Trim(Request.Form("hdnSolId")) <> "" then

	set ObjRsetOrisol = db.execute("select orisol_id , Acl_IDAcessoLogico from cla_solicitacao where sol_id = " & Request.Form("hdnSolId"))
	if not ObjRsetOrisol.eof then
		orisolid = ObjRsetOrisol("orisol_id")
		strIdLogico		= Trim(ObjRsetOrisol("Acl_IDAcessoLogico"))
	end if
	
	'GPON
	set ObjRsetTecID = db.execute("select top 1 cla_acessofisico.Pro_ID,cla_tecnologia.Tec_Nome,cla_acessofisico.Acf_ID, cla_tecnologia.Tec_ID from cla_solicitacao inner join cla_acessologico 		on cla_solicitacao.Acl_IDAcessoLogico = cla_acessologico.Acl_IDAcessoLogico inner join cla_acessologicofisico 	on cla_acessologico.Acl_IDAcessoLogico = cla_acessologicofisico.Acl_IDAcessoLogico inner join cla_acessofisico 		on cla_acessologicofisico.acf_id = cla_acessofisico.acf_id left join cla_tecnologia 		on cla_acessofisico.tec_id = cla_tecnologia.tec_id where ( ( cla_acessologicofisico.alf_flagAux is null  ) or (cla_acessologicofisico.alf_flagAux is not null and pro_id = 145) ) and Acf_DtCancAcessoFis is null and Acf_DtDesatAcessoFis is null and sol_id = " & Request.Form("hdnSolId"))
	if not ObjRsetTecID.eof then
		strTecnologia = ObjRsetTecID("Tec_Nome")
		strTecId = ObjRsetTecID("Tec_ID")
		DblAcf_ID = ObjRsetTecID("Acf_ID")
		strProId = ObjRsetTecID("Pro_ID")
	end if
	
	
	
		'GPON
	If request("hdnAlocar")="Alocar" Then
		Vetor_Campos(1)="adInteger,8,adParamInput," & DblAcf_ID
		Vetor_Campos(2)="adInteger,8,adParamInput,"& ucase(Trim(request("cboTipoONT")))
		Vetor_Campos(3)="adInteger,2,adParamOutput,0"
	
		Call APENDA_PARAM("CLA_sp_upd_PortaONT",3,Vetor_Campos)
		ObjCmd.Execute'pega dbaction
		DBAction = ObjCmd.Parameters("RET").value
	End if

	dblSolId = Request.Form("hdnSolId")
	dblPedId = Request.Form("hdnPedId")
	
	strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

	'<!-- Projeto ABW - Incio -->'
	set rsProjABW = db.execute (" select top 1 aprovisi_id, indicador_migracao_pabx_virtual , indicador_tecnologia_ip " &_
								" from cla_aprovisionador " &_
								" where sol_id = " & dblSolId & " order by 1 desc")
			
	if Not rsProjABW.Eof and Not rsProjABW.Bof then

		strMigraPabx_Virtual = Trim(rsProjABW("indicador_migracao_pabx_virtual"))

		IF ISNULL(strMigraPabx_Virtual) THEN
			strMigraPabx_Virtual = ""
		END IF
	
		IF strMigraPabx_Virtual = "NULL" THEN
			strMigraPabx_Virtual = ""
		END IF

		IF strMigraPabx_Virtual = "N" THEN
			strMigraPabx_Virtual = "Não"
		END IF

		IF strMigraPabx_Virtual = "S" THEN
			strMigraPabx_Virtual = "Sim"
		END IF

		strTecnologia_IP = Trim(rsProjABW("indicador_tecnologia_ip"))

		IF ISNULL(strTecnologia_IP) THEN
			strTecnologia_IP = ""
		END IF
		
		IF strTecnologia_IP = "NULL" THEN
			strTecnologia_IP = ""
		END IF

		IF strTecnologia_IP = "N" THEN
			strTecnologia_IP = "Não"
		END IF

		IF strTecnologia_IP = "S" THEN
			strTecnologia_IP = "Sim"
		END IF

	end if
	'<!-- Projeto ABW - Fim -->'


	'If trim(strLoginRede) = "EDAR"  Then
	
		'response.write "<script>alert('"&dblSolId&"')</script>"
		'response.write "<script>alert('"&dblPedId&"')</script>"
		'	response.write "<script>alert('"&strTecnologia&"')</script>"
	
	'End IF
	if dblPedId = "" and TRIM(strTecnologia) <> "GPON" then 'jcartus
		Response.Write "<script language=javascript>alert('Pedido indisponível.');window.location.replace('facilidade_main.asp');</script>"
		Response.End
	End if
	

	'set ObjFO = db.execute("select top 1 cla_acessofisico.Pro_ID from cla_solicitacao WITH (NOLOCK) inner join cla_acessologico 		WITH (NOLOCK) on cla_solicitacao.Acl_IDAcessoLogico = cla_acessologico.Acl_IDAcessoLogico inner join cla_acessologicofisico 	WITH (NOLOCK) on cla_acessologico.Acl_IDAcessoLogico = cla_acessologicofisico.Acl_IDAcessoLogico inner join cla_acessofisico 		WITH (NOLOCK) on cla_acessologicofisico.acf_id = cla_acessofisico.acf_id where cla_acessologicofisico.alf_flagAux is null and cla_acessofisico.pro_id=145 and sol_id = " & Request.Form("hdnSolId"))
	set ObjFO = db.execute("select top 1 cla_acessofisico.Pro_ID from cla_solicitacao WITH (NOLOCK) inner join cla_acessologico 		WITH (NOLOCK) on cla_solicitacao.Acl_IDAcessoLogico = cla_acessologico.Acl_IDAcessoLogico inner join cla_acessologicofisico 	WITH (NOLOCK) on cla_acessologico.Acl_IDAcessoLogico = cla_acessologicofisico.Acl_IDAcessoLogico inner join cla_acessofisico 		WITH (NOLOCK) on cla_acessologicofisico.acf_id = cla_acessofisico.acf_id where cla_acessofisico.pro_id=145 and sol_id = " & Request.Form("hdnSolId"))
	pro_id_fo = ""
	if not ObjFO.Eof and not ObjFO.Bof then
		pro_id_fo = ObjFO("Pro_ID")
	end if

If trim(strLoginRede) = "JCARTUS"  Then
	'response.write "<script>alert('"&strTecnologia&"')</script>"
end if

	if TRIM(strTecnologia) = "GPON" OR pro_id_fo = "145" then  
	  Set objRS = db.execute("CLA_sp_view_pedido_gpon " & dblSolId) 		
	else
		Set objRS = db.execute("CLA_sp_view_pedido_gpon null,null,null,null,null,null," & dblPedId)
  end if			
 
	
	if objRS.Eof and objRS.Bof then 'Não existe registros
		Response.Write "<script language=javascript>alert('Pedido indisponível.');window.location.replace('facilidade_main.asp');</script>"
		Response.End 
	End if
	strProId = Request.Form("cboProvedor")
	if strProId = "" then  strProId	= Trim(objRS("Pro_id"))

	set ObjV = db.execute("select cvlan , svlan from cla_assocprovedortecnologia where pro_id = " & strProId )
	disabledCvlan = ""
	disabledSvlan = ""
	if not ObjV.Eof and not ObjV.Bof then
		if ObjV("cvlan") = false then
			disabledCvlan = true
		else
			disabledCvlan = false
		end if 
		
		if ObjV("svlan")  = false then
			disabledSvlan = true
		else
			disabledSvlan = false
		end if 		
	end if
		
	strPabx_Virtual     = Trim(objRS("PABX_VIRTUAL"))	
	

	strExigePortadora = Trim(objRS("Pro_ExigePortadora"))
	strNroSolic		= dblSolId 'Trim(objRS("Sol_id"))
	strPropAcesso	= Trim(objRS("Acf_Proprietario"))
	
	'Set objRSol = db.execute("select Acl_IDAcessoLogico from cla_solicitacao where sol_id = " & strNroSolic)
	'if not objRSol.Eof and not objRSol.bof then
	'	strIdLogico		= Trim(objRSol("Acl_IDAcessoLogico"))
 	'end if 
	strDataSolic	= Formatar_Data(Trim(objRS("Sol_Data")))
	strDM			= UCASE(objRS("Ped_Prefixo") & "-" & right("00000" & objRS("Ped_Numero"),5) & "/" & objRS("Ped_Ano"))
	strDataPedido	= Formatar_Data(objRS("Ped_Data"))
	strSatus		= Trim(objRS("Sts_Desc"))
	strCliente		= Trim(objRS("Cli_Nome"))
	intTipoProcesso = objRS("Tprc_idLog")
	strObs			= Trim(objRS("Sol_Obs"))
	PedObs			= Trim(objRS("Ped_Obs"))
	strAcao			= AcaoPedido(objRS("Tprc_id"))
	intTipoAcao		= objRS("Tprc_id")
	Id_Acesso		= objRS("Id_Acesso")
	dblNroSev		= objRS("Sol_SevSeq")

	set ObjFOdm = db.execute("select top 1 Ped_Prefixo, Ped_Numero , Ped_Ano  from cla_pedido where ped_id =  " & dblPedId )
	if not ObjFOdm.Eof and not ObjFOdm.Bof then
		strDM			= UCASE(ObjFOdm("Ped_Prefixo") & "-" & right("00000" & ObjFOdm("Ped_Numero"),5) & "/" & ObjFOdm("Ped_Ano"))
	end if


	Set objRSEndPto = db.execute("CLA_sp_view_Ponto null," & dblPedId)
	if not objRSEndPto.Eof and not objRSEndPto.bof then
		strEndereco		= objRSEndPto("Tpl_Sigla") & " " & objRSEndPto("End_NomeLogr") & ", " & objRSEndPto("End_NroLogr") & " " & objRSEndPto("Aec_Complemento") & " • " & objRSEndPto("End_Bairro") & " • " & objRSEndPto("End_Cep") & " • " & objRSEndPto("Cid_Desc") & " • " & objRSEndPto("Est_Sigla")
	End if
	Set objRSEndPto = Nothing

	strCidSigla	= Trim(objRS("Cid_Sigla"))
	strUfSigla	= Trim(objRS("Est_Sigla"))
	strTplSigla = Trim(objRS("Tpl_Sigla"))
	strNomeLogr	= Trim(objRS("End_NomeLogr"))
	strNroEnd	= Trim(objRS("End_NroLogr"))
	strCep		= Trim(objRS("End_Cep"))

	strNroServico	= Trim(objRS("Acl_NContratoServico"))
	strDesigServico = Trim(objRS("Acl_DesignacaoServico"))
	strServico		= Trim(objRS("Ser_Desc"))
	strVelServico	= Trim(objRS("DescVelAcessoLog"))
	strVelAcessoFis	= Trim(objRS("DescVelAcessoFis"))

	strPrmId		= Trim(objRS("Prm_id"))
	strRegId		= Trim(objRS("Reg_id"))

	strLocalInstala = Request.Form("estacao")
	if strLocalInstala = "" then  strLocalInstala	= Trim(objRS("Esc_IdEntrega"))
	strLocalConfig	= Trim(objRS("Esc_IdConfiguracao"))
	strRecurso		= Trim(objRS("Rec_IDEntrega"))

	strDistrib = Request.Form("distribuidor")
	if strDistrib = "" then  strDistrib	= Trim(objRS("Dst_Id"))

	'JCARTUS
	'if Request.Form("cboProvedor") <> "" then
	'	strRede = 1
	'Else
		strRede = objRS("Sis_ID")
	'End if

	
		
	If (strRede = 4 or strRede = 5) And strProId = 154 Then
		strRede = 13
		'response.write ">>>>>>>>>>PROV=" & strProId	
	End If
	'End if
'response.write ">>>>>> strRede=" & strRede
	set objRSFis = db.execute("CLA_sp_view_acessofisico null," & dblPedId)

	if Not objRSFis.Eof and not objRSFis.Bof then
		strTecnologia	= objRSFis("Tec_Sigla")
		strVelFis	=	objRSFis("Vel_Desc")
		intTipoVel	=	objRSFis("Acf_TipoVel")
	End if

	'GPON
	if TRIM(strTecnologia) <> "GPON" or isnull(strTecnologia) then
		DblAcf_ID = objRSFis("Acf_ID")
	end if
	
	set rsSwicth = db.execute (" select distinct cla_recurso.sis_id , cla_switch.switch_id  " &_
		" from  cla_Acessofisico		inner join cla_facilidade		on cla_Acessofisico.acf_id = cla_facilidade.acf_id " &_
		" inner join cla_ontporta			on cla_Acessofisico.acf_id = cla_ontporta.acf_id " &_
		" inner join cla_switch			on cla_ontporta.switch_id = cla_switch.switch_id " &_
		" inner join cla_ontvlan			on cla_ontporta.ontporta_id = cla_ontvlan.ontporta_id " &_
		" inner join cla_recurso			on cla_facilidade.rec_id = cla_recurso.rec_id " &_		
		" where cla_Acessofisico.acf_id = " & DblAcf_ID)
		
	if not rsSwicth.Eof then ' and objRSSolic.Bof) then  
		dblswitchID 	= rsSwicth("switch_ID") 
	end if 
	
  dim hdnNroAcessoEbt
  hdnNroAcessoEbt = ""
	set ObjTemp = db.execute ("select Acf_NroAcessoPtaEbt, pro_id from cla_acessofisico with (nolock) where acf_ID=" & DblAcf_ID)
	if not ObjTemp.eof and not ObjTemp.Bof then	 
	  	if ObjTemp("pro_id") = 145 and NOT isnull(ObjTemp("Acf_NroAcessoPtaEbt")) then 'FO ETHERNET
		  	hdnNroAcessoEbt = ObjTemp("Acf_NroAcessoPtaEbt")		  		  	
	    end if
  end if
	
	if strTecnologia = "FO EDD" then
		Set objRSSwitch = db.execute("select switch, switch_porta from cla_switch inner join cla_acessologico on cla_switch.acl_idacessologico = cla_acessologico.acl_idacessologico where cla_acessologico.acl_idacessologico = " & strIdLogico)
		if not objRSSwitch.Eof and not objRSSwitch.bof then
			strSwitch		= objRSSwitch("switch")
			strSwitchPorta	= objRSSwitch("switch_porta")
		end if 
	end if 
	
	set objRSPla = db.execute("CLA_sp_sel_facilidade null,null,null, " & DblAcf_ID )
	'Response.Write "CLA_sp_sel_facilidade null,null,null, " & objRSFis("Acf_ID")
	
	'Response.End
	if Not objRSPla.Eof and not objRSPla.Bof then
		strPla					= objRSPla("Pla_ID")
		strDesignacaoContrato	= objRSPla("DesignacaoContrato")
		if isnull(strPla) then strPla	= "0"
	else
		strPla	= "0"
	End if

	Set objRSFis = Nothing

Else
	Response.Write "<script language=javascript>window.location.replace('facilidade_main.asp')</script>"
	Response.End
End if

strUserGicL		= strUserName

'Deterministico
blnGravado = false

'*** Para atender a alocação a partir do controle de rede determinística
if Trim(strRede)		= "" or isNull(strRede)			then strRede	= Request.Form("cboRede")
if Trim(strProId)		= "" or isNull(strProId)		then strProId	= Request.Form("cboProvedor")
if Trim(strLocalInstala)= "" or isNull(strLocalInstala) then strLocalInstala = Request.Form("cboLocalInstala")
intCount = 0

'set objRS = db.execute("CLA_sp_sel_SwitchRec " & recID )

		'response.write "<script>alert('"&strLocalInstala&"')</script>"
		'response.write "<script>alert('"&Request.Form("cboLocalInstala")&"')</script>"
		'response.write "<script>alert('"&Request.Form("cboProvedor")&"')</script>"
		
		Vetor_Campos(1)="adInteger,8,adParamInput, " & strLocalInstala 'Request.Form("cboLocalInstala")
		Vetor_Campos(2)="adInteger,8,adParamInput, " & Request.Form("cboDistLocalInstala")
		Vetor_Campos(3)="adInteger,8,adParamInput, 12 " '& Request.Form("cboRede")
		Vetor_Campos(4)="adInteger,8,adParamInput, " & strProId 'Request.Form("cboProvedor")
		
		strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_SwitchRec",4,Vetor_Campos)
		'response.write strSqlRet
		Set objRS = db.Execute(strSqlRet)
		
		'onchange=ResgatarSwitchRadioIp()
		cbo = ""
		cbo = cbo & " Designação do SWITCH Concentrador <select name=cboSwtchRec onchange=ResgatarSwitchRadioIp() > "
		cbo = cbo & " <Option value="""">:: SWITCH </Option> "
		
		While Not objRS.eof
		  strItemSel = ""
		  if Trim(dblswitchID) = Trim(objRS("switch_ID")) then strItemSel = " Selected " End if
		  	'cbo = cbo & " <Option value='" & objRS("switch_ID") & "'" & strItemSel & ">" & Trim(objRS("switch")) & "</Option> "
		  
		   cbo = cbo & " <Option value=""" &  objRS("switch_ID") & """" & strItemSel & "> " & Trim(objRS("switch")) &  " </Option> "
		  objRS.MoveNext
		Wend
		strItemSel = ""
		
		cbo = cbo & " </select> "



strCamposBsodRadioIP = " <table cellspacing=1 cellpadding=0 width=760 border=0> "
strCamposBsodRadioIP = strCamposBsodRadioIP & "	<tr class=clsSilver> "
'strCamposBsodRadioIP = strCamposBsodRadioIP & "		<td colspan=5><font class=clsObrig>::</font>&nbsp;Designação do SWITCH Concentrador&nbsp;<input type=text name=txtSWITCH size=20  class=text></td> "

strCamposBsodRadioIP = strCamposBsodRadioIP & "		<td colspan=5><font class=clsObrig>::</font>&nbsp;  "

strCamposBsodRadioIP = strCamposBsodRadioIP & cbo

strCamposBsodRadioIP = strCamposBsodRadioIP & " </td> "
strCamposBsodRadioIP = strCamposBsodRadioIP & "	</tr> "
strCamposBsodRadioIP = strCamposBsodRadioIP & " </table> "


				set objRS = db.execute (" select distinct equi.equi_id , tipoEqui.tipoEqui_id , prop.prop_id  " &_
										" from cla_equipamento	equi " &_
										" inner join  cla_tipo_equipamento  tipoEqui			on equi.TipoEqui_id = tipoEqui.TipoEqui_id " &_
										" inner join  cla_propriedade		  prop				on equi.prop_id = prop.prop_id " &_
										" inner join  cla_ontporta	ontporta				on equi.equi_id = ontporta.equi_id " &_
										" inner join  cla_porta_alocada ptAlocada				on ontporta.ontporta_id = ptAlocada.ontporta_id " &_		
										" where ptAlocada.acl_idacessologico = " &  strIdLogico)
				While Not objRS.eof
				  strItemSel = ""
				  
					
				  
					if "2" = Trim(objRS("tipoEqui_id")) then 
						 dblswitchMetro 	= objRS("equi_id")
						 'response.write "<script>alert('"&dblswitchMetro&"')</script>"
					End if
					
					if "3" = Trim(objRS("tipoEqui_id")) then 
						dblswitchInterconexao 	= objRS("equi_id")
						propswitchInterconexao	= objRS("prop_id") 
					End if
					
					if "4" = Trim(objRS("tipoEqui_id")) then 
						dblswitchEDD 	= objRS("equi_id")
						propswitchEDD	= objRS("prop_id")
					End if
					
					if "9" = Trim(objRS("tipoEqui_id")) then 
						
						dblswitchPE 	= objRS("equi_id")
						propswitchPE	= objRS("prop_id")
						'response.write "<script>alert('"&dblswitchPE&"')</script>"
					End if
					
					
					
				  objRS.MoveNext
				Wend


				Vetor_Campos(1)="adInteger,8,adParamInput, " & strLocalInstala 'Request.Form("cboLocalInstala")
				Vetor_Campos(2)="adInteger,8,adParamInput, 3 " 
				Vetor_Campos(3)="adInteger,8,adParamInput, 14 " 
				Vetor_Campos(4)="adInteger,8,adParamInput, " & strProId 'Request.Form("cboProvedor")
				
				strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_DesignacaoSwitch",4,Vetor_Campos)
				'response.write strSqlRet
				Set objRS = db.Execute(strSqlRet)
				
				'onchange=ResgatarSwitchRadioIp()
				cboInterconexao = ""
				cboInterconexao = cboInterconexao & " Designação do SWITCH INTERCONEXÃO <select name=cboSwitchInterconexao disabled=true onchange=ResgatarEthernetInterconexao() > "
				cboInterconexao = cboInterconexao & " <Option value="""">:: SWITCH INTERCONEXÃO</Option> "
				
				While Not objRS.eof
				  strItemSel = ""
				  if Trim(dblswitchInterconexao) = Trim(objRS("equi_id")) then strItemSel = " Selected " End if
							  
						cboInterconexao = cboInterconexao & " <Option value=""" &  objRS("equi_id") & """" & strItemSel & "> " & Trim(objRS("designacao")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
				
				cboInterconexao = cboInterconexao & " </select> "

'				strCamposEthernetInterconexao = strCamposEthernetInterconexao & " <script LANGUAGE=javascript>ResgatarEthernetInterconexao()</script> "

				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "	<table  cellspacing=1 cellpadding=0 width=760 border=0> "
				
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "		<tr>"
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & " 			<th  colspan=5><font class=clsObrig>::</font>&nbsp; SWITCH INTERCONEXÃO</th> "
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "		</tr> "
		
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "		<tr class=clsSilver> "
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "	 		<td colspan=5><font class=clsObrig>::</font>&nbsp;Propriedade do Switch INTERCONEXÃO &nbsp;&nbsp;&nbsp;&nbsp; "
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "				<select name=cboPropriedadeInter onchange=""parent.HabilitaDesignacao(this , \'cboSwitchInterconexao\' )""> "
				
				set objRS = db.execute("CLA_sp_sel_Propriedade ")
				
							
				While Not objRS.eof
				  strItemSel = ""
				  if Trim(propswitchInterconexao) = Trim(objRS("PROP_ID")) then strItemSel = " Selected " End if
				  strCamposEthernetInterconexao = strCamposEthernetInterconexao & " 		<Option value=""" &  objRS("PROP_ID") & """" & strItemSel & "> " & Trim(objRS("PROP_SIGLA")) &  " </Option> "
				  objRS.MoveNext
				Wend
				
				
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & " 				</select> "
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & " 			</td> "
				
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "	  	</tr> "
				
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "	  	<tr class=clsSilver> "


				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "			<td colspan=5><font class=clsObrig>::</font>"

				strCamposEthernetInterconexao = strCamposEthernetInterconexao & 				cboInterconexao

				strCamposEthernetInterconexao = strCamposEthernetInterconexao & " 			</td> "
				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "	  	</tr> "

				strCamposEthernetInterconexao = strCamposEthernetInterconexao & "  </table> "
				
				
				'Response.write " ResgatarEthernetMetro(); "
				'Response.write " ResgatarEthernetEDD(); "
				'Response.write " ResgatarEthernetPE(); "
				
				
				
				
				Vetor_Campos(1)="adInteger,8,adParamInput, " & strLocalInstala 'Request.Form("cboLocalInstala")
				Vetor_Campos(2)="adInteger,8,adParamInput, 2 " 
				Vetor_Campos(3)="adInteger,8,adParamInput, 14 " 
				Vetor_Campos(4)="adInteger,8,adParamInput, " & strProId 'Request.Form("cboProvedor")
				
				strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_DesignacaoSwitch",4,Vetor_Campos)
				'response.write strSqlRet
				Set objRS = db.Execute(strSqlRet)
				
				'onchange=ResgatarSwitchRadioIp()
				cboMetro = ""
				cboMetro = cboMetro & " Designação do SWITCH METRO <select name=cboSwitchMetro   onchange=ResgatarEthernetMetro() > "
				cboMetro = cboMetro & " <Option value="""">:: SWITCH METRO</Option> "
				
				While Not objRS.eof
				  strItemSel = ""
				  if Trim(dblswitchMetro) = Trim(objRS("equi_id")) then strItemSel = " Selected " End if
							'response.write "<script>alert('"&dblswitchMetro&"')</script>"
							'response.write "<script>alert('"&objRS("equi_id")&"')</script>"
									
						cboMetro = cboMetro & " <Option value=""" &  objRS("equi_id") & """" & strItemSel & "> " & Trim(objRS("designacao")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
				
				cboMetro = cboMetro & " </select> "
				
				
				strCamposEthernetMetro = strCamposEthernetMetro & "	<table  cellspacing=1 cellpadding=0 width=760 border=0> "
				
				strCamposEthernetMetro = strCamposEthernetMetro & "		<tr>"
				strCamposEthernetMetro = strCamposEthernetMetro & " 			<th  colspan=5><font class=clsObrig>::</font>&nbsp; SWITCH METRO</th> "
				strCamposEthernetMetro = strCamposEthernetMetro & "		</tr> "
		
				'strCamposEthernetMetro = strCamposEthernetMetro & "		<tr class=clsSilver> "
				'strCamposEthernetMetro = strCamposEthernetMetro & "	 		<td colspan=5><font class=clsObrig>::</font>&nbsp;Propriedade do Switch METRO &nbsp;&nbsp;&nbsp;&nbsp; "
				'strCamposEthernetMetro = strCamposEthernetMetro & "				<select name=cboPropriedade> "
				
				'set objRS = db.execute("CLA_sp_sel_Propriedade ")
				
							
				'While Not objRS.eof
				'  strItemSel = ""
				'  strCamposEthernetMetro = strCamposEthernetMetro & " <Option value=""" &  objRS("PROP_ID") & """" & strItemSel & "> " & Trim(objRS("PROP_SIGLA")) &  " </Option> "
				'  objRS.MoveNext
				'Wend
				
				
				'strCamposEthernetMetro = strCamposEthernetMetro & " 				</select> "
				'strCamposEthernetMetro = strCamposEthernetMetro & " 			</td> "
				
				'strCamposEthernetMetro = strCamposEthernetMetro & "	  </tr> "
				
				
				strCamposEthernetMetro = strCamposEthernetMetro & "	  	<tr class=clsSilver> "


				strCamposEthernetMetro = strCamposEthernetMetro & "			<td colspan=5><font class=clsObrig>::</font>"

				strCamposEthernetMetro = strCamposEthernetMetro & 				cboMetro

				strCamposEthernetMetro = strCamposEthernetMetro & " 			</td> "
				strCamposEthernetMetro = strCamposEthernetMetro & "	  	</tr> "
				
				strCamposEthernetMetro = strCamposEthernetMetro & " </table> "
				
				
				Vetor_Campos(1)="adInteger,8,adParamInput, " & strLocalInstala 'Request.Form("cboLocalInstala")
				Vetor_Campos(2)="adInteger,8,adParamInput, 4 " 
				Vetor_Campos(3)="adInteger,8,adParamInput, 14 " 
				Vetor_Campos(4)="adInteger,8,adParamInput, " & strProId 'Request.Form("cboProvedor")
				
				strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_DesignacaoSwitch",4,Vetor_Campos)
				'response.write strSqlRet
				Set objRS = db.Execute(strSqlRet)
				
				'onchange=ResgatarSwitchRadioIp()
				cboEDD = ""
				cboEDD = cboEDD & " Designação do SWITCH EDD <select name=cboSwitchEDD disabled=true onchange=ResgatarEthernetEDD() > "
				cboEDD = cboEDD & " <Option value="""">:: SWITCH EDD</Option> "
				
				While Not objRS.eof
				  strItemSel = ""
				  if Trim(dblswitchEDD) = Trim(objRS("equi_id")) then strItemSel = " Selected " End if
							  
						cboEDD = cboEDD & " <Option value=""" &  objRS("equi_id") & """" & strItemSel & "> " & Trim(objRS("designacao")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
				
				cboEDD = cboEDD & " </select> "
				
				
				strCamposEthernetEDD = strCamposEthernetEDD & "	<table  cellspacing=1 cellpadding=0 width=760 border=0> "
				
				strCamposEthernetEDD = strCamposEthernetEDD & "		<tr>"
				strCamposEthernetEDD = strCamposEthernetEDD & " 			<th  colspan=5><font class=clsObrig>::</font>&nbsp; SWITCH EDD</th> "
				strCamposEthernetEDD = strCamposEthernetEDD & "		</tr> "
		
				strCamposEthernetEDD = strCamposEthernetEDD & "		<tr class=clsSilver> "
				strCamposEthernetEDD = strCamposEthernetEDD & "	 		<td colspan=5><font class=clsObrig>::</font>&nbsp;Propriedade do Switch EDD &nbsp;&nbsp;&nbsp;&nbsp; "
				strCamposEthernetEDD = strCamposEthernetEDD & "				<select name=cboPropriedadeEDD onchange=""parent.HabilitaDesignacao(this , \'cboSwitchEDD\' )""> " 
				
				set objRS = db.execute("CLA_sp_sel_Propriedade ")
				
							
				While Not objRS.eof
				  strItemSel = ""
				   if Trim(propswitchEDD) = Trim(objRS("PROP_ID")) then strItemSel = " Selected " End if
				   
					strCamposEthernetEDD = strCamposEthernetEDD & " <Option value=""" &  objRS("PROP_ID") & """" & strItemSel & "> " & Trim(objRS("PROP_SIGLA")) &  " </Option> "
					objRS.MoveNext
				Wend
				
				
				strCamposEthernetEDD = strCamposEthernetEDD & " 				</select> "
				strCamposEthernetEDD = strCamposEthernetEDD & " 			</td> "
				
				strCamposEthernetEDD = strCamposEthernetEDD & "	  </tr> "
				
				strCamposEthernetEDD = strCamposEthernetEDD & "	  	<tr class=clsSilver> "


				strCamposEthernetEDD = strCamposEthernetEDD & "			<td colspan=5><font class=clsObrig>::</font>"

				strCamposEthernetEDD = strCamposEthernetEDD & 				cboEDD

				strCamposEthernetEDD = strCamposEthernetEDD & " 			</td> "
				strCamposEthernetEDD = strCamposEthernetEDD & "	  	</tr> "
				
				
				strCamposEthernetEDD = strCamposEthernetEDD & " </table> "
				
				
				
				Vetor_Campos(1)="adInteger,8,adParamInput, " & strLocalInstala 'Request.Form("cboLocalInstala")
				Vetor_Campos(2)="adInteger,8,adParamInput, 9 " 
				Vetor_Campos(3)="adInteger,8,adParamInput, 14 " 
				Vetor_Campos(4)="adInteger,8,adParamInput, " & strProId 'Request.Form("cboProvedor")
				
				strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_DesignacaoSwitch",4,Vetor_Campos)
				'response.write strSqlRet
				Set objRS = db.Execute(strSqlRet)
				
				'onchange=ResgatarSwitchRadioIp()
				cboPE = ""
				cboPE = cboPE & " Designação do SWITCH PE <select name=cboSwitchPE onchange=ResgatarEthernetPE() > "
				cboPE = cboPE & " <Option value="""">:: SWITCH PE</Option> "
				
				While Not objRS.eof
				  strItemSel = ""
				  if Trim(dblswitchPE) = Trim(objRS("equi_id")) then strItemSel = " Selected " End if
							  
						cboPE = cboPE & " <Option value=""" &  objRS("equi_id") & """" & strItemSel & "> " & Trim(objRS("designacao")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
				
				cboPE = cboPE & " </select> "
				
				
				strCamposEthernetPE = strCamposEthernetPE & "	<table  cellspacing=1 cellpadding=0 width=760 border=0> "
				
				strCamposEthernetPE = strCamposEthernetPE & "		<tr>"
				strCamposEthernetPE = strCamposEthernetPE & " 			<th  colspan=5><font class=clsObrig>::</font>&nbsp; PE</th> "
				strCamposEthernetPE = strCamposEthernetPE & "		</tr> "
		
				'strCamposEthernetPE = strCamposEthernetPE & "		<tr class=clsSilver> "
				'strCamposEthernetPE = strCamposEthernetPE & "	 		<td colspan=5><font class=clsObrig>::</font>&nbsp;Propriedade do PE &nbsp;&nbsp;&nbsp;&nbsp; "
				'strCamposEthernetPE = strCamposEthernetPE & "				<select name=cboPropriedadePE onchange=""parent.HabilitaDesignacao(this , \'cboSwitchPE\' )""> " 
				
				'set objRS = db.execute("CLA_sp_sel_Propriedade ")
				
							
				'While Not objRS.eof
				'  strItemSel = ""
				'   if Trim(propswitchPE) = Trim(objRS("PROP_ID")) then strItemSel = " Selected " End if
				'  strCamposEthernetPE = strCamposEthernetPE & " <Option value=""" &  objRS("PROP_ID") & """" & strItemSel & "> " & Trim(objRS("PROP_SIGLA")) &  " </Option> "
				'  objRS.MoveNext
				'Wend
				
				
				'strCamposEthernetPE = strCamposEthernetPE & " 				</select> "
				'strCamposEthernetPE = strCamposEthernetPE & " 			</td> "
				
				'strCamposEthernetPE = strCamposEthernetPE & "	  </tr> "
				
				
				strCamposEthernetPE = strCamposEthernetPE & "	  	<tr class=clsSilver> "


				strCamposEthernetPE = strCamposEthernetPE & "			<td colspan=5><font class=clsObrig>::</font>"

				strCamposEthernetPE = strCamposEthernetPE & 				cboPE

				strCamposEthernetPE = strCamposEthernetPE & " 			</td> "
				strCamposEthernetPE = strCamposEthernetPE & "	  	</tr> "
				
				
				strCamposEthernetPE = strCamposEthernetPE & " </table> "
				
				
				strCamposEthernetVlan = strCamposEthernetVlan & "<table cellspacing=1 cellpadding=0 width=760 border=0>"
				strCamposEthernetVlan = strCamposEthernetVlan & "	<tr><th colspan=10>&nbsp;•&nbsp;Finalização da Alocação</th></tr>"
				strCamposEthernetVlan = strCamposEthernetVlan & "	<tr class=clsSilver>"
				
				strCamposEthernetVlan = strCamposEthernetVlan & "		<td width=200 ><font class=clsObrig>:: </font>Designação do Acesso Terceiro</td>"
				strCamposEthernetVlan = strCamposEthernetVlan & "		<td ><input type=text class=text name=txtNroAcessoEbtEthernet maxlength=50 size=50 ></td>"
				
				'''strCamposEthernetVlan = strCamposEthernetVlan & "		<td >&nbsp;<font class=clsObrig>:: </font>CVLAN Cliente &nbsp;<input type=text class=text maxlength=5 size=7 name=txtVlan onKeyUp=""ValidarTipo(this,0)"" ></td>"
				'''strCamposEthernetVlan = strCamposEthernetVlan & "		<td colspan=2 >&nbsp;<font class=clsObrig></font>SVLAN Serviço&nbsp;<input type=text class=text maxlength=4 size=7 name=txtSvlan onKeyUp=""ValidarTipo(this,0)""></td>"

				strCamposEthernetVlan = strCamposEthernetVlan & "	</tr>"
				
				
				
				strCamposEthernetVlan = strCamposEthernetVlan & "	<tr> "
				strCamposEthernetVlan = strCamposEthernetVlan & "		<td colspan=5 align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnFinalizarAlocar"" style=""width:150px;height:22px"" value=""Finalizar Alocação"" onclick=GravarEthernet() ></td> "
				strCamposEthernetVlan = strCamposEthernetVlan & "	</tr> "
				strCamposEthernetVlan = strCamposEthernetVlan & " </table> "

strCamposBsodVia = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposBsodVia = strCamposBsodVia & "<th width=200>&nbsp;Designacao do Acesso</th>"
strCamposBsodVia = strCamposBsodVia & "<th width=80 nowrap>&nbsp;VLAN</th>"
strCamposBsodVia = strCamposBsodVia & "<th width=200>&nbsp;PE</th>"
strCamposBsodVia = strCamposBsodVia & "<th width=140>&nbsp;Porta</th>"
strCamposBsodVia = strCamposBsodVia & "<th width=110>&nbsp;Link</th></table>"

strCamposBsod = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposBsod = strCamposBsod & "<th width=170>&nbsp;Designacao do Acesso</th>"
strCamposBsod = strCamposBsod & "<th width=80 nowrap>&nbsp;CVLAN</th>"
strCamposBsod = strCamposBsod & "<th width=150>&nbsp;PE</th>"
strCamposBsod = strCamposBsod & "<th width=140>&nbsp;Porta</th>"
strCamposBsod = strCamposBsod & "<th width=90>&nbsp;SVLAN</th>"
strCamposBsod = strCamposBsod & "<th width=100>&nbsp;Tipo Porta</th></table>"


strBsodVia = strBsodVia & "<table rules=groups border=0 cellspacing=1 cellpadding=1 bordercolorlight=#003388 bordercolordark=#ffffff width=760>"
strBsodVia = strBsodVia & "<tr><th colspan=10>&nbsp;•&nbsp;Informações da Facilidade</th></tr>"
strBsodVia = strBsodVia & "<tr class=clsSilver>"
strBsodVia = strBsodVia & "<td width=170 ><font class=clsObrig>:: </font>Designação do Acesso</td>"
strBsodVia = strBsodVia & "<td ><input type=text class=text name=txtNroAcessoEbt maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}"" >"
strBsodVia = strBsodVia & "</td>"
strBsodVia = strBsodVia & "<td >&nbsp;<font class=clsObrig>:: </font>VLAN &nbsp;<input type=text class=text maxlength=6 size=7 name=txtVlan onKeyUp=""ValidarTipo(this,0)"" ></td>"
strBsodVia = strBsodVia & "<td colspan=2 >&nbsp;<font class=clsObrig></font>Link &nbsp;<input type=text class=text maxlength=25 size=25 name=txtLink></td>"

strBsodVia = strBsodVia & "</tr>"
strBsodVia = strBsodVia & "<tr class=clsSilver>"

strBsodVia = strBsodVia & "<td width=120>&nbsp;<font class=clsObrig>:: </font>PE</td>"
strBsodVia = strBsodVia & "<td ><input type=text class=text maxlength=30 size=32 name=txtPE></td>"

strBsodVia = strBsodVia & "<td colspan=3>&nbsp;<font class=clsObrig>:: </font>Porta &nbsp;&nbsp;<input type=text class=text maxlength=10 size=15 name=txtPorta></td>"
'strBsodVia = strBsodVia & "<td  >&nbsp;</td>"
strBsodVia = strBsodVia & "</tr>"
strBsodVia = strBsodVia & "<tr class=clsSilver2>"
strBsodVia = strBsodVia & "<td colspan=10 align=right>"
strBsodVia = strBsodVia & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar"" onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strBsodVia = strBsodVia & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strBsodVia = strBsodVia & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"
strBsodVia = strBsodVia & "</td>"
strBsodVia = strBsodVia & "</tr>"
strBsodVia = strBsodVia & "</table>"


strBsod = strBsod & "<table rules=groups border=0 cellspacing=1 cellpadding=1 bordercolorlight=#003388 bordercolordark=#ffffff width=760>"
strBsod = strBsod & "<tr><th colspan=10>&nbsp;•&nbsp;Informações da Facilidade</th></tr>"
strBsod = strBsod & "<tr class=clsSilver>"
strBsod = strBsod & "<td width=170 ><font class=clsObrig>:: </font>Designação do Acesso</td>"
strBsod = strBsod & "<td ><input type=text class=text name=txtNroAcessoEbt maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}"" >"
strBsod = strBsod & "</td>"
strBsod = strBsod & "<td >&nbsp;<font class=clsObrig>:: </font>CVLAN &nbsp;<input type=text class=text maxlength=5 size=7 name=txtVlan onKeyUp=""ValidarTipo(this,0)"" ></td>"
strBsod = strBsod & "<td colspan=2 >&nbsp;<font class=clsObrig></font>SVLAN &nbsp;<input type=text class=text maxlength=4 size=7 name=txtSvlan onKeyUp=""ValidarTipo(this,0)""></td>"

strBsod = strBsod & "</tr>"
strBsod = strBsod & "<tr class=clsSilver>"

strBsod = strBsod & "<td width=120>&nbsp;<font class=clsObrig>:: </font>PE</td>"
strBsod = strBsod & "<td ><input type=text class=text maxlength=30 size=20 name=txtPE></td>"

strBsod = strBsod & "<td colspan=3>&nbsp;<font class=clsObrig>:: </font>Porta &nbsp;&nbsp;<input type=text class=text maxlength=30 size=15 name=txtPorta> "
'strBsod = strBsod & "&nbsp;<font class=clsObrig>:: </font>T. Porta &nbsp;&nbsp;<input type=text class=text maxlength=10 size=15 name=cboTipoPorta> "
strBsod = strBsod & "&nbsp;<font class=clsObrig>:: </font>Tipo Porta &nbsp;&nbsp;"



strBsod = strBsod & "<select name=cboTipoPorta>"
strBsod = strBsod & "	<option value=""""></option>"
		Set objRS = db.Execute("select * from CLA_TipoPorta_HFCBSoD")
		While Not objRS.eof		   
		   strBsod = strBsod & " <option value=" & objRS("Cod_tipoPorta") & ">" & objRS("Desc_tipoPorta") & "</option>"
		  objRS.MoveNext
		Wend

' strBsod = strBsod & "	<option value=""""></option>"
' strBsod = strBsod & "	<option value=""E"">ETHERNET</option>"
' strBsod = strBsod & "	<option value=""FE"">FAST ETHERNET</option>"
' strBsod = strBsod & "	<option value=""GE"">GIGABIT ETHERNET</option>"
' strBsod = strBsod & "	<option value=""TGE"">TEN GIGABIT ETHERNET</option>"

strBsod = strBsod & "	</select>"

strBsod = strBsod & "</td>"
'strBsodVia = strBsod & "<td  >&nbsp;</td>"
strBsod = strBsod & "</tr>"
strBsod = strBsod & "<tr class=clsSilver2>"
strBsod = strBsod & "<td colspan=10 align=right>"
strBsod = strBsod & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar"" onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strBsod = strBsod & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strBsod = strBsod & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"
strBsod = strBsod & "</td>"
strBsod = strBsod & "</tr>"
strBsod = strBsod & "</table>"

strCamposFoEtherNet = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposFoEtherNet = strCamposFoEtherNet & "<th width=200>&nbsp;Designacao do Acesso</th>"
strCamposFoEtherNet = strCamposFoEtherNet & "<th width=80 nowrap>&nbsp;VLAN</th>"
strCamposFoEtherNet = strCamposFoEtherNet & "<th width=200>&nbsp;PE</th>"
strCamposFoEtherNet = strCamposFoEtherNet & "<th width=80>&nbsp;Porta</th>"
strCamposFoEtherNet = strCamposFoEtherNet & "<th width=170>&nbsp;SVLAN</th></table>"

strFoEtherNet = strFoEtherNet & "<table rules=groups border=0 cellspacing=1 cellpadding=1 bordercolorlight=#003388 bordercolordark=#ffffff width=760>"
strFoEtherNet = strFoEtherNet & "<tr><th colspan=10>&nbsp;•&nbsp;Informações da Facilidade</th></tr>"
strFoEtherNet = strFoEtherNet & "<tr class=clsSilver>"
strFoEtherNet = strFoEtherNet & "<td width=170 ><font class=clsObrig>:: </font>Designação do Acesso</td>"
'if strDesignacaoContrato <> "" then
'	strFoEtherNet = strFoEtherNet & "<td ><input type=text class=text name=txtNroAcessoEbt value=" & strDesignacaoContrato & " maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}"" >"
'else
	strFoEtherNet = strFoEtherNet & "<td ><input type=text class=text name=txtNroAcessoEbt maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}"" >"
'end if

strFoEtherNet = strFoEtherNet & "</td>"
strFoEtherNet = strFoEtherNet & "<td colspan=7 >&nbsp;<font class=clsObrig>:: </font>VLAN &nbsp;<input type=text class=text maxlength=5 size=7 name=txtVlan_FO onKeyUp=""ValidarTipo(this,0)"" ></td>"
strFoEtherNet = strFoEtherNet & "</tr>"
strFoEtherNet = strFoEtherNet & "<tr class=clsSilver>"

strFoEtherNet = strFoEtherNet & "<td width=170>&nbsp;<font class=clsObrig>:: </font>PE</td>"
strFoEtherNet = strFoEtherNet & "<td ><input type=text class=text maxlength=16 size=20 name=txtPE_FO></td>"

strFoEtherNet = strFoEtherNet & "<td width=150 ><font class=clsObrig>:: </font>Porta &nbsp;&nbsp;<input type=text class=text maxlength=10 size=11 name=txtPorta_FO></td>"
strFoEtherNet = strFoEtherNet & "<td colspan=2 >&nbsp;<font class=clsObrig></font>SVLAN &nbsp;<input type=text class=text maxlength=5 size=7 name=txtSvlan_FO onKeyUp=""ValidarTipo(this,0)""></td>"
strFoEtherNet = strFoEtherNet & "</tr>"
strFoEtherNet = strFoEtherNet & "<tr class=clsSilver2>"
strFoEtherNet = strFoEtherNet & "<td colspan=10 align=right>"
strFoEtherNet = strFoEtherNet & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar"" onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strFoEtherNet = strFoEtherNet & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strFoEtherNet = strFoEtherNet & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"

strFoEtherNet = strFoEtherNet & "<input type=""hidden"" name=""hdnOrigem"" value="& orisolid & ">"
strFoEtherNet = strFoEtherNet & "<input type=""Hidden"" name=""hdnId_AcessoAsms"" value="& Id_Acesso&">"
strFoEtherNet = strFoEtherNet & "</td>"
strFoEtherNet = strFoEtherNet & "</tr>"
strFoEtherNet = strFoEtherNet & "</table>"

strCamposBsodNet = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposBsodNet = strCamposBsodNet & "<th width=200>&nbsp;Designacao do Acesso</th>"
strCamposBsodNet = strCamposBsodNet & "<th width=530 nowrap>&nbsp;VLAN</th></table>"

strBsodNet = strBsodNet & "<table rules=groups border=0 cellspacing=1 cellpadding=1 bordercolorlight=#003388 bordercolordark=#ffffff width=760>"

strBsodNet = strBsodNet & "<tr><th colspan=10>&nbsp;•&nbsp;Informações da Facilidade</th></tr>"
strBsodNet = strBsodNet & "<tr class=clsSilver>"
strBsodNet = strBsodNet & "<td width=170 ><font class=clsObrig>:: </font>Designação do Acesso</td>"
strBsodNet = strBsodNet & "<td ><input type=text class=text name=txtNroAcessoEbt maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}"">"
strBsodNet = strBsodNet & "</td>"
strBsodNet = strBsodNet & "<td colspan=7 >&nbsp;<font class=clsObrig>:: </font>VLAN &nbsp;<input type=text class=text maxlength=5 size=7 name=txtVlan onKeyUp=""ValidarTipo(this,0)"" ></td>"
strBsodNet = strBsodNet & "</tr>"

strBsodNet = strBsodNet & "<tr class=clsSilver2>"
strBsodNet = strBsodNet & "<td colspan=10 align=right>"
strBsodNet = strBsodNet & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar"" onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strBsodNet = strBsodNet & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strBsodNet = strBsodNet & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"
strBsodNet = strBsodNet & "</td>"
strBsodNet = strBsodNet & "</tr>"

strBsodNet = strBsodNet & "</table>"



'Bsod Light
strCamposBsodLight = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposBsodLight = strCamposBsodLight & "<th width=200>&nbsp;Designacao do Acesso</th>"
strCamposBsodLight = strCamposBsodLight & "<th width=150 nowrap>&nbsp;IP</th>"
strCamposBsodLight = strCamposBsodLight & "<th width=150>&nbsp;Gateway</th>"
strCamposBsodLight = strCamposBsodLight & "<th width=210>&nbsp;Máscara</th></table>"



strBsodLight = strBsodLight & "<table rules=groups border=0 cellspacing=1 cellpadding=1 bordercolorlight=#003388 bordercolordark=#ffffff width=760>"

strBsodLight = strBsodLight & "<tr><th colspan=10>&nbsp;•&nbsp;Informações da Facilidade</th></tr>"
strBsodLight = strBsodLight & "<tr class=clsSilver>"
strBsodLight = strBsodLight & "<td width=170 ><font class=clsObrig>:: </font>Designação do Acesso</td>"
strBsodLight = strBsodLight & "<td ><input type=text class=text name=txtNroAcessoEbt maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}"">"
strBsodLight = strBsodLight & "</td>"
strBsodLight = strBsodLight & "<td ><font class=clsObrig>:: </font>IP</td>"
strBsodLight = strBsodLight & "<td ><input type=text class=text name=txtIP maxlength=15 size=18 OnKeyUp=""formatarIP(this);ValidarTipo(this,8)"">"
strBsodLight = strBsodLight & "</td>"
strBsodLight = strBsodLight & "</tr>"

strBsodLight = strBsodLight & "<tr class=clsSilver>"

strBsodLight = strBsodLight & "<td width=170 ><font class=clsObrig>:: </font>Gateway</td>"
strBsodLight = strBsodLight & "<td ><input type=text class=text name=txtGateway maxlength=15 size=18 OnKeyUp=""formatarIP(this);ValidarTipo(this,8)"">"
strBsodLight = strBsodLight & "</td>"
strBsodLight = strBsodLight & "<td width=170 ><font class=clsObrig>:: </font>Máscara</td>"
strBsodLight = strBsodLight & "<td ><input type=text class=text name=txtMascara maxlength=15 size=18 OnKeyUp=""formatarIP(this);ValidarTipo(this,8)"">"
strBsodLight = strBsodLight & "</td>"
strBsodLight = strBsodLight & "</tr>"

strBsodLight = strBsodLight & "<tr class=clsSilver2>"
strBsodLight = strBsodLight & "<td colspan=10 align=right>"
strBsodLight = strBsodLight & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar"" onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strBsodLight = strBsodLight & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strBsodLight = strBsodLight & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"
strBsodLight = strBsodLight & "</td>"
strBsodLight = strBsodLight & "</tr>"

strBsodLight = strBsodLight & "</table>"


strCamposDet = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposDet = strCamposDet & "<th width=100>&nbsp;Provedor</th>"
strCamposDet = strCamposDet & "<th width=100 nowrap>&nbsp;Nº Acesso Pta Ebt</th>"
strCamposDet = strCamposDet & "<th width=80>&nbsp;Fila</th>"
strCamposDet = strCamposDet & "<th width=80>&nbsp;Bastidor</th>"
strCamposDet = strCamposDet & "<th width=80>&nbsp;Régua</th>"
strCamposDet = strCamposDet & "<th width=80>&nbsp;Posição</th>"
strCamposDet = strCamposDet & "<th width=80>&nbsp;Timeslot</th>"
strCamposDet = strCamposDet & "<th width=130>&nbsp;Nº Acesso Pta CLI</th></table>"

strDet = strDet & "<table rules=groups border=0 cellspacing=1 cellpadding=1 bordercolorlight=#003388 bordercolordark=#ffffff width=760>"
strDet = strDet & "<tr>"
strDet = strDet & "<th colspan=7>&nbsp;•&nbsp;Informações da Facilidade</th>"
strDet = strDet & "</tr>"
strDet = strDet & "<tr class=clsSilver2>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Cod.Prov</td>"
strDet = strDet & "<td>&nbsp;Nº Acesso Pta Ebt</td>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Fila</td>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Bastidor</td>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Régua</td>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Posição</td>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Timeslot</td>"
strDet = strDet & "	</tr>"

strDet = strDet & "<tr class=clsSilver>"
strDet = strDet & "<td><span id=spnCodProv><select name=cboCodProv>"
strDet = strDet & "	<option value=""""></option>"
if Trim(Request.Form("cboProvedor")) <> "" then
	Set objRS = db.execute("CLA_sp_sel_Provedor null," & Request.Form("cboProvedor"))
	While Not objRS.Eof
		strDet = strDet & "<Option value=" & Trim(objRS("Pro_ID")) & ">" & TratarAspasJS(objRS("Pro_Cod")) & "</Option>"
		objRS.MoveNext
	Wend
End if

strDet = strDet & "</select></span></td>"
strDet = strDet & "	<td ><input type=text class=text name=txtNroAcessoEbt maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}""></td>"
strDet = strDet & "	<td ><input type=text class=text readonly name=txtFila maxlength=25 size=10 value="""&Request.Form("fila")&""" >(A5)</td>"
strDet = strDet & "	<td ><input type=text class=text readonly name=txtBastidor maxlength=3 size=4 value="""&Request.Form("bastidor")&""" onKeyUp=""ValidarTipo(this,0)"">(N3)</td>"
strDet = strDet & "	<td ><input type=text class=text readonly name=txtRegua maxlength=2 size=3 value="""&Request.Form("regua")&""" >(A2)</td>"
strDet = strDet & "	<td ><input type=text class=text readonly name=txtPosicao maxlength=3 size=5 value="""&Request.Form("posicao")&""" >(A3)</td>"
strDet = strDet & "	<td ><input type=text class=text name=txtTimeslot maxlength=9 size=10 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N  value="""&Request.Form("timeslot")&""">(N4-N4)</td>"
strDet = strDet & "</tr>"

strDet = strDet & "<tr class=clsSilver2>"
strDet = strDet & " <td>Nº Acesso Pta Cli</td>"
strDet = strDet & "	<td nowrap>CCTO Provedor</td>"
strDet = strDet & "	<td nowrap>IU / URB</td>"
strDet = strDet & "	<td nowrap>CNL Ponta A</td>"
strDet = strDet & "	<td nowrap>CNL Ponta B</td>"

''@davif
''response.write "Portadora " & strExigePortadora & "davi"
if (strExigePortadora) = "S" Then

	strDet = strDet & "	<td nowrap colspan=2>Portadora</td>"
else
	strDet = strDet & "	<td nowrap colspan=2></td>"
end if
''/@davif

strDet = strDet & "</tr>"

strDet = strDet & "<tr class=clsSilver>"
strDet = strDet & "	<td ><input type=text class=text name=txtNroAcessoCli size=15 maxlength=25 >(A25)</td>"
strDet = strDet & "<td><input type=text size=8 maxlength=8 class=text name=txtCCTOPro >(A8)</td>"
strDet = strDet & "<td >"
strDet = strDet & "	<input type=radio name=rdoUrbano value=I Index=0  onClick=""AdicionarCNL(this)"">IU&nbsp;"
strDet = strDet & "	<input type=radio name=rdoUrbano value=U Index=1 onClick=""AdicionarCNL(this)"">URB"
strDet = strDet & "</td>"
strDet = strDet & "<td><input type=text size=4 maxlength=4 class=text name=txtCNLPtaA >(A4)</td>"
strDet = strDet & "<td ><input type=text size=4 maxlength=4 class=text name=txtCNLPtaB >(A4)</td>"

''@davif
if strExigePortadora = "S" then
	strDet = strDet & "<td valign=top colspan=2> "
	strDet = strDet & "		<select name=cboPortadora>"
	strDet = strDet & "			<Option value=""""></Option>"
	strDet = strDet & "			<Option value=Constante>Constante</Option>"
	strDet = strDet & "			<Option value=PseudoControlada>PseudoControlada</Option>"
	strDet = strDet & "		</select> </td>"
Else
	strDet = strDet & "<td valign=top colspan=2> </td>"
End if
''/@davif

strDet = strDet & "</tr>"

strDet = strDet & "<tr class=clsSilver2>"
strDet = strDet & "	<td><font class=clsObrig>:: </font>Prop Modem</td>"
strDet = strDet & "	<td><font class=clsObrig>:: </font>Qtde Modem</td>"
if strPropAcesso <> "EBT" then
	strDet = strDet & "	<td><div id=""divFaturaLabel""><font class=clsObrig>:: </font>Fatura</div></td>"
	strDet = strDet & "	<td colspan=4>Obs</td>"
Else
	strDet = strDet & "	<td colspan=5>Obs</td>"
End if
strDet = strDet & "</tr>"
strDet = strDet & "<tr class=clsSilver>"
strDet = strDet & "	<td valign=top>"
strDet = strDet & "		<select name=cboPropModem>"
strDet = strDet & "			<Option value=""""></Option>"
strDet = strDet & "			<Option value=CLI >CLI</Option>"
if strPropAcesso = "EBT" then
	strDet = strDet & "			<Option value=EBT selected>EBT</Option>"
Else
	strDet = strDet & "			<Option value=EBT >EBT</Option>"
End if
strDet = strDet & "			<Option value=TER >TER</Option>"
strDet = strDet & "		</select>"
strDet = strDet & "	</td>"
strDet = strDet & "	<td valign=top><input type=text size=4 maxlength=2 class=text name=txtQtdeModem onKeyup=""ValidarTipo(this,0)"">(N2)</td>"

if strPropAcesso <> "EBT" then
	strDet = strDet & "	<td valign=top><div id=""divFatura"">"
	strDet = strDet & "		<input type=radio name=rdoFatura Index=0 value=S >SIM&nbsp;"
	strDet = strDet & "		<input type=radio name=rdoFatura Index=1 value=N >Não</div>"
	strDet = strDet & "	</td>"
	strDet = strDet & "	<td valign=top colspan=4>"
	strDet = strDet & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strDet = strDet & "	</td>"
	strDet = strDet & "</tr>"
Else
	strDet = strDet & "	<td valign=top colspan=5>"
	strDet = strDet & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strDet = strDet & "	</td>"
	strDet = strDet & "</tr>"
End if
strDet = strDet & "<tr class=clsSilver2>"
strDet = strDet & "<td colspan=7 align=right>"
strDet = strDet & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar"" onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strDet = strDet & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strDet = strDet & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"
strDet = strDet & "</td>"
strDet = strDet & "</tr>"
strDet = strDet & "</table>"

'NÃO DETERMINISTICO
strCamposNdet = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposNdet = strCamposNdet & "<th  width=100>&nbsp;Provedor</th>"
strCamposNdet = strCamposNdet & "<th  width=100 nowrap>&nbsp;Nº Acesso Pta Ebt</th>"
strCamposNdet = strCamposNdet & "<th  width=100>&nbsp;Tronco</th>"
strCamposNdet = strCamposNdet & "<th  width=100>&nbsp;Par</th>"
strCamposNdet = strCamposNdet & "<th  width=310>&nbsp;Nº Acesso Pta CLI</th></table>"

strNDet = ""
strNDet = strNDet & "<table border=0 cellspacing=1 cellpadding=1 width=760>"
strNDet = strNDet & "	<tr >"
strNDet = strNDet & "		<th colspan=5>&nbsp;•&nbsp;Informações da Facilidade</th>"
strNDet = strNDet & "	</tr>"
strNDet = strNDet & "	<tr class=clsSilver2>"
strNDet = strNDet & "		<td nowrap><font class=clsObrig>::</font>&nbsp;Cód.Prov</td>"
strNDet = strNDet & "		<td nowrap>&nbsp;Nº Acesso Pta Ebt</td>"
strNDet = strNDet & "		<td nowrap><font class=clsObrig>::</font>&nbsp;Tronco</td>"
strNDet = strNDet & "		<td nowrap><font class=clsObrig>::</font>&nbsp;Par</td>"
strNDet = strNDet & "		<td nowrap>Nº Acesso Pta CLI</td>"
strNDet = strNDet & "	</tr>"

strNDet = strNDet & "<tr class=clsSilver>"
strNDet = strNDet & "<td><span id=spnCodProv><select name=cboCodProv>"
strNDet = strNDet & "<option value=""""></option>"
strNDet = strNDet & "</select></span></td>"
strNDet = strNDet & "<td><input type=text class=text name=txtNroAcessoEbt maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}"" ></td>"
strNDet = strNDet & "<td><input type=text class=text name=txtTronco maxlength=20 size=10></td>"
strNDet = strNDet & "<td><input type=text class=text name=txtPar maxlength=9 size=10 onKeyUp=""if(this.value.length > 1) ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N  ></td>"
strNDet = strNDet & "<td><input type=text class=text name=txtNroAcessoCli maxlength=25 size=15></td>"
strNDet = strNDet & "</tr>"

strNDet = strNDet & "<tr class=clsSilver2>"
strNDet = strNDet & "	<td nowrap>CCTO Provedor</td>"
strNDet = strNDet & "	<td nowrap>IU / URB</td>"
strNDet = strNDet & "	<td nowrap>CNL Ponta A</td>"
strNDet = strNDet & "	<td nowrap>CNL Ponta B</td>"

''@davi
if strExigePortadora = "S" then
	strNDet = strNDet & "	<td nowrap colspan=1>Portadora </td>"
Else
	strNDet = strNDet & "	<td nowrap colspan=1> </td>"
End if

strNDet = strNDet & "</tr>"

strNDet = strNDet & "<tr class=clsSilver>"
strNDet = strNDet & "<td><input type=text size=8 maxlength=8 class=text name=txtCCTOPro></td>"
strNDet = strNDet & "<td >"

strNDet = strNDet & "	<input type=radio name=rdoUrbano value=I Index=0  onClick=""AdicionarCNL(this)"">IU&nbsp;"
strNDet = strNDet & "	<input type=radio name=rdoUrbano value=U Index=1 onClick=""AdicionarCNL(this)"">URB"
strNDet = strNDet & "</td>"
strNDet = strNDet & "<td><input type=text size=4 maxlength=4 class=text name=txtCNLPtaA></td>"
strNDet = strNDet & "<td><input type=text size=4 maxlength=4 class=text name=txtCNLPtaB></td>"

''@davif
if strExigePortadora = "S" then
	strNDet = strNDet & "<td nowrap colspan=2> "
	strNDet = strNDet & "		<select name=cboPortadora>"
	strNDet = strNDet & "			<Option value=""""></Option>"
	strNDet = strNDet & "			<Option value=Constante >Constante</Option>"
	strNDet = strNDet & "			<Option value=PseudoControlada >PseudoControlada</Option>"
	strNDet = strNDet & "		</select> </td>"
Else
	strNDet = strNDet & "<td nowrap colspan=2>  </td>"
End if
''/@davif

strNDet = strNDet & "</tr>"

strNDet = strNDet & "<tr class=clsSilver2>"
strNDet = strNDet & "	<td><font class=clsObrig>:: </font>Prop Modem</td>"
strNDet = strNDet & "	<td><font class=clsObrig>:: </font>Qtde Modem</td>"
if strPropAcesso <> "EBT" then
	strNDet = strNDet & "	<td><div id=""divFaturaLabel""><font class=clsObrig>:: </font>Fatura</div></td>"
	strNDet = strNDet & "	<td colspan=2>Obs</td>"
Else
	strNDet = strNDet & "	<td colspan=3>Obs</td>"
End if
strNDet = strNDet & "</tr>"
strNDet = strNDet & "<tr class=clsSilver>"
strNDet = strNDet & "	<td valign=top>"
strNDet = strNDet & "		<select name=cboPropModem>"
strNDet = strNDet & "			<Option value=""""></Option>"
strNDet = strNDet & "			<Option value=CLI >CLI</Option>"
if strPropAcesso = "EBT" then
	strNDet = strNDet & "			<Option value=EBT selected>EBT</Option>"
Else
	strNDet = strNDet & "			<Option value=EBT >EBT</Option>"
End if
strNDet = strNDet & "			<Option value=TER >TER</Option>"
strNDet = strNDet & "		</select>"
strNDet = strNDet & "	</td>"
strNDet = strNDet & "	<td valign=top><input type=text size=4 maxlength=4 class=text name=txtQtdeModem onKeyup=""ValidarTipo(this,0)""></td>"

if strPropAcesso <> "EBT" then
	strNDet = strNDet & "	<td valign=top><div id=""divFatura"">"
	strNDet = strNDet & "		<input type=radio name=rdoFatura Index=0 value=S >SIM&nbsp;"
	strNDet = strNDet & "		<input type=radio name=rdoFatura Index=1 value=N >Não</div>"
	strNDet = strNDet & "	</td>"
	strNDet = strNDet & "	<td valign=top colspan=4>"
	strNDet = strNDet & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strNDet = strNDet & "	</td>"
	strNDet = strNDet & "</tr>"
Else
	strNDet = strNDet & "	<td valign=top colspan=5>"
	strNDet = strNDet & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strNDet = strNDet & "	</td>"
	strNDet = strNDet & "</tr>"
End if

strNDet = strNDet & "<tr class=clsSilver2>"
strNDet = strNDet & "<td colspan=5 align=right>"
strNDet = strNDet & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar""  onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strNDet = strNDet & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strNDet = strNDet & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"
strNDet = strNDet & "</td>"
strNDet = strNDet & "</tr>"
strNDet = strNDet & "</table>"

'ADE
strCamposAde = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposAde = strCamposAde & "<th width=100>&nbsp;Nº Acesso</th>"
strCamposAde = strCamposAde & "<th width=100>&nbsp;Cabo</th>"
strCamposAde = strCamposAde & "<th width=100>&nbsp;Par</th>"
strCamposAde = strCamposAde & "<th nowrap width=100>&nbsp;PADE</th>"
strCamposAde = strCamposAde & "<th width=100>&nbsp;Derivação</th>"
strCamposAde = strCamposAde & "<th nowrap width=100>&nbsp;T. Cabo</th></table>"

strAde = strAde & "<table rules=groups border=0 cellspacing=1 cellpadding=1 bordercolorlight=#003388 bordercolordark=#ffffff width=760>"
strAde = strAde & "	<tr>"
strAde = strAde & "		<th colspan=6>&nbsp;•&nbsp;Informações da Facilidade</td>"
strAde = strAde & "	</tr>"
strAde = strAde & "	<tr class=clsSilver2>"
strAde = strAde & "		<td width=150>&nbsp;Nº Acesso</td>"
strAde = strAde & "		<td width=100><font class=clsObrig>::</font>&nbsp;Cabo</td>"
strAde = strAde & "		<td width=120><font class=clsObrig>::</font>&nbsp;Par</td>"
strAde = strAde & "		<td nowrap><font class=clsObrig></font>&nbsp;PADE</td>"
strAde = strAde & "		<td width=100><font class=clsObrig></font>&nbsp;Derivação</td>"
strAde = strAde & "		<td nowrap width=100><font class=clsObrig></font>&nbsp;T. Cabo</td>"
strAde = strAde & "	</tr>"
strAde = strAde & "<tr class=clsSilver>"
strAde = strAde & "<td width=150><input type=text class=text name=txtNroAcessoEbt	maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}""></td>"
strAde = strAde & "<td width=100><input type=text class=text name=txtCabo		maxlength=2	 size=5	 onKeyUp=""ValidarTipo(this,0)"" ></td>"
strAde = strAde & "<td width=120><input type=text class=text name=txtPar		maxlength=9  size=10 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N ></td>"
strAde = strAde & "<td><input type=text class=text name=txtCaixaEmenda  maxlength=3 size=5 ></td>"
strAde = strAde & "<td width=100><input type=text class=text name=txtLateral	maxlength=4  size=5></td>"
strAde = strAde & "<td width=100><select name=cboTipoCabo >"
strAde = strAde & "			<option value=""""></option>"
strAde = strAde & "			<option value=""H"" >H</option>"
strAde = strAde & "			<option value=""X"" >X</option>"
'strAde = strAde & "			<option value=XDSL >XDSL</option>"
strAde = strAde & "			<option value=""C"" >C</option>"
strAde = strAde & "		</select></td>"
strAde = strAde & "</tr>"

strAde = strAde & "<tr class=clsSilver2>"
strAde = strAde & "	<td><font class=clsObrig>:: </font>Prop Modem</td>"
strAde = strAde & "	<td><font class=clsObrig>:: </font>Qtde Modem</td>"
if strPropAcesso <> "EBT" then
	strAde = strAde & "	<td><div id=""divFaturaLabel""><font class=clsObrig>:: </font>Fatura</div></td>"
	strAde = strAde & "	<td colspan=3>Obs</td>"
Else
	strAde = strAde & "	<td colspan=4>Obs</td>"
End if
strAde = strAde & "</tr>"
strAde = strAde & "<tr class=clsSilver>"
strAde = strAde & "	<td valign=top>"
strAde = strAde & "		<select name=cboPropModem>"
strAde = strAde & "			<Option value=""""></Option>"
strAde = strAde & "			<Option value=CLI >CLI</Option>"
if strPropAcesso = "EBT" then
	strAde = strAde & "			<Option value=EBT selected>EBT</Option>"
Else
	strAde = strAde & "			<Option value=EBT >EBT</Option>"
End if
strAde = strAde & "			<Option value=TER >TER</Option>"
strAde = strAde & "		</select>"
strAde = strAde & "	</td>"
strAde = strAde & "	<td valign=top><input type=text size=4 maxlength=4 class=text name=txtQtdeModem onKeyup=""ValidarTipo(this,0)""></td>"

if strPropAcesso <> "EBT" then
	strAde = strAde & "	<td valign=top><div id=""divFatura"">"
	strAde = strAde & "		<input type=radio name=rdoFatura Index=0 value=S >SIM&nbsp;"
	strAde = strAde & "		<input type=radio name=rdoFatura Index=1 value=N >Não</div>"
	strAde = strAde & "	</td>"
	strAde = strAde & "	<td valign=top colspan=4>"
	strAde = strAde & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strAde = strAde & "	</td>"
	strAde = strAde & "</tr>"
Else
	strAde = strAde & "	<td valign=top colspan=5>"
	strAde = strAde & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strAde = strAde & "	</td>"
	strAde = strAde & "</tr>"
End if

strAde = strAde & "<tr class=clsSilver2>"
strAde = strAde & "<td colspan=6 align=right>"
strAde = strAde & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar""  onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strAde = strAde & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strAde = strAde & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"
strAde = strAde & "</td>"
strAde = strAde & "</tr>"
strAde = strAde & "</table>"

Dim strStatus
strStatus = ""

Dim strAcfObs
Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
objXmlDados.loadXml("<xDados/>")

Set objRS = db.Execute("CLA_SP_Sel_Facilidade " & dblPedId)
	if not objRS.eof then
		session("ss_Acf_NroAcessoPtaEbt") = objRS("Acf_NroAcessoPtaEbt")
	end if 
if intTipoProcesso = 4 then 'Cancelamento
	Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,"R")
Else
	Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,"A")
End if
Set objNode = objXmlDados.SelectNodes("//Facilidade")

intFac = objNode.length
if intFac > 0 then blnGravado = true
strXmlFac = FormatarXml(objXmlDados)

Function MontaXmlFacOrigem(objRSFacOrigem)

	Dim objNodeFac
	Dim objXMLFacOrigem

	Set objXMLFacOrigem = Server.CreateObject("Microsoft.XMLDOM")
	objXMLFacOrigem.loadXml("<xDados/>")

	if not objRSFacOrigem.Eof and not objRSFacOrigem.bof then

		While not objRSFacOrigem.Eof

			Set objNodeFac = objXMLFacOrigem.createNode("element", "Facilidade", "")
			objXMLFacOrigem.documentElement.appendChild (objNodeFac)

			For intIndex=0 to objRSFacOrigem.fields.count-1
				Call AddElemento(objXMLFacOrigem,objNodeFac,objRSFacOrigem.fields(intIndex).name,objRSFacOrigem.fields(intIndex).value)
			Next
			Call AddElemento(objXMLFacOrigem,objNodeFac,"intFacSel",0)
			objRSFacOrigem.MoveNext
		Wend
		objRSFacOrigem.MoveFirst
	End if
	Set MontaXmlFacOrigem = objXMLFacOrigem
End function
%>
<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
<SCRIPT LANGUAGE=javascript>

	function LimparVlanSwitch(IDrdo)
	{
		vlan = "desigRadioIP" + IDrdo;
		document.Form2.hdnIDrdoVlanSwitch.value = IDrdo ; 		
		
		for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
		 { 
				var elemento = document.Form2.elements[intIndex]; 
				if (elemento.name == vlan ){ 
						
						elemento.disabled=false; 
						elemento.value = "";
						elemento.focus();  
				} 
		} 	
		
	}
	
	function LimparDSLAM(IDrdo)
	{
		
		
		
		
		/**
		Cabo 		= "txtCabo" + document.Form2.hdnIDDSLAM.value;
		Lateral 	= "txtLateral" + document.Form2.hdnIDDSLAM.value;
		TipoCabo 	= "cboTipoCabo" + document.Form2.hdnIDDSLAM.value;
		CaixaEmenda	= "txtCaixaEmenda" + document.Form2.hdnIDDSLAM.value;
		Par 		= "txtPar" + document.Form2.hdnIDDSLAM.value;
		//ObsFac 		= "txtObsFac" + IDrdo;
		
		
		
		//alert(Cabo)
		//document.Form2.hdnIDDSLAM.value = IDrdo ;
		
		 		
		//txtNroAcessoEbt.value = ''
		//txtCabo1.value = ''
		//txtLateral.value = ''
		//cboTipoCabo.value = ''
		//txtCaixaEmenda.value = ''
		//txtPar.value = ''
		//txtObsFac.value = ''
		
		for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
		 { 
				var elemento = document.Form2.elements[intIndex]; 
				if (elemento.name == portaUplink ) { 
				
					//alert(Cabo);
					if (!elemento.checked){
								//alert(Lateral);
						//if (elemento.name == Cabo ){ 
								
								document.Form2.Cabo.disabled=false; 
								document.Form2.Cabo.value = "";
								//elemento.focus();  
						//} 
						//if (elemento.name == Lateral ){ 
								
								document.Form2.Lateral.disabled=false; 
								document.Form2.Lateral.value = "";
								//elemento.focus();  
						//} 
						//if (elemento.name == TipoCabo ){ 
								
								document.Form2.TipoCabo.disabled=false; 
								document.Form2.TipoCabo.value = "";
								//elemento.focus();  
						//} 
						//if (elemento.name == CaixaEmenda ){ 
								
								document.Form2.CaixaEmenda.disabled=false; 
								document.Form2.CaixaEmenda.value = "";
								//elemento.focus();  
						//} 
						//if (elemento.name == Par ){ 
								
								document.Form2.Par.disabled=false; 
								document.Form2.Par.value = "";
								//elemento.focus();  
						//} 
					}
				}
		} 
		**/
		
		
		
		Cabo 		= "txtCabo" + IDrdo;
		Lateral 	= "txtLateral" + IDrdo;
		TipoCabo 	= "cboTipoCabo" + IDrdo;
		CaixaEmenda	= "txtCaixaEmenda" + IDrdo;
		Par 		= "txtPar" + IDrdo;
		cboPropriedadeCabo	=	"cboPropriedadeCabo" + IDrdo;
		portaUplink = "rdoPortaUplinkIDDSLAM"  ;
		
		propriedade = "";
		//ObsFac 		= "txtObsFac" + IDrdo;
		
		
		document.Form2.hdnIDDSLAM.value = IDrdo ;
		
		 		
		//txtNroAcessoEbt.value = ''
		//txtCabo1.value = ''
		//txtLateral.value = ''
		//cboTipoCabo.value = ''
		//txtCaixaEmenda.value = ''
		//txtPar.value = ''
		//txtObsFac.value = ''
		
		//alert(document.Form2.txtCabo1.value);
		
		for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
		 { 
				var elemento = document.Form2.elements[intIndex]; 
				
				//if (elemento.name == portaUplink ) { 
				
					
					//if (!elemento.checked){
						if (elemento.name == Cabo ){ 
								//alert(elemento.name);
								elemento.readonly=false; 
								elemento.value = "";
								//elemento.focus();  
						} 
						if (elemento.name == Lateral ){ 
								
								elemento.disabled=false; 
								elemento.value = "";
								//elemento.focus();  
						} 
						if (elemento.name == TipoCabo ){ 
								
								elemento.disabled=false; 
								elemento.value = "";
								//elemento.focus();  
						} 
						if (elemento.name == CaixaEmenda ){ 
								
								elemento.disabled=false; 
								elemento.value = "";
								//elemento.focus();  
						} 
						if (elemento.name == Par ){ 
								
								elemento.disabled=false; 
								elemento.value = "";
								//elemento.focus();  
						} 
						
						if (elemento.name == cboPropriedadeCabo ){ 
								
								propriedade = elemento.value;
								//elemento.focus();  
						} 
						
						
					//}
				//}
		} 		
		
		
		
		//alert(propriedade);
		
		//alert(IDrdo)
		//Limpar posições do array
		objAryFac[0] = new Array("","")
		objAryFac[1] = new Array("","")
		objAryFac[2] = new Array("","")
		objAryFac[3] = new Array("","")
		objAryFac[4] = new Array("","")
		objAryFac[5] = new Array("","")
		
		if (propriedade == "2"){
			PosicoesLivreDSLAM('L')
		}
		
		
	}
	
	function PosicoesLivreDSLAM(strPagina)
	{
		var TipoPlataforma
		with (document.forms[0])
		{
			if (!ValidarCampos(cboLocalInstala,"Local de Entrega")) return
			if (!ValidarCampos(cboDistLocalInstala,"Distribuidor")) return
			if (!ValidarCampos(cboRede,"Rede")) return
			if (!ValidarCampos(cboProvedor,"Provedor")) return
			
			switch (strPagina)
			{
				case "L": //livres
					try{
						objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=L&cboLocalInstala='+cboLocalInstala.value+'&cboDistLocalInstala='+cboDistLocalInstala.value+'&cboProvedor='+cboProvedor.value+'&strPagina='+strPagina+'&cboSistema='+cboRede.value+'&cboPlataforma='+cboPlataforma.value ,objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
					}
					catch(e){
						objAryFacRet = window.showModalDialog('ConsultarFacilidades.asp?hdnAcao=Posicoes&strStsFac=L&cboLocalInstala='+cboLocalInstala.value+'&cboDistLocalInstala='+cboDistLocalInstala.value+'&cboProvedor='+cboProvedor.value+'&strPagina='+strPagina+'&cboSistema='+cboRede.value, objAryFac,"dialogHeight: 450px; dialogWidth: 570px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
					}
					break
				
			}

			Cabo 		= "txtCabo" + document.Form2.hdnIDDSLAM.value ;
			Par 		= "txtPar" + document.Form2.hdnIDDSLAM.value ;
			Lateral 	= "txtLateral" + document.Form2.hdnIDDSLAM.value ;
			TipoCabo 	= "cboTipoCabo" + document.Form2.hdnIDDSLAM.value ;
			CaixaEmenda = "txtCaixaEmenda" + document.Form2.hdnIDDSLAM.value ;
			FACID		= "hdnFACID" + document.Form2.hdnIDDSLAM.value ;
			
			
			//alert(objAryFacRet.length)
			//alert(Cabo)
			
			try{
				
					for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
					 { 
							
							var elemento = document.Form2.elements[intIndex]; 
							if (elemento.name == Cabo ){ 
								//alert(Cabo)
								//alert(objAryFac[intIndex2][0])
								//elemento.readonly=true; 
								if (objAryFac[0][0] != "" && objAryFac[0][0] != "undefined")
								{
									elemento.value = objAryFac[0][0];
									elemento.focus();  
								}
							} 
							if (elemento.name == Lateral ){ 
									
								if (objAryFac[0][2] != "" )
								{
									//alert(objAryFac.length)
									//alert(objAryFac[0][2])
									
									elemento.disabled=false; 
									elemento.value = objAryFac[0][2];
									
									if (elemento.value == "undefined") {
										elemento.value = "";
									}
								}
								
								//elemento.focus();  
							} 
							if (elemento.name == TipoCabo ){ 
							
								
									
								if (objAryFac[0][3] != "" )
								{	
									
									elemento.disabled=false; 
									elemento.value = objAryFac[0][3];
								}
								
								if (objAryFac[0][3] == ""){
									//alert(objAryFac[0][3]);
									elemento.disabled=false; 
									elemento.value = "";
								}
									//elemento.focus();  
							} 
							if (elemento.name == CaixaEmenda ){ 
								if (objAryFac[0][4] != "" && objAryFac[0][4] != "undefined")
								{	
									elemento.disabled=false; 
									elemento.value = objAryFac[0][4];
									//elemento.focus();  
									
									if (elemento.value == "undefined"){
										elemento.value = "";
									}
								}
							} 
							if (elemento.name == Par ){ 
								if (objAryFac[0][1] != "" && objAryFac[0][1] != "undefined")
								{	
									elemento.disabled=false; 
									elemento.value = objAryFac[0][1];
									//elemento.focus();  
								}
							} 
							
							if (elemento.name == FACID ){ 
								if (objAryFac[0][5] != "" )
								{	
									elemento.value = objAryFac[0][5];
									//elemento.focus();  
								}
							} 
							
							
					} 
					/**
					if (objAryFac[0][5] != "" && objAryFac[0][5] != "undefined")
					{
							//alert(objAryFac[0][5])
							document.Form2.hdnFACID.value = objAryFac[0][5];
					}
					**/
				
			}
			catch(e){}	
		}	
	}
	
	
	function LimparUplink(IDrdo , equi)
	{
		uplink = "uplink" + equi + IDrdo;
		
		if ( equi == "METRO" )
		{
			document.Form2.hdnIDUplinkMETRO.value = IDrdo ;
		}
		if ( equi == "EDD" )
		{
			document.Form2.hdnIDUplinkEDD.value = IDrdo ;
		}
		if ( equi == "Interconexao" )
		{
			document.Form2.hdnIDUplinkInterconexao.value = IDrdo ;
		}
		 		
		
		for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
		 { 
				var elemento = document.Form2.elements[intIndex]; 
				if (elemento.name == uplink ){ 
						
						elemento.disabled=false; 
						elemento.value = "";
						elemento.focus();  
				} 
		} 	
		
	}
	
	
	
	function GravarPortaUplink (equi)
	{
		
		with (document.Form2)
		{
			campo = "";
			portaUplink = "";
			
			if ( equi == "METRO" )
			{
				campo = hdnIDUplinkMETRO.value;
				
				hdnUplinkMETRO.value = document.getElementById("uplink" + equi + campo).value;
			}
			
			if ( equi == "EDD" )
			{
				campo = hdnIDUplinkEDD.value;
				hdnUplinkEDD.value = document.getElementById("uplink" + equi + campo).value;
			}
			
			if ( equi == "Interconexao" )
			{
				campo = hdnIDUplinkInterconexao.value;
				hdnUplinkInterconexao.value = document.getElementById("uplink" + equi + campo).value;
				
			}	
			
			if ( equi == "DSLAM" )
			{
				if ( IsEmpty(txtDesigONT.value) ){
					alert("Favor informar a Designação ONT");
					return;
				}
				campo = hdnIDDSLAM.value;
				hdnUplinkDSLAM.value = hdnOntDesig.value;
				hdnRede.value = "3";
				
			}	

			if ( IsEmpty(campo) ){
				alert("Favor selecionar uma porta");
				return;
			}
			
			portaUplink = "rdoPortaUplinkID" + equi ;
			
						
			for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
			 { 
					var elemento = document.Form2.elements[intIndex]; 
					if (elemento.name == portaUplink ) { //}"rdoPortaUplinkID" ){ 
					
						if (elemento.checked){
						
							if ( equi == "METRO" )
							{
								hdnrdoPortaUplinkIDMETRO.value = elemento.value ;
							}
							if ( equi == "EDD" )
							{
								hdnrdoPortaUplinkIDEDD.value = elemento.value ;
							}
							if ( equi == "Interconexao" )
							{
								hdnrdoPortaUplinkIDInterconexao.value = elemento.value ;
							}
							
							if ( equi == "DSLAM" )
							{
								hdnrdoPortaDSLAM.value = elemento.value ;
							}
							
							//alert(elemento.value)
						}
					} 
			} 	
			
			hdnEqui.value = equi		
			target = "IFrmProcesso"
			
			action = "AlocarPortaUplink.asp"
			submit()
			
		}
	}
	function GravarUplinkDSLAM ()
	{
		with (document.Form2)
		{
			hdnOntDesig.value = txtDesigONT.value;
			//alert(hdnUplinkDSLAM.value);
		}
	}
	
	function GravarPortaDSLAM (equi)
	{
		
		with (document.Form2)
		{
			//alert(hdnUplinkDSLAM.value);
			//return;
			
			campo = "";
			portaUplink = "";
			
			portaUplink = "rdoPortaUplinkID" + equi ;
			 var blnAchou = false
			for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
			{ 
			
				var elemento = document.Form2.elements[intIndex]; 
				if (elemento.name == portaUplink ) {
				
					if (elemento.checked){
						blnAchou = true
					}
					
					
					
				/**
					for (var intIndex2=0;intIndex2<elemento.length;intIndex2++)
					{
						if (elemento[intIndex2].checked)
						{
							blnAchou = true
						}
					}
					if (!blnAchou)
					{
						alert(" DSLAM é um campo obrigatório.")
					}
					**/
				}
				
				//if (!ValidarCampos(rdoPortaUplinkIDDSLAM,"Favor informar o DSLAM")) return false
				
			}
			
			if (!blnAchou)
			{
				alert(" DSLAM é um campo obrigatório.")
				return
			}
			
			if (ConfirmarDSLAM())
			{
				if ( equi == "DSLAM" )
				{
					
					
					<%
						
						'strSQL = "  select top 1  cla_ont.ont_desig as ont_desig from cla_ontvlan  inner join cla_ontporta		on cla_ontvlan.ontporta_id = cla_ontporta.ontporta_id inner join cla_ont				on cla_ontporta.ont_id = cla_ont.ont_id  where acl_idacessologico =  " & strIdLogico 
						'set objRS = db.execute (strSQL)
						
						'While Not objRS.eof
						
						
							
						%>
							
							
							//hdnUplinkDSLAM.value = '<%= objRS("ont_desig") %>'
							//strUplink = hdnUplinkDSLAM.value;
							//strUplink = strUplink.replace(" " , "");
							
						<%
						
							 ' objRS.MoveNext
						'Wend

					%>		
					//alert(hdnOntDesig.value);
					//return;
					if ( IsEmpty(hdnOntDesig.value) ){
						alert("Favor informar a Designação ONT");
						return;
					}
					//return;
					//alert(hdnUplinkDSLAM.value);
					//return;
					//campo = hdnIDDSLAM.value;
					//hdnUplinkDSLAM.value = hdnOntDesig.value;
					hdnRede.value = "3";
					
				}	
				
				strXML = "<root>"
				//alert(strXML);
				for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
				{ 
						var elemento = document.Form2.elements[intIndex]; 
						if (elemento.name == portaUplink ) { //}"rdoPortaUplinkID" ){ 
						
							if (elemento.checked){
							
								if ( equi == "DSLAM" )
								{
									strDados = elemento.value;
									arrayCampos = strDados.split("&&");
									strID = arrayCampos[0]
									portaID = arrayCampos[1]
									//alert(document.getElementById("hdnFacID2").value)
									//alert(elemento.value)
									//alert(strID)
									//alert(portaID)
									facID = "hdnFACID" + strID
									PropriedadeCabo = "cboPropriedadeCabo" + strID
									facValue = document.getElementById(facID).value
									PropCaboValue = document.getElementById(PropriedadeCabo).value
									
									strXML = strXML + "<PortaDslam>"
									
									strXML = strXML + "<hdnFacID>" +  facValue + "</hdnFacID>"
									strXML = strXML + "<hdnrdoPortaDSLAM>" + portaID + "</hdnrdoPortaDSLAM>" 
									strXML = strXML + "<hdnUplinkDSLAM>" + hdnOntDesig.value + "</hdnUplinkDSLAM>"
									strXML = strXML + "<hdncboPropriedadeCabo>" + PropCaboValue + "</hdncboPropriedadeCabo>"
									strXML = strXML + "<hdnRede>3</hdnRede>"
									
									strXML = strXML + "</PortaDslam>"
									
									
									//hdnrdoPortaDSLAM.value = elemento.value ;
								}
								
							}
						}
				}
					
				strXML = strXML + "</root>"
				
				hdnXMLDSLAM.value = strXML;
				
				//alert(strXML);
				//return;
				
				hdnEqui.value = equi		
				target = "IFrmProcesso"
				
				action = "AlocarPortaUplink.asp"
				submit()
			}
		}
	}
	
	
	function CheckPE(IDrdo)
	{
			
		
		CVLAN_ETHERNET = "CVLAN_ETHERNET" + IDrdo;
		SVLAN_ETHERNET = "SVLAN_ETHERNET" + IDrdo;
		VLAN_PortaOLT = "VLAN_PortaOLT" + IDrdo;
		document.Form2.hdnIDPE.value = IDrdo ; 		
		
		for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
		 { 
				var elemento = document.Form2.elements[intIndex]; 
				if (elemento.name == CVLAN_ETHERNET ){ 
						
						elemento.disabled=false; 
						elemento.value = "";
						elemento.focus();  
				} 
				if (elemento.name == SVLAN_ETHERNET ){ 
						
						elemento.disabled=false; 
						elemento.value = "";
						//elemento.focus();  
				} 
				if (elemento.name == VLAN_PortaOLT ){ 
						
						elemento.disabled=false; 
						elemento.value = "";
						//elemento.focus();  
				} 
		} 	
						
	}
	
	
	function GravarPortaPE ()
	{
		
		with (document.Form2)
		{
			
			campo = hdnIDPE.value;
			
			
			if ( IsEmpty(campo) ){
				alert("Favor selecionar uma porta");
				return;
			}
			
			
			hdnCVLAN_ETHERNET.value = document.getElementById("CVLAN_ETHERNET" + campo).value;
			
			hdnSVLAN_ETHERNET.value = document.getElementById("SVLAN_ETHERNET" + campo).value;
			
			hdnVLAN_PortaOLT.value = document.getElementById("VLAN_PortaOLT" + campo).value;
			
			//hdnUplink.value = document.getElementById("uplink" + campo).value;
			
			//alert(hdnCVLAN_ETHERNET.value);
			//alert(hdnSVLAN_ETHERNET.value);
			
			if (hdnCVLAN_ETHERNET.value==null || Trim(hdnCVLAN_ETHERNET.value)==""){
				alert("Favor informar a VLAN do cliente");
				return;
			}
				
			if (hdnSVLAN_ETHERNET.value==null || Trim(hdnSVLAN_ETHERNET.value)==""){
				alert("Favor informar a VLAN do serviço");
				return;
			}
				
			if (hdnVLAN_PortaOLT.value==null || Trim(hdnVLAN_PortaOLT.value)==""){
				alert("Favor informar a Porta PE");
				return;
			}
									
			
			for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) 
			 { 
					var elemento = document.Form2.elements[intIndex]; 
					if (elemento.name == "rdoPortaPEID" ){ 
							if (elemento.checked){
								hdnrdoPortaPEID.value = elemento.value ;
								//alert(hdnrdoPortaPEID.value)
							}
					} 
			} 	
			
						
			target = "IFrmProcesso"
			action = "AlocarPortaPE.asp"
			submit()
			
		}
	}


	//function Trim(str){return str.replace(/^\s+|\s+$/g,"");}

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
			
			//alert(document.getElementById("vlanSwitch" + campo).value);
			//alert(document.getElementById("portaPE" + campo).value);
			//alert(document.getElementById("pe" + campo).value);
			//alert(document.getElementById("svlan" + campo).value);
			//alert(document.getElementById("desigRadioIP" + campo).value);
			
			hdnvlanSwitch.value = document.getElementById("vlanSwitch" + campo).value;
			hdnportaoltSwitch.value = document.getElementById("portaPE" + campo).value;
			hdnpeSwitch.value = document.getElementById("pe" + campo).value;
			hdnSvlanSwitch.value = document.getElementById("svlan" + campo).value;
			hdndesigRadioIP.value = document.getElementById("desigRadioIP" + campo).value;
			
			hdnportaSwitchLadoMetro.value = document.getElementById("portaSwitchLadoMetro").value;
			 
			//alert(hdndisabledSvlan.value)
			//alert(hdndisabledCvlan.value)
			if (hdndisabledSvlan.value == "False"){	
						
				if (hdnSvlanSwitch.value==null || Trim(hdnSvlanSwitch.value)==""){
					alert("Favor informar a SVLAN");
					return;
				}
			}
			
				//alert(hdndisabledCvlan.value)
			if (hdnportaSwitchLadoMetro.value == "False"){	
						
				if (hdnportaSwitchLadoMetro.value==null || Trim(hdnportaSwitchLadoMetro.value)==""){
					alert("Favor informar a porta do Switch lado Metro");
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
	
		

function ResgatarSwitchRadioIp() {
	Form2.hdnSwitchIP.value = document.getElementById('cboSwtchRec').value
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboSwtchRec').value
		param2 = document.getElementById('hdnIdLog').value
		param3 = document.getElementById('hdnTipoProcesso').value
		
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoSwitchRadioIp;
        xmlhttp.open("POST", "../Ajax/AJX_ResgatarSwitch.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><param2>"+param2+"</param2><param3>"+param3+"</param3><texto>S</texto></dados>"
        xmlhttp.send(strXML);
		
		
		//xmlDoc.loadXML(strXML);
		//alert(strXML);
    }
}

function processadorMudancaEstadoSwitchRadioIp () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			//alert( "xml: " + xmlhttp ); 
			document.getElementById("spnSwitchRadioIP").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}

/**
eyc
**/
function ResgatarEthernetInterconexao() {
	Form2.hdnSwitchInterconexao.value = document.getElementById('cboSwitchInterconexao').value
    try { 
        xmlhttpEthernetInterconexao = new ActiveXObject("Msxml2.XMLHTTP"); 
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
    } catch (e) { 
        try { 
            xmlhttpEthernetInterconexao = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttpEthernetInterconexao = false; 
        } 
    } 

    if  (!xmlhttpEthernetInterconexao && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttpEthernetInterconexao = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttpEthernetInterconexao = false ; 
        } 
    }

    if (xmlhttpEthernetInterconexao) {
		param = document.getElementById('cboSwitchInterconexao').value
		param2 = document.getElementById('hdnIdLog').value
		param3 = document.getElementById('hdnTipoProcesso').value
		
		if (param==0){param=0}
        xmlhttpEthernetInterconexao.onreadystatechange = processadorMudancaEstadoEthernetInterconexao;
        xmlhttpEthernetInterconexao.open("POST", "../Ajax/AJX_ResgatarEthernetInterconexao.asp");
        xmlhttpEthernetInterconexao.setRequestHeader('Content-Type','text/xml');
        xmlhttpEthernetInterconexao.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><param2>"+param2+"</param2><param3>"+param3+"</param3><param4>Interconexao</param4><texto>S</texto></dados>"
        xmlhttpEthernetInterconexao.send(strXML);
		
		
		//xmlDoc.loadXML(strXML);
		//alert(strXML);
    }
}

function processadorMudancaEstadoEthernetInterconexao () {
    if ( xmlhttpEthernetInterconexao.readyState == 4) { // Completo 
        if ( xmlhttpEthernetInterconexao.status == 200) { // resposta do servidor OK 
			//alert( "xml: " + xmlhttp ); 
			document.getElementById("spnEthernetInterconexao").innerHTML = xmlhttpEthernetInterconexao.responseText;
        } else { 
            alert( "Erro: " + xmlhttpEthernetInterconexao.statusText ); 
			return 
        } 
    }
}

function ResgatarEthernetMetro() {
	Form2.hdnSwitchMetro.value = document.getElementById('cboSwitchMetro').value
    try { 
        xmlhttpEthernetMetro = new ActiveXObject("Msxml2.XMLHTTP"); 
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
    } catch (e) { 
        try { 
            xmlhttpEthernetMetro = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttpEthernetMetro = false; 
        } 
    } 

    if  (!xmlhttpEthernetMetro && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttpEthernetMetro = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttpEthernetMetro = false ; 
        } 
    }

    if (xmlhttpEthernetMetro) {
		param = document.getElementById('cboSwitchMetro').value
		param2 = document.getElementById('hdnIdLog').value
		param3 = document.getElementById('hdnTipoProcesso').value
		
		if (param==0){param=0}
        xmlhttpEthernetMetro.onreadystatechange = processadorMudancaEstadoEthernetMetro;
        xmlhttpEthernetMetro.open("POST", "../Ajax/AJX_ResgatarEthernetInterconexao.asp");
        xmlhttpEthernetMetro.setRequestHeader('Content-Type','text/xml');
        xmlhttpEthernetMetro.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><param2>"+param2+"</param2><param3>"+param3+"</param3><param4>METRO</param4><texto>S</texto></dados>"
        xmlhttpEthernetMetro.send(strXML);
		
		
		//xmlDoc.loadXML(strXML);
		//alert(strXML);
    }
}

function processadorMudancaEstadoEthernetMetro () {
    if ( xmlhttpEthernetMetro.readyState == 4) { // Completo 
        if ( xmlhttpEthernetMetro.status == 200) { // resposta do servidor OK 
			//alert( "xml: " + xmlhttp ); 
			document.getElementById("spnEthernetMetro").innerHTML = xmlhttpEthernetMetro.responseText;
        } else { 
            alert( "Erro: " + xmlhttpEthernetMetro.statusText ); 
			return 
        } 
    }
}

function ResgatarDslam() {
	Form2.hdnSwitchMetro.value = document.getElementById('cboSwitchDSLAM').value
    try { 
        xmlhttpDSLAM = new ActiveXObject("Msxml2.XMLHTTP"); 
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
    } catch (e) { 
        try { 
            xmlhttpDSLAM = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttpDSLAM = false; 
        } 
    } 

    if  (!xmlhttpDSLAM && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttpDSLAM = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttpDSLAM = false ; 
        } 
    }

    if (xmlhttpDSLAM) {
		param = document.getElementById('cboSwitchDSLAM').value
		param2 = document.getElementById('hdnIdLog').value
		param3 = document.getElementById('hdnTipoProcesso').value
		param5 = document.getElementById('txtDesigONT').value
		
		if (param==0){param=0}
        xmlhttpDSLAM.onreadystatechange = processadorMudancaEstadoDSLAM;
        xmlhttpDSLAM.open("POST", "../Ajax/AJX_ResgatarDSLAM.asp");
        xmlhttpDSLAM.setRequestHeader('Content-Type','text/xml');
        xmlhttpDSLAM.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><param2>"+param2+"</param2><param3>"+param3+"</param3><param4>DSLAM</param4><param5>"+param5+"</param5><texto>S</texto></dados>"
        xmlhttpDSLAM.send(strXML);
		
		
		//xmlDoc.loadXML(strXML);
		//alert(strXML);
    }
}

function processadorMudancaEstadoDSLAM () {
    if ( xmlhttpDSLAM.readyState == 4) { // Completo 
        if ( xmlhttpDSLAM.status == 200) { // resposta do servidor OK 
			//alert( "xml: " + xmlhttp ); 
			document.getElementById("spnDSLAM").innerHTML = xmlhttpDSLAM.responseText;
        } else { 
            alert( "Erro: " + xmlhttpDSLAM.statusText ); 
			return 
        } 
    }
}






function ResgatarEthernetEDD() {
	Form2.hdnSwitchEDD.value = document.getElementById('cboSwitchEDD').value
    try { 
        xmlhttpEthernetEDD = new ActiveXObject("Msxml2.XMLHTTP"); 
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
    } catch (e) { 
        try { 
            xmlhttpEthernetEDD = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttpEthernetEDD = false; 
        } 
    } 

    if  (!xmlhttpEthernetEDD && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttpEthernetEDD = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttpEthernetEDD = false ; 
        } 
    }

    if (xmlhttpEthernetEDD) {
		param = document.getElementById('cboSwitchEDD').value
		param2 = document.getElementById('hdnIdLog').value
		param3 = document.getElementById('hdnTipoProcesso').value
		
		if (param==0){param=0}
        xmlhttpEthernetEDD.onreadystatechange = processadorMudancaEstadoEthernetEDD;
        xmlhttpEthernetEDD.open("POST", "../Ajax/AJX_ResgatarEthernetInterconexao.asp");
        xmlhttpEthernetEDD.setRequestHeader('Content-Type','text/xml');
        xmlhttpEthernetEDD.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><param2>"+param2+"</param2><param3>"+param3+"</param3><param4>EDD</param4><texto>S</texto></dados>"
        xmlhttpEthernetEDD.send(strXML);
		
		
		//xmlDoc.loadXML(strXML);
		//alert(strXML);
    }
}

function processadorMudancaEstadoEthernetEDD () {
    if ( xmlhttpEthernetEDD.readyState == 4) { // Completo 
        if ( xmlhttpEthernetEDD.status == 200) { // resposta do servidor OK 
			//alert( "xml: " + xmlhttp ); 
			document.getElementById("spnEthernetEDD").innerHTML = xmlhttpEthernetEDD.responseText;
        } else { 
            alert( "Erro: " + xmlhttpEthernetEDD.statusText ); 
			return 
        } 
    }
}


function ResgatarEthernetPE() {
	Form2.hdnSwitchPE.value = document.getElementById('cboSwitchPE').value
    try { 
        xmlhttpEthernetPE = new ActiveXObject("Msxml2.XMLHTTP"); 
		var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
    } catch (e) { 
        try { 
            xmlhttpEthernetPE = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttpEthernetPE = false; 
        } 
    } 

    if  (!xmlhttpEthernetPE && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttpEthernetPE = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttpEthernetPE = false ; 
        } 
    }

    if (xmlhttpEthernetPE) {
		param = document.getElementById('cboSwitchPE').value
		param2 = document.getElementById('hdnIdLog').value
		param3 = document.getElementById('hdnTipoProcesso').value
		
		if (param==0){param=0}
        xmlhttpEthernetPE.onreadystatechange = processadorMudancaEstadoEthernetPE;
        xmlhttpEthernetPE.open("POST", "../Ajax/AJX_ResgatarEthernetPE.asp");
        xmlhttpEthernetPE.setRequestHeader('Content-Type','text/xml');
        xmlhttpEthernetPE.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><param2>"+param2+"</param2><param3>"+param3+"</param3><param4>PE</param4><texto>S</texto></dados>"
        xmlhttpEthernetPE.send(strXML);
		
		
		//xmlDoc.loadXML(strXML);
		//alert(strXML);
    }
}

function processadorMudancaEstadoEthernetPE () {
    if ( xmlhttpEthernetPE.readyState == 4) { // Completo 
        if ( xmlhttpEthernetPE.status == 200) { // resposta do servidor OK 
			//alert( "xml: " + xmlhttp ); 
			document.getElementById("spnEthernetPE").innerHTML = xmlhttpEthernetPE.responseText;
        } else { 
            alert( "Erro: " + xmlhttpEthernetPE.statusText ); 
			return 
        } 
    }
}

/**
 EYC
**/

/*----------------------------------------------------------------------------
Formatação para qualquer mascara
-----------------------------------------------------------------------------*/
function formatarIP(src){
  mask = "###.###.###.###"
  var i = src.value.length;
  var saida = mask.substring(0,1);
  var texto = mask.substring(i)
	if (texto.substring(0,1) != saida)
  {
    src.value += texto.substring(0,1);
  }
}


<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
var objXmlReturn = new ActiveXObject("Microsoft.XMLDOM")

var intIndice = 0
var objAryObjs = new Array()
var objAryFac = new Array()
var objAryFacRet

<%if strPropAcesso <> "EBT" then%>
	//alert("Proprietario Não EBT")
	<% If strExigePortadora = "S" then %>
		//alert("Proprietario Não EBT - Portadora = S")
		objAryObjs[1] = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "cboPortadora", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
		objAryObjs[2] = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "cboPortadora", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
	<% Else %>
		objAryObjs[1] = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB",  "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
		objAryObjs[2] = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")

	<%End if %>
	objAryObjs[3] = new Array("txtNroAcessoEbt","txtCabo","txtPar","txtLateral","cboTipoCabo","txtCaixaEmenda","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
	//objAryObjs[4] = new Array("txtDesignacao", "txtVlan","cboRede")
	objAryObjs[4] = new Array("txtNroAcessoEbt", "txtVlan","cboRede")
	//objAryObjs[5] = new Array("txtDesignacao", "txtVlan", "txtPE" , "txtPorta" , "txtLink","cboRede")
	objAryObjs[5] = new Array("txtNroAcessoEbt", "txtVlan", "txtPE" , "txtPorta" , "txtLink","cboRede")
	objAryObjs[13] = new Array("txtNroAcessoEbt", "txtVlan", "txtPE" , "txtPorta" , "txtSvlan","cboRede","cboTipoPorta")
	//Bsod Light
	objAryObjs[10] = new Array("txtNroAcessoEbt","cboRede","txtIP","txtGateway","txtMascara")
	
	objAryObjs[11] = new Array("txtNroAcessoEbt", "txtVlan_FO", "txtPE_FO" , "txtPorta_FO" , "txtSvlan_FO", "cboRede" , "hdnOrigem" , "hdnId_AcessoAsms")
	
	objAryObjs[12] = new Array("txtSWITCH")
<%Else%>
	//alert("Proprietario EBT")
	objAryObjs[1] = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","txtObsFac")
	objAryObjs[2] = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB",  "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","txtObsFac")
	objAryObjs[3] = new Array("txtNroAcessoEbt","txtCabo","txtPar","txtLateral","cboTipoCabo","txtCaixaEmenda","cboRede","Ped_Id","cboPropModem","txtQtdeModem","txtObsFac")
	
	
	//objAryObjs[4] = new Array("txtDesignacao", "txtVlan","cboRede")
	objAryObjs[4] = new Array("txtNroAcessoEbt", "txtVlan","cboRede")
	//objAryObjs[5] = new Array("txtDesignacao", "txtVlan", "txtPE" , "txtPorta" , "txtLink","cboRede")
	objAryObjs[5] = new Array("txtNroAcessoEbt", "txtVlan", "txtPE" , "txtPorta" , "txtLink","cboRede")
	objAryObjs[13] = new Array("txtNroAcessoEbt", "txtVlan", "txtPE" , "txtPorta" , "txtSvlan","cboRede","cboTipoPorta")
	//Bsod Light
	objAryObjs[10] = new Array("txtNroAcessoEbt","cboRede","txtIP","txtGateway","txtMascara")
	
	objAryObjs[11] = new Array("txtNroAcessoEbt", "txtVlan_FO", "txtPE_FO" , "txtPorta_FO" , "txtSvlan_FO","cboRede" , "hdnOrigem" , "hdnId_AcessoAsms")

	objAryObjs[12] = new Array("txtSWITCH")

<%End if%>

objAryFac[0] = new Array("","")

//JCARTUS
function LimpaDistribuidor()
{
	spnDistLocalInstala.innerHTML="<select name='cboDistLocalInstala' style='width:200px'><option value=''></option></select>"
}

function RetornaDistribuidor()
{
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var strXML
	strXML = "<root>"
	strXML = strXML + "<disid>" +  document.forms[0].cboLocalInstala.value + "</disid>"
	strXML = strXML + "</root>"
	xmlDoc.loadXML(strXML);
	xmlhttp.Open("POST","RetornaDistribuidor.asp" , false);
	xmlhttp.Send(xmlDoc.xml);
	strXML = xmlhttp.responseText;
	spnDistLocalInstala.innerHTML = strXML
}
//FIM JCARTUS

function CarregarLista()
{
	objXmlGeral.onreadystatechange = CheckStateXml;
	objXmlGeral.resolveExternals = false;
	if (parseInt(<%=intFac%>) != 0){
		objXmlGeral.loadXML("<%=strXmlFac%>")
	}else{
		var objXmlRoot = objXmlGeral.createNode("element","xDados","")
		
		objXmlGeral.appendChild (objXmlRoot)
	}
}
//Verifica se o Xml já esta carregado
function CheckStateXml()
{
  var state = objXmlGeral.readyState;

  if (state == 4)
  {
    var err = objXmlGeral.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    }
    else
    {
		AtualizarListaFac()
	}
  }
}


function CarregarDoc()
{
	window.name = "Facilidade.asp"
	document.onreadystatechange = CheckStateDoc;
	document.resolveExternals = false;
}

function CheckStateDoc()
{
	
  var state = document.readyState;

  if (state == "complete")
  {
	ResgatarInfoRede()
	PopularCamposFac()
	//alert(document.forms[0].cboRede.value)
	//alert('<%=strPla%>')
	RetornaCboPlataforma(document.forms[0].cboRede.value,<%=strPla%>)
	
  }
}


function VoltarOrigem()
{
	var strPagina = new String("<%=Request.Form("hdnPaginaOrig")%>")
	with (document.forms[0])
	{
		target = self.name
		if (strPagina == "" || strPagina.toUpperCase().indexOf("FACILIDADE.ASP") != -1)
		{
			action = "facilidade_main.asp"

		}else
		{
			action = strPagina
		}
		submit()
	}
}

function RetornaCboPlataforma(sisID, PlaID)
{
	if (sisID == '') {
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		return
	}
	if (sisID == 4 || sisID == 5 || sisID == 13) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		return
	}
	if ( sisID == 10 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		return
	}
	if ( sisID == 11 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		return
	}
	
	if ( sisID == 12 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		Form2.btnConsEstoque.disabled = true
		Form2.btnConsIdFis.disabled = true
		Form2.btnLimparEstoque.disabled = true
		Form2.btnOK.disabled = true
		
		return
	}
	
	if ( sisID == 14 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		Form2.btnConsEstoque.disabled = true
		Form2.btnConsIdFis.disabled = true
		Form2.btnLimparEstoque.disabled = true
		Form2.btnOK.disabled = true
		
		return
	}
	
	if (sisID == 2 || sisID == 3) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = false
		Form2.ocupados.disabled = false
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		return
	}
	if (sisID == 1)
	{
		Form2.btnRedeDet.disabled = false
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true	
		//Form2.btnEmailPro.disabled = false
	}
	
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var strXML

	strXML = "<root>"
	strXML = strXML + "<plaid>" +  PlaID + "</plaid>"
	strXML = strXML + "<funcao></funcao>"
	strXML = strXML + "</root>"

	xmlDoc.loadXML(strXML);
	xmlhttp.Open("POST","RetornaPlataforma.asp" , false);
	xmlhttp.Send(xmlDoc.xml);

	strXML = xmlhttp.responseText;
	spnPlataforma.innerHTML = strXML
	if(PlaID == 0){
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true	
	}else{
	
		Form2.consultar.disabled = false
		Form2.ocupados.disabled = false
		
	}
	if (sisID == '') {
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		return
	}
	if (sisID == 4 || sisID == 5 || sisID == 13) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		return
	}
	if ( sisID == 10 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		return
	}
	if ( sisID == 11 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		return
	}
	
	if ( sisID == 12 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		Form2.btnConsEstoque.disabled = true
		Form2.btnConsIdFis.disabled = true
		Form2.btnLimparEstoque.disabled = true
		Form2.btnOK.disabled = true
		return
	}
	
	if ( sisID == 14 ) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = true
		Form2.btnConsEstoque.disabled = true
		Form2.btnConsIdFis.disabled = true
		Form2.btnLimparEstoque.disabled = true
		Form2.btnOK.disabled = true
		return
	}
	
	if (sisID == 2 || sisID == 3) {
		spnPlataforma.innerHTML = ""
		Form2.consultar.disabled = false
		Form2.ocupados.disabled = false
		Form2.btnRedeDet.disabled = true
		//Form2.btnEmailPro.disabled = false
		
		return
	}
	
	if(sisID == 1){
		Form2.btnRedeDet.disabled = false
		Form2.consultar.disabled = true
		Form2.ocupados.disabled = true	
		//Form2.btnEmailPro.disabled = false
	}
	

}

function CheckEstacaoUsuFac(objCNL,objCompl,usu,origemEst)
{
	with (document.forms[0])
	{
	
		if (objCNL.value != "" && objCompl.value != "")
		{
			hdnCNLEstUsu.value = objCNL.value
			hdnComplEstUsu.value = objCompl.value
			hdnOrigemEst.value = origemEst
			hdnUsuario.value = usu
			hdnAcao.value = "CheckEstacaoUsuFac"
			target = "IFrmProcesso"
			action = "ProcessoSolic.asp"
			submit()
		}
	}
}

<%'GPON
if strTecnologia <> "GPON" or isnull(strTecnologia) then%>
  CarregarDoc()
<%end if%>
//-->



//GPON
function ResgatarFabONT() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboTecnologia').value
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoFabONT;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_FabONT.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<param>"+param+"</param>"
        xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoFabONT () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spncboFabricante").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}

function ResgatarTipoONT() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
		param = document.getElementById('cboFabricanteONT').value
		if (param==0){param=0}
        xmlhttp.onreadystatechange = processadorMudancaEstadoTipoONT;
        xmlhttp.open("POST", "../Ajax/AJX_Resgatar_TipoONT.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
		strXML = "<dados><param>"+param+"</param><texto>N</texto></dados>"
        xmlhttp.send(strXML);
    }
}

function processadorMudancaEstadoTipoONT () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
			document.getElementById("spncboTipoONT").innerHTML = xmlhttp.responseText;
        } else { 
            alert( "Erro: " + xmlhttp.statusText ); 
			return 
        } 
    }
}

function EsconderFabONT()
{
	document.getElementById("spncboFabricante").innerHTML = "";
}

function EsconderTipoONT()
{
	document.getElementById("spncboTipoONT").innerHTML = "";
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
<input type="hidden" name="Ped_Id" value="<%=dblPedId%>">
<input type="hidden" name="hdnAcao">
<input type=hidden	 name="hdnPlataforma" value="<%=strPla%>" >
<input type="hidden" name="hdnSubAcao">
<input type="hidden" name="hdnRede" value="<%=strRede%>">
<input type="hidden" name="hdnPedId"		value="<%=Request.Form("hdnPedId")%>">
<input type="hidden" name="hdnSolId"		value="<%=Request.Form("hdnSolId")%>">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type="hidden" name="txtRazaoSocial"	value="<%=strCliente%>">
<input type="hidden" name="txtEndCid"		value="<%=strCidSigla%>">
<input type="hidden" name="cboUFEnd"		value="<%=strUfSigla%>">
<input type="hidden" name="cboLogrEnd"		value="<%=strTplSigla%>">
<input type="hidden" name="txtEnd"			value="<%=strNomeLogr%>">
<input type="hidden" name="txtNroEnd"		value="<%=strNroEnd%>">
<input type="hidden" name="txtCepEnd"		value="<%=strCep%>">
<input type="hidden" name="hdnTipoProcesso"	value="<%=intTipoProcesso%>">
<input type="hidden" name="hdnTipoAcao"	value="<%=intTipoAcao%>">
<input type="hidden" name="hdnFacDetid">
<input type="hidden" name="hdnOEOrigem" value="<%=orisolid%>">
<input type="hidden" name="hdnDstId">

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

<input type=Hidden name=hdnCVLAN_ETHERNET> 
<input type=Hidden name=hdnSVLAN_ETHERNET> 
<input type=Hidden name=hdnVLAN_PortaOLT> 

<input type=Hidden name=hdnXMLDSLAM> 



<input type=hidden name="hdndisabledSvlan" value="<%=disabledSvlan%>">
<input type=hidden name="hdndisabledCvlan" value="<%=disabledCvlan%>">

<input type=Hidden name=hdnportaoltSwitch>  
<input type=Hidden name=hdnpeSwitch> 
 <input type=Hidden name=hdnQtdPorta >
  <input type=Hidden name=hdndesigRadioIP >
  
   <input type=Hidden name=hdnportaSwitchLadoMetro >
  <input type=Hidden name=hdnSwitchIP >
   <input type=Hidden name=hdnrdoPortaSwitchID >
   
<input type=Hidden name=hdnSwitchInterconexao >
<input type=Hidden name=hdnSwitchMetro >
<input type=Hidden name=hdnSwitchEDD >
<input type=Hidden name=hdnSwitchPE >

<input type=Hidden name=hdnIDUplinkMETRO >
<input type=Hidden name=hdnUplinkMETRO >

<input type=Hidden name=hdnIDDSLAM >
<input type=Hidden name=hdnFACID >


<input type=Hidden name=hdnUplinkEDD >
<input type=Hidden name=hdnIDUplinkEDD >

<input type=Hidden name=hdnUplinkInterconexao >
<input type=Hidden name=hdnIDUplinkInterconexao >
<input type=Hidden name=hdnEqui >

<input type=Hidden name=hdnUplinkDSLAM >




<input type=Hidden name=hdnrdoPortaUplinkIDMETRO >
<input type=Hidden name=hdnrdoPortaUplinkIDEDD >
<input type=Hidden name=hdnrdoPortaUplinkIDInterconexao >
<input type=Hidden name=hdnrdoPortaPEID >
<input type=Hidden name=hdnIDPE >

<input type=Hidden name=hdnrdoPortaDSLAM >




  
  


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

<tr><th colspan=6><p align=center>Alocação de Facilidade</p></th></tr>

<tr>
	<th colspan=5 >&nbsp;•&nbsp;Informações Gerais</th>
	<th align=rigth >
		<a href="javascript:DetalharFac()"><font color=white>Mais...</font></a>
	</th>
</tr>
<tr class=clsSilver>
	<td nowrap width=170>Solicitação de Acesso Nº</td>
	<td>&nbsp;<%=strNroSolic%></td>
	<td align=right >Id Lógico</td>
	<td>&nbsp;<%=strIdLogico%></td>
	<td align=right >Data</td>
	<td>&nbsp;<%=strDataSolic%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Pedido de acesso</td>
	<td>&nbsp;<%=strDM%></td>
	<td align=right>Data</td>
	<td colspan=3 nowrap>&nbsp;<%=strDataPedido%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Ação</td>
	<td colspan=5>&nbsp;<%=strAcao%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Status</td>
	<td colspan=5>&nbsp;<%=strSatus%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Cliente</td>
	<td colspan="5">&nbsp;<%=strCliente%></td>
</tr>

<tr class=clsSilver>
	<td width=170>Endereço</td>
	<td colspan="5">&nbsp;<%=strEndereco%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Nº Contrato Serviço</td>
	<td nowrap>&nbsp;<%=strNroServico%></td>
	<td align=right>Designação do Serviço</td>
	<td colspan=3>&nbsp;<%=strDesigServico%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Serviço</td>
	<td >&nbsp;<%=strServico%> <%IF strPabx_Virtual <> "" then %><br><BR>PABX VIRTUAL BROADSOFT: <%= strPabx_Virtual %><br><%END IF%></td>
	<td align=right>Velocidade do Serviço</td>
	<td colspan=3>&nbsp;<%=strVelServico%></td>
</tr>

<!-- Projeto ABW - Incio -->
<tr class=clsSilver>
	<td width=170>Migração do PABX Virtual</td>
	<td>&nbsp;<%=strMigraPabx_Virtual%></td>
	<td align=right>Tecnologia IP</td>
	<td colspan=3 nowrap>&nbsp;<%=strTecnologia_IP%></td>
</tr>
<!-- Projeto ABW - Fim -->

<tr class=clsSilver>
	<td width=170>Velocidade do Acesso Físico</td>
	<td colspan=5>&nbsp;<%=strVelFis%>&nbsp;<%=TipoVel(intTipoVel)%></td>
</tr>
<%
if (intTipoProcesso = 4 and intTipoAcao = 4) or (intTipoProcesso = 3 and intTipoAcao = 3)then %>
	<tr class="clsSilver">
		<td width=170>Observações</td>        
		<td colspan="6"><textarea name="txtObsProvedor" cols="50" rows="2"><%=PedObs%></textarea></td>
	</tr>
<%Else
 if intTipoProcesso = 1 and intTipoAcao = 1 then%>	
	<tr class="clsSilver">
		<td width=170>Observações</td>        
		<td colspan="6"><textarea name="txtObsProvedor" cols="50" rows="2"><%=strObs%></textarea></td>
	</tr>
<%End if%>
<%End if%>
</table>

<%
'Localiza informações da solução SSA
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblNroSev
	Vetor_Campos(2)="adInteger,2,adParamOutput,0"
	'Call APENDA_PARAM("CLA_sp_sel_solucao_ssa",2,Vetor_Campos)
	Call APENDA_PARAM("CLA_sp_sel_historico_solucao_ssa",2,Vetor_Campos)
	
	Set objRSCli = ObjCmd.Execute
	DBAction = ObjCmd.Parameters("RET").value

	if DBAction = 0 then
	
		If Not objRSCli.eof and  Not objRSCli.bof then

			%>
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
			
			<%
			While Not objRSCli.eof
				%>
				
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
			Wend
			%>
			
									<%IF strPabx_Virtual <> "" then %>
		<tr>
			<th colspan=10 class="clsSilver">&nbsp;PABX VIRTUAL BROADSOFT: <%= strPabx_Virtual %></th>
		</tr>
	<%END IF%>	
				</table> 
					
			</table> 
			
			
				
		
			
<%
		End if
	End if

%>

<%
if intTipoProcesso = 3 and intTipoAcao = 3 then

	Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblPedId
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_logEvento",2,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)
'Response.Write "Sol" & dblSolId
'Response.Write "Ped" & dblPedId

	if not objRS.Eof then
		
		if objRS("Retorno") = "Verdadeiro" then%>
			<table cellpadding=0 cellspacing=1 width=760>
			<tr>
				<tr>
					<th colspan=4>&nbsp;•&nbsp;Informações de Origem</th>
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
		<%
		End if
	End if
End if%>

<%'GPON
if isnull(strTecnologia) then
	strTecnologia = ""
end if

if strTecnologia <> "GPON" AND strTecnologia <> "FO EDD" AND strTecnologia <> "ADE DSLAM" then%>
	<!--#include file="../inc/facilidade_nao_gpon.asp"-->
<%else
	Vetor_Campos(1)="adInteger,2,adParamInput," & DblAcf_ID
	'Vetor_Campos(2)="adInteger,2,adParamInput," & strIdLogico
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_FacilidadeGPON",1,Vetor_Campos)
	'if strloginrede = "PRSSILV" THEN
		'RESPONSE.WRITE strSqlRet
	'END IF
	
	Set objRS = db.Execute(strSqlRet)
	
	if not objRS.eof then
		strTont_ID = objRS("Tont_ID")
		strFont_ID = objRS("Font_ID")
	else
		strTont_ID = 0
		strFont_ID = 0
	end if
	
	'Se não houver porta alocada
					'set objRSedit = db.execute("select top 1 * from cla_acessofisico where 1=2") 'testando !!!!
					'1) buscar a query abaixo sem top 1
					'2) se for > 1, bloquear
					'3) se for = 0, liberar
					'4) se =1, verificar o acesso lógico:
					'pegar o acesso fisico, ir na cla_ontporta, pela ONTPorta_ID ir na cla_ontvlan e buscar a acesso lógico
					'se acesso lógico buscado for igual ao strIdLogico, LIBERAR, caso não, BLOQUEAR
				'set objRSedit = db.execute("select top 1 cla_ontporta.ONTPorta_ID from cla_ontporta inner join cla_ontvlan	on cla_ontporta.ONTPorta_ID = cla_ontvlan.ONTPorta_ID where acf_id = " & DblAcf_ID)
  
  
  'JCARTUS - EDIÇÃO FABRICANTE E NODELO Para primeira porta alocada
  bloqueia_edit = "N"
  SET objRSedit = Server.CreateObject("ADODB.Recordset")
  objRSedit.CursorLocation = 3
	set objRSedit = db.execute("SELECT dbo.CLA_ONTVlan.Acl_IDAcessoLogico FROM dbo.CLA_ONTPorta INNER JOIN dbo.CLA_ONTVlan ON dbo.CLA_ONTPorta.ONTPorta_ID = dbo.CLA_ONTVlan.ONTPorta_ID WHERE (dbo.CLA_ONTPorta.Acf_ID = " & DblAcf_ID )
	
	if objRSedit.recordCount > 1 then
		bloqueia_edit = "S"
	elseif objRSedit.recordCount = 0 then
			bloqueia_edit = "N"
	elseif objRSedit.recordCount = 1 then
		if objRSedit("Acl_IDAcessoLogico") = strIdLogico then
			bloqueia_edit = "N"
		else
			bloqueia_edit = "S"
		end if
	end if
	
	%>
	<script>
	function Trim(str){return str.replace(/^\s+|\s+$/g,"");}

	function GravarPortas ()
	{
		with (document.forms[0])
			{
			
			<%if bloqueia_edit = "S" then
'			if objRSedit.eof then%>
			btnAlt.disabled=true;
			cboFabricanteONT.disabled=true;
			cboTipoONT.disabled=true;
			<%end if%>
			campo = hdnIDrdoVlan.value;
			//alert(txtDesigONT.value);
			
			if ( IsEmpty(campo) ){
				alert("Favor selecionar uma porta");
				return;
			}
			
			hdnvlan.value = document.getElementById(campo).value;
			hdnportaolt.value = document.getElementById("portaolt" + campo).value;
			hdnpe.value = document.getElementById("pe" + campo).value;
			hdnSvlan.value = document.getElementById("svlan" + campo).value;
			//hdnDesigOnt.value = txtDesigONT.value;
			//alert(hdnDesigOnt.value)
			//alert("hdnportaolt.value="+hdnportaolt.value)
			//alert("hdnpe.value="+hdnpe.value)
			
			
			if ( IsEmpty(txtDesigONT.value) ){
				alert("Favor informar a Designação ONT");
				return;
			}
			if ( hdnTecnologia.value ==  "ADE DSLAM" )  {
				hdnTipoServico.value = document.getElementById("cboTipoServico" + campo).value;
			}
			
			//alert(hdnTecnologia.value)
			//alert(document.getElementById("txtSwitch").value)
			//alert(document.getElementById("txtPortaSwitch").value)
			
			if ( hdnTecnologia.value ==  "FO EDD" )  {
			
			
				if ( IsEmpty(txtSwitch.value) ){
					alert("Favor informar Switch");
					return;
				}
				
				if ( IsEmpty(txtPortaSwitch.value) ){
					alert("Favor informar a Porta do Switch");
					return;
				}
				
			}
			
			if ( hdnTecnologia.value !=  "ADE DSLAM" )  {
				if (hdnvlan.value==null || Trim(hdnvlan.value)==""){
					alert("Favor informar a VLAN");
					return;
				}
			}
			if (hdnportaolt.value==null || Trim(hdnportaolt.value)==""){
				alert("Favor informar a Porta PE");
				return;
			}
			if (hdnpe.value==null || Trim(hdnpe.value)==""){
				alert("Favor informar a PE/SW");
				return;
			}
			if (hdnSvlan.value==null || Trim(hdnSvlan.value)==""){
				alert("Favor informar a SVLAN");
				return;
			}
			
			
						
			target = "IFrmProcesso"
			action = "AlocarPortaGPON.asp"
			submit()
			}
	}
	
	function LimparVlan(IDrdo)
	{
		
		document.forms[0].btnOK.disabled=false;
		document.forms[0].hdnIDrdoVlan.value = IDrdo;
		
		<%
			strSQL = "select sum(OPqtd_Qtd) as Tont_QtdPorta from CLA_OntQtdPorta where Tont_ID = " & strTont_ID
			Set objRSqtd = db.Execute(strSQL)
		%>		
	
		qtd = <%=objRSqtd("Tont_QtdPorta")%>
		
		for (i=1;i<=qtd+1;i++)
		{
			try
  			{
  				document.getElementById(i).value = "";
				document.getElementById(i).disabled=true;
  			}
			catch(err)
			{
  				//err
			}
				
		}
		with (document.forms[0]) 
		{
			//alert(hdnTecnologia.value);
			if ( hdnTecnologia.value ==  "ADE DSLAM" )  {
			
				qtd = <%=objRSqtd("Tont_QtdPorta")%>
				
				
		
				for (i=1;i<=qtd;i++)
				{
					try
					{
						Svlan 		=	 'svlan'+i;
						Portaolt	=	 'portaolt'+i;
						Pe			=	 'pe'+i;
						
						//alert(Svlan);
						//alert(Portaolt);
						//alert(Pe);
						//document.getElementById(i).value = "";
						document.getElementById(Svlan).disabled=true;
						document.getElementById(Portaolt).disabled=true;
						document.getElementById(Pe).disabled=true;
						document.getElementById(i).disabled=true;
					}
					catch(err)
					{
						//err
					}
						
				}
		
				campoSvlan 		= 		'svlan'+IDrdo;
				campoPortaolt	=	 	'portaolt'+IDrdo;
				campoPe			=	 	'pe'+IDrdo;
				//alert(campoSvlan);
				combocboTipoServico = 'cboTipoServico'+IDrdo;
				//alert(combocboTipoServico);
				//alert(document.getElementById(combocboTipoServico).value);
				
				if (document.getElementById(combocboTipoServico).value==null || Trim(document.getElementById(combocboTipoServico).value)==""){
					alert("Favor informar o L2/L3");
					return;
				}
				if (document.getElementById(combocboTipoServico).value == "1")
				{
					//alert(IDrdo);
					document.getElementById(IDrdo).disabled=true;
					document.getElementById(campoSvlan).disabled=false;
					document.getElementById(campoPortaolt).disabled=false;
					document.getElementById(campoPe).disabled=false;
					document.getElementById(campoSvlan).focus();	
					
								
					
				}
				if (document.getElementById(combocboTipoServico).value == "2")
				{
					document.getElementById(IDrdo).disabled=false;
					document.getElementById(campoSvlan).disabled=false;
					document.getElementById(campoPortaolt).disabled=false;
					document.getElementById(campoPe).disabled=false;
					document.getElementById(IDrdo).focus();	
					
					
				}
			}
			else{
				document.getElementById(IDrdo).disabled=false;
				document.getElementById(IDrdo).focus();		
			}
		}
	}
	
	
	</script>	
	<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
		<tr>
			<th colspan=10 class="clsSilver">&nbsp;•&nbsp; Controle de Ocupação da ONT | EDD </th>
		</tr>
		
		<%
		if bloqueia_edit = "S" then
'		if objRSedit.eof then
		'GPON
		%>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Fabricante</td><td>&nbsp;&nbsp; <span ID=spncboFabricante>
			<%
			set objRScbo = db.execute("CLA_sp_sel_FabricanteONT 0 , null , null , " & strTecId)

			cbo = ""
			cbo = cbo & "<select name=cboFabricanteONT onchange='ResgatarTipoONT();document.forms[0].btnAlt.disabled=false;'>"
			cbo = cbo & "<Option value=''>:: FABRICANTE ONT</Option>"
			
			While Not objRScbo.eof
			  strItemSel = ""
			  if Trim(dblFontID) = Trim(objRScbo("Font_ID")) then strItemSel = " Selected " End if
			  cbo = cbo & "<Option value='" & objRScbo("Font_ID") & "'"
			  
			  if objRScbo("Font_ID") = strFont_ID then 
			  	cbo = cbo & " selected "
			  end if
			  cbo = cbo & ">" & Trim(objRScbo("Font_Nome")) & "</Option>"
			  objRScbo.MoveNext
			Wend
			strItemSel = ""
			
			cbo = cbo & "</select>"
			
			response.write cbo
			%>
			</span></td><td>&nbsp;&nbsp;</td>
		</tr>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Modelo</td><td>&nbsp;&nbsp; <span ID=spncboTipoONT>
			<%
			set objRScbo = db.execute("CLA_sp_sel_TipoONT null," & strFont_ID)

			cbo = ""
			cbo = cbo & "<select name=cboTipoONT onchange=javascript:document.forms[0].btnAlt.disabled=false;>"
			cbo = cbo & "<Option value=''>:: MODELO ONT</Option>"
				
			While Not objRScbo.eof
				strItemSel = ""
				if Trim(dblTontID) = Trim(objRScbo("Tont_ID")) then strItemSel = " Selected " End if
					cbo = cbo & "<Option value='" & objRScbo("Tont_ID") & "'"
					
					if objRScbo("Tont_ID") = strTont_ID then 
			  			cbo = cbo & " selected "
					end if
					
					cbo = cbo & ">" & Trim(objRScbo("Tont_Modelo")) & "</Option>"
				objRScbo.MoveNext
			Wend
			strItemSel = ""
			
			cbo = cbo & "</select>"
			response.write cbo
			%>			
			</span><input type="Button" class="button" name="btnAlt" value="Alterar" disabled onclick="AlterarPorta()"></td><td>&nbsp;&nbsp;</td>
		</tr>
		<%else%>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Fabricante</td><td>&nbsp;&nbsp; <%=objRS("Font_Nome")%></td><td>&nbsp;&nbsp;</td>
		</tr>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Modelo</td><td>&nbsp;&nbsp; <%=objRS("Tont_Modelo")%></td><td>&nbsp;&nbsp;</td>
		</tr>
		<%end if%>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Portas</td><td>&nbsp;&nbsp; <%=objRSqtd("Tont_QtdPorta")%></td><td>&nbsp;&nbsp;</td>
		</tr>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Designação</td><td>&nbsp;&nbsp; <input type="text" class="text" name="txtDesigONT" value="<%=trim(objRS("Ont_Desig"))%>" size="25" maxlength="20" ></td><td>&nbsp;&nbsp;</td>
		</tr>
		<input type="Hidden" name="hdnOntDesig" value="<%=objRS("Ont_Desig")%>">
		
		<% if strTecnologia = "FO EDD" then %>
			<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
				<tr>
					<th colspan=10 class="clsSilver">&nbsp;•&nbsp; Controle do Switch</th>
				</tr>
				
		    	<tr class="clsSilver">
					<td>&nbsp;&nbsp;Switch Concentrador</td><td>&nbsp;&nbsp; <input type="text" class="text" name="txtSwitch" value="<%=strSwitch%>" size="25" maxlength="20" ></td><td>&nbsp;&nbsp;</td>
				</tr>
				<tr class="clsSilver">
					<td>&nbsp;&nbsp;Porta do Switch</td><td>&nbsp;&nbsp; <input type="text" class="text" name="txtPortaSwitch" value="<%=strSwitchPorta%>" size="25" maxlength="10"></td><td>&nbsp;&nbsp;</td>
				</tr>
			</table>
					
		<%End if %>
				
			</table>
	<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
		<tr>
			<th colspan=11>&nbsp;•&nbsp;</th>
		</tr>
		<tr>
			<th width=10>&nbsp;&nbsp;Porta</th>
			<th>&nbsp;&nbsp;Tipo</th>
			<th>&nbsp;&nbsp;Serviço</th>
			<th>&nbsp;&nbsp;Designação Serviço</th>
			<th>&nbsp;&nbsp;ID Lógico</th>
			<th>&nbsp;&nbsp;VLAN</th>
			<th>&nbsp;&nbsp;SVLAN</th>
			<th>&nbsp;&nbsp;Porta PE</th>			
			<th >&nbsp;&nbsp;PE/SW</th>
			<% IF strTecnologia = "ADE DSLAM" then %>
				<th >&nbsp;&nbsp;L2/L3</th>
			<% end if %>
			<th >&nbsp;&nbsp;</th>
			
		</tr>
		<input type="Hidden" name="hdnOntID" value="<%=objRS("ONT_ID")%>">
		<input type="Hidden" name="hdnOrisol_ID" value="<%=orisolid%>">
		<input type="Hidden" name="hdnAcl_IDAcessoLogico" value="<%=strIdLogico%>">		
		<input type="Hidden" name="hdnId_Acesso" value="<%=Id_Acesso%>">
		<%
		i=1
		Vetor_Campos(1)="adInteger,2,adParamInput," & DblAcf_ID
		Vetor_Campos(2)="adInteger,2,adParamInput," '& DblAcf_ID
		Vetor_Campos(3)="adWChar,50,adParamInput," & Id_Acesso
		
		strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_FacilidadeGPON",3,Vetor_Campos)
		'response.write strSqlRet
		Set objRS = db.Execute(strSqlRet)
	If Not objRS.eof and  Not objRS.bof then	
		While Not objRS.Eof
			if blnCor then
				response.write "<tr class=clsSilver>"
				blnCor = false
			Else
				response.write "<tr class=clsSilver2>"
				blnCor = true
			End if
				'response.write"<script>alert('"&objRS("Acl_IDAcessoLogico")&"')</script>"
				'if isnull(objRS("OntVlan_Nome")) then
				strVlanAlocado = objRS("Acl_IDAcessoLogico") 
				if strVlanAlocado = "" or  isnull(strVlanAlocado) then 
						'response.write "<script>alert('"&DblAcf_ID&"')</script>"
						if isnull(objRS("ONTPorta_Porta")) then
							set ObjRSPorta = db.execute("select top 1 CLA_ONTPorta.ONTPorta_ID,ONTPorta_Porta,	TPPorta_Abrev from cla_acessofisico inner join CLA_ONTPorta			on CLA_AcessoFisico.Acf_ID = CLA_ONTPorta.Acf_ID inner join CLA_TipoPorta		on CLA_ONTPorta.TPPorta_ID = CLA_TipoPorta.TPPorta_ID inner join CLA_OntVlan			on CLA_ONTPorta.ONTPorta_ID = CLA_OntVlan.ONTPorta_ID where cla_acessofisico.acf_id = " & DblAcf_ID )
							if not ObjRSPorta.eof then
							
								ONTPorta_ID    = ObjRSPorta("ONTPorta_ID")
								ONTPorta_Porta = ObjRSPorta("ONTPorta_Porta")
								TPPorta_Abrev  = ObjRSPorta("TPPorta_Abrev")
								'response.write "<script>alert('"&ONTPorta_ID&"')</script>"
							end if
                        else
							ONTPorta_ID    = objRS("ONTPorta_ID")
							ONTPorta_Porta = objRS("ONTPorta_Porta")
							TPPorta_Abrev  = objRS("TPPorta_Abrev")
						end if
					%>
					<td>&nbsp;&nbsp;<%=ONTPorta_Porta%></td>
					<td>&nbsp;&nbsp;<%=TPPorta_Abrev%></td>
					<td>&nbsp;&nbsp;</td>
					<td>&nbsp;&nbsp;</td>
					<td>&nbsp;&nbsp;</td>
					<td>&nbsp;&nbsp;					
						<input type=text class=text name='<%=i%>' ID='<%=i%>' size='5' maxlength='5' disabled value="" onKeyUp="ValidarTipo(this,0)">
					</td>
					<td>&nbsp;&nbsp;
						<input type=text class=text name='svlan<%=i%>' ID='svlan<%=i%>' size='5' maxlength='5'  value="" onKeyUp="ValidarTipo(this,0)"> 
					</td>
					<td>&nbsp;&nbsp;
						<input type=text class=text name='portaolt<%=i%>' ID='portaolt<%=i%>' size='16' maxlength='16'  value="">
					</td>
					<td>&nbsp;&nbsp;
						<input type=text class=text name='pe<%=i%>' ID='pe<%=i%>' size='13' maxlength='13'  value=""> 
					</td>
					<% IF strTecnologia = "ADE DSLAM" then %>
						<td >
							<select name='cboTipoServico<%=i%>'>  
									<Option value=""></Option>	
									<%
									set objRSts = db.execute("CLA_sp_sel_TipoServico ")
									
												
									While Not objRSts.eof
									  strItemSel = ""
									   if Trim(objRS("TpServico_SIGLA")) = Trim(objRSts("TpServico_SIGLA")) then strItemSel = " Selected " End if
									   
									 %>
										<Option value="<%=objRSts("TpServico_ID") %>" <%=strItemSel%> > <%=trim(objRSts("TpServico_SIGLA")) %> </Option> 
										<%
										objRSts.MoveNext
									Wend
									%> 
									
							</select>
						</td>
					<% end if %>
					<td>&nbsp;&nbsp;
					<input type="radio" name="rdoPortaID" onclick="LimparVlan('<%=i%>')" title="Editar" value="<%=ONTPorta_ID%>">	
					</td>	
					<%
				else
					'se o id logico for igual ao corrente habilitar a edição da vlan. não usar a função de apagar. e deixar o rdo marcado
					%>
					<td>&nbsp;&nbsp;<%=objRS("ONTPorta_Porta")%></td>
					<td>&nbsp;&nbsp;<%=objRS("TPPorta_Abrev")%></td>
					<td>&nbsp;&nbsp;<%=objRS("ser_Desc")%></td>
					<td>&nbsp;&nbsp;<%=objRS("Acl_DesignacaoServico")%></td>
					<td>&nbsp;&nbsp;<%=objRS("Acl_IDAcessoLogico")%></td>
					<td>&nbsp;&nbsp;				
					<%
					strIDLogico678 = "678" & mid(strIDLogico,4,7)
					strIDLogico677 = "677" & mid(strIDLogico,4,7)
					'response.write "<script>alert('"&objRS("Acl_IDAcessoLogico")&"')</script>"
					'response.write "<script>alert('"&strIdLogico&"')</script>"
					'response.write "<script>alert('"&strIDLogico678&"')</script>"
					
					if ( Trim(objRS("Acl_IDAcessoLogico")) = strIDLogico678 or Trim(objRS("Acl_IDAcessoLogico")) = strIDLogico677)then
					%>
						<input type=text class=text name='<%=i%>' ID='<%=i%>' size='5' maxlength='5' value="<%=trim(objRS("OntVLAN_Nome"))%>" onKeyUp="ValidarTipo(this,0)">
						</td>
						<td>&nbsp;&nbsp;
						<input type=text class=text name='svlan<%=i%>' ID='svlan<%=i%>' size='5' maxlength='5' value="<%=trim(objRS("OntSVLAN_Nome"))%>">
						</td>
						<td>&nbsp;&nbsp;
						<input type=text class=text name='portaolt<%=i%>' ID='portaolt<%=i%>' size='16' maxlength='16' value="<%=trim(objRS("OntVLAN_portaolt"))%>"> 
						</td>
						
						<td>&nbsp;&nbsp;
						<!--<input type=text class=text name='pe<%=i%>' ID='pe<%=i%>' size='13' maxlength='13' value="<%=trim(objRS("OntVLAN_Pe"))%>" onKeyUp="ValidarTipo(this,0)"> -->
						
						<input type=text class=text name='pe<%=i%>' ID='pe<%=i%>' size='13' maxlength='13' value="<%=trim(objRS("OntVLAN_Pe"))%>">
						</td>
						<% IF strTecnologia = "ADE DSLAM" then %>
							<td >
								<select name='cboTipoServico<%=i%>'>  
										<Option value=""></Option>	
										<%
										set objRSts = db.execute("CLA_sp_sel_TipoServico ")
										
										'response.write "<script>alert('"&objRS("TpServico_SIGLA")&"')</script>"	
										'response.write "<script>alert('"&objRSts("TpServico_SIGLA")&"')</script>"
										While Not objRSts.eof
										  strItemSel = ""
										   if Trim(objRS("TpServico_SIGLA")) = Trim(objRSts("TpServico_SIGLA")) then strItemSel = " Selected " End if
										   
										   'response.write "<script>alert('"&strItemSel&"')</script>"
										 %>
											<Option  value="<%=objRSts("TpServico_ID") %>" <%=strItemSel%> > <%=trim(objRSts("TpServico_SIGLA")) %> </Option> 
											<%
											objRSts.MoveNext
										Wend
										%> 
										
								</select>
							</td>
						<% end if %>
						<td>
						<input type="radio" name="rdoPortaID" onclick="LimparVlan('<%=i%>')" title="Editar" value="<%=trim(objRS("ONTPorta_ID"))%>" checked>
						</td>
                    <%
						IDrdoVlan = i
					else
					%>
						<input type=text class=text name='<%=i%>' ID='<%=i%>' size='5' maxlength='5' value="<%=trim(objRS("OntVLAN_Nome"))%>" disabled onKeyUp="ValidarTipo(this,0)">
						</td>
						<td>&nbsp;&nbsp;
						<input type=text class=text name='svlan<%=i%>' ID='svlan<%=i%>' size='5' maxlength='5' value="<%=trim(objRS("OntSVLAN_Nome"))%>" disabled onKeyUp="ValidarTipo(this,0)"> 
						</td>
						<td>&nbsp;&nbsp;
						<input type=text class=text name='portaolt<%=i%>' ID='portaolt<%=i%>' size='16' maxlength='16' value="<%=trim(objRS("OntVLAN_PortaOLT"))%>" disabled>
						</td>						
						<td>&nbsp;&nbsp;
						<input type=text class=text name='pe<%=i%>' ID='pe<%=i%>' size='13' maxlength='13' value="<%=trim(objRS("OntVLAN_PE"))%>" disabled> 
						</td>
							<% IF strTecnologia = "ADE DSLAM" then %>
								<td >
									<select name='cboTipoServico<%=i%>'>  
											<Option value=""></Option>	
											<%
											set objRSts = db.execute("CLA_sp_sel_TipoServico ")
											
														
											While Not objRSts.eof
											  strItemSel = ""
											  if Trim(objRS("TpServico_SIGLA")) = Trim(objRSts("TpServico_SIGLA")) then strItemSel = " Selected " End if
											   
											 %>
												<Option value="<%=objRSts("TpServico_ID") %>" <%=strItemSel%> > <%=trim(objRSts("TpServico_SIGLA")) %> </Option> 
												<%
												objRSts.MoveNext
											Wend
											%> 
											
									</select>
								</td>
							<% end if %>
						<td>
						<input type="radio" name="rdoPortaID" onclick="LimparVlan('<%=i%>')" title="Editar" value="<%=objRS("ONTPorta_ID")%>" disabled>
						</td>
						<%=objRS("OntVLAN_Nome")%> 
						<input type=hidden class=text name='<%=i%>' ID='<%=i%>' size='5' maxlength='5' value="<%=trim(objRS("OntVLAN_Nome"))%>">
						<input type="Hidden" name="hdnrdoPortaID" value="<%=objRS("ONTPorta_ID")%>">
					<%
						IDrdoVlan = i
						i=i+1
						if blnCor then
							response.write "<tr class=clsSilver>"
							blnCor = true
						Else
							response.write "<tr class=clsSilver2>"
							blnCor = false
						End if
						if isnull(objRS("ONTPorta_Porta")) then
							set ObjRSPorta = db.execute("select top 1 CLA_ONTPorta.ONTPorta_ID,ONTPorta_Porta,	TPPorta_Abrev from cla_acessofisico inner join CLA_ONTPorta			on CLA_AcessoFisico.Acf_ID = CLA_ONTPorta.Acf_ID inner join CLA_TipoPorta		on CLA_ONTPorta.TPPorta_ID = CLA_TipoPorta.TPPorta_ID inner join CLA_OntVlan			on CLA_ONTPorta.ONTPorta_ID = CLA_OntVlan.ONTPorta_ID where cla_acessofisico.acf_id = " & DblAcf_ID )
							if not ObjRSPorta.eof then
							
								ONTPorta_ID    = ObjRSPorta("ONTPorta_ID")
								ONTPorta_Porta = ObjRSPorta("ONTPorta_Porta")
								TPPorta_Abrev  = ObjRSPorta("TPPorta_Abrev")
								'response.write "<script>alert('"&ONTPorta_ID&"')</script>"
							end if
                        else
							ONTPorta_ID    = objRS("ONTPorta_ID")
							ONTPorta_Porta = objRS("ONTPorta_Porta")
							TPPorta_Abrev  = objRS("TPPorta_Abrev")
						end if
					%>
						<td>&nbsp;&nbsp;<%=ONTPorta_Porta%></td>
					<td>&nbsp;&nbsp;<%=TPPorta_Abrev%></td>
					<td>&nbsp;&nbsp;</td>
					<td>&nbsp;&nbsp;</td>
					<td>&nbsp;&nbsp;</td>
					<td>&nbsp;&nbsp;					
						<input type=text class=text name='<%=i%>' ID='<%=i%>' size='5' maxlength='5' disabled value="" onKeyUp="ValidarTipo(this,0)">
					</td>
					<td>&nbsp;&nbsp;
						<input type=text class=text name='svlan<%=i%>' ID='svlan<%=i%>' size='5' maxlength='5'  value="" onKeyUp="ValidarTipo(this,0)"> 
					</td>
					<td>&nbsp;&nbsp;
						<input type=text class=text name='portaolt<%=i%>' ID='portaolt<%=i%>' size='16' maxlength='16'  value="">
					</td>
					<td>&nbsp;&nbsp;
						<input type=text class=text name='pe<%=i%>' ID='pe<%=i%>' size='13' maxlength='13'  value=""> 
					</td>
					<% IF strTecnologia = "ADE DSLAM" then %>
						<td >
							<select name='cboTipoServico<%=i%>'>  
									<Option value=""></Option>	
									<%
									set objRSts = db.execute("CLA_sp_sel_TipoServico ")
									
												
									While Not objRSts.eof
									  strItemSel = ""
									  if Trim(objRS("TpServico_SIGLA")) = Trim(objRSts("TpServico_SIGLA")) then strItemSel = " Selected " End if
									   
									 %>
										<Option value="<%=objRSts("TpServico_ID") %>" <%=strItemSel%> > <%=trim(objRSts("TpServico_SIGLA")) %> </Option> 
										<%
										objRSts.MoveNext
									Wend
									%> 
									
							</select>
						</td>
					<% end if %>
					<td>&nbsp;&nbsp;
					<input type="radio" name="rdoPortaID" onclick="LimparVlan('<%=i%>')" title="Editar" value="<%=ONTPorta_ID%>">	
					</td>
						
						
					<%	
					end if%>		
					
					<%						
				end if
			%></td>
		</tr>		
			<%
			i=i+1
			objRS.MoveNext
		Wend
	end if
        %>
		<input type="Hidden" name="hdnIDrdoVlan" value="<%=IDrdoVlan%>">
		<input type="Hidden" name="hdnvlan">
		<input type="Hidden" name="hdnSvlan">
		<input type="Hidden" name="hdnDesigOnt">
		<input type="Hidden" name="hdnportaolt">
		<input type="Hidden" name="hdnpe">
		<input type="Hidden" name="hdnTipoServico">
		
	</table><br>
	<table width="760"
	
		<%if intTipoProcesso = 3 and intTipoAcao = 3 and orisolid = 9 then%>
			<tr>
				<td align=center class=clsSilver2>&nbsp;<input type="button" class="button" name="btnSair" value="Sair" style="width:150px"  onClick="sair()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');"></td>
				<td align=center class=clsSilver2>&nbsp;<input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Liberar para Serviço" onclick="return liberarServico()" accesskey="U" onmouseover="showtip(this,event,'Liberar para Serviço(Alt+U)');"></td>
			</tr>
		<%Else%>
			<tr>
				<td align=center class=clsSilver2>&nbsp;<input type="button" class="button" name="btnSair" value="Sair" style="width:150px"  onClick="sair()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');"></td>
				<td align=center class=clsSilver2>&nbsp;<input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Alocar Porta(s)" onclick="return GravarPortas()" accesskey="I" onmouseover="showtip(this,event,'Alocar Porta(s)(Alt+I)');" ></td>
			</tr>
		<%End if%>
		
		
		
		<% IF strTecnologia = "ADE DSLAM" then %>
		
		<!-- DSLAM >
		
		<tr>
			<th colspan=10>&nbsp;•&nbsp;Recurso</th>
		</tr>

		
			<tr class="clsSilver">
				<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Entrega</td>
				<td colspan="8">
					<%
					set objRS = db.execute("CLA_sp_sel_estacao " & Trim(strLocalInstala))
					%> 
					<input type="Hidden" name="cboLocalInstala" value="<%=strLocalInstala%>">
				  <input type="text" class="text" name="txtCNLLocalEntrega" value=<%=objRS("Cid_Sigla")%> maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
				  &nbsp;<input type="text" class="text" name="txtComplLocalEntrega" value=<%=objRS("Esc_Sigla")%> maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsuFac(document.Form2.txtCNLLocalEntrega,document.Form2.txtComplLocalEntrega,<%=dblUsuId%>,1);" TIPO="A">
				</td> 
				<td colspan="1">&nbsp;</td>
			</tr>
			<tr class="clsSilver">
				<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Configuração</td>
				<td colspan="8">
					<%
					set objRS = db.execute("CLA_sp_sel_estacao " & Trim(strLocalConfig))
					%>
					<input type="Hidden" name="cboLocalConfig" value="<%=strLocalConfig%>">
				  <input type="text" class="text" name="txtCNLLocalConfig" value=<%=objRS("Cid_Sigla")%> maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
				  &nbsp;<input type="text" class="text" name="txtComplLocalConfig" value=<%=objRS("Esc_Sigla")%> maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsuFac(document.Form2.txtCNLLocalConfig,document.Form2.txtComplLocalConfig,<%=dblUsuId%>,2);" TIPO="A">
				</td>
				<td colspan="1">&nbsp;</td>
			</tr>
		
			<tr class=clsSilver>
				<td width=170><font class="clsObrig">:: </font>Distribuidor</td>
				<td colspan="9">
					<span id="spnDistLocalInstala">
					<select name="cboDistLocalInstala" style="width:200px">
					<option value=""></option>
					<%	if Trim(strLocalInstala) <> "" then
							set objRS = db.execute("CLA_sp_view_recursodistribuicao " & strLocalInstala )
							While not objRS.Eof
								strItemSel = ""
								if Trim(strDistrib) = Trim(objRS("Dst_ID")) then strItemSel = " Selected " End if
								Response.Write "<Option value=" & objRS("Dst_ID") & strItemSel & ">" & objRS("Dst_Desc") & "</Option>"
								objRS.MoveNext
							Wend
							strItemSel = ""
						End if
						%>
					</select>
					</span>
				</td>
			</tr>

			<tr class=clsSilver>
				<td width=170><font class="clsObrig">:: </font>Rede</td>
				<td colspan="9" >
					<select name="cboRede" onChange="RetornaCboPlataforma(cboRede.value,<%=strPla%>)" onBlur="ControleRede(this);"  >
						<option value=""></option>
						<%
							strRede = 3
							Set objRS = db.execute("CLA_sp_sel_sistema 3")
							While not objRS.Eof
								strItemSel = ""
								if Trim(strRede) = Trim(objRS("Sis_ID")) then strItemSel = " Selected " End if
								Response.Write "<Option value='" & objRS("Sis_ID")& "'" & strItemSel & ">" & Ucase(objRS("Sis_Desc")) & "</Option>"
								objRS.MoveNext
							Wend
							strItemSel = ""

						%>
					</select>
					<span ID = spnPlataforma>
					</span>
				</td>
			</tr>

			<tr class=clsSilver>
				<td width=170 ><font class="clsObrig">:: </font>Provedor</td>
				<td colspan="9" >
					<%
					set objRS = db.execute("CLA_sp_sel_provedor " & Trim(strProId))
					%>
					<select name="cboProvedor" style="width:250px" readonly>
						<%Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "' tag_provedor=" & strCartaProv & strItemSel & ">" & objRS("Pro_Nome") & "</Option>"%>
					</select>
				</td>
			</tr>
		
		<tr>
		 			
					<th colspan=10>&nbsp;•&nbsp;DSLAM</th>
		</tr> 
		
		<tr class=clsSilver> 


			<td colspan=10><font class=clsObrig>::</font>
				
				<%
				Vetor_Campos(1)="adInteger,8,adParamInput, " & strLocalInstala 'Request.Form("cboLocalInstala")
				Vetor_Campos(2)="adInteger,8,adParamInput, 11 " 
				Vetor_Campos(3)="adInteger,8,adParamInput, 16 " 
				Vetor_Campos(4)="adInteger,8,adParamInput, " & strProId 'Request.Form("cboProvedor")
				
				strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_DesignacaoSwitch",4,Vetor_Campos)
				'response.write strSqlRet
				Set objRS = db.Execute(strSqlRet)
				%>
				Designação do DSLAM <select name=cboSwitchDSLAM   onchange=ResgatarDslam() > 
				<Option value="">:: Concentrador DSLAM</Option> 
				<%
				While Not objRS.eof
				  strItemSel = ""
				  if Trim(dblswitchMetro) = Trim(objRS("equi_id")) then strItemSel = " Selected " End if
							'response.write "<script>alert('"&dblswitchMetro&"')</script>"
							'response.write "<script>alert('"&objRS("equi_id")&"')</script>"
				%>				
						<Option value="<%=objRS("equi_id") %>" > <%=Trim(objRS("designacao"))%> </Option>
				  <%
				  objRS.MoveNext
				Wend
				%>
				</select> 

			</td> 
		</tr>
		
		<% end if %>
		
	</table><br><br>
	<span id=spnDSLAM></span>
	
	<table  cellspacing=1 cellpadding=0 width=760 border=0> 
				
		
		
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
<%end if 'GPON%>
<%if strloginrede = "PRSSILV" THEN%>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso"
	    width       = "300"
	    height      = "100"
	    frameborder = "1"
	    scrolling   = "YES"
	    align       = "left">
</iFrame><iframe	id			= "IFrmProcesso2"
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
<%ELSE%>
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
<%END IF%>
</body>
</html>
<SCRIPT LANGUAGE=javascript>
function sair(){
	window.returnValue=0;
	if(window.history.length < 1){
		window.close()
	}else{
		window.navigate(window.document.referrer)
	}
	
}
</script>
<%'GPON

if strTecnologia <> "GPON" AND strTecnologia <> "FO EDD" AND strTecnologia <> "ADE DSLAM"then%>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlFacOrigem = new ActiveXObject("Microsoft.XMLDOM")
objXmlFacOrigem.loadXML("<%=strXmlFacOrigem%>")
RetornaCboPlataforma(document.forms[0].cboRede.value,<%=strPla%>)
if (document.forms[0].cboRede.disabled) document.forms[0].cboRede.disabled = true
//-->
</SCRIPT>
<%end if 'GPON%>
