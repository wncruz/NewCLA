<%
<!--#include file="inc/Data.asp"-->
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))


Function FormatarDataHora12H(data)
    Dim hora, minuto, segundo, ampm, hora12, dataFormatada

    If IsNull(data) Or Not IsDate(data) Then
        FormatarDataHora12H = ""
        Exit Function
    End If

    hora = Hour(data)
    minuto = Right("0" & Minute(data), 2)
    segundo = Right("0" & Second(data), 2)

    ' Determina AM ou PM
    If hora >= 12 Then
        ampm = "PM"
        If hora > 12 Then
            hora12 = hora - 12
        Else
            hora12 = hora
        End If
    Else
        ampm = "AM"
        If hora = 0 Then
            hora12 = 12
        Else
            hora12 = hora
        End If
    End If

    hora12 = Right("0" & hora12, 2)

    dataFormatada = Right("0" & Day(data), 2) & "/" & _
                    Right("0" & Month(data), 2) & "/" & _
                    Year(data) & " " & _
                    hora12 & ":" & minuto & ":" & segundo & " " & ampm

    FormatarDataHora12H = dataFormatada
End Function

Function EnviarEntregarAprov(dblIdLogico)

		Dim proprietario ' TER ou EBT
		Dim tecnologia ' 0 - terceiro 1 - radio 2 - fibra otica 3 -ade 4 - satelite 5 - cabo interno 
		Dim idTarefa ' identificador do sistema aprovisionador
		Dim oe_numero 		 
		Dim	oe_ano			 
		Dim	oe_item			 
		Dim	idLogico		 
		Dim	acao			 
		Dim OrigemSolicitacao
		Dim rede
		Dim oriSol_id
		Dim configAcesso

   		if dblIdLogico <> "" then

			
			Vetor_Campos(1)="adWChar,15,adParamInput," & dblIdLogico
			strSqlRet = APENDA_PARAMSTR("CLA_sp_view_solicitacaoAprov",1,Vetor_Campos)
			
			Set objRSDadosCla = db.Execute(strSqlRet)
			
			'If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
			if not objRSDadosCla.Eof then
			
			
				solid 		=  objRSDadosCla("sol_id")
				
				Vetor_Campos(1)="adInteger,4,adParamInput," & solid
				Vetor_Campos(2)="adInteger,4,adParamInput, 274"
				Vetor_Campos(3)="adInteger,4,adParamInput," & strloginrede
				Vetor_Campos(4)="adVarchar,1,adParamInput,"
				Vetor_Campos(5)="adVarchar,100,adParamInput,STATUS AUTOMATICO"  
				Vetor_Campos(6)="adVarchar,1,adParamInput,M"
				
  				strSqlRet = APENDA_PARAMSTR("CLA_sp_ins_StatusSolicitacao",6,Vetor_Campos)
				
				db.Execute(strSqlRet)
				
				Aprovisi_ID			= objRSDadosCla("Aprovisi_ID")
				idTarefa 			= objRSDadosCla("id_tarefa")
				oe_numero 			= objRSDadosCla("oe_numero")
				oe_ano				= objRSDadosCla("oe_ano")
				oe_item				= objRSDadosCla("oe_item")
				idLogico			= objRSDadosCla("acl_idacessologico")
				acao				= objRSDadosCla("acao")
				OrigemSolicitacao	= objRSDadosCla("oriSol_Descricao")
				oriSol_id			= objRSDadosCla("oriSol_ID")
				oriSol_Descricao 	= objRSDadosCla("oriSol_Descricao")
				
				Strxml1 = Strxml1 & 	"	<retorno-cla>" & vbnewline
				Strxml1 = Strxml1 & 	"		<acao>"& objRSDadosCla("acao") &"</acao>" & vbnewline
				Strxml1 = Strxml1 & 	"		<origem>"& objRSDadosCla("oriSol_Descricao") &"</origem>" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-tarefa>"& objRSDadosCla("id_tarefa")& "</id-tarefa>" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-logico>"& objRSDadosCla("acl_idacessologico") &"</id-logico>" & vbnewline
				Strxml1 = Strxml1 & 	"		<id-solicitacao>"& objRSDadosCla("sol_id") &"</id-solicitacao>"  & vbnewline
				Strxml1 = Strxml1 & 	"		<qtd-acessos>"& objRSDadosCla("QTDFisico") &"</qtd-acessos>" & vbnewline
				Strxml1 = Strxml1 & 	"		<acessos-fisicos>" & vbnewline
				
				xmlAcesso  = ""
				
				tecnologia   =  objRSDadosCla("tec_id")
				
				if tecnologia = "8" then
					
					set objRS = db.execute (" select	distinct	ltrim(rtrim(cla_equipamento.Designacao)) as Designacao , " &_
														  " 					ltrim(rtrim(FABONT.Font_Nome)) as Font_Nome, " &_
														  " 					ltrim(rtrim(TIPOONT.Tont_Modelo)) as Tont_Modelo, " &_
														  " 					ltrim(rtrim(cla_propriedade.prop_desc)) as  prop_desc, " &_
														  " 					ltrim(rtrim(cla_tipo_equipamento.tipoEqui_id)) as  tipoEqui_id " &_
											" from cla_Acessofisico_equipamento " &_	
											" inner join cla_equipamento						on cla_Acessofisico_equipamento.equi_id = cla_equipamento.equi_id " &_
											" inner join cla_tipo_equipamento					on cla_equipamento.tipoequi_id = cla_tipo_equipamento.tipoequi_id " &_
											" inner join cla_propriedade						on cla_equipamento.prop_id = cla_propriedade.prop_id " &_
											" left join  cla_ontporta	ontporta				on cla_equipamento.equi_id = ontporta.equi_id " &_
											" left JOIN  CLA_TIPOONT	TIPOONT					ON ontporta.TONT_ID = TIPOONT.TONT_ID " &_
											" left JOIN  CLA_FABONT		FABONT				ON TIPOONT.FONT_ID = FABONT.FONT_Id " &_
											" where cla_Acessofisico_equipamento.acf_id = " &  objRSDadosCla("acf_id"))
								
					While Not objRS.eof
												  
						'response.write "<script>alert('"&objRS("DESIGNACAO")&"')</script>"
					  
						if "11" = Trim(objRS("tipoEqui_id")) then 
						
							
							DESIGNACAO_DSLAM 				= objRS("DESIGNACAO")
							
							
						End if
					
					  objRS.MoveNext
					  
					Wend
					
					'''set objRS = db.execute (" select distinct ltrim(rtrim(tipoEqui.tipoEqui_id)) as  tipoEqui_id, ltrim(rtrim(equi.DESIGNACAO)) as DESIGNACAO, ltrim(rtrim(ontporta.OntPorta_Porta)) as OntPorta_Porta , ltrim(rtrim(FABONT.Font_Nome)) as Font_Nome, ltrim(rtrim(TIPOONT.Tont_Modelo)) as Tont_Modelo, ltrim(rtrim(TPPORTA.TPPORTA_ABREV)) as TPPORTA_ABREV , ltrim(rtrim(prop.prop_desc)) as  prop_desc  , ltrim(rtrim(uplink)) as uplink " &_
					'''			" from cla_equipamento	equi " &_
					'''			" inner join  cla_tipo_equipamento  tipoEqui			on equi.TipoEqui_id = tipoEqui.TipoEqui_id " &_
					'''			" inner join  cla_propriedade		  prop				on equi.prop_id = prop.prop_id " &_
					'''			" inner join  cla_ontporta	ontporta				on equi.equi_id = ontporta.equi_id " &_
					'''			" inner join  cla_porta_alocada ptAlocada				on ontporta.ontporta_id = ptAlocada.ontporta_id " &_	
					'''			" INNER JOIN  CLA_TIPOONT	TIPOONT					ON ontporta.TONT_ID = TIPOONT.TONT_ID " &_
					'''			" INNER JOIN  CLA_FABONT		FABONT					ON TIPOONT.FONT_ID = FABONT.FONT_ID " &_
					'''			" INNER JOIN  CLA_TIPOPORTA	TPPORTA				ON ontporta.TPPORTA_ID = TPPORTA.TPPORTA_ID " &_
					'''			" where ptAlocada.acl_idacessologico = " &  idLogico)
					
					'''While Not objRS.eof
												  
						'response.write "<script>alert('"&objRS("DESIGNACAO")&"')</script>"
					  
					'''	if "11" = Trim(objRS("tipoEqui_id")) then 
						
							'FABRICANTE_ETHERNET					 	= objRS("Font_Nome")
							'MODELO_ETHERNET			 				= objRS("Tont_Modelo")
					'''		PORTA_MODEM 	= objRS("OntPorta_Porta")
							
							
					'''	End if
						
					'''  objRS.MoveNext
					  
					''' Wend
					
					'response.write "<script>alert('"&objRSDadosCla("Acf_DtConstrAcessoFis")&"')</script>"
				
					xmlAcesso = xmlAcesso & 	"	<acesso> " & vbnewline
					xmlAcesso = xmlAcesso & 	"		<id-acessoFisico>"& objRSDadosCla("acf_idacessofisico") &"</id-acessoFisico>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<tecnologia>"& objRSDadosCla("tec_id") &"</tecnologia>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<tipo-contrato>"& objRSDadosCla("reg_id") &"</tipo-contrato>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<velocidade>"& objRSDadosCla("vel_desc") &"</velocidade>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<vel-conversao>"& objRSDadosCla("vel_conversao") &"</vel-conversao>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<proprietario>"& objRSDadosCla("acf_proprietario") &"</proprietario>" & vbnewline
					'xmlAcesso = xmlAcesso & 	"		<dt-construcao>"& objRSDadosCla("Acf_DtConstrAcessoFis") &"</dt-construcao>" & vbnewline
					
					xmlAcesso = xmlAcesso & 	"		<dt-construcao>"& FormatarDataHora12H(objRSDadosCla("Acf_DtConstrAcessoFis")) &"</dt-construcao>" & vbnewline
					
					if objRSDadosCla("Acf_DtAceite") = "" then
						strAcf_DtAceite = date()
					else
						strAcf_DtAceite = objRSDadosCla("Acf_DtAceite")
					end if
					xmlAcesso = xmlAcesso & 	"		<dt-aceite>"& FormatarDataHora12H(strAcf_DtAceite) &"</dt-aceite>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<interface>"& objRSDadosCla("Acf_Interface") &"</interface>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<tp-vel>"& objRSDadosCla("acf_tipovel") &"</tp-vel>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<cnl>"& objRSDadosCla("Acf_SiglaEstEntregaFisico") &"</cnl>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<compl-estacao>"& objRSDadosCla("Acf_ComplSiglaEstEntregaFisico") &"</compl-estacao>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<interface-estacao>"& objRSDadosCla("Acf_InterfaceEstEntregaFisico") &"</interface-estacao>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<SiglaCentroCliente>"& objRSDadosCla("Aec_SiglaCentroCliente") &"</SiglaCentroCliente>" & vbnewline
					''' Retirada da interface que trata o paralelismo do SGAPlus na fase tratar acesso
					''' Solicitação de CTERESA conforme email 
					''' Assunto : Não paralelismo da fase Tratar Acesso com Configuração
					''' Data: 04/05/2010 
					''' Inicio
					xmlAcesso = xmlAcesso & 	"		<EstacaoConfiguracao>"& objRSDadosCla("est_config") &"</EstacaoConfiguracao>" & vbnewline
					''' FIM
					
					if oriSol_id = 6 then
						set objRS = db.execute ("select pro_id from cla_provedorConfig where Proconf_Visivel = 'S' and pro_id = " &  objRSDadosCla("pro_id") )
						While Not objRS.eof
						
							configAcesso = "N" 	
							objRS.MoveNext
						  
						Wend
						
						if configAcesso = "N" then
						
							xmlAcesso = xmlAcesso & 	"		<SAIP>N</SAIP>" & vbnewline
						else
							xmlAcesso = xmlAcesso & 	"		<SAIP>S</SAIP>" & vbnewline
						end if 
						
					end if
												
					xmlAcesso = xmlAcesso & 	"		<numero-acesso>"& objRSDadosCla("Acf_NroAcessoPtaEbt") &"</numero-acesso>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<ccto>"& objRSDadosCla("Acf_CCTOFatura") &"</ccto>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<qtd-modem>"& objRSDadosCla("Acf_QtdEquip") &"</qtd-modem>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<prop-modem>"& objRSDadosCla("Acf_ProprietarioEquip") &"</prop-modem>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<tp-aceite>"& objRSDadosCla("TP_TipoAceite") &"</tp-aceite>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<distribuidor>"& objRSDadosCla("Dst_desc") &"</distribuidor>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<rede>"& objRSDadosCla("Sis_Desc") &"</rede>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<cabo>"& objRSDadosCla("fac_tronco") &"</cabo>" & vbnewline ' tipo de acesso ADE fac_tronco e o cabo
					xmlAcesso = xmlAcesso & 	"		<par>"& objRSDadosCla("fac_par") &"</par>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<tp-cabo>"& objRSDadosCla("fac_tipoCabo")&"</tp-cabo>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<obs-ade>"& objRSDadosCla("fac_obs")&"</obs-ade>" & vbnewline
					
					xmlAcesso = xmlAcesso & 	"		<fabricante>"& trim(objRSDadosCla("Font_Nome")) &"</fabricante>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<modelo>"& trim(objRSDadosCla("Tont_Modelo")) &"</modelo>" & vbnewline
					
					xmlAcesso = xmlAcesso & 	"		<tronco-uplink-dslam>"& trim(objRSDadosCla("Ont_Desig")) &"</tronco-uplink-dslam>" & vbnewline
					
					xmlAcesso = xmlAcesso & 	"		<designacao_dslam>"& trim(DESIGNACAO_DSLAM) &"</designacao_dslam>" & vbnewline
					
					xmlAcesso = xmlAcesso & 	"		<vlan>"& trim(objRSDadosCla("OntVlan_Nome")) &"</vlan>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<portaPE>"& trim(objRSDadosCla("OntVlan_PortaOLT")) &"</portaPE>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<pe>"& trim(objRSDadosCla("OntVlan_PE")) &"</pe>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<Svlan>"& objRSDadosCla("OntSVlan_Nome") &"</Svlan>" & vbnewline
					
					xmlAcesso = xmlAcesso & 	"		<porta></porta>" & vbnewline
					
					xmlAcesso = xmlAcesso & 	"		<SWITCH-CONCENTRADOR>"& trim(DESIGNACAO_DSLAM) &"</SWITCH-CONCENTRADOR>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<PORTA-SWITCH-CONCENTRADOR>"& trim(objRSDadosCla("OntPorta_Porta")) &"</PORTA-SWITCH-CONCENTRADOR>" & vbnewline
					
					xmlAcesso = xmlAcesso & 	"		<PROPRIEDADE-DSLAM></PROPRIEDADE-DSLAM>" & vbnewline
					xmlAcesso = xmlAcesso & 	"		<PROPRIEDADE-REDE></PROPRIEDADE-REDE>" & vbnewline
				
					xmlAcesso = xmlAcesso & 	"	</acesso> "
				
			else
			
				while not objRSDadosCla.eof
				
					  proprietario =  objRSDadosCla("acf_proprietario")
					  tecnologia   =  objRSDadosCla("tec_id")
					  rede 		=  objRSDadosCla("sis_id")
					  solid 		=  objRSDadosCla("sol_id")
					  
					  if oriSol_id = 7 and tecnologia = "9" then
							tecnologia = "0"
						end if 
						
					  if proprietario = "TER" then
						
						'response.write "<script>alert('"&rede&"')</script>"
						if ( objRSDadosCla("Sis_Desc") = "" or isnull(objRSDadosCla("Sis_Desc"))) then
							rede_Desc = "ETHERNET"
								
						else 
				
							rede_Desc = objRSDadosCla("Sis_Desc")
		
						end if
						'response.write "<script>alert('"&rede_Desc&"')</script>"
						'response.end
						
						xmlAcesso = xmlAcesso & 	"	<acesso> " & vbnewline
						xmlAcesso = xmlAcesso & 	"		<id-acessoFisico>"& objRSDadosCla("acf_idacessofisico") &"</id-acessoFisico>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<tecnologia>"& objRSDadosCla("tec_id") &"</tecnologia>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<provedor>"& objRSDadosCla("pro_nome") &"</provedor>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<tipo-contrato>"& objRSDadosCla("reg_id") &"</tipo-contrato>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<velocidade>"& objRSDadosCla("vel_desc") &"</velocidade>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<vel-conversao>"& objRSDadosCla("vel_conversao") &"</vel-conversao>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<proprietario>"& objRSDadosCla("acf_proprietario") &"</proprietario>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<dt-construcao>"& FormatarDataHora12H(objRSDadosCla("Acf_DtConstrAcessoFis")) &"</dt-construcao>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<dt-aceite>"& FormatarDataHora12H(objRSDadosCla("Acf_DtAceite")) &"</dt-aceite>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<interface>"& objRSDadosCla("Acf_Interface") &"</interface>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<tp-vel>"& objRSDadosCla("acf_tipovel") &"</tp-vel>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<cnl>"& objRSDadosCla("Acf_SiglaEstEntregaFisico") &"</cnl>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<compl-estacao>"& objRSDadosCla("Acf_ComplSiglaEstEntregaFisico") &"</compl-estacao>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<interface-estacao>"& objRSDadosCla("Acf_InterfaceEstEntregaFisico") &"</interface-estacao>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<rede>"& rede_Desc &"</rede>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<numero-acesso>"& objRSDadosCla("Acf_NroAcessoPtaEbt") &"</numero-acesso> " & vbnewline 
						xmlAcesso = xmlAcesso & 	"		<numero-acesso-cli>"& objRSDadosCla("Acf_NroAcessoPtaCli") &"</numero-acesso-cli>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<ccto>"& objRSDadosCla("Acf_CCTOFatura") &"</ccto>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<cnla>"& objRSDadosCla("Acf_CnlPTA") &"</cnla>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<cnlb>"& objRSDadosCla("Acf_CnlPTB") &"</cnlb>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<tipo-circuito>"& objRSDadosCla("Acf_CCTOTipo") &"</tipo-circuito>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<qtd-modem>"& objRSDadosCla("Acf_QtdEquip") &"</qtd-modem>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<prop-modem>"& objRSDadosCla("Acf_ProprietarioEquip") &"</prop-modem>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<tp-aceite>"& objRSDadosCla("TP_TipoAceite") &"</tp-aceite>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<distribuidor>"& objRSDadosCla("Dst_desc") &"</distribuidor>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<plataforma>"& objRSDadosCla("Pla_TipoPlataforma") &"</plataforma>" & vbnewline
						xmlAcesso = xmlAcesso & 	"		<SiglaCentroCliente>"& objRSDadosCla("Aec_SiglaCentroCliente") &"</SiglaCentroCliente>" & vbnewline
						
						''' Retirada da interface que trata o paralelismo do SGAPlus na fase tratar acesso
						''' Solicitação de CTERESA conforme email 
						''' Assunto : Não paralelismo da fase Tratar Acesso com Configuração
						''' Data: 04/05/2010 
						''' Inicio
						xmlAcesso = xmlAcesso & 	"		<EstacaoConfiguracao>"& objRSDadosCla("est_config") &"</EstacaoConfiguracao>" & vbnewline
						''' FIM
						
						if oriSol_id = 6 then
							set objRS = db.execute ("select pro_id from cla_provedorConfig where Proconf_Visivel = 'S' and pro_id = " &  objRSDadosCla("pro_id") )
							While Not objRS.eof
							
								configAcesso = "N" 	
								objRS.MoveNext
							  
							Wend
							
							if configAcesso = "N" then
							
								xmlAcesso = xmlAcesso & 	"		<SAIP>N</SAIP>" & vbnewline
							else
								xmlAcesso = xmlAcesso & 	"		<SAIP>S</SAIP>" & vbnewline
							end if 
							
						end if
							
						if rede = "1" then
							xmlAcesso = xmlAcesso & 	"		<fila>"& objRSDadosCla("fac_fila") &"</fila>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<bastidor>"& objRSDadosCla("fac_bastidor") &"</bastidor>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<regua>"& objRSDadosCla("fac_regua") &"</regua>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<posicao>"& objRSDadosCla("fac_posicao") &"</posicao>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<dominio>"& objRSDadosCla("fac_dominio") &"</dominio>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<no>"& objRSDadosCla("fac_no") &"</no>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<slot>"& objRSDadosCla("fac_slot") &"</slot>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<porta>"& objRSDadosCla("fac_porta") &"</porta>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<time-slot>"& objRSDadosCla("fac_timeslot") &"</time-slot>" & vbnewline
						end if
						if rede = "2" then
							xmlAcesso = xmlAcesso & 	"		<tronco>"& objRSDadosCla("fac_tronco") &"</tronco>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<par>"& objRSDadosCla("fac_par") &"</par>" & vbnewline
						end if
							
						if rede = "4"  then
							xmlAcesso = xmlAcesso & 	"		<vlan>"& objRSDadosCla("fac_vlan") &"</vlan>" & vbnewline
						end if
						if rede = "5" then
							xmlAcesso = xmlAcesso & 	"		<vlan>"& objRSDadosCla("fac_vlan") &"</vlan>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<pe>"& objRSDadosCla("fac_pe") &"</pe>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<link>"& objRSDadosCla("fac_link") &"</link>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<porta>"& objRSDadosCla("fac_porta") &"</porta>" & vbnewline
						end if
						
						if rede = "13" then
							xmlAcesso = xmlAcesso & 	"		<vlan>"& objRSDadosCla("fac_vlan") &"</vlan>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<Svlan>"& objRSDadosCla("fac_svlan") &"</Svlan>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<pe>"& objRSDadosCla("fac_pe") &"</pe>" & vbnewline							
							xmlAcesso = xmlAcesso & 	"		<porta>"& objRSDadosCla("fac_porta") &"</porta>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<tipoporta>"& objRSDadosCla("fac_tipoporta") &"</tipoporta>" & vbnewline
						end if						
						
						if rede = "10" then
							xmlAcesso = xmlAcesso & 	"		<ip>"& objRSDadosCla("fac_IP") &"</ip>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<gateway>"& objRSDadosCla("fac_Gateway") &"</gateway>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<mascara>"& objRSDadosCla("fac_Mascara") &"</mascara>" & vbnewline
						end if
						if rede = "11" then
							'xmlAcesso = xmlAcesso & 	"		<vlan>"& objRSDadosCla("fac_vlan") &"</vlan>" & vbnewline
							'xmlAcesso = xmlAcesso & 	"		<pe>"& objRSDadosCla("fac_pe") &"</pe>" & vbnewline
							'xmlAcesso = xmlAcesso & 	"		<link>"& objRSDadosCla("fac_link") &"</link>" & vbnewline
							'xmlAcesso = xmlAcesso & 	"		<porta>"& objRSDadosCla("fac_porta") &"</porta>" & vbnewline
							'xmlAcesso = xmlAcesso & 	"		<Svlan>"& objRSDadosCla("fac_Svlan") &"</Svlan>" & vbnewline
							
							xmlAcesso = xmlAcesso & 	"		<vlan>"& objRSDadosCla("OntVlan_Nome") &"</vlan>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<pe>"& objRSDadosCla("OntVlan_PE") &"</pe>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<link></link>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<porta>"& objRSDadosCla("OntVlan_PortaOLT") &"</porta>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<Svlan>"& objRSDadosCla("OntSVlan_Nome") &"</Svlan>" & vbnewline
						end if
						
						if rede = "12" then
						
							'xmlAcesso = xmlAcesso & 	"		<vlan>"& objRSDadosCla("fac_vlan") &"</vlan>" & vbnewline
							'xmlAcesso = xmlAcesso & 	"		<pe>"& objRSDadosCla("fac_pe") &"</pe>" & vbnewline
							'xmlAcesso = xmlAcesso & 	"		<link>"& objRSDadosCla("fac_link") &"</link>" & vbnewline
							'xmlAcesso = xmlAcesso & 	"		<porta>"& objRSDadosCla("fac_porta") &"</porta>" & vbnewline
							'xmlAcesso = xmlAcesso & 	"		<Svlan>"& objRSDadosCla("fac_Svlan") &"</Svlan>" & vbnewline
							
							xmlAcesso = xmlAcesso & 	"		<fabricante>"& trim(objRSDadosCla("Font_Nome")) &"</fabricante>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<modelo>"& trim(objRSDadosCla("Tont_Modelo")) &"</modelo>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<designacao_SWITCH>"& objRSDadosCla("Switch") &"</designacao_SWITCH>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<porta>"& trim(objRSDadosCla("OntPorta_Porta")) &"</porta>" & vbnewline
							
							xmlAcesso = xmlAcesso & 	"		<vlan>"& trim(objRSDadosCla("OntVlan_Nome")) &"</vlan>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<Svlan>"& trim(objRSDadosCla("OntSVlan_Nome")) &"</Svlan>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<pe>"& trim(objRSDadosCla("OntVlan_PE")) &"</pe>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<portaPE>"& trim(objRSDadosCla("OntVlan_PortaOLT")) &"</portaPE>" & vbnewline
							
							xmlAcesso = xmlAcesso & 	"		<PORTA-SWITCH-CONCENTRADOR-LADO-METRO>"& trim(objRSDadosCla("portaSwitchLadoMetro")) &"</PORTA-SWITCH-CONCENTRADOR-LADO-METRO>" & vbnewline
							
							
						end if
						
						if ( objRSDadosCla("Sis_Desc") = "" or isnull(objRSDadosCla("Sis_Desc"))) then 
						
						'if rede = "14" then
						
						   set objRS = db.execute (" select distinct ltrim(rtrim(tipoEqui.tipoEqui_id)) as  tipoEqui_id, ltrim(rtrim(equi.DESIGNACAO)) as DESIGNACAO, ltrim(rtrim(ontporta.OntPorta_Porta)) as OntPorta_Porta , ltrim(rtrim(FABONT.Font_Nome)) as Font_Nome, ltrim(rtrim(TIPOONT.Tont_Modelo)) as Tont_Modelo, ltrim(rtrim(TPPORTA.TPPORTA_ABREV)) as TPPORTA_ABREV , ltrim(rtrim(prop.prop_desc)) as  prop_desc  , ltrim(rtrim(uplink)) as uplink " &_
										" from cla_equipamento	equi " &_
										" inner join  cla_tipo_equipamento  tipoEqui			on equi.TipoEqui_id = tipoEqui.TipoEqui_id " &_
										" inner join  cla_propriedade		  prop				on equi.prop_id = prop.prop_id " &_
										" inner join  cla_ontporta	ontporta				on equi.equi_id = ontporta.equi_id " &_
										" inner join  cla_porta_alocada ptAlocada				on ontporta.ontporta_id = ptAlocada.ontporta_id " &_	
										" INNER JOIN  CLA_TIPOONT	TIPOONT					ON ontporta.TONT_ID = TIPOONT.TONT_ID " &_
										" INNER JOIN  CLA_FABONT		FABONT					ON TIPOONT.FONT_ID = FABONT.FONT_ID " &_
										" INNER JOIN  CLA_TIPOPORTA	TPPORTA				ON ontporta.TPPORTA_ID = TPPORTA.TPPORTA_ID " &_
										" where ptAlocada.acl_idacessologico = " &  idLogico)
							
							While Not objRS.eof
														  
								'response.write "<script>alert('"&objRS("DESIGNACAO")&"')</script>"
							  
								if "2" = Trim(objRS("tipoEqui_id")) then 
								
									'FABRICANTE_ETHERNET					 	= objRS("Font_Nome")
									'MODELO_ETHERNET			 				= objRS("Tont_Modelo")
									PORTA_SWITCH_CONCENTRADOR_LADO_METRO 	= objRS("OntPorta_Porta")
									'DESIGNACAO_SWITCH_METRO 				= objRS("DESIGNACAO")
									PORTA_SWITCH_METRO					 	= objRS("OntPorta_Porta")
									TIPO_PORTA_SWITCH_METRO				 	= objRS("TPPORTA_ABREV")
									
								End if
								
								if "3" = Trim(objRS("tipoEqui_id")) then 
								
									'FABRICANTE-ETHERNET					 	= objRS("Font_Nome")
									'MODELO-ETHERNET			 				= objRS("Tont_Modelo")
									'PORTA-SWITCH-CONCENTRADOR-LADO-METRO 	= objRS("OntPorta_Porta")
									
									'DESIGNACAO_SWITCH_INTERCONEXAO			= objRS("DESIGNACAO")
									PORTA_SWITCH_INTERCONEXAO			 	= objRS("OntPorta_Porta")
									UPLINK_SWITCH_INTERCONEXAO			 	= objRS("uplink")
									'PROPRIEDADE_SWITCH_INTERCONEXAO		 	= objRS("prop_desc")
									
									'propswitchInterconexao	= objRS("prop_id") 
								End if
								
								if "4" = Trim(objRS("tipoEqui_id")) then 
									'DESIGNACAO_EDD_CLIENTE					= objRS("DESIGNACAO")
									PORTA_EDD_CLIENTE			 			= objRS("OntPorta_Porta")
									UPLINK_EDD_CLIENTE			 			= objRS("uplink")
									'PROPRIEDADE_EDD_CLIENTE				 	= objRS("prop_desc")
								End if
								
								if "9" = Trim(objRS("tipoEqui_id")) then 
									
									'DESIGNACAO_pe							= objRS("DESIGNACAO")
									PORTA_pe			 					= objRS("OntPorta_Porta")
									
								End if
								
								
								
							  objRS.MoveNext
							  
							Wend
							
							
							
							  set objRS = db.execute (" select	distinct	ltrim(rtrim(cla_equipamento.Designacao)) as Designacao , " &_
													  " 					ltrim(rtrim(FABONT.Font_Nome)) as Font_Nome, " &_
													  " 					ltrim(rtrim(TIPOONT.Tont_Modelo)) as Tont_Modelo, " &_
													  " 					ltrim(rtrim(cla_propriedade.prop_desc)) as  prop_desc, " &_
													  " 					ltrim(rtrim(cla_tipo_equipamento.tipoEqui_id)) as  tipoEqui_id " &_
										" from cla_Acessofisico_equipamento " &_	
										" inner join cla_equipamento						on cla_Acessofisico_equipamento.equi_id = cla_equipamento.equi_id " &_
										" inner join cla_tipo_equipamento					on cla_equipamento.tipoequi_id = cla_tipo_equipamento.tipoequi_id " &_
										" inner join cla_propriedade						on cla_equipamento.prop_id = cla_propriedade.prop_id " &_
										" left join  cla_ontporta	ontporta				on cla_equipamento.equi_id = ontporta.equi_id " &_
										" left JOIN  CLA_TIPOONT	TIPOONT					ON ontporta.TONT_ID = TIPOONT.TONT_ID " &_
										" left JOIN  CLA_FABONT		FABONT				ON TIPOONT.FONT_ID = FABONT.FONT_Id " &_
										" where cla_Acessofisico_equipamento.acf_id = " &  objRSDadosCla("acf_id"))
							
							While Not objRS.eof
														  
								'response.write "<script>alert('"&objRS("DESIGNACAO")&"')</script>"
							  
								if "2" = Trim(objRS("tipoEqui_id")) then 
								
									FABRICANTE_ETHERNET					 	= objRS("Font_Nome")
									MODELO_ETHERNET			 				= objRS("Tont_Modelo")
									'PORTA_SWITCH_CONCENTRADOR_LADO_METRO 	= objRS("OntPorta_Porta")
									DESIGNACAO_SWITCH_METRO 				= objRS("DESIGNACAO")
									'PORTA_SWITCH_METRO					 	= objRS("OntPorta_Porta")
									'TIPO_PORTA_SWITCH_METRO				 	= objRS("TPPORTA_ABREV")
									
								End if
								
								if "3" = Trim(objRS("tipoEqui_id")) then 
								
									'FABRICANTE-ETHERNET					 	= objRS("Font_Nome")
									'MODELO-ETHERNET			 				= objRS("Tont_Modelo")
									'PORTA-SWITCH-CONCENTRADOR-LADO-METRO 	= objRS("OntPorta_Porta")
									DESIGNACAO_SWITCH_INTERCONEXAO			= objRS("DESIGNACAO")
									'PORTA_SWITCH_INTERCONEXAO			 	= objRS("OntPorta_Porta")
									'UPLINK_SWITCH_INTERCONEXAO			 	= objRS("uplink")
									PROPRIEDADE_SWITCH_INTERCONEXAO		 	= objRS("prop_desc")
									
									'propswitchInterconexao	= objRS("prop_id") 
								End if
								
								if "4" = Trim(objRS("tipoEqui_id")) then 
									DESIGNACAO_EDD_CLIENTE					= objRS("DESIGNACAO")
									'PORTA_EDD_CLIENTE			 			= objRS("OntPorta_Porta")
									'UPLINK_EDD_CLIENTE			 			= objRS("uplink")
									PROPRIEDADE_EDD_CLIENTE				 	= objRS("prop_desc")
								End if
								
								if "9" = Trim(objRS("tipoEqui_id")) then 
									
									DESIGNACAO_pe							= objRS("DESIGNACAO")
									'PORTA_pe			 					= objRS("OntPorta_Porta")
									
								End if
								
								
								
							  objRS.MoveNext
							  
							Wend
							
							xmlAcesso = xmlAcesso & 	"		<PORTA-SWITCH-CONCENTRADOR-LADO-METRO>"& PORTA_SWITCH_CONCENTRADOR_LADO_METRO &"</PORTA-SWITCH-CONCENTRADOR-LADO-METRO>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<FABRICANTE-ETHERNET>"& FABRICANTE_ETHERNET &"</FABRICANTE-ETHERNET>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<MODELO-ETHERNET>"& MODELO_ETHERNET &"</MODELO-ETHERNET>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<DESIGNACAO-SWITCH-INTERCONEXAO>"& DESIGNACAO_SWITCH_INTERCONEXAO &"</DESIGNACAO-SWITCH-INTERCONEXAO>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<PROPRIEDADE-SWITCH-INTERCONEXAO>"& PROPRIEDADE_SWITCH_INTERCONEXAO &"</PROPRIEDADE-SWITCH-INTERCONEXAO>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<PORTA-SWITCH-INTERCONEXAO>"& PORTA_SWITCH_INTERCONEXAO &"</PORTA-SWITCH-INTERCONEXAO>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<UPLINK-SWITCH-INTERCONEXAO>"& UPLINK_SWITCH_INTERCONEXAO &"</UPLINK-SWITCH-INTERCONEXAO>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<DESIGNACAO-SWITCH-METRO>"& DESIGNACAO_SWITCH_METRO &"</DESIGNACAO-SWITCH-METRO>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<PORTA-SWITCH-METRO>"& PORTA_SWITCH_METRO &"</PORTA-SWITCH-METRO>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<TIPO-PORTA-SWITCH-METRO>"& TIPO_PORTA_SWITCH_METRO &"</TIPO-PORTA-SWITCH-METRO>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<PROPRIEDADE-EDD-CLIENTE>"& PROPRIEDADE_EDD_CLIENTE &"</PROPRIEDADE-EDD-CLIENTE>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<DESIGNACAO-EDD-CLIENTE>"& DESIGNACAO_EDD_CLIENTE &"</DESIGNACAO-EDD-CLIENTE>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<PORTA-EDD-CLIENTE>"& PORTA_EDD_CLIENTE &"</PORTA-EDD-CLIENTE>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<UPLINK-EDD-CLIENTE>"& UPLINK_EDD_CLIENTE &"</UPLINK-EDD-CLIENTE>" & vbnewline
						
							xmlAcesso = xmlAcesso & 	"		<vlan>"& objRSDadosCla("OntVlan_Nome") &"</vlan>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<pe>"& DESIGNACAO_pe &"</pe>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<porta>"& objRSDadosCla("OntVlan_PortaOLT")  &"</porta>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<Svlan>"& objRSDadosCla("OntSVlan_Nome") &"</Svlan>" & vbnewline
							
							
							
							
						end if
						
						xmlAcesso = xmlAcesso & 	"	</acesso> "
				   else
				   					
							xmlAcesso = xmlAcesso & 	"	<acesso> " & vbnewline
							xmlAcesso = xmlAcesso & 	"		<id-acessoFisico>"& objRSDadosCla("acf_idacessofisico") &"</id-acessoFisico>" & vbnewline
							'''xmlAcesso = xmlAcesso & 	"		<tecnologia>"& objRSDadosCla("tec_id") &"</tecnologia>" & vbnewline
							
							xmlAcesso = xmlAcesso & 	"		<tecnologia>"& tecnologia &"</tecnologia>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<tipo-contrato>"& objRSDadosCla("reg_id") &"</tipo-contrato>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<velocidade>"& objRSDadosCla("vel_desc") &"</velocidade>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<vel-conversao>"& objRSDadosCla("vel_conversao") &"</vel-conversao>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<proprietario>"& objRSDadosCla("acf_proprietario") &"</proprietario>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<dt-construcao>"& FormatarDataHora12H(objRSDadosCla("Acf_DtConstrAcessoFis")) &"</dt-construcao>" & vbnewline
							
							if objRSDadosCla("Acf_DtAceite") = "" then
								strAcf_DtAceite = date()
							else
								strAcf_DtAceite = objRSDadosCla("Acf_DtAceite")
							end if
							xmlAcesso = xmlAcesso & 	"		<dt-aceite>"& FormatarDataHora12H(strAcf_DtAceite) &"</dt-aceite>" & vbnewline
							if oriSol_id <> 7 and tecnologia <> "9" then
								xmlAcesso = xmlAcesso & 	"		<interface>"& objRSDadosCla("Acf_Interface") &"</interface>" & vbnewline
								xmlAcesso = xmlAcesso & 	"		<tp-vel>"& objRSDadosCla("acf_tipovel") &"</tp-vel>" & vbnewline
							end if 
							xmlAcesso = xmlAcesso & 	"		<cnl>"& objRSDadosCla("Acf_SiglaEstEntregaFisico") &"</cnl>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<compl-estacao>"& objRSDadosCla("Acf_ComplSiglaEstEntregaFisico") &"</compl-estacao>" & vbnewline
							if oriSol_id <> 7 and tecnologia <> "9" then
								xmlAcesso = xmlAcesso & 	"		<interface-estacao>"& objRSDadosCla("Acf_InterfaceEstEntregaFisico") &"</interface-estacao>" & vbnewline
							end if 





							xmlAcesso = xmlAcesso & 	"		<SiglaCentroCliente>"& objRSDadosCla("Aec_SiglaCentroCliente") &"</SiglaCentroCliente>" & vbnewline
							''' Retirada da interface que trata o paralelismo do SGAPlus na fase tratar acesso
							
							''' Solicitação de CTERESA conforme email 
							''' Assunto : Não paralelismo da fase Tratar Acesso com Configuração
							''' Data: 04/05/2010 
							''' Inicio
							xmlAcesso = xmlAcesso & 	"		<EstacaoConfiguracao>"& objRSDadosCla("est_config") &"</EstacaoConfiguracao>" & vbnewline
							''' FIM
						if oriSol_id = 6 then
								set objRS = db.execute ("select pro_id from cla_provedorConfig where Proconf_Visivel = 'S' and pro_id = " &  objRSDadosCla("pro_id") )
								While Not objRS.eof
								
									configAcesso = "N" 	
									objRS.MoveNext
								  
								Wend
								
								if configAcesso = "N" then
								
									xmlAcesso = xmlAcesso & 	"		<SAIP>N</SAIP>" & vbnewline
								else
									xmlAcesso = xmlAcesso & 	"		<SAIP>S</SAIP>" & vbnewline
								end if 
								
							end if	
						
						if tecnologia = "1" then
							' alteração uat SGAPlus 
							'xmlAcesso = xmlAcesso & 	"		<desig-banda-basica>"& objRSDadosCla("designacaoTronco") &"</desig-banda-basica>" & vbnewline
							' Envio correta da banda básica							                  
							if trim(objRSDadosCla("Acf_NroAcessoPtaEbt")) = "" then
							xmlAcesso = xmlAcesso & 	"		<desig-banda-basica>"& objRSDadosCla("designacaoTronco") &"</desig-banda-basica>" & vbnewline
							else
								xmlAcesso = xmlAcesso & 	"		<desig-banda-basica>"& objRSDadosCla("Acf_NroAcessoPtaEbt") &"</desig-banda-basica>" & vbnewline
							end if 
							
							




						 end if
						
						if tecnologia = "3" then
							xmlAcesso = xmlAcesso & 	"		<numero-acesso>"& objRSDadosCla("Acf_NroAcessoPtaEbt") &"</numero-acesso>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<ccto>"& objRSDadosCla("Acf_CCTOFatura") &"</ccto>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<qtd-modem>"& objRSDadosCla("Acf_QtdEquip") &"</qtd-modem>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<prop-modem>"& objRSDadosCla("Acf_ProprietarioEquip") &"</prop-modem>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<tp-aceite>"& objRSDadosCla("TP_TipoAceite") &"</tp-aceite>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<distribuidor>"& objRSDadosCla("Dst_desc") &"</distribuidor>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<rede>"& objRSDadosCla("Sis_Desc") &"</rede>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<cabo>"& objRSDadosCla("fac_tronco") &"</cabo>" & vbnewline ' tipo de acesso ADE fac_tronco e o cabo
							xmlAcesso = xmlAcesso & 	"		<par>"& objRSDadosCla("fac_par") &"</par>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<tp-cabo>"& objRSDadosCla("fac_tipoCabo")&"</tp-cabo>" & vbnewline
						end if
							
						'GPON
						if tecnologia = "6" then
							xmlAcesso = xmlAcesso & 	"		<fabricante>"& trim(objRSDadosCla("Font_Nome")) &"</fabricante>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<modelo>"& trim(objRSDadosCla("Tont_Modelo")) &"</modelo>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<designacao_ont>"& trim(objRSDadosCla("Ont_Desig")) &"</designacao_ont>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<porta>"& trim(objRSDadosCla("OntPorta_Porta")) &"</porta>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<vlan></vlan>" & vbnewline
							if oriSol_id = 7 then
								xmlAcesso = xmlAcesso & 	"		<vlan_voz>"& trim(objRSDadosCla("OntVlan_Nome")) &"</vlan_voz>" & vbnewline
							else
								xmlAcesso = xmlAcesso & 	"		<vlan_ont>"& trim(objRSDadosCla("OntVlan_Nome")) &"</vlan_ont>" & vbnewline
							end if
							xmlAcesso = xmlAcesso & 	"		<portaolt>"& trim(objRSDadosCla("OntVlan_PortaOLT")) &"</portaolt>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<peolt>"& trim(objRSDadosCla("OntVlan_PE")) &"</peolt>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<Svlan>"& objRSDadosCla("OntSVlan_Nome") &"</Svlan>" & vbnewline
						end if
						
						if tecnologia = "7" then
							xmlAcesso = xmlAcesso & 	"		<fabricante>"& trim(objRSDadosCla("Font_Nome")) &"</fabricante>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<modelo>"& trim(objRSDadosCla("Tont_Modelo")) &"</modelo>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<designacao_ont>"& trim(objRSDadosCla("Ont_Desig")) &"</designacao_ont>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<porta>"& trim(objRSDadosCla("OntPorta_Porta")) &"</porta>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<vlan>"& trim(objRSDadosCla("OntVlan_Nome")) &"</vlan>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<portaolt>"& trim(objRSDadosCla("OntVlan_PortaOLT")) &"</portaolt>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<peolt>"& trim(objRSDadosCla("OntVlan_PE")) &"</peolt>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<Svlan>"& objRSDadosCla("OntSVlan_Nome") &"</Svlan>" & vbnewline
							
							xmlAcesso = xmlAcesso & 	"		<switch>"& objRSDadosCla("Switch") &"</switch>" & vbnewline
							xmlAcesso = xmlAcesso & 	"		<porta-switch>"& objRSDadosCla("Switch_Porta") &"</porta-switch>" & vbnewline
						end if
						
							xmlAcesso = xmlAcesso & 	"	</acesso> "
				
					
					end if				   
				
			        objRSDadosCla.movenext
				wend
				
				end if
				
				Strxml1 = Strxml1 & 	xmlAcesso & vbnewline
				
				Strxml1 = Strxml1 & 	"		</acessos-fisicos>" & vbnewline
				Strxml1 = Strxml1 & 	"	</retorno-cla>" & vbnewline
			'db.execute("insert into cla_temp values('" & Strxml1 & "')")
			
				set ConSGA = Server.CreateObject("ADODB.Command")
				
				If Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO913" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.21" or Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO912" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.17" then
  					StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_Tipo = 'PRD' and OriSol_ID = " & OriSol_ID
				else
  					StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_Tipo = 'DSV' and OriSol_ID = " & OriSol_ID
				end if
				
				Set objRS = db.Execute(StrSQL)
				If Not objRS.eof and  not objRS.Bof Then
					objConSGA = objRS("Conn_Desc")
				End if
				
				ConSGA.ActiveConnection = objConSGA
				
				if oriSol_id = 6 then
					ConSGA.CommandText = "sgaplus_adm.pck_sgap_interface_cla.pc_retorno_solicitacao_cla"
				end if
				if oriSol_id = 7 then
					ConSGA.CommandText = "sgav_vips.sp_sgav_interface_cla"
				end if

				ConSGA.CommandType = adCmdStoredProc
				
				'*** Carregando parâmetros de entrada
				Set objParam = ConSGA.CreateParameter("p1", adNumeric, adParamInput, 10, idTarefa)
				ConSGA.Parameters.Append objParam
				
				Set objParam = ConSGA.CreateParameter("p2", adLongVarWChar, adParamInput, 1073741823, Strxml1)
				ConSGA.Parameters.Append objParam
				
				'*** Configurando variável que receberá o retorno
				Set objParam = ConSGA.CreateParameter("Ret1", adNumeric, adParamOutput, 10)
				ConSGA.Parameters.Append objParam
				
				Set objParam = ConSGA.CreateParameter("Ret2", adVarChar, adParamOutput, 100 )
				ConSGA.Parameters.Append objParam
				
				'Tratamento de erro crítico:
				On error resume next
				
				'*** Executando a stored procedure
				ConSGA.Execute
				
				if err.number <> 0 then
					strxmlResp = "ERRO Critico: " & err.number & " - " & err.description
					On Error GoTo 0
					
					'Checa se serviço é 0800 - E.
					Vetor_Campos(1)="adVarchar,4,adParamInput," 	& OE_Ano
					Vetor_Campos(2)="adVarchar,7,adParamInput," 	& OE_Numero
					Vetor_Campos(3)="adVarchar,3,adParamInput," 	& OE_Item
					Vetor_Campos(4)="adVarchar,20,adParamInput," 	& idTarefa
					Vetor_Campos(5)="adVarchar,20,adParamInput," 	& OriSol_Descricao
					Vetor_Campos(6)="adVarchar,10,adParamInput," 	& Acao
					Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
					Vetor_Campos(8)="adVarchar,200,adParamInput," 	& strxmlResp
					Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& Strxml1
					Vetor_Campos(10)="adInteger,1,adParamInput,4" 'Construir Return
					Vetor_Campos(11)="adInteger,1,adParamInput,1"
					Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
					
					strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
					db.Execute(strSqlRet)
					
				    response.write "<script>alert('Erro crítico no " & OriSol_Descricao & "')</script>"
					response.end
				end if
				
				cod_retorno  = ConSGA.Parameters("Ret1").value
				desc_retorno = ConSGA.Parameters("Ret2").value
				
				strxmlResp = cod_retorno & " - " & desc_retorno
				
				if cod_retorno = 0 then
					Vetor_Campos(1)="adInteger,4,adParamInput," & solid
					Vetor_Campos(2)="adInteger,4,adParamInput, 267"
					Vetor_Campos(3)="adInteger,4,adParamInput," & strloginrede
					Vetor_Campos(4)="adVarchar,1,adParamInput,"
					Vetor_Campos(5)="adVarchar,100,adParamInput,STATUS AUTOMATICO"  
					Vetor_Campos(6)="adVarchar,1,adParamInput,M"
					
  					strSqlRet = APENDA_PARAMSTR("CLA_sp_ins_StatusSolicitacao",6,Vetor_Campos)
					
					db.Execute(strSqlRet)
					
					Vetor_Campos(1)="adInteger,4,adParamInput," & Aprovisi_ID
					Vetor_Campos(2)="adVarchar,20,adParamInput, Entregar"
					strSqlRet = APENDA_PARAMSTR("CLA_sp_interface_status",2,Vetor_Campos)
					db.Execute(strSqlRet)
					
					'Checa se serviço é 0800.
					Vetor_Campos(1)="adVarchar,4,adParamInput," 	& OE_Ano
					Vetor_Campos(2)="adVarchar,7,adParamInput," 	& OE_Numero
					Vetor_Campos(3)="adVarchar,3,adParamInput," 	& OE_Item
					Vetor_Campos(4)="adVarchar,20,adParamInput," 	& idTarefa
					Vetor_Campos(5)="adVarchar,20,adParamInput," 	& OriSol_Descricao
					Vetor_Campos(6)="adVarchar,10,adParamInput," 	& Acao
					Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
					Vetor_Campos(8)="adVarchar,200,adParamInput," 	& strxmlResp
					Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& Strxml1
					Vetor_Campos(10)="adInteger,1,adParamInput,4" 'Construir Return
					Vetor_Campos(11)="adInteger,1,adParamInput,0"
					Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
					
					strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
					db.Execute(strSqlRet)
				else
					'Checa se serviço é 0800 - E.
					Vetor_Campos(1)="adVarchar,4,adParamInput," 	& OE_Ano
					Vetor_Campos(2)="adVarchar,7,adParamInput," 	& OE_Numero
					Vetor_Campos(3)="adVarchar,3,adParamInput," 	& OE_Item
					Vetor_Campos(4)="adVarchar,20,adParamInput," 	& idTarefa
					Vetor_Campos(5)="adVarchar,20,adParamInput," 	& OriSol_Descricao
					Vetor_Campos(6)="adVarchar,10,adParamInput," 	& Acao
					Vetor_Campos(7)="adInteger,4,adParamInput," 	& Aprovisi_ID
					Vetor_Campos(8)="adVarchar,200,adParamInput," 	& strxmlResp
					Vetor_Campos(9)="adVarchar,8000,adParamInput," 	& Strxml1
					Vetor_Campos(10)="adInteger,1,adParamInput,4" 'Construir Return
					Vetor_Campos(11)="adInteger,1,adParamInput,1"
					Vetor_Campos(12)="adNumeric,10,adParamInput," & dblIdLogico
					
					strSqlRet = APENDA_PARAMSTR("CLA_sp_check_servico2",12,Vetor_Campos)
					db.Execute(strSqlRet)
					
					Set objMail = CreateObject("CDONTS.NewMail")
					gMailTo1 =  "prssilv@embratel.com.br,edar@embratel.com.br"
					objMail.To = gMailTo1
					objMail.From = From
					objMail.Subject = Sbj4
					objMail.Body = Data & " - " & Hora & " > - " & dblIdLogico & " - " & strxmlResp
					objMail.Send
					Set objMail = Nothing
				end if
				
			end if
		   
	 end if
	 	   
End Function
%>