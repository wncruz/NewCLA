<!--#include file="../inc/data.asp"-->
<Html>
	<Head>
		<link rel=stylesheet type="text/css" href="../css/cla.css">
		<script language='javascript' src="../javascript/cla.js"></script>
		<meta http-equiv="content-type" content="text/html; charset=UTF-8">
	</Head>
	<Body topmargin=0 leftmargin=0>
	
			<SCRIPT LANGUAGE=javascript>
			
				function GravarDadosTecnico()
				{
					with (document.forms[0])
					{
						
						
						if (!ValidarCampos(cboVelocidade,"Taxa de Transmissão")) return
						if (!ValidarCampos(txtQtdLinhas,"Quantidade de Linhas")) return

						if (!ValidarCampos(cboFinalidade,"Finalidade")) return
						if (!ValidarCampos(cboPrazContr,"Prazo de Contratação")) return
						if (!ValidarCampos(cboVelocidade,"Taxa de Transmição")) return
						if (!ValidarCampos(cboCaracTec,"Característica Técnica")) return
						if (!ValidarCampos(cboAplicacao,"Aplicação")) return
						if (!ValidarCampos(cboMeioPref,"Meio Preferencial")) return
						
						hdnAcao.value = "GravarDadosTecnico"
						hdnGravouDadosTecnico.value = "Sim"
						target = "IFrmDadosTecnicoOperadora"
						action = "DadosTecnicoOperadora.asp"
						submit()
					}
				}

			
			</SCRIPT>
	
			<form name=Form1 method=Post >
				<input type=hidden name=hdnAcao>
				<input type=hidden name=hdnGravouDadosTecnico>
				<input type=hidden name=hdnAcf_ID value="<%=Request.QueryString("dblAcfId")%>">
				<input type=hidden name=hdnSol_ID value="<%=Request.QueryString("dblSolId")%>">
				<input type=hidden name=hdnEILD value="<%=Request.QueryString("dblEild")%>">
				<input type=hidden name=hdnPonta value="<%=Request.QueryString("dblPonta")%>">
				
				<%
				
					If	Trim(Request.form("hdnAcao")) = "GravarDadosTecnico" Then 
					
					
						str_acf_id = Request.Form("hdnAcf_ID")
						str_vel_idTransmissao = Trim(Request.Form("cboVelocidade")) 
						str_aplicacao_id = Trim(Request.Form("cboAplicacao"))
						str_meioPreferencial_id = Trim(Request.Form("cboMeioPref"))
						str_caracteristicaTecnica_id = Trim(Request.Form("cboCaracTec"))
						str_QtdLinha = Trim(Request.Form("txtQtdLinhas"))
						str_QinQ = Trim(Request.Form("rdoQinQ"))
						str_Prazo_contratacao_id = Trim(Request.Form("cboPrazContr"))
						str_Finalidade_id = Trim(Request.Form("cboFinalidade"))
						
					
						'NOVAS PROCEDURES - SNOA - INICIO'
						Vetor_Campos(1)="adInteger,10,adParamInput," & str_Finalidade_id
						Vetor_Campos(2)="adInteger,10,adParamInput," & str_Prazo_contratacao_id
						Vetor_Campos(3)="adInteger,10,adParamInput," & str_vel_idTransmissao
						Vetor_Campos(4)="adInteger,10,adParamInput," & str_caracteristicaTecnica_id
						Vetor_Campos(5)="adWChar,1,adParamInput," & str_QinQ
						Vetor_Campos(6)="adInteger,4,adParamInput," & str_QtdLinha
						Vetor_Campos(7)="adInteger,10,adParamInput," & str_aplicacao_id
						Vetor_Campos(8)="adInteger,10,adParamInput," & str_meioPreferencial_id
						Vetor_Campos(9)="adInteger,10,adParamInput," & str_acf_id
						Vetor_Campos(10)="adInteger,2,adParamInput," 
						Vetor_Campos(11)="adWChar,100,adParamInput,null" 
						Vetor_Campos(12)="adInteger,2,adParamInput," 
						Vetor_Campos(13)="adWChar,100,adParamInput,null"
						Vetor_Campos(14)="adInteger,2,adParamInput,"
						Vetor_Campos(15)="adInteger,2,adParamInput," 
						Vetor_Campos(16)="adInteger,2,adParamInput," 
						Vetor_Campos(17)="adInteger,4,adParamOutput,0"
						Vetor_Campos(18)="adWChar,100,adParamOutput,"
						
						
						Call APENDA_PARAM("CLA_sp_ins_AssocFisicoTecnico",18,Vetor_Campos)
						ObjCmd.Execute'pega dbaction
						
						DBErro = ObjCmd.Parameters("RET").value
						
						if DBErro = 0 then
							Response.Write "<script language=javascript>alert('Registro Gravado com Sucesso !');</script>"
							
						end if
						
					
					else 
					
						Dim AcfID
						Dim strDtPrevEntregaProvedor
						Dim strDtReaprPrevEntregaProvedor
						Dim StrProjetoEspecial
						Dim StrProjEsp_NumeroProposta
						Dim StrProjEsp_Prazo

						AcfID = Request.QueryString("dblAcfId")
						
						'NOVAS PROCEDURES - SNOA - INICIO'
						Vetor_Campos(1)="adInteger,2,adParamInput," & AcfID
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

						'<!-- CH-83646VWR - Inicio -->
						if trim(str_QinQ) = "" then str_QinQ="N"
					
					end if
				
				%>
				
				<iframe	id			= "IFrmDadosTecnicoOperadora"
						name        = "IFrmDadosTecnicoOperadora"
						width       = "0"
						height      = "0"
						frameborder = "0"
						scrolling   = "auto"
						align       = "left">
				</iFrame>

				<table  border=0 cellspacing="1" cellpadding="0"width="100%">
					<tr>
						<th style="FONT-SIZE: 14px" colspan=6><font class="clsObrig">:: </font>Dados Técnicos</th>
					</tr>

					<tr class=clsSilver>
						<td nowrap width=170><font class="clsObrig">:: </font>Taxa de Transmissão</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboVelocidade">
							<% if bbloqueia="" then%>
								<select name="cboVelocidade" style="width:200px">
							<%else%>
								<select name="cboVelocidade" style="width:200px" disabled="disabled">
							<%end if%>
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

						<td align=right ><font class="clsObrig">:: </font>Característica Técnica &nbsp;</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboCaracTec">
							<% if bbloqueia="" then%>
								<select name="cboCaracTec" style="width:200px">
							<%else%>
								<select name="cboCaracTec" style="width:200px" disabled="disabled">
							<%end if%>
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
						<td nowrap width=170><font class="clsObrig">:: </font>&nbsp;Q in Q</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboQinQ">
							<% if bbloqueia="" then%>
								<input type="radio" name="rdoQinQ" onClick="javascript:document.Form1.hdnQinQ.value = 'S';" value="S"
								<%if str_QinQ = "S" then%> checked <%end if%>>&nbsp; Sim
			
								<input type="radio" name="rdoQinQ" onClick="javascript:document.Form1.hdnQinQ.value = 'N';" value="N"
								<%if str_QinQ <> "S" then%> checked <%end if%>>&nbsp; Não
							<%else%>
								<input type="radio" name="rdoQinQ" value="S" disabled="disabled"
								<%if str_QinQ = "S" then%> checked <%end if%>>&nbsp; Sim
			
								<input type="radio" name="rdoQinQ" value="N" disabled="disabled"
								<%if str_QinQ <> "S" then%> checked <%end if%>>&nbsp; Não
							<%end if%>
							</span>
						</td>

						<td align=right ><font class="clsObrig">:: </font>Quantidade de Linhas</td>
						<td>&nbsp;
							<input type="text" class="text" name="txtQtdLinhas" value="1" size="30" maxlength="20"  readonly="true"
							onBlur="ValidarTipo(this,0)">
						</td>
					</tr>

					<tr class=clsSilver>
						<td nowrap width=170><font class="clsObrig">:: </font>&nbsp;Meio Preferencial</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboMeioPref">
							<% if bbloqueia="" then%>
								<select name="cboMeioPref" style="width:200px">
							<%else%>
								<select name="cboMeioPref" style="width:200px" disabled="disabled">
							<%end if%>
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

						<td align=right ><font class="clsObrig">:: </font>Aplicação &nbsp;</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboAplicacao">
							<% if bbloqueia="" then%>
								<select name="cboAplicacao" style="width:200px">
							<%else%>
								<select name="cboAplicacao" style="width:200px" disabled="disabled">
							<%end if%>
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
						<td width=170><font class="clsObrig">:: </font>Finalidade</td>
						<td>&nbsp;
							<span id="solicPedSnoaCboFinalidade">
							<% if bbloqueia="" then%>
								<select name="cboFinalidade" style="width:200px">
							<%else%>
								<select name="cboFinalidade" style="width:200px" disabled="disabled">
							<%end if%>
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

						<td align=right><font class="clsObrig">:: </font>Prazo de Contratação</td>
						<td colspan=3 nowrap>&nbsp;
							<span id="solicPedSnoaCboPrazContr">
							<% if bbloqueia="" then%>
								<select name="cboPrazContr" style="width:200px">
							<%else%>
								<select name="cboPrazContr" style="width:200px" disabled="disabled">
							<%end if%>
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
						<td colspan=6>
							<P style="FONT-SIZE: 9pt">
								A Característica Técnica definida pela fornecedora conforme disponibilidade constante no contrato, e não necessariamente, a que foi requisitada neste pedido.
								<br>
								A Característica Técnica deve seguir conforme disponibilidade da Oferta de Referência.
							</p>
						</td>
					</tr>
					<tr class=clsSilver >
						<td align=center colspan=6>
							<input type=button name=btnGravar value="Gravar Dados Técnicos" class=button style="width:150px;height:22px" onclick="GravarDadosTecnico()" >
							<!--onclick="EnviarEmail(0,document.forms[0].cboProvedor(document.forms[0].cboProvedor.selectedIndex).tag_provedor)" >
							 -->
						</td>
					</tr>
				</table>
			</Form>
	</Body>
</Html>