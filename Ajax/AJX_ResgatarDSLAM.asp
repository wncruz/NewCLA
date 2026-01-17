<!--#include file="../inc/data.asp"-->


<%

	Response.Expiresabsolute = Now() - 2
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	
	Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")
	Set objXmlRetorno = Server.CreateObject("Microsoft.XMLDOM")
	objXmlDadosForm.load(Request)
	
	
	paramID = objXmlDadosForm.selectSingleNode("//param").text
	strIDLogico = objXmlDadosForm.selectSingleNode("//param2").text
	strProcesso = objXmlDadosForm.selectSingleNode("//param3").text
	strDesignacao = objXmlDadosForm.selectSingleNode("//param5").text
	
	
	
	strEqui		= objXmlDadosForm.selectSingleNode("//param4").text
	
	propswitchEDD = ""
	
	strUplink = "uplink" + strEqui
	
	rdoPortaUplinkID = "rdoPortaUplinkID" + strEqui
	
	'response.write paramID
				
	set objRS = db.execute("CLA_sp_sel_Equipamento " & paramID )
	
	'response.write "<script language=javascript> parent.AddChk();</script>"
	'response.write "<script language=javascript> alert(1)</script>"
	
	if not objRS.eof then	
	
		if objRS("cvlan") = false then
			disabledCvlan = " disabled "
		else
			disabledCvlan = "  "
		end if 
		
		if objRS("svlan")  = false then
			disabledSvlan = " disabled  "
		else
			disabledSvlan = "  "
		end if 
	
		strRet = ""
		
		strRet = strRet & " <script> "
		
	
		
		strRet = strRet & " function LimparUplink(IDrdo)"
		
		strRet = strRet & "  { "
			
							
		strRet = strRet & "		for (var intIndex=0;intIndex<document.Form2.elements.length;intIndex++) "
		strRet = strRet & " 		{ "
		strRet = strRet & "				var elemento = document.Form2.elements[intIndex]; "
		strRet = strRet & "				if (elemento.name == IDrdo ){ "
			
		strRet = strRet & "					elemento.disabled=false; "
		strRet = strRet & "					elemento.focus();  "
		strRet = strRet & "				} "
		strRet = strRet & "			} "


					
		strRet = strRet & " } " 

		strRet = strRet & " </script> "
		
	
	
	
		strRet = strRet & " <table rules=groups  border=0 cellspacing=1 cellpadding=0 bordercolorlight=#003388 bordercolordark=#ffffff width=760 > "
		
		
	
	
	
		strRet = strRet & "		<tr> "
		strRet = strRet & "			<th colspan=10 class=clsSilver >&nbsp;&nbsp; Controle de Ocupacao</th> "
		strRet = strRet & "		</tr> "
		strRet = strRet & "		<tr class=clsSilver > "
		strRet = strRet & "			<td>&nbsp;&nbsp;Fabricante</td><td>&nbsp;&nbsp; " &  objRS("Font_Nome") & " </td><td>&nbsp;&nbsp;</td> "
		strRet = strRet & "		</tr> "
		strRet = strRet & "		<tr class=clsSilver > "
		strRet = strRet & "			<td>&nbsp;&nbsp;Modelo</td><td>&nbsp;&nbsp; " & objRS("Tont_Modelo") & " </td><td>&nbsp;&nbsp;</td> "
		strRet = strRet & "		</tr> "
		strRet = strRet & "		<tr class=clsSilver > "
		strRet = strRet & "			<td>&nbsp;&nbsp;Portas</td><td>&nbsp;&nbsp; " & objRS("Tont_QtdPorta") & " </td><td>&nbsp;&nbsp;</td> "
		strRet = strRet & "		</tr> "			
		
		strRet = strRet & " </table> "
		
		
		strRet = strRet & " <table rules=groups  border=0 cellspacing=1 cellpadding=0 bordercolorlight=#003388 bordercolordark=#ffffff width=760> "
		strRet = strRet & " 	<tr> "
		strRet = strRet & " 		<th width=10>&nbsp;&nbsp;Porta DSLAM</th> "
		
		strRet = strRet & " 		<th>&nbsp;&nbsp;Número Acesso</th> "
		strRet = strRet & " 		<th>&nbsp;&nbsp;Cabo</th> "
		strRet = strRet & " 		<th>&nbsp;&nbsp;Par</th> "
		strRet = strRet & " 		<th>&nbsp;&nbsp;PADE</th> "
		strRet = strRet & " 		<th>&nbsp;&nbsp;Derivação</th> "
		strRet = strRet & " 		<th>&nbsp;&nbsp;T. Cabo</th> "
		strRet = strRet & " 		<th>&nbsp;&nbsp;Propriedade</th> "
		'strRet = strRet & " 		<th>&nbsp;&nbsp;OBS</th> "
		strRet = strRet & " 		<th >&nbsp;&nbsp;</th> "
				
		strRet = strRet & " 	</tr> "
		
		i=1
		blnPorta = false
		
		While Not objRS.Eof
				if blnCor then
					strRet = strRet & " <tr class=clsSilver> "
					blnCor = false
				Else
					strRet = strRet & " <tr class=clsSilver2>"
					blnCor = true
				End if
			
				'strUplink = strUplink  
			
				IF ISNULL(objRS("Uplink")) THEN
				
					strRet = strRet & " <input type=Hidden name=hdnFACID"& i & " > "
					
					strPorta = ""
					strPorta = i & "&&" & objRS("ONTPorta_ID")
				
					strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
					strRet = strRet & " <td>&nbsp;&nbsp; </td> "
					strRet = strRet & " <td width=100><input type=text class=text name=txtCabo"& i & " ID=txtCabo" & i & "	readonly=true maxlength=2	 size=5	 onKeyUp=""ValidarTipo(this,0)"" ></td> "
					strRet = strRet & "<td width=120><input type=text class=text name=txtPar"& i & "		maxlength=9  size=10 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N ></td>"
					strRet = strRet & "<td><input type=text class=text name=txtCaixaEmenda"& i & "  maxlength=3 size=5 ></td>"
					strRet = strRet & "<td width=100><input type=text class=text name=txtLateral"& i & "	maxlength=4  size=5></td>"
					strRet = strRet & "<td width=100><select name=cboTipoCabo"& i & " >"
					
					if Trim(objRS("fac_tipoCabo")) = "" OR ISNULL(Trim(objRS("fac_tipoCabo")))  then 
							strRet = strRet &  "<Option value="""" ></Option>"
							strRet = strRet &  "<Option value=""H"" >H</Option>"
							strRet = strRet &  "<Option value=""X"" >X</Option>"
							strRet = strRet &  "<Option value=""C"" >C</Option>"
					End if						
					if Trim(objRS("fac_tipoCabo")) = "H" then 
							strItemSel = " Selected " 
							strRet = strRet &  "<Option value="""" ></Option>"
							strRet = strRet &  "<Option value=""H"" " & strItemSel & ">H</Option>"
							strRet = strRet &  "<Option value=""X"" >X</Option>"
							strRet = strRet &  "<Option value=""C"" >C</Option>"
					End if	
					
					if Trim(objRS("fac_tipoCabo")) = "X" then 
							strItemSel = " Selected " 
							strRet = strRet &  "<Option value="""" ></Option>"
							strRet = strRet &  "<Option value=""H"" >H</Option>"
							strRet = strRet &  "<Option value=""X"" " & strItemSel & ">X</Option>"
							strRet = strRet &  "<Option value=""C"" >C</Option>"
					End if	
					
					if Trim(objRS("fac_tipoCabo")) = "C" then 
							strItemSel = " Selected " 
							strRet = strRet &  "<Option value="""" ></Option>"
							strRet = strRet &  "<Option value=""H"" >H</Option>"
							strRet = strRet &  "<Option value=""X"" >X</Option>"
							strRet = strRet &  "<Option value=""C""  " & strItemSel & ">C</Option>"
					End if	
					
					
					'''strRet = strRet & "			<option value=""""></option>"
					'''strRet = strRet & "			<option value=""H"" >H</option>"
					'''strRet = strRet & "			<option value=""X"" >X</option>"
					'strAde = strAde & "			<option value=XDSL >XDSL</option>"
					'''strRet = strRet & "			<option value=""C"" >C</option>"
					
					strRet = strRet & "		</select></td>"
					
					
					strRet = strRet & " <td>"
					strRet = strRet & "		<select name=cboPropriedadeCabo"& i & " >  "
										
										
										set objRSProp = db.execute("CLA_sp_sel_PropriedadeDslam ")
										
										'strRet = strRet & "			<option value=""""></option>"			
										While Not objRSProp.eof
											strItemSel = ""
											if Trim(objRSProp("PROP_ID")) = "2" then strItemSel = " Selected " End if
										   if Trim(objRS("propriedade_porta")) = Trim(objRSProp("PROP_ID")) then strItemSel = " Selected " End if
										   
										
											strRet = strRet & "	<Option " & strItemSel & " value= " & objRSProp("PROP_ID") &  " > " & trim(objRSProp("PROP_SIGLA")) &  " </Option> "
											
											objRSProp.MoveNext
										Wend
										
					strRet = strRet & "		</select> "
					strRet = strRet & " </td> "
					
				'	strRet = strRet & " <td> <input type=text class=text disabled=true name="& strUplink & i & " ID=uplink" & i & " size=20 maxlength=20 value=" & objRS("Uplink") & " ></td> " 
									
					strRet = strRet & " <td> "
					
					'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )"";objAryFac[0][0]='';PosicoesLivre('L')"" > "
					'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""objAryFac[0][0]='';PosicoesLivre('L')"" > "
					
					strRet = strRet & " <input type=checkbox name=" & rdoPortaUplinkID & " title=Editar value=" & strPorta & " onclick=""parent.LimparDSLAM(" & i & ");AddSelecaoChk(this)"" > "
					strRet = strRet & " </td> "
					
					'onblur="CompletarCampo(this);hdnOrderEntryAno.value=this.value;" onkeyup="ValidarTipo(this,0)" 
					
					strRet = strRet & " </tr> "
					
				else
					
					strIDLogico678 = "678" & mid(strIDLogico,4,7)
					strIDLogico677 = "677" & mid(strIDLogico,4,7)
						
					if ( Trim(objRS("Acl_IDAcessoLogico")) = strIDLogico678 or Trim(objRS("Acl_IDAcessoLogico")) = strIDLogico677  or  Trim(objRS("uplink")) = trim(strDesignacao) ) then
					
					
						strRet = strRet & " <input type=Hidden name=hdnFACID"& i & " value= " & objRS("fac_id") & "> "
						
						strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
						strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("uplink")  & " </td> "
						strRet = strRet & " <td width=100><input type=text class=text name=txtCabo"& i & " ID=txtCabo" & i & "	readonly=true maxlength=2	 size=5	 onKeyUp=""ValidarTipo(this,0)"" value=" & objRS("fac_tronco") & "></td> "
						strRet = strRet & "<td width=120><input type=text class=text name=txtPar"& i & "		maxlength=9  size=10 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N value=" & objRS("fac_par") & " ></td>"
						strRet = strRet & "<td><input type=text class=text name=txtCaixaEmenda"& i & "  maxlength=3 size=5 value=" & objRS("fac_cxemenda") & " ></td>"
						strRet = strRet & "<td width=100><input type=text class=text name=txtLateral"& i & "	maxlength=4  size=5 value=" & objRS("fac_lateral") & "></td>"
						strRet = strRet & "<td width=100><select name=cboTipoCabo"& i & " >"
						
						if Trim(objRS("fac_tipoCabo")) = "" OR ISNULL(Trim(objRS("fac_tipoCabo"))) then 
								strRet = strRet &  "<Option value="""" ></Option>"
								strRet = strRet &  "<Option value=""H"" >H</Option>"
								strRet = strRet &  "<Option value=""X"" >X</Option>"
								strRet = strRet &  "<Option value=""C"" >C</Option>"
						End if						
						if Trim(objRS("fac_tipoCabo")) = "H" then 
								strItemSel = " Selected " 
								strRet = strRet &  "<Option value="""" ></Option>"
								strRet = strRet &  "<Option value=""H"" " & strItemSel & ">H</Option>"
								strRet = strRet &  "<Option value=""X"" >X</Option>"
								strRet = strRet &  "<Option value=""C"" >C</Option>"
						End if	
						
						if Trim(objRS("fac_tipoCabo")) = "X" then 
								strItemSel = " Selected " 
								strRet = strRet &  "<Option value="""" ></Option>"
								strRet = strRet &  "<Option value=""H"" >H</Option>"
								strRet = strRet &  "<Option value=""X"" " & strItemSel & ">X</Option>"
								strRet = strRet &  "<Option value=""C"" >C</Option>"
						End if	
						
						if Trim(objRS("fac_tipoCabo")) = "C" then 
								strItemSel = " Selected " 
								strRet = strRet &  "<Option value="""" ></Option>"
								strRet = strRet &  "<Option value=""H"" >H</Option>"
								strRet = strRet &  "<Option value=""X"" >X</Option>"
								strRet = strRet &  "<Option value=""C""  " & strItemSel & ">C</Option>"
						End if	
						
						'''strRet = strRet & "			<option value=""""></option>"
						'''strRet = strRet & "			<option value=""H"" >H</option>"
						'''strRet = strRet & "			<option value=""X"" >X</option>"
						'strAde = strAde & "			<option value=XDSL >XDSL</option>"
						'''strRet = strRet & "			<option value=""C"" >C</option>"
						strRet = strRet & "		</select></td>"
						
						
						strRet = strRet & " <td>"
						strRet = strRet & "		<select name=cboPropriedadeCabo"& i & " >  "
											
											
											set objRSProp = db.execute("CLA_sp_sel_PropriedadeDslam ")
											
											'strRet = strRet & "			<option value=""""></option>"			
											While Not objRSProp.eof
												strItemSel = ""
												if Trim(objRSProp("PROP_ID")) = 2 then strItemSel = " Selected " End if
												
												if Trim(objRS("propriedade_porta")) = Trim(objRSProp("PROP_ID")) then strItemSel = " Selected " End if
											   
												
												strRet = strRet & "	<Option " & strItemSel & " value= " & objRSProp("PROP_ID") &  " > " & trim(objRSProp("PROP_SIGLA")) &  " </Option> "
												 
												objRSProp.MoveNext
											Wend
											
						strRet = strRet & "		</select> "
						strRet = strRet & " </td> "
						
						'strRet = strRet & " <td> "
						'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar checked value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )"" )  > "
						'strRet = strRet & " </td> "
						
						strRet = strRet & " <td> "
					
						strPorta = ""
						strPorta = i & "&&" & objRS("ONTPorta_ID")
						
						'alert(strPorta)
						'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )"";objAryFac[0][0]='';PosicoesLivre('L')"" > "
						'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""objAryFac[0][0]='';PosicoesLivre('L')"" > "
						
						''''strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar checked value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparDSLAM(" & i & ")"" > "
						
						'strRet = strRet & " <input type=checkbox name=" & rdoPortaUplinkID & " title=Editar checked value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparDSLAM(" & i & ");AddSelecaoChk(this)"" > "
						
						'parent.parent.AddChk()
						
						strRet = strRet & " <input type=checkbox name=" & rdoPortaUplinkID & " title=Editar checked value=" & strPorta & " onclick=""parent.LimparDSLAM(" & i & ");AddSelecaoChk(this)"" > "
						
						strRet = strRet & " </td> "
						
						strRet = strRet & " </tr> "
						
						
						'strRet = strRet & " <script> AddChk(); </script> "
					
					else
						
						strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
						strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("uplink")  & " </td> "
						strRet = strRet & " <td width=100>&nbsp;&nbsp;" & objRS("fac_tronco")  & "</td> "
						strRet = strRet & "<td width=120>&nbsp;&nbsp;" & objRS("fac_par")  & "</td>"
						strRet = strRet & "<td width=100>&nbsp;&nbsp;" & objRS("fac_cxEmenda")  & "</td>"
						strRet = strRet & "<td>&nbsp;&nbsp;" & objRS("fac_lateral")  & "</td>"						
						strRet = strRet & "<td width=100>&nbsp;&nbsp;" & objRS("fac_tipoCabo")  & "</td>"
						strRet = strRet & "<td width=100>&nbsp;&nbsp;" & objRS("PROP_SIGLA")  & "</td>"
						
						'strRet = strRet & " <td> "
						
						'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar disabled value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )"" )  > "
						'strRet = strRet & " </td> "
						
						'strRet = strRet & " <td> "
						'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar checked value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )"" )  > "
						'strRet = strRet & " </td> "
						
						strRet = strRet & " <td> "
					
						'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )"";objAryFac[0][0]='';PosicoesLivre('L')"" > "
						'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""objAryFac[0][0]='';PosicoesLivre('L')"" > "
						
						'''strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar disabled value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparDSLAM(" & i & ")"" > "
						
						'strRet = strRet & " <input type=checkbox name=" & rdoPortaUplinkID & " title=Editar disabled value=" & objRS("ONTPorta_ID") &  " onclick=""parent.LimparDSLAM(" & i & ");AddSelecaoChk(this)"" > "
						
						strPorta = ""
						strPorta = i & "&&" & objRS("ONTPorta_ID")
						
						'response.write "<script>alert('2')</script>"
						'%> 
						'	<script>
					'			alert(2);
					'			parent.parent.AddChk();
						'	</script>
						
					'	<%
					
						
						
						'strRet = strRet & " <input type=checkbox name=" & rdoPortaUplinkID & " title=Editar disabled value=" & objRS("ONTPorta_ID") &  " onclick=""parent.LimparDSLAM(" & i & ");AddSelecaoChk(this)"" > "
						
						strRet = strRet & " <input type=checkbox name=" & rdoPortaUplinkID & " title=Editar disabled value=" & strPorta & " onclick=""parent.LimparDSLAM(" & i & ");AddSelecaoChk(this)""  > "
						
						strRet = strRet & " </td> "
						
						strRet = strRet & " </tr> "
						
						'strRet = strRet & " <script> AddChk(); </script> "
					
					end if		
						
					'''' retirada UAT DSLAM i=i+1
					
					'''' retirada UAT DSLAM  if blnCor then
					'''' retirada UAT DSLAM  	strRet = strRet & " <tr class=clsSilver> "
					'''' retirada UAT DSLAM  	blnCor = false
					'''' retirada UAT DSLAM  Else
					'''' retirada UAT DSLAM  	strRet = strRet & " <tr class=clsSilver2>"
					'''' retirada UAT DSLAM  		blnCor = true
					'''' retirada UAT DSLAM  End if
					
					'''' retirada UAT DSLAM  strRet = strRet & " <td>&nbsp;&nbsp;" & objRS("ONTPorta_Porta")  & " </td> "
					'''' retirada UAT DSLAM  strRet = strRet & " <td>&nbsp;&nbsp; </td> "
					'''' retirada UAT DSLAM  strRet = strRet & " <td width=100><input type=text class=text name=txtCabo"& i & " ID=txtCabo" & i & "	readonly=true maxlength=2	 size=5	 onKeyUp=""ValidarTipo(this,0)"" ></td> "
					'''' retirada UAT DSLAM  	strRet = strRet & "<td width=120><input type=text class=text name=txtPar"& i & "		maxlength=9  size=10 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N ></td>"
					'''' retirada UAT DSLAM  strRet = strRet & "<td><input type=text class=text name=txtCaixaEmenda"& i & "  maxlength=3 size=5 ></td>"
					'''' retirada UAT DSLAM  strRet = strRet & "<td width=100><input type=text class=text name=txtLateral"& i & "	maxlength=4  size=5></td>"
					'''' retirada UAT DSLAM  strRet = strRet & "<td width=100><select name=cboTipoCabo"& i & " >"
					'''' retirada UAT DSLAM  strRet = strRet & "			<option value=""""></option>"
					'''' retirada UAT DSLAM  strRet = strRet & "			<option value=""H"" >H</option>"
					'''' retirada UAT DSLAM  strRet = strRet & "			<option value=""X"" >X</option>"
					'strAde = strAde & "			<option value=XDSL >XDSL</option>"
					'''' retirada UAT DSLAM  strRet = strRet & "			<option value=""C"" >C</option>"
					'''' retirada UAT DSLAM  strRet = strRet & "		</select></td>"
					
					
					'''' retirada UAT DSLAM  strRet = strRet & " <td>"
					'''' retirada UAT DSLAM  strRet = strRet & "		<select name=cboPropriedadeCabo >  "
										
										
					'''' retirada UAT DSLAM  					set objRSProp = db.execute("CLA_sp_sel_Propriedade ")
										
													
					'''' retirada UAT DSLAM  					While Not objRSProp.eof
					'''' retirada UAT DSLAM  						strItemSel = ""
					'''' retirada UAT DSLAM  					   if Trim(propswitchEDD) = Trim(objRSProp("PROP_ID")) then strItemSel = " Selected " End if
										   
										
					'''' retirada UAT DSLAM  						strRet = strRet & "	<Option value= " & objRSProp("PROP_ID") & " > " & trim(objRSProp("PROP_SIGLA")) &  " </Option> "
											
					'''' retirada UAT DSLAM  						objRSProp.MoveNext
					'''' retirada UAT DSLAM  					Wend
										
					'''' retirada UAT DSLAM  strRet = strRet & "		</select> "
					'''' retirada UAT DSLAM  strRet = strRet & " </td> "
					
				'	strRet = strRet & " <td> <input type=text class=text disabled=true name="& strUplink & i & " ID=uplink" & i & " size=20 maxlength=20 value=" & objRS("Uplink") & " ></td> " 
									
					'''' retirada UAT DSLAM  strRet = strRet & " <td> "
					
					'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparUplink(" & i & " , '" & strEqui & "' )"";objAryFac[0][0]='';PosicoesLivre('L')"" > "
					'strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""objAryFac[0][0]='';PosicoesLivre('L')"" > "
					
					'''strRet = strRet & " <input type=radio name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparDSLAM(" & i & ")"" > "
					
					'''' retirada UAT DSLAM  strRet = strRet & " <input type=checkbox name=" & rdoPortaUplinkID & " title=Editar value=" & objRS("ONTPorta_ID") & " onclick=""parent.LimparDSLAM(" & i & ");AddSelecaoChk(this)"" > "
					'''' retirada UAT DSLAM  strRet = strRet & " </td> "
					
					'onblur="CompletarCampo(this);hdnOrderEntryAno.value=this.value;" onkeyup="ValidarTipo(this,0)" 
					
					'''' retirada UAT DSLAM  strRet = strRet & " </tr> "
					
					
					
				end if
				
				'strRet = strRet & " </tr> "
				
				i=i+1		
				
				
				
		
			objRS.MoveNext
		Wend
		strRet = strRet & " </table> "
		
		strRet = strRet & " <table width=""760"" > "
		
			if strProcesso = "4" then
			
				strRet = strRet & "	<tr> "
				strRet = strRet & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnRemoverUplink"" value=""Remover DSLAM"" style=""width:150px;height:22px""  onClick=""RemoverUplink()"" ></td> "
			'	strRet = strRet & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnGravarUplink"" style=""width:150px;height:22px"" value=""Concluir Cancelamento"" onclick= parent.GravarFacilidade() ></td> "
				strRet = strRet & "	</tr> "
				
			else 
				
				
				strRet = strRet & "	<tr> "
				'strRet = strRet & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnSairSwitch"" value=""Sair"" style=""width:150px""  onClick=""sair()"" ></td> "
				
				strRet = strRet & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnGravarUplink"" style=""width:150px;height:22px"" value=""Alocar DSLAM"" onclick=""parent.GravarPortaDSLAM('" & strEqui & "')"" ></td> "
				
				''''strRet = strRet & "		<td align=center class=clsSilver2>&nbsp;<input type=""button"" class=""button"" name=""btnGravarUplink"" style=""width:150px;height:22px"" value=""Alocar DSLAM"" onclick=""IncluirDSLAM()"" ></td> "
				
				strRet = strRet & "	</tr> "
			end if 
			
			
			strRet = strRet & " </table> "
		
	
	else
		
		strRet = ""
		
	end if	
		
		
		
		
%>
<%=strRet%>

