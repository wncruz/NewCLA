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
strRede = "14"
'strRede = objXmlDadosForm.selectSingleNode("//texto").text

'texto = objXmlDadosForm.selectSingleNode("//texto").text
'set objRS = db.execute("CLA_sp_sel_TipoONT null," & paramID)

'cbo = ""
'if texto <> "N" then
'	cbo = cbo & "Modelo ONT | EDD"
'end if
'cbo = cbo & "<select name=cboTipoONT>"
'cbo = cbo & "<Option value=''>:: MODELO</Option>"
	
'While Not objRS.eof
'	strItemSel = ""
'	if Trim(dblTontID) = Trim(objRS("Tont_ID")) then strItemSel = " Selected " End if
'		cbo = cbo & "<Option value='" & objRS("Tont_ID") & "'" & strItemSel & ">" & Trim(objRS("Tont_Modelo")) & "</Option>"
'	objRS.MoveNext
'Wend
'strItemSel = ""

'cbo = cbo & "</select>"

		if ( paramID = "2" or paramID = "3" or paramID = "4" or paramID = "9" or paramID = "10" or paramID = "11" or paramID = "12") then

			strRet = ""


			strRet = strRet & "	<table  border=0 cellspacing=1 cellpadding=1 width=760> "
		
			if (paramID = "4") then
			
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "	 		<td colspan=5><font class=clsObrig>::</font>&nbsp;Propriedade do Switch EDD &nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "				<select name=cboPropriedade> "
				
				'set objRS = db.execute("CLA_sp_sel_Propriedade ")
				
							
				'While Not objRS.eof
				'  strItemSel = ""
				'  strRet = strRet & " <Option value=""" &  objRS("PROP_ID") & """" & strItemSel & "> " & Trim(objRS("PROP_SIGLA")) &  " </Option> "
				'  objRS.MoveNext
				'Wend
				
				strRet = strRet & " <Option value=2>EBT</Option> "
				
				strRet = strRet & " 			</select> "
				strRet = strRet & " 		</td> "
				
				strRet = strRet & "	</tr> "
				
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "		<td colspan=5><font class=clsObrig>::</font>&nbsp;Designação do SWITCH EDD &nbsp;<input type=text name=txtDesignacao size=20  class=text></td> "
				strRet = strRet & "	</tr> "
				
				

				
				strRet = strRet & "	 <tr class=clsSilver> "
				strRet = strRet & "		<th></th> "
				strRet = strRet & "		<th></th> "
				strRet = strRet & "		<th></th> "
				strRet = strRet & "	</tr> "
				strRet = strRet & " <tr class=clsSilver> "
				
				strRet = strRet & "	 <td nowrap><font class=clsObrig>::</font>&nbsp;Fabricante EDD &nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "	<select name=cboFabricanteONT onchange=ResgatarTipoONT() > "
				
				strRet = strRet & "	<option value=""""></option> "
				
				
				set objRS = db.execute("CLA_sp_sel_FabricanteONT 0 , null , null , null , " &  int("0" & strRede) )
				
							
				While Not objRS.eof
				  strItemSel = ""
				  strRet = strRet & " <Option value=""" &  objRS("Font_ID") & """" & strItemSel & "> " & Trim(objRS("Font_Nome")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
				
				strRet = strRet & " </select> "
				strRet = strRet & " </td> "
				
				strRet = strRet & " <td nowrap colspan=2> <span ID=spncboTipoONT></span> "
				strRet = strRet & "	</td>"
				strRet = strRet & " </tr> "
			
			
			
				'strRet = strRet & " <tr class=clsSilver> "
				'strRet = strRet & "	<td nowrap><font class=clsObrig>::</font>&nbsp;PORTA EDD CLIENTE"
				'strRet = strRet & "	<input type=text class=text name=txtPortaPE maxlength=20 size=20>(A20)</td> "
				'strRet = strRet & " <td nowrap colspan=2><font class=clsObrig>::</font>&nbsp;UPLINK EDD CLIENTE&nbsp; "
				'strRet = strRet & "	<input type=text class=text name=txtPE  maxlength=13 size=13>(A13)</td> "
				'strRet = strRet & " </tr> "
				
			elseIF (paramID = "3") then 
			
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "	 		<td colspan=5><font class=clsObrig>::</font>&nbsp;Propriedade do Switch de Interconexão Terceiros&nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "				<select name=cboPropriedade> "
				
				set objRS = db.execute("CLA_sp_sel_Propriedade ")
				
							
				'While Not objRS.eof
				'  strItemSel = ""
				'  strRet = strRet & " <Option value=""" &  objRS("PROP_ID") & """" & strItemSel & "> " & Trim(objRS("PROP_SIGLA")) &  " </Option> "
				'  objRS.MoveNext
				'Wend
				
				strRet = strRet & " <Option value=2>EBT</Option> "
				
				strRet = strRet & " 			</select> "
				strRet = strRet & " 		</td> "
				
				strRet = strRet & "	</tr> "
				
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "		<td colspan=5><font class=clsObrig>::</font>&nbsp;Designação do SWITCH INTERCONEXÃO TERCEIROS&nbsp;<input type=text name=txtDesignacao size=30  class=text></td> "
				strRet = strRet & "	</tr> "
				
				

				
				'strRet = strRet & "	 <tr class=clsSilver> "
				'strRet = strRet & "		<th></th> "
				'strRet = strRet & "		<th></th> "
				'strRet = strRet & "		<th></th> "
				'strRet = strRet & "	</tr> "
				
				strRet = strRet & " <tr class=clsSilver> "
				
				strRet = strRet & "	 <td nowrap><font class=clsObrig>::</font>&nbsp;Fabricante&nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "	<select name=cboFabricanteONT onchange=ResgatarTipoONT() > "
				
				strRet = strRet & "	<option value=""""></option> "
				
				
				set objRS = db.execute("CLA_sp_sel_FabricanteONT 0 , null , null , null , " &  int("0" & strRede) )
				
							
				While Not objRS.eof
				  strItemSel = ""
				  strRet = strRet & " <Option value=""" &  objRS("Font_ID") & """" & strItemSel & "> " & Trim(objRS("Font_Nome")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
				
				strRet = strRet & " </select> "
				strRet = strRet & " </td> "
				
				strRet = strRet & " <td nowrap colspan=2> <span ID=spncboTipoONT></span> "
				strRet = strRet & "	</td>"
				strRet = strRet & " </tr> "

			
				'strRet = strRet & " <tr class=clsSilver> "
				'strRet = strRet & "	<td nowrap><font class=clsObrig>::</font>&nbsp;PORTA SWITCH INTERCONEXÃO TER "
				'strRet = strRet & "	<input type=text class=text name=txtPortaSwitchInterconexao maxlength=20 size=20>(A20)</td> "
				'strRet = strRet & " <td nowrap colspan=2><font class=clsObrig>::</font>&nbsp;UPLINK SWITCH INTERCONEXÃO TER&nbsp; "
				'strRet = strRet & "	<input type=text class=text name=txtUplinkSwitchInterconexao  maxlength=13 size=13>(A13)</td> "
			
			elseIF (paramID = "9") then 
			
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "	 		<td colspan=5><font class=clsObrig>::</font>&nbsp;Propriedade do PE&nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "				<select name=cboPropriedade> "
				
				'set objRS = db.execute("CLA_sp_sel_Propriedade ")
				
							
				'While Not objRS.eof
				'  strItemSel = ""
				'  strRet = strRet & " <Option value=""" &  objRS("PROP_ID") & """" & strItemSel & "> " & Trim(objRS("PROP_SIGLA")) &  " </Option> "
				'  objRS.MoveNext
				'Wend
				
				strRet = strRet & " <Option value=2>EBT</Option> "
				
				strRede = "15"
				
				strRet = strRet & " 			</select> "
				strRet = strRet & " 		</td> "
				
				strRet = strRet & "	</tr> "
				
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "		<td colspan=5><font class=clsObrig>::</font>&nbsp;Designação do PE&nbsp;<input type=text name=txtDesignacao size=30  class=text></td> "
				strRet = strRet & "	</tr> "
				
				

				
				'strRet = strRet & "	 <tr class=clsSilver> "
				'strRet = strRet & "		<th></th> "
				'strRet = strRet & "		<th></th> "
				'strRet = strRet & "		<th></th> "
				'strRet = strRet & "	</tr> "
				
				strRet = strRet & " <tr class=clsSilver> "
				
				strRet = strRet & "	 <td nowrap><font class=clsObrig>::</font>&nbsp;Fabricante&nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "	<select name=cboFabricanteONT onchange=ResgatarTipoONT() > "
				
				strRet = strRet & "	<option value=""""></option> "
				
				
				set objRS = db.execute("CLA_sp_sel_FabricanteONT 0 , null , null , null , " &  int("0" & strRede) )
				
							
				While Not objRS.eof
				  strItemSel = ""
				  strRet = strRet & " <Option value=""" &  objRS("Font_ID") & """" & strItemSel & "> " & Trim(objRS("Font_Nome")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
				
				strRet = strRet & " </select> "
				strRet = strRet & " </td> "
				
				strRet = strRet & " <td nowrap colspan=2> <span ID=spncboTipoONT></span> "
				strRet = strRet & "	</td>"
				strRet = strRet & " </tr> "
				
			elseIF (paramID = "10") then	
				

				strRet = strRet & "	<tr class=clsSilver >"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Cabo</th>"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Par</th>"
				strRet = strRet & "		<th nowrap><font class=clsObrig></font>&nbsp;Tipo do Cabo</th>"
				strRet = strRet & "		<th><font class=clsObrig></font>&nbsp;Derivação</th>"
				strRet = strRet & "		<th nowrap><font class=clsObrig></font>&nbsp;PADE</th>"
				strRet = strRet & "	</tr>"
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "<td><input type=text class=text name=txtTronco	maxlength=2 size=2 onKeyUp=""parent.ValidarTipo(this,0)""  >(N2)</td>"
				strRet = strRet & "<td><input type=text class=text name=txtPar 	maxlength=9 size=9 onKeyUp=""parent.ValidarNTipo(this,0,4,4,1,0,4)"">(N4-N4)</td>"
				strRet = strRet & "<td><select name=cboTipoCabo >"
				strRet = strRet & "			<option value=""""></option>"
				strRet = strRet & "			<option value=H >H</option>"
				strRet = strRet & "			<option value=X >X</option>"
				'strRet = strRet & "			<option value=XDSL >XDSL</option>"
				strRet = strRet & "			<option value=C >C</option>"
				strRet = strRet & "		</select></td>"
				strRet = strRet & "<td><input type=text class=text name=txtLateral  onKeyUp=""parent.ValidarTipo(this,2)""	maxlength=4 size=5>(A4)</td>"
				strRet = strRet & "<td><input type=text class=text name=txtCaixaEmenda onKeyUp=""parent.ValidarTipo(this,2)""  maxlength=3 size=5 >(A3)</td>"
				strRet = strRet & "</tr>"

				strRet = strRet & "<tr>"
				strRet = strRet & "<th colspan=5>&nbsp;Obs</th>"
				strRet = strRet & "	</tr>"

				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td colspan=5><textarea name=txtObsFac onkeydown=""MaxLength(this,300);"" cols=40 rows=2></textarea>(A300)</td>"
				strRet = strRet & "</tr>"

						
	        elseIF (paramID = "12") then
			
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "	 		<td colspan=5><font class=clsObrig>::</font>&nbsp;Propriedade do Switch EDD &nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "				<select name=cboPropriedade> "
				
				'set objRS = db.execute("CLA_sp_sel_Propriedade ")
				
							
				'While Not objRS.eof
				'  strItemSel = ""
				'  strRet = strRet & " <Option value=""" &  objRS("PROP_ID") & """" & strItemSel & "> " & Trim(objRS("PROP_SIGLA")) &  " </Option> "
				'  objRS.MoveNext
				'Wend
				
				strRet = strRet & " <Option value=2>EBT</Option> "
				
				strRet = strRet & " 			</select> "
				strRet = strRet & " 		</td> "
				
				strRet = strRet & "	</tr> "
				
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "		<td colspan=5><font class=clsObrig>::</font>&nbsp;Designação DSLAM &nbsp;<input type=text name=txtDesignacao size=20  class=text></td> "
				strRet = strRet & "	</tr> "
				
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "		<td colspan=5><font class=clsObrig>::</font>&nbsp;Tronco UPLINK DSLAM &nbsp;<input type=text name=txtUplink size=20  class=text></td> "
				strRet = strRet & "	</tr> "
				
				

				
				strRet = strRet & "	 <tr class=clsSilver> "
				strRet = strRet & "		<th></th> "
				strRet = strRet & "		<th></th> "
				strRet = strRet & "		<th></th> "
				strRet = strRet & "	</tr> "
				strRet = strRet & " <tr class=clsSilver> "
				
				strRet = strRet & "	 <td nowrap><font class=clsObrig>::</font>&nbsp;Fabricante &nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "	<select name=cboFabricanteONT onchange=ResgatarTipoONT() > "
				
				strRet = strRet & "	<option value=""""></option> "
				
				
				set objRS = db.execute("CLA_sp_sel_FabricanteONT 0 , null , null , null , " &  int("0" & strRede) )
				
							
				While Not objRS.eof
				  strItemSel = ""
				  strRet = strRet & " <Option value=""" &  objRS("Font_ID") & """" & strItemSel & "> " & Trim(objRS("Font_Nome")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
				
				strRet = strRet & " </select> "
				strRet = strRet & " </td> "
				
				strRet = strRet & " <td nowrap colspan=2> <span ID=spncboTipoONT></span> "
				strRet = strRet & "	</td>"
				strRet = strRet & " </tr> "
			
			
			
				'strRet = strRet & " <tr class=clsSilver> "
				'strRet = strRet & "	<td nowrap><font class=clsObrig>::</font>&nbsp;PORTA EDD CLIENTE"
				'strRet = strRet & "	<input type=text class=text name=txtPortaPE maxlength=20 size=20>(A20)</td> "
				'strRet = strRet & " <td nowrap colspan=2><font class=clsObrig>::</font>&nbsp;UPLINK EDD CLIENTE&nbsp; "
				'strRet = strRet & "	<input type=text class=text name=txtPE  maxlength=13 size=13>(A13)</td> "
				'strRet = strRet & " </tr> "
			
				
			else  
			
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "	 		<td colspan=5><font class=clsObrig>::</font>&nbsp;Propriedade do Switch METRO&nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "				<select name=cboPropriedade> "
				
				'set objRS = db.execute("CLA_sp_sel_Propriedade ")
				
							
				'While Not objRS.eof
				'  strItemSel = ""
				'  strRet = strRet & " <Option value=""" &  objRS("PROP_ID") & """" & strItemSel & "> " & Trim(objRS("PROP_SIGLA")) &  " </Option> "
				'  objRS.MoveNext
				'Wend
				strRet = strRet & " <Option value=2>EBT</Option> "
				
				strRet = strRet & " 			</select> "
				strRet = strRet & " 		</td> "
				
				strRet = strRet & "	</tr> "
				
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "		<td colspan=5><font class=clsObrig>::</font>&nbsp;Designação do SWITCH METRO&nbsp;<input type=text name=txtDesignacao size=30  class=text></td> "
				strRet = strRet & "	</tr> "
				
				

				
				'strRet = strRet & "	 <tr class=clsSilver> "
				'strRet = strRet & "		<th></th> "
				'strRet = strRet & "		<th></th> "
				'strRet = strRet & "		<th></th> "
				'strRet = strRet & "	</tr> "
				
				strRet = strRet & " <tr class=clsSilver> "
				
				strRet = strRet & "	 <td nowrap><font class=clsObrig>::</font>&nbsp;Fabricante&nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "	<select name=cboFabricanteONT onchange=ResgatarTipoONT() > "
				
				strRet = strRet & "	<option value=""""></option> "
				
				
				set objRS = db.execute("CLA_sp_sel_FabricanteONT 0 , null , null , null , " &  int("0" & strRede) )
				
							
				While Not objRS.eof
				  strItemSel = ""
				  strRet = strRet & " <Option value=""" &  objRS("Font_ID") & """" & strItemSel & "> " & Trim(objRS("Font_Nome")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
				
				strRet = strRet & " </select> "
				strRet = strRet & " </td> "
				
				strRet = strRet & " <td nowrap colspan=2> <span ID=spncboTipoONT></span> "
				strRet = strRet & "	</td>"
				strRet = strRet & " </tr> "
				'
			end if 

			
			
			'strRet = strRet & " </tr> "
			strRet = strRet & " </table> "
		
		else
		
			strRet = ""
			
		end if





%>
<%=strRet%>