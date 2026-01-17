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

		if ( paramID = "12") or ( paramID = "11") then

			strRet = ""


			strRet = strRet & "	<table  border=0 cellspacing=1 cellpadding=1 width=760> "
		
					
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
				
				
				set objRS = db.execute("CLA_sp_sel_FabricanteONT 0 , null , null , 8 , null " )
				
							
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
			
				
		

			
			
			'strRet = strRet & " </tr> "
			strRet = strRet & " </table> "
		
		else
		
			strRet = ""
			
		end if





%>
<%=strRet%>