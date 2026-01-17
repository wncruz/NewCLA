
<!--#include file="../inc/data.asp"-->
<html>
<HEAD>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
</HEAD>
<Form name=Form1 method=Post onsubmit="return false">
<input type=hidden name=hdnProvedor>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnFacIDAtual>
<input type=hidden name=hdnDelete value ="">
<input type=hidden name=hdnXmlReturn>
<input type=hidden name=hdnQtdLinha value="">

<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: IFrmManobra.asp
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: IFrame da tela de Manobra

%>
<body topmargin="0" leftmargin="0">
<%

dim teste
dim strLocalEntrega
dim	strLocalConfiguracao
dim	strDistribuidor
dim	strRede
dim	strProvedor
dim strPro
dim intIndex
dim strTimeSlot
dim strPar
dim strSair
dim strHdnFac
dim strSql
dim strHdnGrid
dim strRdoFac
dim strPed
dim Pro_ExigePortadora
dim strPropAcesso
dim strLocalID
dim strDisID
dim strProID
dim strRedeID
dim strPlaID
dim strRepresentacao
dim strAcfID
dim strListRep
dim strRecID
dim strEncontrado
dim FacIDAntes

		Dim strNroAcesso, strNroPedido, strAcessoFisico

		strListRep = ""
		FacIDAntes = 0
		strNroAcesso = Request.Form("txtNumeroAcesso")
		strNroPedido = Request.Form("txtNumeroPedido")
		strAcessoFisico = Request.Form("txtAcessoFisico")

		Vetor_Campos(1)="AdWChar,25,adParamInput," & Replace(strNroAcesso,"'","")
		Vetor_Campos(2)="AdWChar,15,adParamInput," & Replace(strAcessoFisico,"'","")
		Vetor_Campos(3)="AdWChar,25,adParamInput," & Replace(strNroPedido,"'","")

		strSql= APENDA_PARAMSTR("CLA_SP_sel_manobra",3,Vetor_Campos)

		'ObjCmd.Execute

		intPageSize = 100

		Call PaginarRS(0,strSql)
		
		if not objRSPag.Eof and not objRSPag.Bof then

			strLocalEntrega = objRSPag("Sigla_Entrega")
			strLocalConfiguracao = objRSPag("Sigla_Conf")
			strDistribuidor = objRSPag("Dst_Desc")
			strRede = objRSPag("Sis_Desc")
			strProvedor = objRSPag("Pro_Nome")
			strPlataforma = objRSPag("Pla_TipoPlataforma")
			strPlaID = objRSPag("Pla_ID")
			strPro = objRSPag("Pro_ID")
			strPed = objRSPag("Ped_ID")
			strExigePortadora = Trim(objRSPag("Pro_ExigePortadora"))
			strPropAcesso	= Trim(objRSPag("Acf_Proprietario"))
			strLocalID = Trim(objRSPag("Esc_ID"))
			strDisID =Trim(objRSPag("Dst_ID"))
			strRedeID =Trim(objRSPag("Sis_ID"))
			strPlaID =Trim(objRSPag("Pla_ID"))
			strAcfID = objRSPag("Acf_ID")
			strRecID = objRSPag("Rec_ID")
			%>
			
			<%
			Set objRS = db.execute("CLA_sp_sel_usuarioctfc2 null," & dblUsuId)
			Set objDicCef = Server.CreateObject("Scripting.Dictionary")
			Set objDicPag = Server.CreateObject("Scripting.Dictionary")

			While Not objRS.Eof
				if Not objDicCef.Exists(Ucase(Trim(objRS("Age_Sigla")))) then
					Call objDicCef.Add (Ucase(Trim(objRS("Age_Sigla"))),Ucase(Trim(objRS("Ctfc_ID"))))
				End if
				objRS.MoveNext
			Wend

			if objRSPag("Sts_ID") = "38" and not objDicCef.Exists("PST") then
			%>
				<input type=hidden name=hdnEncontrado value="1">
				<table width=730 border=0 cellspacing=0 cellpadding=0 valign=top>
				<tr>
				<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Favor alterar a facilidade pela Alocação de Facilidades.</font></td>
				</tr>
				</table>
			<%
			response.end
			end if
			
			if isNull(objRSPag("Sis_Desc")) = true then
				strRede = ""
				strEncontrado = true
			%>
				<input type=hidden name=hdnEncontrado value="1">
				<table width=730 border=0 cellspacing=0 cellpadding=0 valign=top>
				<tr>
				<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Pedido não disponível para manobrar facilidades.</font></td>
				</tr>
				</table>
			<%

			elseif strRede <> "ADE" and strProvedor = "EMBRATEL" then
				strRede = ""
				strEncontrado = true
			%>
				<input type=hidden name=hdnEncontrado value="1">
				<table width=730 border=0 cellspacing=0 cellpadding=0 valign=top>
				<tr>
				<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Não é possível alterar facilidade de Acesso Embratel.</font></td>
				</tr>
				</table>
			<%
			else
			%>
			<!-- =======================================================================
			================ INÍCIO DAS INFORMAÇÕES DO RECURSO  ===================
			======================================================================= -->
			<input type=hidden name=hdnPedDM value ="<%=strPed%>">
			<input type=hidden name=hdnProvedor2 value ="<%=strPro%>">
			<input type=hidden name=hdnLocalInstala value ="<%=strLocalID%>">
			<input type=hidden name=hdnDisID value ="<%=strDisID%>">
			<input type=hidden name=hdnRedeID value ="<%=strRedeID%>">
			<input type=hidden name=hdnPlaID value ="<%=Pla_ID%>">
			<input type="hidden" name=hdnPedId value="<%=strPed%>">
			<table leftmargin=0 border="0" cellspacing="" cellpadding="0" width="760" >
			<tr class=clsSilver>
				<th colspan=2>&nbsp;•&nbsp;Recurso</th>
			</tr>
			<tr class=clsSilver>
				<td width="25%"><font class="clsObrig">:: </font>Local de Entrega</td>
				<td><input type="text" name="txtLocalEntrega" size="15"  value="<%=strLocalEntrega%>" class="text" readonly maxlength="15"></td>
			</tr>
			<tr class=clsSilver>
				<td><font class="clsObrig">:: </font>Local de Configuração</td>
				<td><input type="text" name="txtLocalConfiguracao" size="15"  value="<%=strLocalConfiguracao%>" class="text" readonly maxlength="15"></td>
			</tr>
			<tr class=clsSilver>
				<td><font class="clsObrig">:: </font>Distribuidor</td>
				<td><input type="text" name="txtDistribuidor" size="35"  value="<%=strDistribuidor%>" class="text" readonly maxlength="30"></td>
			</tr>
			<tr class=clsSilver>
				<td ><font class="clsObrig">:: </font>Rede</td>
				<td >
					<input type="text" name="txtRede" size="25"  value="<%=strRede%>" class="text" readonly maxlength="15">
					<label id="lblPlataforma" style="visibility: visible;"> - Plataforma </label>
					<input type="text" name="txtPlataforma" size="15" value="<%=strPlataforma%>" class="text" maxlength="15" readonly>
					<input type="hidden" name="hdnPlataforma" value="<%= strPlaID %>">

				</td>
			</tr>
			<tr class=clsSilver>
				<td ><font class="clsObrig">:: </font>Provedor</td>
				<td ><input type="text" name="txtProvedor" size="40"  value="<%=strProvedor%>" class="text" readonly maxlength="40"></td>
			</tr>
			</table>
			<input type=hidden name=hdnEncontrado value="0">

			<!-- =======================================================================
			================== FIM DAS INFORMAÇÕES DO RECURSO =====================
			======================================================================= -->
			<%
			strEncontrado = false
			end if
		Else
				%>
				<input type=hidden name=hdnEncontrado value="1">
				<table width=730 border=0 cellspacing=0 cellpadding=0 valign=top>
				<tr>
				<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>
				</tr>
				</table>
				<%
				strEncontrado = true
		End if

		if strRede = "DETERMINISTICO" then%>
			<SCRIPT LANGUAGE=javascript>
			var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
			var objXmlReturn = new ActiveXObject("Microsoft.XMLDOM")

			var intIndice = 0
			var objAryObjs = new Array()
			var objAryFac = new Array()
			var objAryFacRet

			objAryObjs[1] = new Array("txtNumAcessoPtaEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOProvedor","txtCNLPontaA","txtCNLPontaB", "rdoUrbano[0]","rdoUrbano[1]","txtRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtAreaObs")
			objAryObjs[2] = new Array("txtNumAcessoPtaEbt2","txtTronco","txtPar2","txtNroAcessoCli2","cboCodProv2","txtCCTOPro2","txtCNLPontaA2","txtCNLPontaB2", "rdoUrbano2[0]","rdoUrbano2[1]","txtRede","Ped_Id","cboPropModem2","txtQtdeModem2","rdoFatura2[0]","rdoFatura2[1]","txtAreaObs2")
			objAryObjs[3] = new Array("txtNumAcesso3","txtCabo3","txtPar3","txtDerivacao3","cboTCabo3","txtPADE3","txtRede","Ped_Id","cboPropModem3","txtQtdeModem3","txtAreaObs3")
			objAryFac[0] = new Array("","")
			</SCRIPT>

				<!-- =======================================================================
				 =========================== DETERMINISTICO ============================
				 ======================================================================= -->


			<!-- =======================================================================
				 ============= INÍCIO DAS FACILIDADES DO RECURSO DETERM ================
				 ======================================================================= -->
			<table cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr class=clsSilver>
				<th colspan=2>
					&nbsp;•&nbsp;Facilidades do <%=strNroPedido%>
				</th>
			</tr>
			</table>


			<table class=clsSilver2 id="tblFac" cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr class=clsSilver>
				<th width="5%"></th>
				<th width="10%">Provedor</th>
				<th width="15%">Nº Acesso Pta Ebt</th>
				<th width="10%">Fila</th>
				<th width="10%">Bastidor</th>
				<th width="10%">Régua</th>
				<th width="10%">Posição</th>
				<th width="10%">Timeslot</th>
				<th width="20%">Nº Acesso Pta CLI</th>
			</tr>
			<%
				For intIndex = 0 to objRSPag.RecordCount -1
					if isNull(objRSPag("Fac_ID")) = false then
						if isNull(objRSPag("Fac_Representacao")) then
							strTimeSlot = objRSPag("Fac_TimeSlot")
							strSair = false
							strRepresentacao = ""
						else
							strTimeSlot = objRSPag("Fac_Representacao")
							strSair = true
							strRepresentacao = objRSPag("Fac_Representacao")
							strListRep = strListRep & objRSPag("Fac_ID") & ","
						End if
						strHdnFac = "hdnFacID" & intIndex
						strHdnGrid = "hdnDados" & intIndex
				%>

								<tr class=clssilver2>
								<td><input type="radio" name="rdoSelFacilidade" index="<%=intIndex%>" class="radio" onclick="ExibeManobraDet('<%=objRSPag("Pro_ID")%>','<%=objRSPag("Acf_NroAcessoPtaEbt")%>','<%=objRSPag("Fac_Fila")%>','<%=objRSPag("Fac_Bastidor")%>','<%=objRSPag("Fac_Regua")%>','<%=objRSPag("Fac_Posicao")%>','<%=strTimeSlot%>','<%=objRSPag("Acf_NroAcessoPtaCli")%>','<%=objRSPag("Acf_NroAcessoCCTOProvedor")%>','<%=objRSPag("Acf_CCTOTipo")%>','<%=objRSPag("Acf_CnlPTA")%>','<%=objRSPag("Acf_CnlPTB")%>','<%=objRSPag("Acf_ProprietarioEquip")%>','<%=objRSPag("Acf_QtdEquip")%>','<%=objRSPag("Acf_CCTOFatura")%>','<%=objRSPag("Acf_Obs")%>','<%=objRSPag("Fac_ID")%>','<%=objRSPag("Fac_Link")%>');ResgatarPadraoProvedor(<%=objRSPag("Pro_ID")%>)"></td>
								<td><%=objRSPag("Pro_Cod")%></td>
								<td><%=objRSPag("Acf_NroAcessoPtaEbt")%></td>
								<td><%=objRSPag("Fac_Fila")%></td>
								<td><%=objRSPag("Fac_Bastidor")%></td>
								<td><%=objRSPag("Fac_Regua")%></td>
								<td><%=objRSPag("Fac_Posicao")%></td>
								<td><%=strTimeSlot%></td>
								<td><%=objRSPag("Acf_NroAcessoPtaCli")%></td>
								<td><input type=hidden name="<%=strHdnFac%>" value=<%=objRSPag("Fac_ID")%>></td>
								<td><input type=hidden name="<%=strHdnGrid%>" value="<%=objRSPag("Pro_ID") & "&&" & objRSPag("Acf_NroAcessoPtaEbt") & "&&" & objRSPag("Fac_Fila") & "&&" & objRSPag("Fac_Bastidor") & "&&" & objRSPag("Fac_Regua") & "&&" & objRSPag("Fac_Posicao") & "&&" & strTimeSlot & "&&" &  objRSPag("Acf_NroAcessoCCTOProvedor") & "&&" & objRSPag("Acf_NroAcessoPtaCli") & "&&" &  objRSPag("Acf_CnlPTA") & "&&" & objRSPag("Acf_CnlPTB") & "&&" &  objRSPag("Acf_QtdEquip")  & "&&" & objRSPag("Acf_ProprietarioEquip") & "&&" & objRSPag("Fac_Link") & "&&" & objRSPag("Acf_Obs") & "&&" & objRSPag("Acf_CCTOTipo") & "&&" & objRSPag("Acf_CCTOFatura") & "&&" & objRSPag("Fac_Link")%>"></td>
							<%
							FacIDAntes = objRSPag("Fac_ID")
							objRSPag.MoveNext
							%>
							</tr>
							<%
							if objRSPag.EOF then Exit For
							While strSair = true and not objRSPag.EOF
								if isNull(objRSPag("Fac_ID")) = false then
									if isNull(objRSPag("Fac_Representacao")) then
										strTimeSlot = objRSPag("Fac_TimeSlot")
										strSair = false
										strRepresentacao = ""
										intIndex = intIndex + 1
										strHdnFac = "hdnFacID" & intIndex
										strHdnGrid = "hdnDados" & intIndex
										%>
											<tr class=clssilver2>
											<td><input type="radio" name="rdoSelFacilidade" index="<%=intIndex%>" class="radio" onclick="ExibeManobraDet('<%=objRSPag("Pro_ID")%>','<%=objRSPag("Acf_NroAcessoPtaEbt")%>','<%=objRSPag("Fac_Fila")%>','<%=objRSPag("Fac_Bastidor")%>','<%=objRSPag("Fac_Regua")%>','<%=objRSPag("Fac_Posicao")%>','<%=strTimeSlot%>','<%=objRSPag("Acf_NroAcessoPtaCli")%>','<%=objRSPag("Acf_NroAcessoCCTOProvedor")%>','<%=objRSPag("Acf_CCTOTipo")%>','<%=objRSPag("Acf_CnlPTA")%>','<%=objRSPag("Acf_CnlPTB")%>','<%=objRSPag("Acf_ProprietarioEquip")%>','<%=objRSPag("Acf_QtdEquip")%>','<%=objRSPag("Acf_CCTOFatura")%>','<%=objRSPag("Acf_Obs")%>','<%=objRSPag("Fac_ID")%>','<%=objRSPag("Fac_Link")%>');ResgatarPadraoProvedor(<%=objRSPag("Pro_ID")%>)"></td>
											<td><%=objRSPag("Pro_Cod")%></td>
											<td><%=objRSPag("Acf_NroAcessoPtaEbt")%></td>
											<td><%=objRSPag("Fac_Fila")%></td>
											<td><%=objRSPag("Fac_Bastidor")%></td>
											<td><%=objRSPag("Fac_Regua")%></td>
											<td><%=objRSPag("Fac_Posicao")%></td>
											<td><%=strTimeSlot%></td>
											<td><%=objRSPag("Acf_NroAcessoPtaCli")%></td>
											<td><input type=hidden name="<%=strHdnFac%>" value=<%=objRSPag("Fac_ID")%>></td>
											<td><input type=hidden name="<%=strHdnGrid%>" value=""></td>
										<%
									elseif FacIDAntes + 1 <> objRSPag("Fac_ID") then
										strTimeSlot = objRSPag("Fac_Representacao")
										strSair = true
										strRepresentacao = objRSPag("Fac_Representacao")
										intIndex = intIndex + 1
										strListRep = strListRep & objRSPag("Fac_ID") & ","
										strHdnFac = "hdnFacID" & intIndex
										strHdnGrid = "hdnDados" & intIndex
										%>
											<tr class=clssilver2>
											<td><input type="radio" name="rdoSelFacilidade" index="<%=intIndex%>" class="radio" onclick="ExibeManobraDet('<%=objRSPag("Pro_ID")%>','<%=objRSPag("Acf_NroAcessoPtaEbt")%>','<%=objRSPag("Fac_Fila")%>','<%=objRSPag("Fac_Bastidor")%>','<%=objRSPag("Fac_Regua")%>','<%=objRSPag("Fac_Posicao")%>','<%=strTimeSlot%>','<%=objRSPag("Acf_NroAcessoPtaCli")%>','<%=objRSPag("Acf_NroAcessoCCTOProvedor")%>','<%=objRSPag("Acf_CCTOTipo")%>','<%=objRSPag("Acf_CnlPTA")%>','<%=objRSPag("Acf_CnlPTB")%>','<%=objRSPag("Acf_ProprietarioEquip")%>','<%=objRSPag("Acf_QtdEquip")%>','<%=objRSPag("Acf_CCTOFatura")%>','<%=objRSPag("Acf_Obs")%>','<%=objRSPag("Fac_ID")%>','<%=objRSPag("Fac_Link")%>');ResgatarPadraoProvedor(<%=objRSPag("Pro_ID")%>)"></td>
											<td><%=objRSPag("Pro_Cod")%></td>
											<td><%=objRSPag("Acf_NroAcessoPtaEbt")%></td>
											<td><%=objRSPag("Fac_Fila")%></td>
											<td><%=objRSPag("Fac_Bastidor")%></td>
											<td><%=objRSPag("Fac_Regua")%></td>
											<td><%=objRSPag("Fac_Posicao")%></td>
											<td><%=strTimeSlot%></td>
											<td><%=objRSPag("Acf_NroAcessoPtaCli")%></td>
											<td><input type=hidden name="<%=strHdnFac%>" value=<%=objRSPag("Fac_ID")%>></td>
											<td><input type=hidden name="<%=strHdnGrid%>" value="<%=objRSPag("Pro_ID") & "&&" & objRSPag("Acf_NroAcessoPtaEbt") & "&&" & objRSPag("Fac_Fila") & "&&" & objRSPag("Fac_Bastidor") & "&&" & objRSPag("Fac_Regua") & "&&" & objRSPag("Fac_Posicao") & "&&" & strTimeSlot & "&&" &  objRSPag("Acf_NroAcessoCCTOProvedor") & "&&" & objRSPag("Acf_NroAcessoPtaCli") & "&&" &  objRSPag("Acf_CnlPTA") & "&&" & objRSPag("Acf_CnlPTB") & "&&" &  objRSPag("Acf_QtdEquip")  & "&&" & objRSPag("Acf_ProprietarioEquip") & "&&" & objRSPag("Fac_Link") & "&&" & objRSPag("Acf_Obs") & "&&" & objRSPag("Acf_CCTOTipo") & "&&" & objRSPag("Acf_CCTOFatura") & "&&" & objRSPag("Fac_Link")%>"></td>
										<%
									else
										strListRep = strListRep & objRSPag("Fac_ID") & ","
										strSair = true
										
									End if
								End if			
								FacIDAntes = objRSPag("Fac_ID")
								objRSPag.MoveNext
							Wend
						End if
					if objRSPag.EOF then Exit For
				Next
					 %>					 
			</table>
			<input type=hidden name=hdnListRep value="<%=strListRep%>">
			<!-- =======================================================================
				 =============== FIM DAS FACILIDADES DO RECURSO DETRM. =================
				 ======================================================================= -->



			<!-- =======================================================================
				 ========== INÍCIO INFORMAÇÕES DA FACILIDADE DETERMINISTICO ============
				 ======================================================================= -->
			<table cellspacing="1" cellpadding="0" width="760" border="0">
			<tr class=clssilver>
				<th colspan="2">&nbsp;•&nbsp;Informações da Facilidade</th>
			</tr>
			</table>


			<table cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr class=clssilver2>
				<td width="14%">Cod. Prov</td>
				<td width="14%">Nº Acesso Pta Ebt</td>
				<td width="14%"><font class="clsObrig">:: </font>Fila</td>
				<td width="10%"><font class="clsObrig">:: </font>Bastidor</td>
				<td width="10%"><font class="clsObrig">:: </font>Régua</td>
				<td width="10%"><font class="clsObrig">:: </font>Posição</td>
				<td width="18%"><font class="clsObrig">:: </font>Timeslot</td>
			</tr>
			<tr class=clssilver>
				<td><select name="cboCodProv" onChange="ResgatarPadraoProvedor(this.value)">
					<option value=""></Option>
					<%
						Set objRS = db.execute("CLA_sp_sel_Provedor null," & strPro)
						While Not objRS.Eof
							Response.Write "<Option value=" & Trim(objRS("Pro_ID")) & ">" & TratarAspasJS(objRS("Pro_Cod")) & "</Option>"
							objRS.MoveNext
						Wend
					%>
				    </select>
				</td>
				<td><input type="text" name="txtNumAcessoPtaEbt" maxlength=25 size=15  value="" class="text" onKeyup="ValidarVirgula(this);" onKeyPress="if (ValidaProvedor(this)){IFrmProcesso1.TamanhoMax(this)}"></td>
				<td><input type="text" name="txtFila" maxlength=5 readonly size=10  value="" class="text" onKeyUp="ValidarVirgula(this)">(A5)</td>
				<td><input type="text" name="txtBastidor" maxlength=3 readonly size=4  value="" class="text" onKeyUp="ValidarVirgula(this)">(N3)</td>
				<td><input type="text" name="txtRegua" maxlength=2 readonly size=3  value="" class="text" onKeyUp="ValidarVirgula(this)">(A2)</td>
				<td><input type="text" name="txtPosicao" maxlength=3 readonly size=5  value="" class="text" onKeyUp="ValidarVirgula(this)">(A3)</td>
				<td><input type="text" name="txtTimeslot" size="14"  value="" class="text"  maxlength=9 size=10 onKeyUp="ValidarNTipo(this,0,4,4,1,0,4)" TIPO=N>(N4-N4)</td>
			</tr>
			<tr class=clssilver2>
				<td>Nº Acesso Pta Cli</td>
				<td>CCTO Provedor</td>
				<td>IU/URB</td>
				<td>CNL Ponta A</td>
				<td>CNL Ponta B</td>
				<td colspan="2">Link</td>
			</tr>
			<tr class=clssilver>
				<td><input type="text" name="txtNumAcessoCLI" size=15 maxlength=25 value="" class="text" onKeyUp="ValidarVirgula(this)">(A25)</td>
				<td><input type="text" name="txtCCTOProvedor" size=8 maxlength=8  value="" class="text" onKeyUp="ValidarVirgula(this)">(A8)</td>
				<td>
					<input type="radio" name="rdoUrbano" value="I" index=0 onClick="AdicionarCNLManobra(1)">IU
					<input type="radio" name="rdoUrbano" value="U" index=1 onClick="AdicionarCNLManobra(1)">URB
				</td>
				<td><input type="text" name="txtCNLPontaA" size=4 maxlength=4  value="" class="text" onKeyUp="ValidarVirgula(this);">(A4)"</td>
				<td><input type="text" name="txtCNLPontaB" size=4 maxlength=4  value="" class="text" onKeyUp="ValidarVirgula(this);">(A4)"</td>
				<td colspan="2"><input type="text" readonly name="txtLink" size="25"  value=""  class="text"  maxlength="" onKeyUp="ValidarVirgula(this);"></td>
			</tr>
			<tr class=clssilver2>
				<td><font class="clsObrig">:: </font>Prop Modem</td>
				<td><font class="clsObrig">:: </font>Qtde Modem</td>
				<td><font class="clsObrig">:: </font>Fatura</td>
				<td colspan="4">Obs</td>
			</tr>
			<tr class=clssilver>
				<td><select name="cboPropModem">
						<option value=""></option>
						<Option value=CLI >CLI</Option>
               			<Option value=EBT >EBT</Option>
						<Option value=TER >TER</Option>
				    </select>
				</td>
				<td><input type="text" name="txtQtdeModem" size="3" onKeyUp="ValidarTipo(this,0)" value="" class="text" maxlength="2">(N2)</td>
				<td>
					<input type="radio" name=rdoFatura class="radio" index=0 value="">Sim
					<input type="radio" name=rdoFatura class="radio" index=1 value="">Não
				</td>
				<td colspan="4"><textarea cols = "35" name="txtAreaObs"onkeydown="MaxLength(this, '300');" onkeyup="ValidarVirgula(this);MaxLength(this, '300');"></textarea>(A300)</td>
			</tr>
			</table>
			<!-- =======================================================================
				 ============= FIM INFORMAÇÕES DA FACILIDADE DETERMINISTICO ============
				 ======================================================================= -->

			<!-- =======================================================================
				 ========================= FIM DETERMINISTICO ==========================
				 ======================================================================= -->

			<!-- =======================================================================
				 ============ BOTÕES ADICIONAR/ALTERAR, REMOVER E LIMPAR  ==============
				 ======================================================================= -->
			<table cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr class=clssilver2>
				<td colspan="2" align="right">
					<input type="button" class="button" name="btnAdicionarAlterar" value="Adicionar/Alterar" style="width:120px" onclick="AddManobra(hdnFacIDAtual.value)">&nbsp;
					<input type="button" class="button" disabled=true name="btnRemover" value="Remover" style="width:120px" onclick="RemoveLinhaManobra(hdnFacIDAtual.value)">&nbsp;
					<input type="button" class="button" name="btnLimpar" value="Limpar" onClick="LimparCamposManobra(hdnFacIDAtual.value)" accesskey="Q" onmouseover="showtip(this,event,'Limpar Facilidade(Alt+Q)');">&nbsp;
				</td>
			</tr>
			</table>
		<%

		elseif strRede =  "NAO DETERMINISTICO" then%>
		<SCRIPT LANGUAGE=javascript>
			var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
			var objXmlReturn = new ActiveXObject("Microsoft.XMLDOM")

			var intIndice = 0
			var objAryObjs = new Array()
			var objAryFac = new Array()
			var objAryFacRet

			objAryObjs[1] = new Array("txtNumAcessoPtaEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOProvedor","txtCNLPontaA","txtCNLPontaB", "rdoUrbano[0]","rdoUrbano[1]","txtRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtAreaObs")
			objAryObjs[2] = new Array("txtNumAcessoPtaEbt2","txtTronco","txtPar2","txtNroAcessoCli2","cboCodProv2","txtCCTOPro2","txtCNLPontaA2","txtCNLPontaB2", "rdoUrbano2[0]","rdoUrbano2[1]","txtRede","Ped_Id","cboPropModem2","txtQtdeModem2","rdoFatura2[0]","rdoFatura2[1]","txtAreaObs2")
			objAryObjs[3] = new Array("txtNumAcesso3","txtCabo3","txtPar3","txtDerivacao3","cboTCabo3","txtPADE3","txtRede","Ped_Id","cboPropModem3","txtQtdeModem3","txtAreaObs3")
			objAryFac[0] = new Array("","")
		</SCRIPT>
			<!-- =======================================================================
				 ========================== NÃO DETERMINISTICO =========================
				 ======================================================================= -->

			<!-- =======================================================================
				 =========== INÍCIO DAS FACILIDADES DO RECURSO N. DETRM. ===============
				 ======================================================================= -->
			<table cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr class=clsSilver>
				<th colspan=2>
					&nbsp;•&nbsp;Facilidades do <%=strNroPedido%>
				</th>
			</tr>
			</table>


			<table class=clsSilver2 id="tblFac" cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr>
				<th width="5%"></th>
				<th width="10%">Provedor</th>
				<th width="15%">Nº Acesso Pta Ebt</th>
				<th width="10%">Tronco</th>
				<th width="10%">Par</th>
				<th width="20%">Nº Acesso Pta CLI</th>
			</tr>

			<%
			For intIndex = 0 to objRSPag.RecordCount -1
				if isNull(objRSPag("Fac_ID")) = false then
					if isNull(objRSPag("Fac_Representacao")) then
						strPar = objRSPag("Fac_Par")
						strSair = false
					else
						strPar = objRSPag("Fac_Representacao")
						strSair = true
					End if
					strHdnFac = "hdnFacID" & intIndex
					strHdnGrid = "hdnDados" & intIndex
							%>
							<tr class=clssilver2>
								<td><input type="radio" name="rdoSelFacilidade" index="<%=intIndex%>" class="radio" onclick="ExibeManobraNaoDet('<%=objRSPag("Pro_ID")%>','<%=objRSPag("Acf_NroAcessoPtaEbt")%>','<%=objRSPag("Fac_Tronco")%>','<%=strPar%>','<%=objRSPag("Acf_NroAcessoPtaCli")%>','<%=objRSPag("Acf_NroAcessoCCTOProvedor")%>','<%=objRSPag("Acf_CCTOTipo")%>','<%=objRSPag("Acf_CnlPTA")%>','<%=objRSPag("Acf_CnlPTB")%>','<%=objRSPag("Acf_ProprietarioEquip")%>','<%=objRSPag("Acf_QtdEquip")%>','<%=objRSPag("Acf_CCTOFatura")%>','<%=objRSPag("Acf_Obs")%>','<%=objRSPag("Fac_ID")%>');ResgatarPadraoProvedor(<%=strPro%>)"></td>
								<td><%=objRSPag("Pro_Cod")%></td>
								<td><%=objRSPag("Acf_NroAcessoPtaEbt")%></td>
								<td><%=objRSPag("Fac_Tronco")%></td>
								<td><%=strPar%></td>
								<td><%=objRSPag("Acf_NroAcessoPtaCli")%></td>
								<td><input type=hidden name="<%=strHdnFac%>" value=<%=objRSPag("Fac_ID")%>></td>
								<td><input type=hidden name="<%=strHdnGrid%>" value=""></td>
							<%
							intCount = intCount+1
							objRSPag.MoveNext
							%>
							</tr>
							<%
							if objRSPag.EOF or strSair = true then Exit For
				End if
			Next
				 %>

			</table>
			<input type=hidden name=hdnListRep value="">
			<!-- =======================================================================
				 ============= FIM DAS FACILIDADES DO RECURSO N. DETERM. ===============
				 ======================================================================= -->



			<!-- =======================================================================
				 =============INICIO INFORMAÇÕES DA FACILIDADE NÃO DETERM. =============
				 ======================================================================= -->
			<table cellspacing="1" cellpadding="0" width="760" border="0">
			<tr class=clssilver>
				<th colspan="2">&nbsp;•&nbsp;Informações da Facilidade</th>
			</tr>
			</table>


			<table class=clsSilver2 cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr class=clssilver2>
				<td width="14%">Cod. Prov</td>
				<td width="14%">Nº Acesso Pta Ebt</td>
				<td width="14%"><font class="clsObrig">:: </font>Tronco</td>
				<td width="10%"><font class="clsObrig">:: </font>Par</td>
				<td width="10%">Nº Acesso Ponta CLI</td>
			</tr>
			<tr class=clssilver>
				<td><select name="cboCodProv2" onChange="ResgatarPadraoProvedor(this.value)">
					<option value=""></Option>
						<%
						Set objRS = db.execute("CLA_sp_sel_Provedor null," & strPro)
						While Not objRS.Eof
							Response.Write "<Option value=" & Trim(objRS("Pro_ID")) & ">" & TratarAspasJS(objRS("Pro_Cod")) & "</Option>"
							objRS.MoveNext
						Wend
						%>
				    </select>
				</td>
				<td><input type="text" name="txtNumAcessoPtaEbt2"  maxlength=25 size=15  value="" class="text" onKeyPress="ValidarVirgula(this);if (ValidaProvedor(this)){IFrmProcesso1.TamanhoMax(this)}"></td>
				<td><input type="text" name="txtTronco" readonly size="13"  value="" class="text "maxlength=20 size=10 onKeyUp="ValidarVirgula(this)"></td>
				<td><input type="text" name="txtPar2" readonly size="13"  value="" class="text" maxlength=9 size=10 onKeyUp="if(this.value.length > 1) ValidarNTipo(this,0,4,4,1,0,4)" TIPO=N></td>
				<td><input type="text" name="txtNumAcessoCLI2" size="17"  value="" class="text" maxlength=25 onKeyUp="ValidarVirgula(this)" size=15></td>
			</tr>
			<tr class=clssilver2>
				<td>CCTO Provedor</td>
				<td>IU/URB</td>
				<td>CNL Ponta A</td>
				<td colspan="2">CNL Ponta B</td>
			</tr>
			<tr class=clssilver>
				<td><input type="text" name="txtCCTOProvedor2" size="9"  value="" class="text" onKeyUp="ValidarVirgula(this)" maxlength="8"></td>
				<td>
					<input type="radio" name="rdoUrbano2" value="I" Index=0 onClick="AdicionarCNLManobra(2)">IU
					<input type="radio" name="rdoUrbano2" value="U" Index=1 onClick="AdicionarCNLManobra(2)">URB
				</td>
				<td><input type="text" name="txtCNLPontaA2" size=4 maxlength=4  value="" onKeyUp="ValidarVirgula(this)" class="text"></td>
				<td colspan="2"><input type="text" name="txtCNLPontaB2" size=4 maxlength=4  value="" onKeyUp="ValidarVirgula(this)" class="text"></td>
			</tr>
			<tr class=clssilver2>
				<td><font class="clsObrig">:: </font>Prop Modem</td>
				<td><font class="clsObrig">:: </font>Qtde Modem</td>
				<td><font class="clsObrig">:: </font>Fatura</td>
				<td colspan="4">Obs</td>
			</tr>
			<tr class=clssilver>
				<td><select name="cboPropModem2">
						<option value=""></option>
						<Option value=CLI >CLI</Option>
               			<Option value=EBT >EBT</Option>
						<Option value=TER >TER</Option>
				    </select>
				</td>
				<td><input type="text" name="txtQtdeModem2" size="3" onKeyUp="ValidarTipo(this,0)" value=""  onKeyUp="ValidarVirgula(this)" class="text" maxlength="2"></td>
				<td>
					<input type="radio" name=rdoFatura2 class="radio" index=0 value ="S">Sim
					<input type="radio" name=rdoFatura2 class="radio" index=1 value ="N">Não
				</td>
				<td colspan="4"><textarea cols = "35" name="txtAreaObs2"onkeydown="MaxLength(this, '300');" onkeyup="ValidarVirgula(this);MaxLength(this, '300');"></textarea>(A300)</td>
			</tr>
			</table>
			<!-- =======================================================================
				 ============= FIM INFORMAÇÕES DA FACILIDADE NÃO DETERM. ===============
				 ======================================================================= -->

			<!-- =======================================================================
				 ======================== FIM NÃO DETERMINISTICO =======================
				 ======================================================================= -->

			<!-- =======================================================================
				 ============ BOTÕES ADICIONAR/ALTERAR, REMOVER E LIMPAR  ==============
				 ======================================================================= -->
			<table cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr class=clssilver2>
				<td colspan="2" align="right">
					<input type="button" class="button" name="btnAdicionarAlterar" value="Adicionar/Alterar" style="width:120px" onclick="AddManobra(hdnFacIDAtual.value)">&nbsp;
					<input type="button" class="button" disabled=true name="btnRemover" value="Remover" style="width:120px" onclick="RemoveLinhaManobra(hdnFacIDAtual.value)">&nbsp;
					<input type="button" class="button" name="btnLimpar" value="Limpar" onClick="LimparCamposManobra(hdnFacIDAtual.value)" accesskey="Q" onmouseover="showtip(this,event,'Limpar Facilidade(Alt+Q)');">&nbsp;
				</td>
			</tr>
			</table>
			<!-- =======================================================================
				 ========== FIM BOTÕES ADICIONAR/ALTERAR, REMOVER E LIMPAR  ============
				 ======================================================================= --><%

		elseif strRede = "ADE" then %>
		<SCRIPT LANGUAGE=javascript>
			var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
			var objXmlReturn = new ActiveXObject("Microsoft.XMLDOM")

			var intIndice = 0
			var objAryObjs = new Array()
			var objAryFac = new Array()
			var objAryFacRet

			objAryObjs[1] = new Array("txtNumAcessoPtaEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOProvedor","txtCNLPontaA","txtCNLPontaB", "rdoUrbano[0]","rdoUrbano[1]","txtRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtAreaObs")
			objAryObjs[2] = new Array("txtNumAcessoPtaEbt2","txtTronco","txtPar2","txtNroAcessoCli2","cboCodProv2","txtCCTOPro2","txtCNLPontaA2","txtCNLPontaB2", "rdoUrbano2[0]","rdoUrbano2[1]","txtRede","Ped_Id","cboPropModem2","txtQtdeModem2","rdoFatura2[0]","rdoFatura2[1]","txtAreaObs2")
			objAryObjs[3] = new Array("txtNumAcesso3","txtCabo3","txtPar3","txtDerivacao3","cboTCabo3","txtPADE3","txtRede","Ped_Id","cboPropModem3","txtQtdeModem3","txtAreaObs3")
			objAryFac[0] = new Array("","")
		</SCRIPT>
			<!-- =======================================================================
				 ================================= ADE =================================
				 ======================================================================= -->

			<!-- =======================================================================
				 =============== INÍCIO DAS FACILIDADES DO RECURSO ADE =================
				 ======================================================================= -->
			<table  cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr class=clsSilver>
				<th colspan=2>
					&nbsp;•&nbsp;Facilidades do <%=strNroPedido%>
				</th>
			</tr>
			</table>


			<table id="tblFac" class=clssilver2 cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr class=clsSilver>
				<th width="5%"></th>
				<th>Nº Acesso</th>
				<th>Cabo</th>
				<th>Par</th>
				<th>PADE</th>
				<th>Derivação</th>
				<th>T. Cabo</th>
			</tr>

				<%
				For intIndex = 0 to objRSPag.RecordCount -1
					if isNull(objRSPag("Fac_ID")) = false then
						if isNull(objRSPag("Fac_Representacao")) then
							strPar = objRSPag("Fac_Par")
							strSair = false
						else
							strPar = objRSPag("Fac_Representacao")
							strSair = true
						End if
						strHdnFac = "hdnFacID" & intIndex
						strHdnGrid = "hdnDados" & intIndex
							%>
							<tr class=clssilver2>
								<td><input type="radio" name="rdoSelFacilidade" index="<%=intIndex%>" class="radio" onclick="ExibeManobraADE('<%=objRSPag("Acf_NroAcessoPtaEbt")%>','<%=objRSPag("Fac_Tronco")%>','<%=strPar%>','<%=objRSPag("Fac_CxEmenda")%>','<%=objRSPag("Fac_Lateral")%>','<%=objRSPag("Fac_TipoCabo")%>','<%=objRSPag("Acf_ProprietarioEquip")%>','<%=objRSPag("Acf_QtdEquip")%>','<%=objRSPag("Acf_Obs")%>','<%=objRSPag("Fac_ID")%>','<%=strQtd%>')"></td>
								<td><%=objRSPag("Acf_NroAcessoPtaEbt")%></td>
								<td><%=objRSPag("Fac_Tronco")%></td>
								<td><%=strPar%></td>
								<td><%=objRSPag("Fac_CxEmenda")%></td>
								<td><%=objRSPag("Fac_Lateral")%></td>
								<td><%=objRSPag("Fac_TipoCabo")%></td>
								<td><input type=hidden name="<%=strHdnFac%>" value=<%=objRSPag("Fac_ID")%>></td>
								<td><input type=hidden name="<%=strHdnGrid%>" value=""></td>
							<%
							intCount = intCount+1
							objRSPag.MoveNext
							%>
							</tr>
							<%
							if objRSPag.EOF or strSair = true then Exit For
					End if
				Next
				 %>
			</table>
			<input type=hidden name=hdnListRep value="">
			<!-- =======================================================================
				 ================ FIM DAS FACILIDADES DO RECURSO ADE ===================
				 ======================================================================= -->


			<!-- =======================================================================
				 ================ INICIO INFORMAÇÕES DA FACILIDADE ADE =================
				 ======================================================================= -->
			<table cellspacing="1" cellpadding="0" width="760" border="0">
			<tr class=clssilver>
				<th colspan="2">&nbsp;•&nbsp;Informações da Facilidade</th>
			</tr>
			</table>


			<table cellspacing="1" cellpadding="0" width="760"  border="0">
			<tr class=clssilver2>
				<td width="14%">Nº Acesso</td>
				<td width="14%"><font class="clsObrig">:: </font>Cabo</td>
				<td width="10%"><font class="clsObrig">:: </font>Par</td>
				<td width="10%">PADE</td>
				<td width="10%">Derivação</td>
				<td width="10%">T. Cabo</td>
			</tr>
			<tr class=clssilver>
				<td><input type="text" name="txtNumAcesso3" maxlength=25 size=15  value="" class="text" onKeyPress="ValidarVirgula(this);if (ValidaProvedor(this)){IFrmProcesso1.TamanhoMax(this)}"></td>
				<td><input type="text" name="txtCabo3" readonly maxlength=2	 size=5	onKeyUp="ValidarVirgula(this)" onKeyUp="ValidarTipo(this,0)" class="text"></td>
				<td><input type="text" name="txtPar3" readonly maxlength=9  size=10 onKeyUp="ValidarNTipo(this,0,4,4,1,0,4)" TIPO=N class="text"></td>
				<td><input type="text" name="txtPADE3" readonly maxlength=3 size=5 onKeyUp="ValidarVirgula(this)" value="" class="text"></td>
				<td><input type="text" name="txtDerivacao3" readonly maxlength=4  size=5 value="" onKeyUp="ValidarVirgula(this)" class="text"></td>
				<td><input type="text" name="cboTCabo3" readonly maxlength=4  size=5 value="" onKeyUp="ValidarVirgula(this)" class="text"></td>	
			</tr>
			<tr class=clssilver2>
				<td><font class="clsObrig">:: </font>Prop. Modem</td>
				<td><font class="clsObrig">:: </font>Qtd. Modem</td>
				<td colspan="4">Obs</td>
			</tr>
			<tr class=clssilver>
				<td><select name="cboPropModem3">
						<option value=""></option>
						<Option value=CLI >CLI</Option>
               			<Option value=EBT >EBT</Option>
						<Option value=TER >TER</Option>
				    </select>
				</td>
				<td><input type="text" name="txtQtdeModem3" size="3" onKeyUp="ValidarTipo(this,0)" value="" class="text" maxlength="2"></td>
				<td colspan="4"><textarea cols = "35" name="txtAreaObs3"onkeydown="MaxLength(this, '300');" onkeyup="ValidarVirgula(this);MaxLength(this, '300');"></textarea>(A300)</td>
			</tr>
			</table>
			<!-- =======================================================================
				 ================= FIM INFORMAÇÕES DA FACILIDADE ADE ===================
				 ======================================================================= -->

			<!-- =======================================================================
				 ============================= FIM ADE =================================
				 ======================================================================= -->
				 <%
				 Response.write Request.Form("txtRede")
				 %>
			<!-- =======================================================================
				 ============ BOTÕES ADICIONAR/ALTERAR, REMOVER E LIMPAR  ==============
				 ======================================================================= -->
			<table cellspacing="1" cellpadding="0" width="760"  border="0">
				<tr class=clssilver2>
					<td colspan="2" align="right">
						<input type="button" class="button" name="btnAdicionarAlterar" value="Adicionar/Alterar" style="width:120px" onclick="AddManobra(hdnFacIDAtual.value)">&nbsp;
						<input type="button" class="button" disabled=true name="btnRemover" value="Remover" style="width:120px" onclick="RemoveLinhaManobra(hdnFacIDAtual.value)">&nbsp;
						<input type="button" class="button" name="btnLimpar" value="Limpar" onClick="LimparCamposManobra(hdnFacIDAtual.value)" accesskey="Q" onmouseover="showtip(this,event,'Limpar Facilidade(Alt+Q)');">&nbsp;
					</td>
				</tr>
			</table>
			<!-- =======================================================================
				 ========== FIM BOTÕES ADICIONAR/ALTERAR, REMOVER E LIMPAR  ============
				 ======================================================================= --><%
		end if




%>
<iframe id=IFrmProcesso1
	name=IFrmProcesso1
	align=left
	frameBorder=0
	width="100%"
	height=35px
	BORDER=0
	scrolling="no"
	height=60>
</iframe>
</body>
</form>
<BR>

<!-- =======================================================================
	 ============ BOTÕES POS. DISPON., ALOC. FACILIDADE E SAIR  ============
	 ======================================================================= -->
<input type=hidden name=hdnRepresentacao value="<%=strRepresentacao%>">
<table id="TblBotao" cellpadding="7" width="760" border="0">
<tr class=clssilver>
	<td colspan=2 align="center">
		<input type="button" class="button" name="btnPosicaoDisponivel" value="Posições Disponíveis" style="width:150px" onclick="objAryFac[0][0]='';JanelaPosManobra()">&nbsp;
		<input type="button" class="button" name="btnAlocarFacilidade" value="Alocar Facilidade(s)" style="width:150px" onClick="RealizaCadastroManobra('<%=strAcfID%>','<%=strRecID%>')">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:150px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
	</tr>
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
<% if strEncontrado = true then %>
<SCRIPT LANGUAGE=javascript>
	document.getElementById("btnPosicaoDisponivel").disabled = true;
	document.getElementById("btnAlocarFacilidade").disabled = true;
</SCRIPT>

<%
Else
%>
<SCRIPT LANGUAGE=javascript>
	document.getElementById("hdnQtdLinha").value = document.getElementById('tblFac').rows.length -1
	if (document.getElementById('txtRede').value == "ADE")
		ResgatarPadraoProvedor(document.getElementById('hdnProvedor2').value);
</SCRIPT>
<% End if %>
<!-- =======================================================================
	 ========== FIM BOTÕES POS. DISPON., ALOC. FACILIDADE E SAIR  ==========
	 ======================================================================= -->
</html>