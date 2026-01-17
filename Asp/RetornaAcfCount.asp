<%
  Response.ContentType = "text/html; charset=utf-8"
                Response.Charset = "UTF-8"
%>
<!--#include file="../inc/data.asp"-->
<HTML>
<HEAD>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
</HEAD>
<body leftMargin=0>
<%

  sEst_Sigla    = Request("cboUFEnd")
  sCid_Sigla    = Request("txtEndCid")
  sStatus       = Request("cboStatus")
  sCliente       = Request("txtCliente")
  tec_id       = Request("cboNewTecnologia")
  fac_ID       = Request("cboNewFacilidade")
  
 ' response.write "<script>alert('"&sStatus&"')</script>"
  
  

  %>
  <form method="post" name=Form1 >

  <table border=0 width=760><tr><td colspan=2 align=right>
	<a href="javascript:AbrirXlsAcesso()" onmouseover="showtip(this,event,'Consulta em formato Excel...')"><img src='../imagens/excel.gif' border=0></a>&nbsp;
	

	</table>
	
  
 <table border="0" cellspacing="1" cellpadding=0 width="760">
	
			<% 

			Vetor_Campos(1)="adInteger,2,adParamInput," & fac_ID
			Vetor_Campos(2)="adInteger,2,adParamInput," & tec_id
			Vetor_Campos(3)="adWChar,50,adParamInput," & sStatus
			Vetor_Campos(4)="adWChar,100,adParamInput, " & sEst_Sigla
			Vetor_Campos(5)="adWChar,100,adParamInput, " & sCid_Sigla
			Vetor_Campos(6)="adWChar,100,adParamInput, " & sCliente
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_CountnewAcessoFisico",6,Vetor_Campos)
			
			
			Vetor_Campos(1)="adInteger,2,adParamInput," & fac_ID
			Vetor_Campos(2)="adInteger,2,adParamInput," & tec_id
			Vetor_Campos(3)="adWChar,50,adParamInput," & sStatus
			Vetor_Campos(4)="adWChar,100,adParamInput, " & sEst_Sigla
			Vetor_Campos(5)="adWChar,100,adParamInput, " & sCid_Sigla
			Vetor_Campos(6)="adWChar,100,adParamInput, " & sCliente
			strSqln = APENDA_PARAMSTR("CLA_sp_sel_newAcessoFisico2",6,Vetor_Campos)
			Response.Write "<script language=javascript>parent.document.forms[0].hdnSQLXLS.value ="&chr(34)&strSqln&chr(34)&";</script>"
			
			
			Vetor_Campos(1)="adInteger,2,adParamInput," & fac_ID
			Vetor_Campos(2)="adInteger,2,adParamInput," & tec_id
			Vetor_Campos(3)="adWChar,50,adParamInput," & sStatus
			Vetor_Campos(4)="adWChar,100,adParamInput, " & sEst_Sigla
			Vetor_Campos(5)="adWChar,100,adParamInput, " & sCid_Sigla
			Vetor_Campos(6)="adWChar,100,adParamInput, " & sCliente
			strSqlnCampo = APENDA_PARAMSTR("CLA_sp_ret_NomecamponewAcessoFisico",6,Vetor_Campos)
			Response.Write "<script language=javascript>parent.document.forms[0].hdnCampoSQLXLS.value ="&chr(34)&strSqlnCampo&chr(34)&";</script>"


			'Response.Write "<script language=javascript>parent.document.forms[0].hdnSQLXLS.value ="&chr(34)&strSqlRet&chr(34)&";</script>"

			'Response.Write strSqlRet
			'Response.Write strSqln
			Set objRS = db.Execute(strSqlRet)
			objRS.Close
			objRS.CursorLocation = adUseClient
			objRS.Open
intCount=1 
if not objRS.Eof and not objRS.Bof then  
	'For intIndex = 1 to objRS.PageSize
	While Not objRS.Eof
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class=<%=strClass%>>
			<td  width=223px >&nbsp;<%=TratarAspasHtml(objRS("newfac_nome"))%></td>
			
			<td width=215px ><%=trim(objRS("newtec_nome"))%></td>
			<td width=65px ><%=trim(objRS("est_sigla"))%></td>
			<td  width=100px ><%=trim(objRS("cid_sigla"))%></td>
			<td  width=85px ><%=trim(objRS("status"))%></td>
			
			<td  width=70px><%=trim(objRS("Quantidade"))%></td>
			
		</tr>
		<%
		intCount = intCount+1
		objRS.MoveNext
	Wend
		
End if
%>
		</td>
	</tr>
</table>
<input type="Hidden" name="hdnSQLXLS" value="<%=strSqln%>">
<input type="Hidden" name="hdnCampoSQLXLS" value="<%=strSqlnCampo%>">
  </form>
  
</body>
</html>

 