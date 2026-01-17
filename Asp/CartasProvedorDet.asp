<!--#include file="../inc/data.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: CartasProvedorDet.ASP
'	- Responsável		: Vital
'	- Descrição			: Detalhe da carta enviada ao provedor
Response.Expires		= -1
Response.CacheControl	= "no-cache"

intDoc		= Request("intDoc")
strDocTipo	= Request("strDocTipo")
dblProId	= Request("intProId")

Response.Write "<Html>"
Response.Write "<body " 
if strDocTipo = "T" then Response.Write " scroll=no "
Response.Write " class=TA>"
Response.Write "<Head>"
Response.Write "<Title>CLA - Controle Local de Acesso</Title>"
Response.Write "</head>"

if intDoc <> "" then
	
	Vetor_Campos(1)="adInteger,2,adParamInput," & intDoc
	Vetor_Campos(2)="adInteger,2,adParamInput,"
	Vetor_Campos(3)="adInteger,2,adParamInput,"
	Vetor_Campos(4)="adDate,8,adParamInput,"
	Vetor_Campos(5)="adDate,8,adParamInput,"

	strSql = APENDA_PARAMSTR("CLA_sp_sel_relacaodocumento",5,Vetor_Campos)

	Set objRS = db.Execute(strSql)
	if Not objRS.Eof and Not objRS.Bof then
	
	    session("ss_PedId") = objRS("Ped_Id")
		strLink =  rethtml(objRS("Doc_Conteudo"))
		
		if objRS("Doc_Tipo") = "T" then
			strHtml = ""
			strHtml = strHtml &  "<br><table cellpadding=0 cellspacing=1 width=510 align=center>"
			strHtml = strHtml &  "<tr>"
			strHtml = strHtml &  "	<th nowrap>&nbsp;Enviar Email</th>"
			strHtml = strHtml &  "	<th nowrap>&nbsp;Editar</th>"
			strHtml = strHtml &  "	<th nowrap>&nbsp;Arquivo</th>"
			strHtml = strHtml &  "</tr>"
			strHtml = strHtml &  "<form name=Form0 method=post>"

			strHtml = strHtml &  "<input type=hidden name=hdnEmailProvedor value='" & strProEmail &"'>"
			strHtml = strHtml &  "<input type=hidden name=hdnNomeProvedor value='" & strProNome &"'>"
			strHtml = strHtml &  "<input type=hidden name=hdnPedId value='" & objRS("Ped_Id") &"'>"
			strHtml = strHtml &  "<input type=hidden name=hdnArquivo value='" & strLink &"'>"
			strHtml = strHtml &  "<input type=hidden name=hdnAcao>"

			strHtml = strHtml &  "</form>"
			strHtml = strHtml &  "<tr class=clsSilver height=17px>"
			strHtml = strHtml &  "<td nowrap align=center>"
			strHtml = strHtml &  "<span style=""cursor:hand"" onClick=""EnviarEmail()""><img src=""../imagens/Email.gif"" border=0 Alt=""Enviar Email ao Provedor""></span>&nbsp;<font color=red><i><span id=spnEnviado></span></i></font></td>"
			strHtml = strHtml &  "<td nowrap align=center><span id=spnRef onClick=""javascript:window.open('../" & strLink & "');"" style=""cursor:hand""><font color=#003388><img src='../imagens/Editar.gif' border=0 alt=Editar></font></span></td>"
			strHtml = strHtml &  "<td nowrap align=center>" & strLink & "</td>"
			strHtml = strHtml &  "</tr></table>"
			Response.Write strHtml
		Else
			Response.Write strLink
		End if	
	End if

Else
	response.write "<table align=center>"
	response.write "<tr><td><font color=red>Não foi possível resgatar o e-mail.</font></td></tr>"
	response.write "</table>"
End if

private function RetHtml(strDocumento)
	dim strNomeCarta , strCaminho,  objXSL, objXML , ndArquivo
	strDocumento = TRIM(strDocumento)
	if ucase(mid(trim(strDocumento),1,6)) = "<HTML>"  then
		 RetHtml = strDocumento
		 session("ss_reenviarcarta") = RetHtml
		 
		 Set objRSPro = db.execute("CLA_sp_sel_provedor_reenvia " & dblProId) 
		 if Not objRSPro.Eof and Not objRSPro.bof then
			strProEmail = Trim(objRSPro("CPro_ContratadaEmail"))
			strProNome	= Trim(objRSPro("Pro_Nome"))
			
			session("ss_ProEmail") = Trim(objRSPro("CPro_ContratadaEmail"))
			session("ss_ProNome") =  Trim(objRSPro("Pro_Nome"))
		End if
		 
		 exit function 
	else 
		set objXml = server.CreateObject("Microsoft.XMLDOM")
		set objXsl = server.CreateObject("Microsoft.XMLDOM")
		
		objXml.loadxml(TRIM(strDocumento))
		set ndArquivo =  objXml.selectSingleNode("//arquivo")
		strNomeCarta = ndArquivo.Text

		Set objRSPro = db.execute("CLA_sp_sel_provedor_reenvia " & dblProId) 
		if Not objRSPro.Eof and Not objRSPro.bof then
			strProEmail = Trim(objRSPro("CPro_ContratadaEmail"))
			strProNome	= Trim(objRSPro("Pro_Nome"))
			
			session("ss_ProEmail") = Trim(objRSPro("CPro_ContratadaEmail"))
			session("ss_ProNome") =  Trim(objRSPro("Pro_Nome"))
		End if
		
		if not FileStatus(server.MapPath("..\") & "\asp\" & strNomeCarta & ".asp") then 
			RetHtml = "<table width=100% ><tr><td style=""text-align:center""><font color = red>Modelo cadastrado no banco de dados não é um modelo válido. - " & strNomeCarta  & server.MapPath("..\") & "\" & strNomeCarta & ".asp </font></td></tr></table>" 
			exit function 
			'html informando que não encontrou o modelo.	
		end if 
		strCaminho =  server.MapPath("..\") & "\xsl\" & strNomeCarta & ".xsl"
		
		objXSL.load(strCaminho) 
		objXSL.async = false
		RetHtml = objXml.transformnode(objXSL)  
		RetHtml = REPLACE(RetHtml ,"; charset=UTF-16","")
		session("ss_reenviarcarta") = RetHtml
		set objXSL = nothing 
		set objXML = nothing 

	end if 
	
	
	'Set objRSPro = db.execute("CLA_sp_sel_provedor " & dblProId) 
	'if Not objRSPro.Eof and Not objRSPro.bof then
	'	strProEmail = Trim(objRSPro("pro_email"))
	'	strProNome	= Trim(objRSPro("Pro_Nome"))
	'	if isnull(objRSPro("Pro_PathTemplate")) = true then 
	'		strNomeCarta =	"EnviaProvedor"
	'	else
	'		strNomeCarta =	Trim(objRSPro("Pro_PathTemplate"))
	'	end if 
				
	'End if

	'if strNomeCarta = "EnviaProvedor" or  strNomeCarta = "" or ucase(mid(strDocumento,1,6)) = "<HTML>" then 
	'	 RetHtml =  
	'else

	'	set objXml = server.CreateObject("Microsoft.XMLDOM")
	'	set objXsl = server.CreateObject("Microsoft.XMLDOM")
	'	
	'	objXml.loadXML(strDocumento)				
	'	strCaminho =  server.MapPath("..\")
	'	objXSL.load(strCaminho & "\xsl\" & replace(strNomeCarta, "asp","xsl")) 
	'	objXSL.async = false
	'	RetHtml = objXml.transformnode(objXSL)  
	'	RetHtml = REPLACE(RetHtml ,"; charset=UTF-16","")
	'	set objXSL = nothing 
	'	set objXML = nothing 
	'end if 

end function

Function FileStatus(nameFile)
   Dim fso, msg
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists(nameFile)) Then
      msg = true
   Else
      msg = false
   End If
   set fso = nothing 
   FileStatus = msg
   
End Function


%>

<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
function Imprimir()
{
	window.print()
}
function EnviarEmail()
{
	with (document.forms[0])
	{
		hdnAcao.value = "EnviarEmailProvedor"
		target = "IFrmProcesso"
		action = "ProcessoCartasProv.asp"
		submit()
	}
}
		
function EmailEnviado(strNomePro,strProEmail)
{
	alert("Email enviado com sucesso para " + strNomePro + " (" + strProEmail + ").");
	parent.spnEnviado.innerHTML = '(Email enviado)';
	return true
}
//-->
</SCRIPT>


<script language="VBscript">
Sub btnreenviar_onClick
  returnvalue=MsgBox ("Você realmente deseja reenviar a carta ao provedor?",579,"Confirmação de envio de e-mail")
  
  If returnvalue=6 Then
    document.Form_1.hdnAcao.value = "EnviarEmailProvedor"
	document.Form_1.target = "IFrmProcesso"
	document.Form_1.action = "Reenvia_CartaProvedor.asp"    
	document.Form_1.Submit()
  Else
                        
  End If
End Sub
</script>



<form name="Form_1" method="post">
  <input type="hidden" name="hdnAcao">
  <input type="hidden" name="strEmail" value="<%=replace(session("ss_reenviarcarta"),"""","'")%>">
  <input type="hidden" name="hdnEmailProvedor" value="<%=session("ss_ProEmail")%>">
  <input type="hidden" name="hdnNomeProvedor" value="<%=session("ss_ProNome")%>">
  <input type="hidden" name="hdnPedId" value="<%=session("ss_PedId")%>">
</form>


<table width="100%" border=0>
<tr>
	<td align="center">
	  <%
	  Set objRS = db.execute("CLA_sp_sel_usuarioctfc null," & dblUsuId)
      Set objDicCef = Server.CreateObject("Scripting.Dictionary")

      While Not objRS.Eof
	  if Not objDicCef.Exists(Ucase(Trim(objRS("Age_Sigla")))) then
		Call objDicCef.Add (Ucase(Trim(objRS("Age_Sigla"))),Ucase(Trim(objRS("Ctfc_ID"))))
	    End if
	    objRS.MoveNext
      Wend
      %>
	
		<%if objDicCef.Exists("GAT") then%>
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="button" class="button" name="btnreenviar" value="Reenviar Carta">&nbsp;
		<%end if%>
		<input type="button" class="button" name="btnImprimir" value="Imprimir" onClick="Imprimir()">&nbsp;
		<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.returnValue=0;window.close()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');"><br><br>
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
</body>
</HTML>
