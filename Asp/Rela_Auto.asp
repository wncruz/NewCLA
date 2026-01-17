<!--#include file="../inc/Data.asp"-->
<!--#include file="../inc/EnviarEntregarAprov.asp"-->
<!--#include file="../inc/EnviarEntregarAprovASMS.asp"-->

	<html>
	<HEAD>
		<meta http-equiv="refresh" content="240">
	</HEAD>
	<title>Checagem de Auto.</title>
	<center>
		
	<Body topmargin=0 leftmargin=0 class=TA>
		Last refresh: <%= Now() %> 
	</body>
	<br>
	<br>
	<%
		strSQL = "select top 1 Acl_IDAcessoLogico  from CLA_Aprovisionador  where Rede_Wireless = 'S'  and Aprovisi_dtRetornoStatus is not null  and Aprovisi_dtRetornoEntregar is null  " ' & strIDLogico
		SET objRS =  db.execute(strSQL)
		
		If Not objRS.eof and not objRS.Bof Then
			strIDLogico 		= trim(objRS("Acl_IDAcessoLogico"))
										
		end if
	'strIDLogico = Request("txt_IDLogico")
	'Aprovisi_ID = Request("txt_Aprovisi_ID")
	'rdo_acao = Request("rdo_acao")
	'rdo_Aprov = Request("rdo_Aprov")
	%>
	<form name="Form_1" method="post" action="Rela_Auto.asp">
	   <input type="radio" name="rdo_Aprov" value="CFD" <%if rdo_Aprov = "CFD" then%>checked<%end if%>> CFD
	   <input type="radio" name="rdo_Aprov" value="OUTROS" <%if rdo_Aprov = "OUTROS" then%>checked<%end if%>> OUTROS	
	   <br><br>
	   <input type="text" name="txt_IDLogico" value="<%=strIDLogico%>">
	   <br><br>
	   <input type="submit" name="btnok" value="Entregar Acesso SGAP/SGAV/CFD">	   
	    <input type="button" name="btnlimp" value="Limpar" onclick="Form_1.txt_IDLogico.value=''">
	</form>
	<%
	if strIDLogico <> "" then
	
		'response.write "<b>dblIdLogico</b>: " & strIDLogico & "<br>"
		EnviarEntregarAprov(strIDLogico)
		 
	
	End if
	%>
		
