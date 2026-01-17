<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="monta-sql.asp"-->

<html>
<%



	'*** busca Variaveis
			Filtro				= request.form("Filtro")		
			if Filtro="" then
				Filtro			= "0"
			end if

		   IDOrdena			= request.form("IDOrdena")

   
		   IDestado 			= request("IDestado")	
		   if request("IDestado")<>""  then
			   Filtro				= "1"
		   end if
		   if request.form("IDestado")	<>"" then
			   IDestado 			= request.form("IDestado")		
			end if   
	
   		   IDQtde				= converte_inteiroLongo(request.form("IDQtde"),0)
   		   IDQtde1				= converte_inteiroLongo(request.form("IDQtde1"),0)

			if request.form("IDtecnologia")	<>"" then
			   IDtecnologia = request.form("IDtecnologia")		
			end if   


		   IF IDestado 	="" THEN
			   IDestado 			= request.form("cboUF")		   
			END IF
			
			if IDQtde=0 then 
			   IDQtde				= converte_inteiroLongo(request.form("cboQtde_acessos"),0)
			end if   
		   
			if IDQtde1=0 then 
			   IDQtde1				= converte_inteiroLongo(request.form("cboQtde_acessos1"),0)
			end if   

%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel=stylesheet type="text/css" href="../css/cla.css">
<title>CLA - Relatório de Acesso</title>
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" border="0">
<tr>
<td>
<center><h3>CLA - Controle Local de Acesso</h3><center>
<h4 align="center">Relatório de Consolidado por Proprietário - <%= date() %></h4>
<center>
<Form name="FormRelat" method="Post" action="excel_consolida_acesso_uf.asp" target="_self" >
 <input type=hidden name="filtro" 				value="1" >
 <input type=hidden name="IDEstado" 			value="<%=IDEstado%>" >
 <input type=hidden name="IDProprietario" 	value="<%=IDProprietario%>" >
 <input type=hidden name="IDOrdena" 			value="<%=IDOrdena%>" >
 
</center>
<tr>
<td>
<br>
<!--************ MONTA A TABELA DE RELATÓRIO ****************** !-->
<% Response.ContentType = "application/vnd.ms-excel" %>
<table width="80%" border="1" align="center" class="TableLine">

<th align=center>#</th>
<th>Proprietário</th>
<th>Qtde Acessos Físicos</th>

<% 
 strSQL =  Monta_SQL_consolida_uf_Proprietario() 

if strSQL<> "" then
   qtde =0 
	SET RS= Server.CreateObject("ADODB.Recordset")
	RS.Open strSQL,db	
	

	
   While Not RS.eof
	 qtde = qtde +1 
	 Tqtde  = Tqtde + RS("Valor_total")   
  
      SELECT CASE RS("Proprietario")
             CASE "EBT"
                   NOMEPROP ="EMBRATEL"
             CASE "TER"       
                   NOMEPROP ="TERCEIRO"
			 CASE "CLI"       
                   NOMEPROP ="CLIENTE"
	 END SELECT                   

  %>
<tr>    
 

 <td align="right"><% =qtde %></td> 
 <td><%=NOMEPROP%></td>
 <td align="center"><%=formatnumber(RS("Valor_total"),0)%></td>    
</tr> 
<%    
     RS.MoveNext
  Wend 
  RS.close : set RS = nothing	%>
 		
<% end if %> 
<tr class=clsSilver>
<td colspan="2"></td>  
<td align="center"><%= formatnumber(Tqtde,0) %></td>  
</tr>	 	
</table>
</form>
</body>






