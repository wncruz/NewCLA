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

		   IDestado 			= request("IDestado")
			if IDestado ="" then
			   IDestado 		= request.form("IDestado")				
			end if

		   IDporte				= trim(request.form("IDporte"))				
		   IDOrdena			= request.form("IDOrdena")
   		   IDQtde				= converte_inteiroLongo(request.form("IDQtde"),0)
   		   IDQtde1				= converte_inteiroLongo(request.form("IDQtde1"),0)
   		   IDConta_corrente	= request.form("IDConta_corrente")
   		   IDSubconta			= request.form("IDSubconta")


			if NomeCli ="" then
			   NomeCli 			= Trim(request.form("txtCliente")) 
			end if  
			
			if IDporte="" then
			   IDporte			= trim(request.form("cboPorte"))				
			end if
			
		   if Cliente ="" then
			   Cliente 			= request.form("cbocliente") 
			end if   


		   IF IDestado="" THEN
			   IDestado 			= request.form("cboUF")		   
			END IF
			
			if IDQtde=0 then 
			   IDQtde				= converte_inteiroLongo(request.form("cboQtde_acessos"),0)
			end if   
		   

			if IDQtde1=0 then 
			   IDQtde1				= converte_inteiroLongo(request.form("cboQtde_acessos1"),0)
			end if   
			
			
		strSQL = Monta_SQL_consolida_porte_cliente()  
		SET rs = Server.CreateObject("ADODB.Recordset")

		rs.Open strSQL, db

 		
 

%>


<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel=stylesheet type="text/css" href="../css/cla.css">
<title>CLA - Relatório de Acesso</title>
</head>
<body topmargin="0" leftmargin="0">
<!--************ MONTA A TABELA DE FILTROS ****************** !-->
<table width="100%" border="1">
<tr>
<td>
<table width="100%" border="0">
<tr>
<td>
<center><h3>CLA - Controle Local de Acesso</h3><center>
<h4 align="center">Relatório de Consolidado por cliente  - <%= date() %></h4>
<center>
<Form name="FormRelat" method="Post" action="detalhe_acesso_endereco_servico.asp" target="_self" >
 <input type=hidden name="filtro" 					value="1" >
 <input type=hidden name="NomeCli" 				value="<%=NomeCli%>" >
 <input type=hidden name="IDestado"			 	value="<%=IDestado%>" >
 <input type=hidden name="IDQtde" 					value="<%=IDQtde%>" >
 <input type=hidden name="IDConta_corrente"		value="<%=IDConta_corrente%>" >
 <input type=hidden name="IDporte"					value="<%=IDPorte%>" >
 <input type=hidden name="IDSubconta"				value="<%=IDSubconta%>" >
 <input type=hidden name="IDQtde1" 				value="<%=IDQtde1%>" > 
 <input type=hidden name="IDOrdena" 				value="<%=IDOrdena%>" >
 <input type=hidden name="Npagina" 				value="<%=Npagina%>" >
</center>
</center></center>
<tr>
<td>
<br>
<!--************ MONTA A TABELA DE RELATÓRIO ****************** !-->
<% Response.ContentType = "application/vnd.ms-excel" %>

<table width="80%" border="1" align="center" class="TableLine">
<tr>
 <th align="center">#</th>

 <th>Cliente</th>
 <th>Conta Corrente</th>
 <th>Sub-conta</th>
 <th>Porte</th>
 <th>Estado</th>
 <th>Disponibilidade Teórica</th>
 <th>Qtde Acessos Físicos </th>
 

<% 
  qtde =0 

 
if NOT RS.eof  then
   While Not RS.eof
	 qtde = qtde +1 
	 Tqtde  = Tqtde + RS("Valor_total")  
	 
	 IDConta_corrente		= RS("Conta_corrente")	
	 IDSubconta			= RS("Subconta")
	 Dispon 				= Busca_disponibilidade_cliente()
	 
  %>

<tr>    
 <td align="right"><% =qtde %></td> 
 <td ><%=RS("Razao_social")%></td>  
 <td ><%=RS("Conta_corrente")%></td>  
 <td ><%=RS("Subconta")%></td>  
 <td><%=RS("Porte_Cliente")%></td>  
 <td><%=RS("estado")%></td> 
 <td><%=Dispon%></td>

									
 <td align="center"><%=formatnumber(RS("Valor_total"),0)%></td>    
</tr>
<% 	
	   RS.MoveNext     
	
 Wend 
 RS.close : set RS = nothing	%>
 		
<% end if %> 
<tr class=clsSilver>
<td colspan="7"></td>  
<td align="center"><%= formatnumber(Tqtde,0) %></td>  
</tr>	 	
</table>
</td>
</tr>
</table>
</form>
</body>




