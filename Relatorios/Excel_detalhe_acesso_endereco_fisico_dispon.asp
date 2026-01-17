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


   
  	       IDestado 				= request.form("IDestado")	
 			IDEnd_sigla			= request.form("IDEnd_sigla")
			IDEnd_bairro			= request.form("IDEnd_bairro")			
			IDEnd_Nome				= request.form("IDEnd_Nome")	
			IDproprietario		= request.form("IDproprietario")
			IDtecnologia			= request.form("IDtecnologia")
			IDestacao				= request.form("IDestacao")
			IDTipoestacao			= request.form("IDTipoestacao")

   		    IDQtde					= converte_inteiroLongo(request.form("IDQtde"),0)
   		    IDQtde1				= converte_inteiroLongo(request.form("IDQtde1"),0)


		strSQL =  Monta_SQL_consolida_endereco()
		SET rs = Server.CreateObject("ADODB.Recordset")		
		rs.Open strSQL, db,2,2
		


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
<h4 align="center">Relatório de acessos físicos - <%= date() %></h4>
<center>
<Form name="FormRelat" method="Post" action="excel_detalhe_acesso_endereco_fisico_dispon.asp" target="_self" >
 <input type=hidden name="filtro" 				value="1" >
 <input type=hidden name="IDEstado" 			value="<%=IDEstado%>" >
 <input type=hidden name="IDQtde" 				value="<%=IDQtde%>" >
 <input type=hidden name="IDQtde1" 			value="<%=IDQtde1%>" > 
 <input type=hidden name="IDOrdena" 			value="<%=IDOrdena%>" >
 <input type=hidden name="IDEnd_sigla" 		value="<%=IDEnd_sigla%>" >
 <input type=hidden name="IDEnd_Nome" 		value="<%=IDEnd_Nome%>" >
 <input type=hidden name="IDEnd_bairro" 		value="<%=IDEnd_bairro%>" > 
 <input type=hidden name="IDproprietario" 	value="<%=IDproprietario%>" >  
 <input type=hidden name="IDtecnologia" 		value="<%=IDtecnologia%>" >  
 <input type=hidden name="IDestacao" 			value="<%=IDestacao%>" >  
 <input type=hidden name="IDTipoestacao" 		value="<%=IDTipoestacao%>" >   
 <input type=hidden name="IDacf_id" 			value="<%=IDacf_id%>" > 
 <input type=hidden name="Npagina" 			value="<%=Npagina%>" >
</center>
</center></center>
<tr>
<td>
<br>
<!--************ MONTA A TABELA DE RELATÓRIO ****************** !-->
<%   

Response.ContentType = "application/vnd.ms-excel" 

%>
<table width="100%" border="1" align="center" class="TableLine">
<tr>
 <th align="center">#</th>
 <th>Estado</th>
 <th>Bairro</th>
 <th>Endereço</th>
 <th>Razão Social</th>
 <th>Proprietário</th>
 <th>Tecnologia</th>
 <th>Estação Entrega</th>
 <th>Tipo Estação Entrega</th>  
 <th>Velocidade Físico</th>
 <th>Disponibilidade Teorica</th>
 <th>Qtde Acessos Lógicos</th>
	 
 <% 

 if not  RS.eof then


   While Not RS.eof  

	 qtde = qtde +1 
	 Tqtde  = Tqtde + RS("qtde_logico")   
  %>
<tr>    

 <td align="right"><% =qtde %></td> 
 <td><%=RS("estado")%></td>
 <td><%=RS("end_bairro")%></td>
 <td><%=trim(RS("End_tpl_sigla") & " " & RS("End_Nomelogr") & " " & RS("End_NroLogr"))%></td>
 <td><%=rs("razao_social")%></td>
 <td><%=trim(RS("Proprietario")) %></td>
 <td><%=trim(RS("tecnologia")) %></td>
 <td><%=trim(RS("EstacaoEntrega")) %></td>
 <td><%=trim(RS("TipoEstacaoEntrega")) %></td>
 <td align="right"><%=RS("Vel_fisico") %></td>
 <td align="right"><%=RS("disponibilidade") %></td>
 <td align="center"><%=formatnumber(RS("qtde_logico"),0)%></td>    
</tr> 
<%    
     RS.MoveNext
  Wend 
  RS.close : set RS = nothing	%>
 		
<% end if %> 
<tr class=clsSilver>
<td colspan="11"></td>  
<td align="center"><%= formatnumber(Tqtde,0) %></td>  
</tr>	 	
</table>
</td>
</tr>
</table>
</form>
</body>






