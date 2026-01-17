<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="monta-sql.asp"-->
<html>
<%		
		'*** busca Variaveis
			Filtro				= request.form("Filtro")		
		   IDOrdena			= request.form("IDOrdena")
		   IDestado 			= request.form("IDestado")		
 			IDEnd_sigla		= request.form("IDEnd_sigla")
			IDEnd_bairro		= request.form("IDEnd_bairro")			
			IDEnd_Nome			= request.form("IDEnd_Nome")	
			IDproprietario	= request.form("IDproprietario")
			IDtecnologia		= request.form("IDtecnologia")	
			IDConta_corrente		= request.form("IDConta_corrente")
			IDSubconta				= request.form("IDSubconta")			
 		    IDQtdedet			= converte_inteiroLongo(request.form("IDQtdedet"),0)
   		    IDqtdedet1		= converte_inteiroLongo(request.form("IDqtdedet1"),0)


			strSQL =  Monta_SQL_consolida_endereco_cliente()
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
<SCRIPT LANGUAGE="JavaScript">

function filtrar(){

	mform 		   				= document.Formfiltro;
	mform.action 		 		= "detalhe_acesso_endereco_cliente_dispon.asp"
	mform.IDOrdena.value  		= mform.IDOrdena.value;
	mform.filtro.value			= "1";
	mform.target 				= "_self";
	mform.method 				= "post";
	mform.submit();
}

function enviar(IDacf_id) {

	mform 		   						= document.FormRelat;
	mform.action 						= "detalhe_acesso_servico_cliente.asp";
	mform.filtro.value					= "1";
	mform.IDacf_id.value 				= IDacf_id;
	mform.IDOrdena.value 				= "";
	mform.target = "_self";
	mform.method = "post"; 
	mform.submit();
}



function ordenar(IDOrdena){
	mform 		           = document.FormRelat;
	mform.action 			 ="detalhe_acesso_endereco_cliente_dispon.asp"
	mform.IDOrdena.value  = IDOrdena;
	mform.filtro.value	 = "1";
	mform.target = "_self";
	mform.method = "post";
	mform.submit();
}

function Imprimir()
{
	window.print()
}

// --></script>

<table width="100%" border="1">
<tr>
<td>
<table width="100%" border="0">
<tr>
<td>
<center><h3>CLA - Controle Local de Acesso</h3><center>
<h4 align="center">Relatório de acessos físicos&nbsp;&nbsp; - <%= date() %></h4>
<h5 align="center"><% =rs("razao_social") %></h5>
<center>
<Form name="FormRelat" method="Post" action="detalhe_acesso_endereco.asp" target="_self" >
 <input type=hidden name="filtro" 				value="1" >
 <input type=hidden name="IDEstado" 			value="<%=IDEstado%>" >
 <input type=hidden name="IDQtde" 				value="<%=IDQtde%>" >
 <input type=hidden name="IDQtde1" 			value="<%=IDQtde1%>" > 
 <input type=hidden name="IDOrdena" 			value="<%=IDOrdena%>" >
 <input type=hidden name="IDEnd_sigla" 		value="<%=IDEnd_sigla%>" >
 <input type=hidden name="IDEnd_Nome" 		value="<%=IDEnd_Nome%>" >
 <input type=hidden name="IDEnd_bairro" 		value="<%=IDEnd_bairro%>" > 
 <input type=hidden name="IDproprietario" 	value="<%=IDproprietario%>" >  
 <input type=hidden name="IDConta_corrente" 	value="<%=IDConta_corrente%>" > 
 <input type=hidden name="IDSubconta" 		value="<%=IDSubconta%>" > 
 <input type=hidden name="IDtecnologia" 		value="<%=IDtecnologia%>" >  
 <input type=hidden name="IDacf_id" 			value="<%=IDacf_id%>" > 
 <input type=hidden name="Npagina" 			value="<%=Npagina%>" >
</center>
<tr>
<td>
<br>
<!--************ MONTA A TABELA DE RELATÓRIO ****************** !-->
<%Response.ContentType = "application/vnd.ms-excel"  %>
<table width="100%" border="1" align="center" class="TableLine">
<tr>
 <th align="center">#</th>
 <th>Estado</th>
 <th>Bairro</th>
 <th>Endereço</th>
 <th>Proprietário</th>
 <th>Tecnologia</th>
 <th>Velocidade do Físico</th>
 <th>Velocidade do Lógico</th>
 <th>Disponibilidade</th>
<th>Qtde Acessos Lógicos</th>
 <% if not  RS.eof then

   While Not RS.eof  
 
	 qtde = qtde +1 
	 Tqtde  = Tqtde + RS("qtde_logico")   
  %>
<tr>    

 <td align="right"><% =qtde %></td> 
 <td><%=RS("estado")%></td>
 <td><%=RS("end_bairro")%></td>
 <td><%=trim(RS("End_tpl_sigla") & " " & RS("End_Nomelogr") & " " & RS("End_NroLogr"))%></td>
 <td><%=trim(RS("Proprietario")) %></td>
 <td><%=trim(RS("tecnologia")) %></td>
 <td align="right"><%=RS("Vel_Conversao") %></td>
 <td align="right"><%=RS("vel_logico") %></td>
 <td align="right"><%=RS("disponibilidade") %></td>
 <td align="center"><%=formatnumber(RS("qtde_logico"),0)%></td>    
</tr> 
<%    
     RS.MoveNext
  Wend 
  RS.close : set RS = nothing	%>
 		
<% end if %> 
<tr class=clsSilver>
<td colspan="9"></td>  
<td align="center"><%= formatnumber(Tqtde,0) %></td>  
</tr>	 	
</table>
</td>
</tr>
</table>
</form>
</body>










