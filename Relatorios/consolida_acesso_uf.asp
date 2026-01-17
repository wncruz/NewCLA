<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="RelatoriosCla.asp"-->
<!--#include file="monta-sql.asp"-->

<html>
<%


Dim strXls
Dim strLink
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
<!--************ MONTA A TABELA DE FILTROS ****************** !-->
<SCRIPT LANGUAGE="JavaScript">

function filtrar(){

	mform 		   				= document.Formfiltro;
	mform.action 		 		= "consolida_acesso_uf.asp"
	mform.IDOrdena.value  	= mform.IDOrdena.value;
	mform.filtro.value		= "1";
	mform.target 				= "_self";
	mform.method 				= "post";
	mform.submit();
}

function enviar(IDEstado,IDProprietario) {

	mform 		   						= document.FormRelat;
	mform.action 						= "detalhe_acesso_endereco_fisico_dispon.asp";
	mform.filtro.value				= "1";
	mform.IDEstado.value 			= IDEstado;
	mform.IDProprietario.value 		= IDProprietario;
	mform.IDOrdena.value 			= "";
	mform.target = "_self";
	mform.method = "post"; 
	mform.submit();
}



function ordenar(IDOrdena){
	mform 		           = document.FormRelat;
	mform.action 			 ="consolida_acesso_uf.asp"
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

function RelExcel(){
	mform 		           = document.FormRelat;
	mform.action 			 = "excel_consolida_acesso_uf.asp"
	mform.IDEstado.value  = mform.IDEstado.value;
	mform.filtro.value	 = "1";
	mform.target = "_blank";
	mform.method = "post";
	mform.submit();
}


// --></script>

<Form name="Formfiltro" method="Post" action="consolida_acesso_uf.asp" target="_self">
<input type="hidden" name="filtro" value="1"  >
<input type=hidden name="IDOrdena" 		value="<%=IDOrdena%>" >
<table border="1" width="100%">
<tr>
<td>
<table  width="100%">
<tr><td align="right" width="50%" >
<!--<a target=_self href=javascript:RelExcel()><img src='../imagens/excel.gif' border=0></a>!-->&nbsp
<td align="left" width="50%">
<a target=_self href="javascript:window.print()" ><img src='../imagens/impressora.gif' border=0></a></td>
</tr>
</table>
<table  width="100%">
<tr class=clsSilver>

 		<td ALIGN="RIGHT"><font ></font>UF</td>
		<td COLSPAN="2">
  		  <% strSQL = Monta_SQL_estado()
			SET RS= Server.CreateObject("ADODB.Recordset")
			RS.Open strSQL,db	 %>
			<select name="cboUF">
				<option ></option>
				<% While Not RS.eof %>
					<Option value="<%=RS("est_sigla")%>" <%if IDestado=RS("est_sigla") then %> selected <% end if %>><%=RS("est_sigla")%> - <%=RS("est_desc")%></Option>
				<% RS.MoveNext
				Wend 
				RS.close : set RS = nothing
				%>				
       </select>
		</td>
	</tr>
</table>
<table width="100%">
<tr class=clsSilver>
   <td >
        <p align="right">	<input type=button name="Filtrar" value="Filtrar" class="button" onclick="filtrar()" >
	</p>
	</td>
   <td>
        <p align="left">&nbsp;
        <input type=button name="Voltar" value="Voltar" class="button" onclick="javascript:window.history.back(-1)" ></p>
    </td>		
</tr>
</table>
</tr>
</table>
</form>
<table width="100%" border="0">
<tr>
<td>
<h5 align="center">Relatório de Consolidado por Proprietário</h5>
<center>
<Form name="FormRelat" method="Post" action="detalhe_acesso_endereco_servico.asp" target="_self" >
 <input type=hidden name="filtro" 				value="1" >
 <input type=hidden name="IDEstado" 			value="<%=IDEstado%>" >
 <input type=hidden name="IDProprietario" 	value="<%=IDProprietario%>" >
 <input type=hidden name="IDOrdena" 			value="<%=IDOrdena%>" >
<table border=0 width=760><tr><td colspan=2 align=right>



</td></tr>
</table>
 
</center>
<tr>
<td>
<br>

<!--************ MONTA A TABELA DE RELATÓRIO ****************** !-->

<table width="60%" border="1" align="center" class="TableLine">

<th align=center>#</th>
<th><a class=white target=_self href=javascript:ordenar('Proprietario')>Proprietário</a></th>
<th><a class=white target=_self href=javascript:ordenar('Valor_total')>Qtde Acessos Físicos </A></th>
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
 <td align="center"><a  href="javascript:enviar('<%=RS("Estado")%>','<%=RS("Proprietario")%>');" target="_self"><%=formatnumber(RS("Valor_total"),0)%></a></td>    
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






