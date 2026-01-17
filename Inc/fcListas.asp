<%
'•BULL
'	- Sistema			: CLA
'	- Arquivo			: fcListas.asp
'	- Responsável		: PRSS
'	- Descrição			: Funções para a lista de pendências de terceiros.
'	- Criação			: 13/01/2007
%>
<%
function erro()
  if err.number <> 0 then
	Response.Write "-------------------------<BR>" & chr(10)
	Response.Write err.number & "<BR>" & chr(10)
	Response.Write err.Description & "<BR>" & chr(10)
	Response.Write err.Source & "<BR>" & chr(10)
	err.Raise
	Response.Write "-------------------------<BR>" & chr(10)
	Response.End
  end if
end function

function trocaplics(texto)
  trocaplics = texto
  if len(texto)> 1 then
    trocaplics = replace(trocaplics,"'","")
  End if
end function

function listaTarefas (usuario,usuario_logado,status)
  on error resume next

  dim rs
  dim conn
  dim arrRs
  dim i
  dim Strsql
  dim Sol_id
  dim Indice
  DIM userid
  DIM objRS
  
  DIM LTAREFA
  DIM LSOL
  DIM LUSER
  DIM LDATA
  DIM LTEMPO
  DIM LCLI
  DIM LEND
  DIM LPED
  
  dim tipo
  dim tarefa
  
  SET objRS = Server.CreateObject("ADODB.Recordset")
  Vetor_Campos(1)="adInteger,8,adParamInput,"	& 	usuario
  Vetor_Campos(2)="adWChar,20,adParamInput,"	& 	usuario_logado
  Vetor_Campos(3)="adWChar,20,adParamInput,"	& 	status
  Vetor_Campos(4)="adInteger,8,adParamInput,"	& 	strSol_Id
  Vetor_Campos(5)="adWChar,2,adParamInput,"	& 	uf
  
  strSql = APENDA_PARAMSTRSQL("cla_sp_sel_tarefas ",5,Vetor_Campos)
  
  set objRS = db.execute(STRSQL)
	
  IF not objRS.EOF THEN
  %>
  <TABLE class="EBTFormulario" border="0" width="760">
  <TR class="EBTtitulo">
    <TD align="center" colspan="6" class="EBTtitulo"><p align=center >Tarefas Pendentes</p></TD>
  </TR>
  <tr>
    <td>
    <table border="0" cellspacing="1" cellpadding=0 width="760">
      <tr>
	    <th>&nbsp;Tarefa</th>
		<%IF strtipo = "T" THEN%>
		  <th>&nbsp;Pedido</th>
		<%END IF%>
	    <th nowrap>&nbsp;Nº Solicitação</th>
	    <th>&nbsp;Responsável</th>
	    <th>&nbsp;Data Pedido</th>
	    <th>&nbsp;Tempo Tarefa</th>
	    <th>&nbsp;Cliente</th>
	    <th>&nbsp;Endereço</th>
      </tr>
	  <tr>
	  <td>
  <% 
  Call PaginarRS(0,strSql)
    intCount=1
    if not objRSPag.Eof and not objRSPag.Bof then
	  For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class=<%=strClass%>>
		    <td>
			<%
			select case objRSPag("tarefa")
			Case "AVALIAR"
			  %><a href="javascript:Avaliacao(<%=objRSPag("sol_id")%>,'<%=Ucase(trim(objRSPag("usu_username")))%>')"><%=objRSPag("tarefa")%></a><%
			Case "ENTREGAR"
			  %><a href="javascript:EntregaAtivacao(<%=objRSPag("sol_id")%>,'<%=Ucase(trim(objRSPag("usu_username")))%>')"><%=objRSPag("tarefa")%></a><%
			Case "ATIVAR"
			  %><a href="javascript:EntregaAtivacao(<%=objRSPag("sol_id")%>,'<%=Ucase(trim(objRSPag("usu_username")))%>')"><%=objRSPag("tarefa")%></a><%
			Case "ALOCAR"
			  %><a href="javascript:AlocacaoCancelamento(<%=objRSPag("ped_id")%>,<%=objRSPag("sol_id")%>,'<%=Ucase(trim(objRSPag("usu_username")))%>')"> <%=objRSPag("tarefa")%></a><%
			Case "EXECUTAR"
			  %><a href="javascript:Execucao(<%=objRSPag("ped_id")%>,'<%=trim(objRSPag("usu_username"))%>')"><%=objRSPag("tarefa")%></a><%
			Case "ACEITAR"
			  %><a href="javascript:Aceitacao(<%=objRSPag("ped_id")%>,'<%=Ucase(trim(objRSPag("usu_username")))%>')"><%=objRSPag("tarefa")%></a><%
			Case "CANCELAR"
			  %><a href="javascript:AlocacaoCancelamento(<%=objRSPag("ped_id")%>,<%=objRSPag("sol_id")%>,'<%=Ucase(trim(objRSPag("usu_username")))%>')"> <%=objRSPag("tarefa")%></a><%
			Case "LIBERAR"
			  %><a href="javascript:Liberacao(<%=objRSPag("ped_id")%>,<%=objRSPag("sol_id")%>,'<%=Ucase(trim(objRSPag("usu_username")))%>')"><%=objRSPag("tarefa")%></a><%
			End select
			%>
			</td>
			<%IF strtipo = "T" THEN%>
			  <%var_pedido = objRSPag("ped_prefixo") & "-" & objRSPag("ped_numero") & "/" & objRSPag("ped_ano")%>
			<td nowrap><%=var_pedido%></td>
			<%END IF%>
			<td><%=objRSPag("sol_id")%></td>
			<td><%=Ucase(trocaplics(objRSPag("usu_username")))%></td>
			<td><%=trocaplics(objRSPag("data"))%></td>
			<td><%=trocaplics(objRSPag("tempo"))%></td>
			<td><%=trocaplics(objRSPag("nome"))%></td>
			<td><%=trocaplics(objRSPag("endereco"))%></td>
		</tr>
		<%
		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	  Next
	  session("ss_intCurrentPage") = intCurrentPage
	  %><!--#include file="../inc/ControlesPaginacao.asp"--><%
    End if
    %>
	</td>
    </TR>
    <TR>
      <TD align="center">
	    <br>
	    <br>
      </TD>
    </TR>
    </table>
  </TABLE>
  <%
  else
	Response.write "<center><font size=2 color='ff0000'><li> Registro(s) não encontrado(s).</font></center>"
  end if
end function
%>
<script language="JavaScript">
<%if permissao <> "GIC-L" then%>
  function ValidaPermissao(dblUsuario)
  {
  if(dblUsuario != '' && dblUsuario != '<%=Ucase(strLoginRede)%>' ){
    alert('Não é possivel visualizar uma solicitação de outro usuário!');
    return(false);
    }
  }
<%else%>
  function ValidaPermissao(dblUsuario)
  {
    return(true);
  }
<%end if%>

function AlocacaoCancelamento(dblPedId,dblSolId,dblUsuario)
{
  if (ValidaPermissao(dblUsuario) != false){
    with(document.frmprinc){
	 	hdnSolId.value = dblSolId
	 	hdnPedId.value = dblPedId
     	if (dblUsuario != '')
	  	 {
     	  target = "_new"
     	  action = "Facilidade.asp"
    	 }else{
     	  hdnAcao.value = "AlocacaoGLA"
     	  action = "ProcessoFac.asp"
     	  target = 'IFrmProcesso'
    	 }
		 submit()
     }
   }
}

function Execucao(dblPedId,dblUsuario)
{
  Executar(dblPedId)
}

function Aceitacao(dblPedId,dblUsuario)
{
  Aceitar(dblPedId)
}

function Liberacao(dblPedId,dblSolId,dblUsuario)
{
  with(document.frmprinc){
			cboUsuario.value= "<%=dblUsuId%>"
			hdnSolId.value = dblSolId 
			hdnAcao.value = "AlteracaoCad"
			window.open('AlteracaoCad.asp?SolID='+dblSolId+ ' &libera=1  &provedor=' +dblPedId ,'Edicao','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,width=780,height=540,top=0,left=0'); 
  }
}

function Monitoracao(dblSolId,dblUsuario)
{
  if (ValidaPermissao(dblUsuario) != false){
    with(document.frmprinc){
			cboUsuario.value= "<%=dblUsuId%>"
			hdnSolId.value = dblSolId 
			hdnAcao.value = "AlteracaoCad"
			target = '_new'
			action = "Monitoracao.asp?txtSolId=" + dblSolId
			submit()
	}
  }
}

function Avaliacao(dblSolId,dblUsuario)
{
  if (ValidaPermissao(dblUsuario) != false){
    with(document.frmprinc){
			hdnSolId.value = dblSolId 
			hdnAcao.value = "AlteracaoCad"
			target = '_new'
			action = "AvaliarAcesso.asp"
			submit()
	}
  }
}

function EntregaAtivacao(dblSolId,dblUsuario)
{
  if (ValidaPermissao(dblUsuario) != false){
    with(document.frmprinc){
			cboUsuario.value= "<%=dblUsuId%>"
			txtSolId.value = dblSolId
			hdnLocation.value = "MONITORACAO"
			target = '_new'
			action = "ProcessoEntregaAtivacao.asp"
			submit()
    }
  }
}

function ResgatarUsuarioCtfc(obj)
{
	with (document.frmprinc)
	{
		hdnAcao.value = "ResgatarUsuarioCtfc"
		hdnCboStatus.value = obj.value
		cboUsuario.value= "<%=dblUsuId%>"
		target = "IFrmProcesso"
		action = "PendenciaTerc_combos.asp"
		submit()
	}	
}
</script>