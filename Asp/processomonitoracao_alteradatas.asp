<%
Response.Expires = 0

'	- Sistema			: CLA
'	- Arquivo			: processomonitoracao_alteradatas.asp
'	- Descrição			: Altera as datas para cadatramento de serviços já ativados
'	- Responsável		: PSOUTO 
'	- Data				: 23/11/2005


%>
<!--#include file="../inc/data.asp"-->


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel=stylesheet type="text/css" href="../css/cla.css">
</HEAD>
<BODY topmargin=0 leftmargin=0 onunload="verifica()">

<script language='javascript' src="../javascript/solicitacao.js"></script>
<script language='javascript' src="../javascript/cla.js"></script>
<%


								

%>
<Form name=forme method=Post action="ProcessoCla.asp">
<SCRIPT LANGUAGE=javascript>
<!--

var fecha = 0
function verifica(){
	if(fecha ==0){
		alert('Favor informar as datas');
	}else{
		window.returnValue = "ok"
	}
	
}
function valida(){
	with(document.forme){
	
		if (!ValidarCampos(txtDtEntrAcesServ,"Data da Entrega")) return false
		if (!ValidarTipoInfo(txtDtEntrAcesServ,1,"Data da Entrega")) return;
		
		if (!ValidarCampos(txtDtConstrAcesFis,"Data da Construção do Acesso Físico")) return false
		if (!ValidarTipoInfo(txtDtConstrAcesFis,1,"Data da Construção do Acesso Físico")) return;
		
		if (!ValidarCampos(txtDtAtivServ,"Data da Ativação")) return false
		if (!ValidarTipoInfo(txtDtAtivServ,1,"Data da Ativação")) return;
		
		//if (!ValidarCampos(txtDtConcluSolic,"Data da Conclusão")) return false
			
		
		if(!checa('txtDtConstrAcesFis','<=','txtDtEntrAcesServ','A data de contrução deve ser menor ou igual a data de Entrega!')) return false;
		
		if(!checa('txtDtEntrAcesServ','<=','txtDtAtivServ','A data de Entrega deve ser menor ou igual a data de Ativação!')) return false;
		
		//if(!checa('txtDtConcluSolic','==','txtDtEntrAcesServ','A data de Conclusão deve ser igual a data de Entrega!')) return false;
		
		if(!checa('txtDtAtivServ','>=','txtDtEntrAcesServ','A data de Ativação deve ser maior ou igual a data de Entrega!')) return false;
		fecha = 1
		hdnAcao.value = "alteradatas"
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
		return true
		
	}
	
	
}

function checa(campo1,condicao,campo2,mensagem){
	with(document.forme){
	
		
	
		var d1_str = eval(campo1 + '.value');
		var d2_str = eval(campo2 + '.value');
		
		if(d1_str == "" || IsEmpty(d1_str)) return true
		if(d2_str == "" || IsEmpty(d2_str)) return true

		var d1 = new Date(d1_str.split('/')[2],d1_str.split('/')[1],d1_str.split('/')[0]);
		var d2 = new Date(d2_str.split('/')[2],d2_str.split('/')[1],d2_str.split('/')[0]);
		
		if(!eval(d1.getTime() + condicao + d2.getTime())){
			alert(mensagem);
			eval(campo1 + '.focus()');
			return false
		}else{
			return true
		}
	}
}
//-->
</SCRIPT>

<%


dblSolId =  Trim(Request("sol_id"))

Set objRSSolic =	db.execute("CLA_sp_view_solicitacaomin " & dblSolId)

if Not objRSSolic.Eof and Not objRSSolic.Bof then
	
	strDtEntrAcesServ		= Formatar_Data(Trim(objRSSolic("Acl_DtEntregaServico")))
	strDtConstrAcesFisTer 	= Formatar_Data(Trim(objRSSolic("Sol_DtConclusao"))) 		'Data Construção Acesso Físico Terceiro
	strDtAtivServ			= Formatar_Data(Trim(objRSSolic("Acl_DtAtivacaoServico")))	
	strDtConcluSolic		= Formatar_Data(Trim(objRSSolic("Sol_DtConclusao")))	
	
END IF 
%>
<table rules="groups" border=0 cellspacing="1" cellpadding ="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="350px" align=center>
	<tr>
		<th colspan=2>
			<p align=center>Data da Entrega do  Acesso Físico</p>
		</th>
	</tr>
	<tr class=clsSilver>
		<td>Data</td>
		<td>
			<input type=text class=text name=txtDtEntrAcesServ size=10 value=""  maxlength=10 onKeyPress="OnlyNumbers();AdicionaBarraData(this)" >&nbsp;(dd/mm/aaaa)
		</td>
	</tr>
	<tr>
		<th colspan=2>
			<p align=center>Data da Construção do Acesso Físico</p>
		</th>
	</tr>
	<tr class=clsSilver>
		<td>Data</td>
		<td>
			<input type=text class=text name=txtDtConstrAcesFis size=10 value=""  maxlength=10 onKeyPress="OnlyNumbers();AdicionaBarraData(this)" >&nbsp;(dd/mm/aaaa)
		</td>
	</tr>
	<tr>
		<th colspan=2>
			<p align=center>Data da Ativação do Serviço</p>
		</th>
	</tr>
	<tr class=clsSilver>
		<td>Data</td>
		<td>
			<input type=text class=text name=txtDtAtivServ size=10 value="" maxlength=10 onKeyPress="OnlyNumbers();AdicionaBarraData(this)" >&nbsp;(dd/mm/aaaa)
		</td>
	</tr>
	
</table>

	<input type=hidden name=hdnSolId value="<%=dblSolId%>">
	<input type=hidden name=hdnAcao value="alteradatas">
	
	
	<table cellspacing=1 cellpadding=1  width=350px border=0>
		<tr>
			<td align=center>
			<input type="button" class="button" name="btnGravar" value="Gravar" onClick="valida()" >
			<!--
			<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.close()" >-->
			</td>
		</tr>
	</table>		
</form>	

<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "200"
	    height      = "100"
	    frameborder = "2"
	    scrolling   = "yes" 
	    align       = "left">
</iFrame>
							 
<P>&nbsp;</P>

</BODY>
</HTML>

