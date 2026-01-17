<%
'	- Sistema			: CLA
'	- Arquivo			: cst_estruturaCombo.asp
'	- Responsável		: EDAR
'	- Descrição			: Consulta
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblID
Dim objRSVel
Dim strSel
Dim dblIDAtual

dblID = Request.QueryString("ID")
'dblID = Request.QueryString("ID")
'response.write "<script>alert('"&dblID&"')</script>"

if Trim(dblID) = "" then
	dblID = Request.Form("hdnId")
End if	

If Request.Form("hdnAcao")="Gravar" then

	If dblID="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	End if

	Vetor_Campos(2)="adInteger,2,adParamInput,"& Request.Form("cboEstruturaTecnologia") 
	Vetor_Campos(3)="adWChar,60,adParamInput," & Request.Form("txtLabel") 
	Vetor_Campos(4)="adWChar,5,adParamInput,"& Request.Form("cboFormato") 
	Vetor_Campos(5)="adWChar,5,adParamInput," & Request.Form("txtTamanho") 
	Vetor_Campos(6)="adWChar,5,adParamInput,"& Request.Form("rdoCompartilha") 
	Vetor_Campos(7)="adWChar,5,adParamInput," & Request.Form("rdoObrigatorio") 
	Vetor_Campos(8)="adWChar,5,adParamInput," & Request.Form("rdoStatus") 
	Vetor_Campos(9)="adWChar,10,adParamInput," & strloginrede 
	Vetor_Campos(10)="adInteger,2,adParamOutput,0"  
	
	Call APENDA_PARAM("CLA_sp_ins_EstruturaTecnologia",10,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End if

If dblID <> "" then
	
	'Response.Write "<script language=javascript>alert('"&dblID&"')</script>"
				
	Set objRSAssocTecFac = db.execute("CLA_sp_sel_EstruturaCombo null, " & dblID)
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="consultaestruturaTecnologiaFacilidade.asp" method="post" >
<SCRIPT LANGUAGE="JavaScript">


function GravarEstruturaCombo()
{
	//if (!checa(document.forms[0])) return
	with (document.forms[0])
	{
		action = "estruturaCombo.asp?ID=<%=dblID%>"
		//hdnAcao.value = "Gravar"
		submit()
	}
}
</script>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnId value="<%=dblID%>" >
<tr>
	<td>
		<table border="0" cellspacing="1" cellpadding=0 width="760">

			
		
		
			<tr>
				<th colspan=2><p align="center">Estrutura de Combo</p></th>
			</tr>
			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Combo </font>
				</td>
				
				<td>
					
					<%= TratarAspasHtml(objRSAssocTecFac("newcombo_nome"))   %>
				
					
				</td>
			</tr>
			<tr>
				<td colspan=2 align="center"><br>
					<input type="button" class="button" name="btnGravar" value="ADICIONAR" onClick="GravarEstruturaCombo()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
					<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('estruturaCombo_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');" >
					<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
				</td>
			</tr>
		</table>
		<div class="table-responsive">          
	    <table border="0" cellspacing="1" cellpadding=0 width="760">
			<thead>
			  <tr>
				<th>Editar</th>
				<th>Label</th>
				
				<th>STATUS </th>
						
			  </tr>
			</thead>
			<tbody>
			<%
			if objRSAssocTecFac("estrutura_combo_id") <> "" then
				'AQUI COMEÇA O DO WHILE ONDE TRARÁ AS LINHAS DE ACORDO COM O SELECT FEITO
				do while not objRSAssocTecFac.eof 
				%>
					  <tr>
						<td>
				 <a href="estruturaCombo.asp?assoc=<%=objRSAssocTecFac("estrutura_combo_id")%>">
				<!--<img src="../imagens/Editar.gif">-->  <center>Alterar</center>
						</a>		
						</td>
						<td><%=TratarAspasHtml(objRSAssocTecFac("Label"))%></td>
						<!--<td><%=TratarAspasHtml(objRSAssocTecFac("Status"))%></td> -->
						
						<td> <%if objRSAssocTecFac("Status") = "S" then 
							%>
							ATIVO 
							<% ELSE %> 

							INATIVO 
							<% END IF  %> </td>
					  </tr>
				<%
				'LOOP DEPOIS DA LINHA PARA QUE SE REPITA ENQUANTO TIVER REGISTROS NO SELECT FEITO
				objRSAssocTecFac.movenext
				loop
			'Else
			'	Response.Write "<script language=javascript>alert('Registro(s) não encontrado(s)')</script>"
				'Response.Write "<script language=javascript>parent.spnLinks.innerHTML = ''</script>"
			'	Response.End 
			End if	
				%>	  
					</tbody>
				  </table>

				  </div>
			
		
	</td>
</tr>
</table>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('cbonewTecnologia');
//-->
</SCRIPT>

</html>
<%
Set objRSAssocTecFac = Nothing
DesconectarCla()
%>
