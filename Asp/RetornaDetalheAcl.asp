<!--#include file="../inc/data.asp"-->
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script type="text/javascript">
function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		hdnSolId.value = dblSolId
		hdnAcao.value = "DetalheSolicitacao"
		target = "_New"
		action = "ConsultaGeralDet.asp"
		submit()
	}	
}

var TableBackgroundNormalColor = "#ffffff";
var TableBackgroundMouseoverColor = "#eeeee1";
function ChangeBackgroundColor(row) { row.style.backgroundColor = TableBackgroundMouseoverColor; }
function RestoreBackgroundColor(row) { row.style.backgroundColor = TableBackgroundNormalColor; }
</script>
<style type="text/css">
tr { background-color: white }
tr:hover { background-color: black };
</style>
<Form name=Form1 method=Post> 
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnAcao>
<table border=0 width=760 cellspacing=1 cellpadding=1 >
<tr>
<th>&nbsp;<span id=spnCol3>Id-Lógico</span></th>
<th>&nbsp;<span id=spnCol1>Velocidade</span></th>
<th>&nbsp;<span id=spnCol1>Cliente</span></th>
<th>&nbsp;<span id=spnCol1>Conta-Corrente</span></th>	
</tr>
<%

  sAcf_id    = Request("acf_id")
			 
	Vetor_Campos(1)="adInteger,4,adParamInput," & sAcf_id
	strSql = APENDA_PARAMSTR("CLA_sp_sel_acl_cons",1,Vetor_Campos)		

	Set objRS = db.Execute(strSql)
			If Not objRS.eof and  Not objRS.bof then			
			  strHtml = ""
				While Not objRS.Eof 
				    If strClass = "clsSilver2" then strClass = "clsSilver" else strClass = "clsSilver2" End if 
						'strHtml = strHtml & "<tr onmouseover='ChangeBackgroundColor(this);' onmouseout='RestoreBackgroundColor(this);'><td>" & objRS("Acl_IdAcessoLogico") & "</td><td>" & objRS("vel_desc") &_
						strHtml = strHtml & "<tr class=" & strClass & "><td><a href='javascript:DetalharItem(" & objRS("Sol_id") & ")'>&nbsp;<b>" & objRS("Acl_IdAcessoLogico") & "</b></a></td><td>" & objRS("vel_desc") &_						
															  "</td><td>" & objRS("cli_nome") & "</td><td>" & objRS("Conta_Cliente") & "</td></tr>"		
						objRS.MoveNext
				Wend  
			End If  
  
  Response.Write strHtml
  Response.End    
%>
</table>

</Form>
</body>
</html>



