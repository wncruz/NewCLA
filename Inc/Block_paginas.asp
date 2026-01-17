<%
'•BULL
'	- Sistema			: CLA
'	- Arquivo			: Block_consulta.asp
'	- Responsável		: PRSS
'	- Descrição			: Bloqueia as páginas em um determinado horário.
'	- Criação			: 11/04/2007
%>
<META HTTP-EQUIV="Refresh" CONTENT="900;">
<%
Dim var_hora_bloq_ini,var_hora_bloq_fim,var_mensagem

'Hora início/fim do Bloqueio:
var_hora_bloq_ini = Cdate("07:00:00 AM") '07:00
var_hora_bloq_fim = Cdate("16:00:00 PM") '16:00

'Mensagem de aviso: 
var_mensagem = "Esta consulta está temporariamente indisponível para a monitoração do sistema e será liberada após as "&var_hora_bloq_fim&".<br> "&_
"Em caso de dúvida entrar em contato com EUTIMIO, ramal 521-8564."

'Bloqueio:
if time() >= var_hora_bloq_ini and time() < var_hora_bloq_fim then
  response.write "<center><br><br><br><h4>" & var_mensagem & "</h4><br>"
  response.end
end if
%>