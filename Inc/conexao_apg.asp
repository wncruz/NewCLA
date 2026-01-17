<%
'•BULL
'	- Sistema			: CLA
'	- Arquivo			: conexao_apg.asp
'	- Responsável		: PRSS
'	- Descrição			: Configura a conexão com o APG.
'	- Criação			: 12/03/2007

'If trim(strLoginRede) = "PRSS" or trim(strLoginRede) = "T3PRSS" or trim(strLoginRede) = "T3LBOAS" or trim(strLoginRede) = "T3ANJOS" Then
'If trim(strLoginRede) = "PRSS" or trim(strLoginRede) = "T3PRSS" or trim(strLoginRede) = "T3ENA" or trim(strLoginRede) = "T3LBOAS" Then
  '***Ambiente DSV APG ***
  'AdresserPath = "http://10.2.30.4:9092/dsvapg/services/ApiaWS" 'Servico 'Endereço atual dos testes
  'AdresserPath = "http://10.2.30.4:9095/tstapg/services/ApiaWS"
  'AdresserPath = "http://10.2.13.197:9093/hmgapg/
  'AdresserPath ="http://10.102.2.211:8080/APGDefault/services/ApiaWS"

'else

  '***Ambiente TST APG ***
  'AdresserPath = "http://10.2.30.4:9095/tstapg/services/ApiaWS" 'Servico
  'AdresserPath = "http://10.2.13.197:9092/hmgapg/services/ApiaWS"
  'APG HMG2
  'AdresserPath = "http://10.2.13.197:9093/hmgapg/services/ApiaWS"
'End IF


'select case Ucase(session("ss_amb_apg"))
'Case "TST"
 ' AdresserPath = "http://10.2.30.4:9095/tstapg/services/ApiaWS"
'Case "UAT"
'  AdresserPath = "http://10.2.13.197:6367/hmgapg/services/ApiaWS"
'Case "HMG"
'  AdresserPath = "http://10.2.13.197:9092/hmgapg/services/ApiaWS"
'Case "HMG2"
'  AdresserPath = "http://10.2.13.197:9093/hmgapg/services/ApiaWS"
'Case "DSV"
'  AdresserPath = "http://10.2.30.4:9092/dsvapg/services/ApiaWS" 'Servico 'Endereço atual dos testes
'Case "TRN"
'  AdresserPath = "http://10.2.13.197:9072/trnapg/services/ApiaWS"
'Case "DLV"
'  AdresserPath = "http://10.2.30.2:9090/dlvapg/services/ApiaWS"
'Case "MIG"
'  AdresserPath = "http://10.2.30.2:8484/migapg/programs/login/login.jsp"
'Case "PRD"
 ' AdresserPath = "http://10.4.40.60:8090/prdapg/services/ApiaWS"
  AdresserPath = "http://172.16.4.22:8090/prdapg/services/ApiaWS" 
'End select

'AdresserPath = "http://10.2.30.4:9092/dsvapg/services/ApiaWS" 'OLD
'AdresserPath = "http://10.102.2.9:8080/APGDefaultSimulacao/services/ApiaWS" '"http://10.2.30.4:9092/dsvapg/services/ApiaWS"

StrLogin = "sisecla"
StrSenha = "36-74142-12611467-95-2011442125126-58-15117"

'StrLogin = "admin_ebt"
'StrSenha = "YnVsbDIwMDc="


%>