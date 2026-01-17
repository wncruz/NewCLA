<%
'•BULL
'	- Sistema			: CLA
'	- Arquivo			: conexao_csl.asp
'	- Responsável		: T3ENA
'	- Descrição			: Configura a conexão com o CSL
'	- Criação			: 15/07/2007

'HML/PRD

Set ConnCSL = Server.CreateObject("ADODB.Connection")

'ConnCSL.ConnectionString = "Provider=MSDAORA.1;Password=usrnwcl1;User ID=newcla;Data Source=loc1;Persist Security Info=True"
ConnCSL.ConnectionString = "Provider=MSDAORA.1;Password=usrnwcl1;User ID=newcla;Data Source=PRDSUN;Persist Security Info=True"
'DES
'ConnCSL.ConnectionString = "Provider=MSDAORA.1;Password=usrnwcl1;User ID=newcla;Data Source=csl;Persist Security Info=True"
ConnCSL.open ConnCSL

%>