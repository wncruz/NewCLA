<%
set ConSGA = Server.CreateObject("ADODB.Command")
'' Desenvolvimento
''ConSGA.ActiveConnection = "Provider=MSDAORA.1;Password=newcla21;User ID=USER_CLA;Data Source=SGAGSID1;Persist Security Info=True"
'' Producao
ConSGA.ActiveConnection = "Provider=MSDAORA.1;Password=usr_cl3;User ID=USER_CLA;Data Source=SGAGSI1;Persist Security Info=True"
%>