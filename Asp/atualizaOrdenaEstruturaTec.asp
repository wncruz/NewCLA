
<!--#include file="../inc/data.asp"-->
<%
Dim idAntigo, idNovo, iDirecao, sql, ichave, ichaveNovo
idAntigo = Request.Form("idAntigo")
idNovo = Request.Form("idNovo")
iDirecao = Request.Form("direcao") 
ichave= Request.Form("chave") 
ichaveNovo= Request.Form("chaveNovo") 

' Aqui você pode adicionar a lógica para atualizar os dados no banco de dados ou em outra estrutura

if iDirecao="acima" then
       sql ="update [dbo].[cla_estrutura_tecnologiaFacilidade] set [ordenacao]=" & idNovo & " where [estrutura_tec_fac_id]=" & ichave	& " and ordenacao=" & idAntigo	 		
       db.Execute sql

       sql ="update [dbo].[cla_estrutura_tecnologiaFacilidade] set [ordenacao]=" & idAntigo & " where [estrutura_tec_fac_id]=" & ichaveNovo & " and ordenacao=" & idNovo	 		
       db.Execute sql

else
       sql ="update [dbo].[cla_estrutura_tecnologiaFacilidade] set [ordenacao]=" & idNovo & " where [estrutura_tec_fac_id]=" & ichave	& " and ordenacao=" & idAntigo	 		
        db.Execute sql

       sql ="update [dbo].[cla_estrutura_tecnologiaFacilidade] set [ordenacao]=" & idAntigo & " where [estrutura_tec_fac_id]=" & ichaveNovo & " and ordenacao=" & idNovo	 		
       db.Execute sql
end if 




		 


%>