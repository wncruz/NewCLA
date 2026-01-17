<!-- #include file="../inc/data.asp" -->
<%


if strLoginRede = "PRSSILV" OR strLoginRede = "EDAR" OR strLoginRede = "FEMAG" OR strLoginRede = "EDLUIZ" OR strLoginRede = "RAMON" THEN
  acesso = "ok"
else
  response.write "<center><h3>Solicite seu acesso ao grupo CLA</h3></center>"
  response.end
end if

For Each item in Session.Contents
'Response.Write item & " = " & Session.Contents(item) & "<BR>"
Next
%>
<br>
<br>
<link rel="stylesheet" type="text/css" href="../css/cla.css">
<html>
<title>CLA - Log de acessos</title>

<style>
A:visited
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 11px;
    COLOR: #000000;
    FONT-FAMILY: tahoma,tahoma,sans-serif;
    POSITION: relative;
    TEXT-DECORATION: none
}
A:link
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 11px;
    COLOR: #000000;
    FONT-FAMILY: tahoma,tahoma,sans-serif;
    POSITION: relative;
    TEXT-DECORATION: none
}
A:hover
{
    FONT-SIZE: 11px;
    COLOR: #2194d1;
    FONT-FAMILY: tahoma,tahoma,sans-serif;
    TEXT-DECORATION: none
}
A:active
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 11px;
    COLOR: #186997;
    FONT-FAMILY: tahoma,tahoma,sans-serif;
    POSITION: relative;
    TEXT-DECORATION: none
}
BODY
{
    SCROLLBAR-FACE-COLOR: #dadada;
    SCROLLBAR-HIGHLIGHT-COLOR: #f5f5f5;
    SCROLLBAR-SHADOW-COLOR: #cacaca;
    SCROLLBAR-3DLIGHT-COLOR: #f5f5f5;
    SCROLLBAR-ARROW-COLOR: #8b8b8b;
    SCROLLBAR-TRACK-COLOR: #e3e3e3;
    SCROLLBAR-DARKSHADOW-COLOR: #acacac
}
</style>


<center>
<h2>CLA - Usuários Simultâneos</h2>
<hr width=70%>
<%response.write "Sessões ativas: " & Application("UsuariosAtuais")%>
