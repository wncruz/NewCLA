<%response.write Request.ServerVariables("REMOTE_HOST")%>
<BR><BR>

<%


strCaminho = server.MapPath("..\")

RESPONSE.WRITE strCaminho


dim fs,fname
set fs=Server.CreateObject("Scripting.FileSystemObject")
set fname=fs.CreateTextFile(strCaminho & "\CartasProvedor\test.txt",true)
fname.WriteLine("TESTE DE ESCRITA")
fname.Close
set fname=nothing
set fs=nothing
%>