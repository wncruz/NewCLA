<%@ language=vbscript %>
<%
  response.buffer = true
  response.ContentType = "application/vnd.ms-excel"
  response.AddHeader "content-disposition", "inline; filename=dynamic.xls"

  response.write "<table width=200>"
  response.write "<tr>"
  for i = 1 to 4
    response.write "<td width=40>"
    response.write i + i
    response.write "</td>"
  next
  response.write "<td width=40><b>=sum(A1:D1)</b></td>"
  response.write "</tr>"
  response.write "</table>"

  response.flush
  response.end

%>

