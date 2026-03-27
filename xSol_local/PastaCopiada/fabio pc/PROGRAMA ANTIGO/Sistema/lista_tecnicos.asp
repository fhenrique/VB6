<html>
<head>
<title>X.Sol Arte em Aquecedores - Lista de Técnicos</title>
<link rel="STYLESHEET" type="text/css" href="css/forAll.css">

</head>
<%
pesq_serv=Request.Form("busca")


filepath=Server.MapPath("dados.mdb")
Set con = Server.CreateObject("ADODB.CONNECTION")

con.open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & filepath
set rs=Server.CreateObject("ADODB.Recordset")

sql = "SELECT Codigo, Nome, Razao, Celular FROM Tecnico ORDER BY Nome; "

rs.Open sql,con
On Error Resume Next

%>

<table width="100%" border="1" cellspacing="0" cellpadding="0">
<tr>
	<td class="forTexts" bgcolor="F5C443" width="40%"><div align="center"> EMPRESA 
        / RAZ&Atilde;O SOCIAL</div></td>
	<td class="forTexts" bgcolor="F5C443" width="40%"><div align="center">NOME DO T&Eacute;CNICO</div></td>

	<td class="forTexts" bgcolor="F5C443" width="20%"><div align="center">CELULAR</div></td>
</tr>

   <%if rs.EOF then
       Response.Write ("Não existe registro")
     Else
       rs.MoveFirst
     Do While Not rs.EOF
   %> 
  <tr>
	    <td class="forTexts" width="30%"><a href="visualizar_cliente.asp?Codigo=<%Response.Write(rs.Fields("Codigo"))%>"><%Response.Write(Ucase(rs.Fields("Razao")))%></a></td>
	<td class="forTexts" width="40%"><%Response.Write(Ucase(rs.Fields("Nome")))%></td>								
	<td class="forTexts" width="20%"><%Response.Write(rs.Fields("Celular"))%></td>					
  </tr>  
<%  
  rs.MoveNext  
Loop

End if
rs.Close
Set rs = Nothing

cn.Close
Set cn = Nothing
%>  
</table>
</html>