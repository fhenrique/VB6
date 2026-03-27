<html>
<head>
<title>X.Sol Arte em Aquecedores - Lista de Clientes</title>
<link rel="STYLESHEET" type="text/css" href="css/forAll.css">

</head>
<%
pesq_serv=Request.Form("busca")


filepath=Server.MapPath("dados.mdb")
Set con = Server.CreateObject("ADODB.CONNECTION")

con.open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & filepath
set rs=Server.CreateObject("ADODB.Recordset")

Nome_Pesq = Request.Form("Envio_Nome")

sql = "SELECT * FROM Clientes WHERE Nome_Cli LIKE '%"& Nome_Pesq &"%' ORDER BY Nome_Cli; "

rs.Open sql,con
On Error Resume Next

%>

<table width="100%" border="1" cellspacing="0" cellpadding="0">
<tr>
	<td class="forTexts" bgcolor="F5C443" width="30%"><div align="center">NOME / RAZ&Atilde;O SOCIAL</div></td>
	<td class="forTexts" bgcolor="F5C443" width="40%"><div align="center">ENDERE&Ccedil;O</div></td>
	<td class="forTexts" bgcolor="F5C443" width="15%"><div align="center">FONE</div></td>
	<td class="forTexts" bgcolor="F5C443" width="15%"><div align="center">CELULAR</div></td>
</tr>

   <%if rs.EOF then
       Response.Write ("Não existe registro")
     Else
       rs.MoveFirst
     Do While Not rs.EOF
   %> 
  <tr>
    <td class="forTexts" width="30%"><a href="visualizar_cliente.asp?Codigo=<%Response.Write(rs.Fields("Codigo"))%>"><%Response.Write(Ucase(rs.Fields("Nome_Cli")))%></a></td>
	<td class="forTexts" width="40%"><%Response.Write(Ucase(rs.Fields("End")))%></td>
    <td class="forTexts" width="15%"><%Response.Write(rs.Fields("Fone_Res"))%></td>								
	<td class="forTexts" width="15%"><%Response.Write(rs.Fields("Fone_Celular"))%></td>					
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