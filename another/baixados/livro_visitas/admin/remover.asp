<%
'Verifica se existe uma se��o
if session("login") = "" then
'se n�o existir redireciona o usu�rio para a p�gina principal
response.Redirect("index.asp")
end if

'Pegamos o id da mensagem
id = request("id")

'Abre a conex�o com o banco de dados
Set Conex = Server.CreateObject ("ADODB.Connection") 
Conex.open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="& Server.MapPath("../database/dados.mdb")

'Procuramos a mensagem com o respectivo id e removermos o registro
Set remover = Server.CreateObject("ADODB.Recordset")
sql = "Select * from mensagens where id="&id
remover.open sql, conex, 3,3
remover.delete

'Redirecionamos o usuario para a p�gina principal
response.Redirect("admin.asp")
%>