<%
'Encerramos a se��o criada pelo usu�rio
Session.Abandon()
'redirecionamos ele para a p�gina principal
response.Redirect("index.asp")
%>