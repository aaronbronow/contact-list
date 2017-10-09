<%@ Language=VBScript %>
<% Option Explicit %>

<%
response.cookies("contactlist").expires = NOW - 1
response.redirect("default.asp")
%>
