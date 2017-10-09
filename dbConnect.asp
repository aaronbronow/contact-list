<%
     dim objConn
     Set objConn = Server.CreateObject("ADODB.Connection")
     'objConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("contactlist.mdb"))
     objConn.Open("Driver={Microsoft Access Driver (*.mdb)};Dbq=C:\inetpub\db\contactlist.mdb;")
%>
