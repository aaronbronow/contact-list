<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="adovbs.inc"-->
<!--#include file="dbConnect.asp"-->
<%
   dim objRS, strSQL

   strSQL = "SELECT contact.ContactID, contact.ContactNameLast, contact.ContactNameFirst, contact.ContactEmail, contact.ContactEmail2, contact.ContactPhone, contact.ContactCell, contact.ContactAddress, contact.ContactMemo, contact.ContactLastEdited, contact.ContactLastEditedBy " & _
            "FROM contact;"

   Set objRS = Server.CreateObject("ADODB.Recordset")
   objRS.Open strSQL, objConn, adLockOptimistic, adCmdTable
%>


<html>

<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=iso-8859-1">
<title>The Contact List</title>
<link rel="stylesheet" type="text/css" href="default.css">
</head>


<h1>ContactAdd</h1>

<%  if not request("action") = "add" then %>

<form method="post" action="contact_add.asp" name="frmAdd">
  <input type="hidden" name="action" value="add">

    <table border="0" cellpadding="0" cellspacing="0" class="frmAdd">
     <tr>
      <td><p>Last Name:</p></td> <td>First Name:</td>
     </tr>
     <tr>
      <td><input type="text" name="strNameLast" value="" title="Last Name (255 characters, not required, no HTML)">, </td>
      <td><input type="text" name="strNameFirst" value="" title="First Name (255 characters, not required)"></td>
     </tr>
     <tr>
      <td>Phone:</td> <td>Cell:</td>
     </tr>
     <tr>
      <td><input type="text" name="strPhone" value="" title="Phone (50 characters, not required)">&nbsp;</td>
      <td><input type="text" name="strCell" value="" title="Cell (50 characters, not required)"></td>
     </tr>
     <tr>
      <td>Email:</td> <td>Email2:</td>
     </tr>
     <tr>
      <td><input type="text" name="strEmail" value="" title="Email (255 characters, not required)">&nbsp;</td>
      <td><input type="text" name="strEmail2" value="" title="Email2 (255 characters, not required)"></td>
     </tr>
    </table>
    <table border="0" cellpadding="0" cellspacing="0" class="frmAdd">
     <tr>
      <td>Address:</td>
     </tr>
     <tr>
      <td><textarea rows="4" cols="20" name="strAddress" title="Address (255 characters, not required)"></textarea></td>
     </tr>
     <tr>
      <td>Memo:</td>
     </tr>
     <tr>
      <td><textarea rows="8" cols="40" name="strMemo" title="Memo (65535 characters, not required)"></textarea></td>
     </tr>
     <tr>
      <td><input type="submit" value="Add"><input type="button" value="Cancel" onClick="location.href='default.asp'"></td>
     </tr>
    </table>

</form>

<%  elseif request("action") = "add" then %>

   <p>Adding...</p>
   <ul class="updateField">
<%

   dim strNameLast, tempStr, tagStart, tagEnd, tagLength

   if inStr(request("strNameLast"),"<") > 0 AND inStr(request("strNameLast"),">") > 0 then
     strNameLast = request("strNameLast")

'     response.write("<li>tagStart = " & inStr(strNameLast,"<") & "</li>")
'     response.write("<li>tagEnd = " & inStr(strNameLast,">") & "</li>")
'     response.write("<li>tagLength = " & (inStr(strNameLast,">") - inStr(strNameLast,"<") + 1) & "</li>")

     While inStr(strNameLast,"<") > 0 AND inStr(strNameLast,">") > 0
       tagStart = inStr(strNameLast,"<")
       tagEnd = inStr(strNameLast,">")
       tagLength = tagEnd - tagStart + 1
       tempStr = replace(strNameLast,mid(strNameLast,tagStart,tagLength)," ")
       strNameLast = trim(tempStr)
     Wend
   else
     strNameLast = request("strNameLast")
   end if

   dim validName, validContact
   validName = 2
   validContact = 5
   if strNameLast = "" OR strNameLast = "Not entered" then
     validName = validName - 1
   end if
   if request("strNameFirst") = "" OR request("strNameFirst") = "Not entered" then
     validName = validname - 1
   end if
   if request("strPhone") = "" OR request("strPhone") = "Not entered" then
     validContact = validContact - 1
   end if
   if request("strCell") = "" OR request("strCell") = "Not entered" then
     validContact = validContact - 1
   end if
   if request("strEmail") = "" OR request("strEmail") = "Not entered" then
     validContact = validContact - 1
   end if
   if request("strEmail2") = "" OR request("strEmail") = "Not entered" then
     validContact = validContact - 1
   end if
   if request("strAddress") = "" OR request("strAddress") = "Not entered" then
     validContact = validContact - 1
   end if


     objRS.AddNew
   if strNameLast = "" then
     objRS("ContactNameLast") = "Not entered"
   else
     objRS("ContactNameLast") = strNameLast
     response.write("<li>Last Name</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if request("strNameFirst") = "" then
     objRS("ContactNameFirst") = "Not entered"
   else
     objRS("ContactNameFirst") = request("strNameFirst")
     response.write("<li>First Name</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if request("strPhone") = "" then
     objRS("ContactPhone") = "Not entered"
   else
     objRS("ContactPhone") = request("StrPhone")
     response.write("<li>Phone</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if request("strCell") = "" then
     objRS("ContactCell") = "Not entered"
   else
     objRS("ContactCell") = request("strCell")
     response.write("<li>Cell</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if request("strEmail") = "" then
     objRS("ContactEmail") = "Not entered"
   else
     objRS("ContactEmail") = request("strEmail")
     response.write("<li>Email</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if request("strEmail2") = "" then
     objRS("ContactEmail2") = "Not entered"
   else
     objRS("ContactEmail2") = request("strEmail2")
     response.write("<li>Email2</li>")
   end if
   if request("strAddress") = "" then
     objRS("ContactAddress") = "Not entered"
   else
     objRS("ContactAddress") = request("strAddress")
     response.write("<li>Address</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if request("strMemo") = "" then
     objRS("ContactMemo") = "Not entered"
   else
     objRS("ContactMemo") = request("strMemo")
     response.write("<li>Memo</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   objRS.Update
   objRS.MoveFirst
%>
    </ul>
    <p>... adding complete.</p>

<% if validName < 1 then %>
    <p class="warning">It appears you do not have a valid name. Please enter a first or last name.</p>
<% end if %>
<% if validContact < 1 then %>
    <p class="warning">It appears you do not have any valid methods of contact. Please enter a phone number, email, or address.</p>
<% end if %>

    <p><a href="default.asp">ContactList</a></p>


<%  else %>

   <p class="warning">An error has occured, please return to the <a href="default.asp">ContactList</a>.</p>

<% end if %>

<%
    objRS.Close
    Set objRS = Nothing

    objConn.Close
    Set objConn = Nothing
%>


</body>
</html>
