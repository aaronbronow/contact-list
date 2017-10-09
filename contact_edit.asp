<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="adovbs.inc"-->
<!--#include file="dbConnect.asp"-->
<%
   dim objRS, strSQL

''''If "delete" has been passed in and "id" is not null then make strSQL delete the id

   if request("action") = "delete" AND not request("id") = "" then

   strSQL = "DELETE FROM contact " & _
            "WHERE contact.ContactID = " & cint(request("id")) & ";"

''''Otherwise, load as per usual. If the "id" has not been passed in 
''''the recordset will end and an error message will be displayed in body

   else

   strSQL = "SELECT contact.ContactID, contact.ContactNameLast, contact.ContactNameFirst, contact.ContactEmail, contact.ContactEmail2, contact.ContactPhone, contact.ContactCell, contact.ContactAddress, contact.ContactMemo, contact.ContactLastEdited, contact.ContactLastEditedBy, contact.ContactAvatar " & _
            "FROM contact " & _ 
            "WHERE contact.ContactID = " & cint(request("id")) & ";"
'INNER JOIN groupmembers ON contact.ContactID = groupmembers.ContactID
', groupmembers.GroupID, groupmembers.GroupController
''''End if
   end if
  ' response.write(strSQL)
   Set objRS = Server.CreateObject("ADODB.Recordset")
   objRS.Open strSQL, objConn, adLockOptimistic, adCmdTable
%>
<html>

<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=iso-8859-1">
<title>The Contact List</title>
<link rel="stylesheet" type="text/css" href="default.css">

<script language="JavaScript">

//conf() takes an id string from the Delete button on the edit form and confirms
//deletion. Then changes location to the delete page.

function conf(id)
{
  if(confirm("Are you sure you want to delete this contact?"))
   {
    location.href="contact_edit.asp?id=" + id + "&action=delete";
   }
}

//openChooser() opens the avatar chooser

function openChooser(chooserType,chooserData)
{
  if (chooserType == "avatar")
    window.open('avatar_chooser.asp','chooser','width=540,height=300,menubar=no,resizable=yes,location=no,scrollbars=yes');

  if (chooserType == "group")
    alert("This feature is not implemented yet.");
//    window.open('group_chooser.asp?groupIDs=' + chooserData,'chooser','width=540,height=300,menubar=no,resizable=yes,location=no,scrollbars=yes');
}

</script>

</head>
<body>

<h1>ContactEdit</h1>

<%
'''If "delete" has NOT been passed in, nor has "update" been passed in then proceed with loading the page normally
  if not request("action") = "delete" AND not request("action") = "update" then

'''''If the recordset has not reached the end of file...
    if not objRS.EOF then

  'dim strGroupIDs
  'strGroupIDs = cstr(objRS("GroupID"))

'Do While Not objRS.EOF
' if NOT instr(strGroupIDs,cstr(objRS("GroupID"))) > 0 then
'   strGroupIDs = cstr(strGroupIDs) + "," + cstr(objRS("GroupID"))
' end if
' objRS.MoveNext
'Loop
'objRS.MoveFirst

%>

<form method="post" action="contact_edit.asp" name="frmEdit">
  <input type="hidden" name="action" value="update">
  <input type="hidden" name="id" value="<% =request("id") %>">

    <table border="0" cellpadding="0" cellspacing="0" class="frmEdit">
     <tr>
      <td><p>Last Name:</p></td> <td>First Name:</td>
     </tr>
     <tr>
      <td><input type="text" name="strNameLast" value="<% =objRS("ContactNameLast") %>" title="Last Name (255 characters, not required, no HTML)">, </td>
      <td><input type="text" name="strNameFirst" value="<% =objRS("ContactNameFirst") %>" title="First Name (255 characters, not required)"></td>
     </tr>
     <tr>
      <td>Phone:</td> <td>Cell:</td>
     </tr>
     <tr>
      <td><input type="text" name="strPhone" value="<% =objRS("ContactPhone") %>" title="Phone (50 characters, not required)">&nbsp;</td>
      <td><input type="text" name="strCell" value="<% =objRS("ContactCell") %>" title="Cell (50 characters, not required)"></td>
     </tr>
     <tr>
      <td>Email:</td> <td>Email2:</td>
     </tr>
     <tr>
      <td><input type="text" name="strEmail" value="<% =objRS("ContactEmail") %>" title="Email (255 characters, not required)">&nbsp;</td>
      <td><input type="text" name="strEmail2" value="<% =objRS("ContactEmail2") %>" title="Email2 (255 characters, not required)"></td>
     </tr>
    </table>
    <table border="0" cellpadding="0" cellspacing="0" class="frmEdit">
     <tr>
      <td>Address:</td>
     </tr>
     <tr>
      <td><textarea rows="4" cols="20" name="strAddress" title="Address (255 characters, not required)"><% =objRS("ContactAddress") %></textarea></td>
     </tr>
     <tr>
      <td>Memo:</td>
     </tr>
     <tr>
      <td><textarea rows="8" cols="40" name="strMemo" title="Memo (65535 characters, not required)"><% =objRS("ContactMemo") %></textarea></td>
     </tr>
    </table>
    <table border="0" cellpadding="0" cellspacing="0" class="frmEdit">
     <tr>
      <td>Avatar:</td> <td>Groups:</td>
     </tr>
     <tr>
      <td><input type="text" name="strAvatar" value="<% =objRS("ContactAvatar") %>" title="Your avatar (255 characters, not required) [DoubleClick to activate chooser]" onDblClick="openChooser('avatar','null');"></td>
      <td><input type="text" name="strGroups" value="" title="Your group memberships (seperated by comma, not required) [DoubleClick to activate chooser]" onDblClick="openChooser('group',this.value);" readonly></td>
     </tr>
     <tr>
      <td><input type="submit" value="Edit"><input type="button" value="Cancel" onClick="location.href='default.asp'"><input type="button" value="Delete" onClick="conf(<% =request("id") %>)"></td>
     </tr>
    </table>

</form>

<%
'''''If recordset hits EOF before a chance to load the page (the ContactID was not found)...
    else 
%>

   <p>The record was not found in the database.</p>
   <p><a href="default.asp">ContactList</a></p>

<%
'''''End if
    end if



'''If "update" has been passed in...
  elseif request("action") = "update" then
%>

   <p>Updating...</p>
   <ul class="updateField">
<%

'''' noHTML() takes a string by value and returns the string sans anything between "<" and ">"
function noHTML(ByVal strCheck)

   dim badStr, tempStr, tagStart, tagEnd, tagLength

''''If both a < and > appear within the string then proceed with function
   if inStr(strCheck,"<") > 0 AND inStr(strCheck,">") > 0 then

     badStr = strCheck

''''''Error checking
'     response.write("<li>tagStart = " & inStr(badStr,"<") & "</li>")
'     response.write("<li>tagEnd = " & inStr(badStr,">") & "</li>")
'     response.write("<li>tagLength = " & (inStr(badStr,">") - inStr(badStr,"<") + 1) & "</li>")

     'As long as both < and > appear within the string, keep looping
     While inStr(badStr,"<") > 0 AND inStr(badStr,">") > 0
       tagStart = inStr(badStr,"<")
       tagEnd = inStr(badStr,">")
       tagLength = tagEnd - tagStart + 1
       'Replace all the code with whitespace
       tempStr = replace(badStr,mid(badStr,tagStart,tagLength)," ")
       'Then trim off whitespace
       'Any code between uncoded text will appear as one space (per tag)
       'EX: <foo>bar</foo><bar>foo</bar> will return: "bar  foo"
       badStr = trim(tempStr)
     Wend
     strCheck = badStr

''''If both < and > DO NOT appear in the string then...
   else
     ' do nothing

''''End if
   end if

''''Return the string
   noHTML = strCheck

'''End noHTML()
end function



''''This bunch of code checks for valid data
''''There are first and last name fields, if both are blank the user is prompt to enter one
''''There are 5 contact fields (2 phone, 2 email, 1 address) if all are blank the user is prompted
   dim validName, validContact, strNameLast
   validName = 2
   validContact = 5
   strNameLast = noHTML(request("strNameLast"))

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



''''This is the recordset updating code. Each value is checked to see if it already exists the same
''''way in the database, if yes do nothing, if no set it to the new value, update the ContactLastEdited
''''record and ContactLastEditedBy record, and write the field which was changed to the page.
   if not objRS("ContactNameLast") = strNameLast then
     if strNameLast = "" then
       objRS("ContactNameLast") = "Not entered"
     else
       objRS("ContactNameLast") = strNameLast
     end if
     response.write("<li>Last Name</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if not objRS("ContactNameFirst") = request("strNameFirst") then
     if request("strNameFirst") = "" then
       objRS("ContactNameFirst") = "Not entered"
     else
       objRS("ContactNameFirst") = request("strNameFirst")
     end if
     response.write("<li>First Name</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if not objRS("ContactPhone") = request("strPhone") then
     if request("strPhone") = "" then
       objRS("ContactPhone") = "Not entered"
     else
       objRS("ContactPhone") = request("StrPhone")
     end if
     response.write("<li>Phone</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if not objRS("ContactCell") = request("strCell") then
     if request("strCell") = "" then
       objRS("ContactCell") = "Not entered"
     else
       objRS("ContactCell") = request("strCell")
     end if
     response.write("<li>Cell</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if not objRS("ContactEmail") = request("strEmail") then
     if request("strEmail") = "" then
       objRS("ContactEmail") = "Not entered"
     else
       objRS("ContactEmail") = request("strEmail")
     end if
     response.write("<li>Email</li>")
   end if
   if not objRS("ContactEmail2") = request("strEmail2") then
     if request("strEmail2") = "" then
       objRS("ContactEmail2") = "Not entered"
     else
       objRS("ContactEmail2") = request("strEmail2")
     end if
     response.write("<li>Email2</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if not objRS("ContactAddress") = request("strAddress") then
     if request("strAddress") = "" then
       objRS("ContactAddress") = "Not entered"
     else
       objRS("ContactAddress") = request("strAddress")
     end if
     response.write("<li>Address</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if not objRS("ContactMemo") = request("strMemo") then
     if request("strMemo") = "" then
       objRS("ContactMemo") = "Not entered"
     else
       objRS("ContactMemo") = request("strMemo")
     end if
     response.write("<li>Memo</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   if not objRS("ContactAvatar") = request("strAvatar") then
     if request("strAvatar") = "" then
       objRS("ContactAvatar") = "Not entered"
     else
       objRS("ContactAvatar") = request("strAvatar")
     end if
     response.write("<li>Avatar</li>")
     objRS("ContactLastEdited") = Now
     objRS("ContactLastEditedBy") = Request.ServerVariables("REMOTE_ADDR")
   end if
   objRS.Update

%>
    </ul>
    <p>... updating complete.</p>

<%
''''If there are fewer than 1 valid name
   if validName < 1 then %>
    <p class="warning">It appears you do not have a valid name. Please enter a first or last name.</p>
<%
   end if

''''If there are fewer than 1 valid contact method
   if validContact < 1 then
%>
    <p class="warning">It appears you do not have any valid methods of contact. Please enter a phone number, email, or address.</p>
<%
   end if
%>

    <p><a href="default.asp">ContactList</a></p>
<%
    objRS.Close
    Set objRS = Nothing

    objConn.Close
    Set objConn = Nothing

'''''If "delete" is passed in then print the following (the record has already been deleted by SQL)
    elseif request("action") = "delete" then 
%>

<%'''''If the id is null there has been an error and the user must return to main page.%>
   <% if request("id") = "" then %>
     <p>The contact ID was missing. Please return to the <a href="default.asp">ContactList</a>.</p>
   <% else %>
     <p>Contact <% =request("id") %> successfully deleted.</p>
     <p><a href="default.asp">ContactList</a></p>
   <% end if %>

<%
'''''If none of the aforementioned conditions are met we assume an id was not given or not found in the database.
    else %>

     <p>The requested contact ID was not found in the database.</p>
     <p><a href="default.asp">ContactList</a></p>

<% end if %>


</body>
</html>
