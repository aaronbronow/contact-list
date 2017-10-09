<%@ Language=VBScript %>
<% Option Explicit %>

<%
if Request("splashed") = "1" then
      response.cookies("contactlist").expires = date() + 30
      response.cookies("contactlist")("splashed") = "1"
end if
%>

<%
if NOT Request("splashed") = "1" AND NOT Request.cookies("contactlist")("splashed") = "1" then
'  response.redirect("splash.asp")
end if
%>

<%
Dim isWap, isPalm, isPocketPC, httpAccept, httpUserAgent
httpAccept = LCase(Request.ServerVariables("HTTP_ACCEPT"))
httpUserAgent = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
if Instr(httpAccept,"wap") then
 isWap = True
End if
if Instr(httpUserAgent,"avantgo") OR Instr(httpUserAgent,"palm") then
 isPalm = True
elseif Instr(httpUserAgent,"pocketpc") then
 isPocketPC = True
end if
%>

<!--#include file="adovbs.inc"-->
<!--#include file="dbConnect.asp"-->


<%


   if request("showAllMemos") = "1" then
      response.cookies("contactlist").expires = date() + 30
      response.cookies("contactlist")("showAllMemos") = "1"
   end if

   if request("showAllMemos") = "0" then
      response.cookies("contactlist")("showAllMemos") = "0"
   end if

   if request("showAvatars") = "0" then
      response.cookies("contactlist").expires = date() + 30
      response.cookies("contactlist")("showAvatars") = "0"
   end if

   if request("showAvatars") = "1" AND NOT isPalm = True then
      response.cookies("contactlist")("showAvatars") = "1"
   end if

   if request("showAvatars") = "1" AND isPalm = True then
      response.cookies("contactlist").expires = date() + 30
      response.cookies("contactlist")("showAvatars") = "1"
   end if

   if cint(request("groupID")) > 0 then
      response.cookies("contactlist").expires = date() + 30
      response.cookies("contactlist")("groupID") = request("groupID")
   end if

   if request("groupID") = "0" then
      response.cookies("contactlist")("groupID") = "0"
   end if

   dim bolShowAllMemos, bolShowAvatars, groupID

   if (request.cookies("contactlist")("showAllMemos") = "1" OR request("showAllMemos") = "1") then
      bolShowAllMemos = True
   else
      bolShowAllMemos = False
   end if

   if (request.cookies("contactlist")("showAvatars") = "0" OR request("showAvatars") = "0") OR (isPalm = True AND (NOT request.cookies("contactlist")("showAvatars") = "1" AND NOT request("showAvatars") = "1")) then
      bolShowAvatars = False
   else
      bolShowAvatars = True
   end if

   if (not request.cookies("contactlist")("groupID") = "0" AND NOT request.cookies("contactlist")("groupID") = "") then
      groupID = cint(request.cookies("contactlist")("groupID"))
   else
      groupID = 0
   end if


   dim objRS, strSQL

   if not request("contactID") = "" then

   strSQL = "SELECT contact.ContactID, contact.ContactNameLast, contact.ContactNameFirst, contact.ContactEmail, contact.ContactEmail2, contact.ContactPhone, contact.ContactCell, contact.ContactAddress, contact.ContactMemo, contact.ContactLastEdited, contact.ContactLastEditedBy, contact.ContactAvatar " & _
            "FROM contact " & _
            "WHERE contact.ContactID = " & cint(request("contactID")) & ";"

   elseif groupID > 0 then

   strSQL = "SELECT contact.ContactID, contact.ContactNameLast, contact.ContactNameFirst, contact.ContactEmail, contact.ContactEmail2, contact.ContactPhone, contact.ContactCell, contact.ContactAddress, contact.ContactMemo, contact.ContactLastEdited, contact.ContactLastEditedBy, contact.ContactAvatar, groupmembers.GroupID, groupmembers.GroupController " & _
            "FROM contact INNER JOIN groupmembers ON contact.ContactID = groupmembers.ContactID " & _
            "WHERE ((groupmembers.GroupID) = " & groupID & " )" & _
            "ORDER BY contact.ContactNameLast;"

   else

   strSQL = "SELECT contact.ContactID, contact.ContactNameLast, contact.ContactNameFirst, contact.ContactEmail, contact.ContactEmail2, contact.ContactPhone, contact.ContactCell, contact.ContactAddress, contact.ContactMemo, contact.ContactLastEdited, contact.ContactLastEditedBy, contact.ContactAvatar " & _
            "FROM contact " & _
            "ORDER BY contact.ContactNameLast;"
   end if

   Set objRS = Server.CreateObject("ADODB.Recordset")
   objRS.Open strSQL, objConn



function noHTML(ByVal strCheck)

   dim badStr, tempStr, tagStart, tagEnd, tagLength

   if inStr(strCheck,"<") > 0 AND inStr(strCheck,">") > 0 then
     badStr = strCheck

'     response.write("<li>tagStart = " & inStr(badStr,"<") & "</li>")
'     response.write("<li>tagEnd = " & inStr(badStr,">") & "</li>")
'     response.write("<li>tagLength = " & (inStr(badStr,">") - inStr(badStr,"<") + 1) & "</li>")

     While inStr(badStr,"<") > 0 AND inStr(badStr,">") > 0
       tagStart = inStr(badStr,"<")
       tagEnd = inStr(badStr,">")
       tagLength = tagEnd - tagStart + 1
       tempStr = replace(badStr,mid(badStr,tagStart,tagLength)," ")
       badStr = trim(tempStr)
     Wend
     strCheck = badStr
   else
     ' do nothing
   end if
   noHTML = strCheck

end function

function toPhoneNumber(ByVal strCheck)

   dim badStr, tempStr
   badStr = strCheck

   while inStr(badStr,"-") > 0 OR inStr(badStr,"(") > 0 OR inStr(badStr,")") > 0 OR inStr(badStr," ") > 0
     if inStr(badStr,"-") > 0 then
       tempStr = replace(badStr,"-","")
     end if
     if inStr(badStr,"(") > 0 then
       tempStr = replace(badStr,"(","")
     end if
     if inStr(badStr,")") > 0 then
       tempStr = replace(badStr,")","")
     end if
     if inStr(badStr," ") > 0 then
       tempStr = replace(badStr," ","")
     end if
     badStr = tempStr
   wend

   toPhoneNumber = badStr

end function

if isWap = True then

Response.ContentType = "text/vnd.wap.wml" %>

<?xml version="1.0"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<% if not request("contactID") = "" then %>

 <card id="main" title="ContactList">
  <do type="prev" label="back">
   <go href="default.asp"/>
  </do>
  <p mode="wrap">
   <b><% =noHTML(objRS("ContactNameLast")) %>, <% =noHTML(objRS("ContactNameFirst")) %></b>
  </p>
<% if not objRS("ContactPhone") = "Not entered" then %>
  <p>
   P: <a href="wtai://wp/mc;<% =toPhoneNumber(noHTML(objRS("ContactPhone"))) %>"><% =noHTML(objRS("ContactPhone")) %></a>
  </p>
<% end if %>
<% if not objRS("ContactCell") = "Not entered" then %>
  <p>
   C: <a href="wtai://wp/mc;<% =toPhoneNumber(noHTML(objRS("ContactCell"))) %>"><% =noHTML(objRS("ContactCell")) %></a>
  </p>
<% end if %>
<% if not objRS("ContactEmail") = "Not entered" then %>
  <p>
   E: <% if instr(noHTML(objRS("ContactEmail")),"@") then %>
      <a href="mailto:<% =noHTML(objRS("ContactEmail")) %>"><% =noHTML(objRS("ContactEmail")) %></a>
      <% else %>
      <% =noHTML(objRS("ContactEmail")) %>
      <% end if %>
  </p>
<% end if %>
<% if not objRS("ContactEmail2") = "Not entered" then %>
  <p>
   E2: <% if instr(noHTML(objRS("ContactEmail2")),"@") then %>
      <a href="mailto:<% =noHTML(objRS("ContactEmail2")) %>"><% =noHTML(objRS("ContactEmail2")) %></a>
      <% else %>
      <% =noHTML(objRS("ContactEmail2")) %>
      <% end if %>
  </p>
<% end if %>
<% if not objRS("ContactAddress") = "Not entered" then %>
  <p>
<% dim strWMLAddress
   strWMLAddress = noHTML(objRS("ContactAddress"))
   if inStr(strWMLAddress,VBCr) > 0 then
     strWMLAddress = replace(strWMLAddress,VBCr,"<br/>")
   end if
%>
   A: <% =strWMLAddress %>
  </p>
<% end if %>
<% dim strWMLMemo
   strWMLMemo = objRS("ContactMemo")

   if not strWMLMemo = "Not entered" then %>
  <p>
<% strWMLMemo = noHTML(strWMLMemo)
   if inStr(strWMLMemo,VBCr) > 0 then
     strWMLMemo = replace(strWMLMemo,VBCr,"<br/>")
   end if
%>
   M: <% =strWMLMemo %>
  </p>
<% end if %>
 </card>

<% else %>

 <card id="main" title="ContactList">
<% do while not objRS.EOF %>
  <p mode="nowrap">
   <a href="default.asp?contactID=<% =noHTML(objRS("ContactID")) %>"><% =noHTML(objRS("ContactNameLast")) %>, <% =noHTML(objRS("ContactNameFirst")) %></a>
  </p>
<% objRS.MoveNext
   Loop %>
 </card>

<% end if %>

</wml>

<% response.flush
   response.end %>
<% end if %>


<html>

<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=iso-8859-1">
<title>The Contact List</title>
<link rel="stylesheet" type="text/css" href="default.css" title="Default">

<script language="JavaScript">

function showAllMemo()
{
  var id = 1;

  while(typeof document.getElementById("memo" + id) != "Undefined")
  {
    showMemo(id);
    ++id;
  }    

}

function hideMemo(id)
{
  var memoID = document.getElementById("memo" + id);
  var memoBtnID = document.getElementById("memoBtn" + id);
  if(memoID.style.display == "block")
  {
    memoID.style.display = "none";
    memoBtnID.style.color = "#000000";
  }
  else
    showMemo(id);

}

function showMemo(id)
{
  var memoID = document.getElementById("memo" + id);
  var memoBtnID = document.getElementById("memoBtn" + id);
  if(typeof memoID != "Undefined")
  {
  if(memoID.style.display == "none")
  {
    memoID.style.display = "block";
    memoBtnID.style.color = "#ffffff";
  }
  else
    hideMemo(id);
  }

}

</script>

</head>

<body>


<div class="main">

<div class="banner" style="float: right; margin-right: 20px;">
<script type="text/javascript"><!--
google_ad_client = "pub-3984801897331171";
google_ad_width = 468;
google_ad_height = 60;
google_ad_format = "468x60_as";
google_ad_channel ="";
google_color_border = "ACBCBC";
google_color_bg = "9CACAC";
google_color_link = "0000FF";
google_color_url = "008000";
google_color_text = "000000";
//--></script>
<script type="text/javascript"
  src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script>
</div>

<h1><a href="./" style="text-decoration: none; color: black">ContactList</a></h1>

<p class="nav">
[<a href="contact_add.asp" title="Add a contact.">add</a>]
<% if bolShowAllMemos = True then %>
<span class="active">
[<a href="default.asp?showAllMemos=0">memos</a>] 
</span>
<% else %>
<span class="inactive">
[<a href="default.asp?showAllMemos=1">memos</a>] 
</span>
<% end if %>
<% if bolShowAvatars = False then %>
<span class="inactive">
[<a href="default.asp?showAvatars=1">avatars</a>] 
</span>
<% else %>
<span class="active">
[<a href="default.asp?showAvatars=0">avatars</a>] 
</span>
<% end if %>
</p>

<%
    Do While Not objRS.EOF 
%>

 <table border="0" cellpadding="0" cellspacing="5" class="contact">
  <tr valign="top">
   <td class="left">
    <% if not objRS("ContactAvatar") = "Not entered" AND bolShowAvatars = True then %>
      <p class="avatar"><img src="avatars/<% =objRS("ContactAvatar") %>"></p>
    <% end if %>
    
    <p class="nav">
      [<a href="contact_edit.asp?id=<% =objRS("ContactID") %>" title="Edit the contact data listed above.">edit</a>]
<% if isPalm = True then %>
<br>
<% end if %>
      [<a href="javascript:alert('Last Edited: <% =objRS("ContactLastEdited") %> (<%
                                      if dateDiff("D",objRS("ContactLastEdited"),Now) = 0 then
                                        response.write("Today") 
                                      elseif dateDiff("D",objRS("ContactLastEdited"),Now) = 1 then
                                        response.write("Yesterday")
                                      else 
                                        response.write(dateDiff("D",objRS("ContactLastEdited"),Now) & " days ago")
                                      end if %>)\nLast Edited By: <% =objRS("ContactLastEditedBy") %>');" title="View information on the contact data listed above.">info</a>]
      <span id="memoBtn<% =objRS("ContactID") %>" style="color: #000000; display: none;">[<a href="javascript:showMemo(<% =objRS("ContactID") %>);">memo</a>]</span>
    </p>

   </td>
   <td class="right">

     <% if not objRS("ContactNameLast") = "Not entered" then %>
      <span class="name" title="Last Name">
        <% =objRS("ContactNameLast") %>,
      </span>
     <% end if %>
     <% if not objRS("ContactNameFirst") = "Not entered" then %>
      <span class="name" title="First Name">
        <% =objRS("ContactNameFirst") %>
      </span>
     <% end if %>
   <% if not objRS("ContactPhone") = "Not entered" OR not objRS("ContactCell") = "Not entered" then %>
    <br>
   <% end if %>
     <% if not objRS("ContactPhone") = "Not entered" then %>
      <span class="phone" title="Phone">
        <% =objRS("ContactPhone") %>
      </span>
     <% end if %>
     <% if not objRS("ContactCell") = "Not entered" then %>
      <% if isPalm = True AND NOT objRS("ContactPhone") = "Not entered" then %>
      <br>
      <% end if %>
      <span class="phone" title="Cell">
        <% =objRS("ContactCell") %>
      </span>
     <% end if %>
   <% if not objRS("ContactEmail") = "Not entered" OR not objRS("ContactEmail2") = "Not entered" then %>
    <br>
   <% end if %>
     <% if not objRS("ContactEmail") = "Not entered" then %>
      <span class="email" title="Email">
      <% if instr(objRS("ContactEmail"),"@") then 
           response.write("<a href='mailto:" & objRS("ContactEmail")) & "'>" & objRS("ContactEmail") & "</a>"
         else
           response.write(objRS("ContactEmail"))
         end if %>
      </span>
     <% end if %>
     <% if not objRS("ContactEmail2") = "Not entered" then %>
      <span class="email" title="Email2">
      <% if isPalm = True AND NOT objRS("ContactEmail") = "Not entered" then %>
      <br>
      <% end if %>
      <% if instr(objRS("ContactEmail2"),"@") then 
           response.write("<a href='mailto:" & objRS("ContactEmail2")) & "'>" & objRS("ContactEmail2") & "</a>"
         end if %>
      </span>
     <% end if %>
   <% if not objRS("ContactAddress") = "Not entered" then %>
    <br>
      <span class="address" title="Address">
      <% response.write(replace(noHTML(objRS("ContactAddress")),VBCr,"<br>")) %>
      </span>
    <br>
      <span class="note" title="Google Maps">[<a href="http://maps.google.com/maps?q=<%= Escape(replace(noHTML(objRS("ContactAddress")),VBNewline,"+")) %>">Map</a>]</span>
   <% end if %>

    <% dim strMemo
        strMemo = objRS("ContactMemo")
       if not strMemo = "Not entered" AND bolShowAllMemos = True then %>
      <span class="memo" id="memo<% =objRS("ContactID") %>" title="Memo" style="display: block;" onLoad="if(getCookie('memos') == 'true') showMemo(<% =objRS("ContactID") %>;">
    <br>
      <% response.write(replace(strMemo,VBCr,"<br>")) %>
      </span>
    <% else %>
      <span class="memo" id="memo<% =objRS("ContactID") %>" title="Memo" style="display: none;"></span>
    <% end if %>
   </td>
  </tr>
 </table>


<%
    objRS.MoveNext
    Loop

    objRS.Close
    Set objRS = Nothing

    objConn.Close
    Set objConn = Nothing
%>

</div>

</body>
</html>







