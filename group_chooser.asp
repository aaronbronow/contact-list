<%@ Language=VBScript %>
<% Option Explicit %>

<!--#include file="adovbs.inc"-->
<!--#include file="dbConnect.asp"-->

<%
   dim objRS, strSQL

   strSQL = "SELECT group.GroupID, group.GroupName, group.GroupMemo " & _
            "FROM [group] " & _
            "ORDER BY group.GroupName;"

   Set objRS = Server.CreateObject("ADODB.Recordset")
   objRS.Open strSQL, objConn


// takes two strings
// returns true if the second string is found in the first string seperated by commas
function inGroupIDs(strSource,strTestID)

dim arrayGroupIDs, strGroupID, bolInGroupIDs
arrayGroupIDs = split(strSource,",")

For Each strGroupID In arrayGroupIDs

if strGroupID = strTestID then
  bolInGroupIDs = True
  Exit For
else
  bolInGroupIDs = False
end if

Next

inGroupIDs = bolInGroupIDs

end function

%>


<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=iso-8859-1">
<title>The Contact List - Group Chooser</title>
<link rel="stylesheet" type="text/css" href="default.css">

<script language="JavaScript">


function chooseGroup(groupIDChkBox)
{
//every time
var arrayGroupIDs = document.frmGroups.groupIDs.value.split(",");
var bolChecked;

for(i = 0; i < arrayGroupIDs.length; ++i)
{
 if(arrayGroupIDs[i] == groupIDChkBox.name)
  bolChecked = 1;
 else
  bolChecked = 0;
}


//checked
 if(bolChecked == 0)
 {

  for(i = 0; i < arrayGroupIDs.length; ++i)
  {
    if(arrayGroupIDs[i] <= groupIDChkBox.name)
     document.frmGroups.groupIDs.value += arrayGroupIDs[i];
    else
     document.frmGroups.groupIDs.value += groupIDChkBox.name;
    if(i > 0 && i < arrayGroupIDs.length)
     document.frmGroups.groupIDs.value += ",";
  }
 }
//unchecked
 else 
 {
   var newGroupIDs = arrayGroupIDs;
   var j = 0;

   for(i = 0; i < arrayGroupIDs.length; ++i)
   {
     if(!(arrayGroupIDs[i] == groupIDChkBox.name))
     {
       newGroupIDs[j] = arrayGroupIDs[i];
       ++j;
     }
   }

  document.frmGroups.groupIDs.value = "";
   for(k = 0; k < newGroupIDs.length - 1; ++k)
   {
    document.frmGroups.groupIDs.value += newGroupIDs[k]; 
    if(k > 1 && k < newGroupIDs.length -1)
     document.frmGroups.groupIDs.value += ",";
   }
 }
}
</script>

</head>
<body>

<form name="frmGroups" onSubmit="window.close(self)">
<input type="text" name="groupIDs" value="<% =request("groupIDs") %>">
<% Do While Not objRS.EOF %>

<p class="contact">
<input type="checkbox" name="<% =objRS("GroupID") %>" <% if inGroupIDs(request("groupIDs"),cstr(objRS("GroupID"))) = True then
                                                           response.write("checked ")
                                                         end if %> onClick="chooseGroup(this);"><% =objRS("GroupName") %> (ID: <% =objRS("GroupID") %>)
<br>
<% =objRS("GroupMemo") %>
</p>

<% objRS.MoveNext 
   Loop

   set objRS = Nothing
   set objConn = Nothing
%>

<p>
<input type="submit" value="Done" onClick="window.opener.document.frmEdit.strGroups.value=document.frmGroups.groupIDs.value;">
</p>

</form>

</body>
</html>
