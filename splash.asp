<html>

<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=iso-8859-1">
<meta http-equiv="refresh" content="5; url=default.asp?splashed=1">
<title>The Contact List</title>
<link rel="stylesheet" type="text/css" href="default.css" title="Default">

<script type="text/javascript" src="fading.js"></script>

<script type="text/javascript">

function fade()
{
setFromColor(156, 172, 172);
setTimeout("fadeIn(document.getElementById('coming'));", 0);

setTimeout("setFromColor(156, 172, 172);", 2000);
setTimeout("fadeOut(document.getElementById('contactlist'));", 2000 );
}
</script>

</head>

<body onload="fade();">

<table border="0" width="100%" height="100%">
 <tr>
  <td style="text-align: center;">
   <h2 style="color: #9cacac" id="coming">It's coming...</h2>
   <h1 style="color: #9cacac" id="contactlist">ContactList 2.0</h1>
  </td>
 </tr>
 <tr style="text-align: right;">
  <td valign="bottom">
   <a href="default.asp?splashed=1">Skip</a>
  </td>
 </tr>
</table>

</body>
</html>
