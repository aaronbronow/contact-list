<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=iso-8859-1">
<title>The Contact List - Avatar Chooser</title>
<link rel="stylesheet" type="text/css" href="default.css">

<script language="JavaScript">

// getHeight takes an image's filename, checks and returns the height
// Currently used to resize images which are larger than average (100px)

function getHeight(img)
{

  myImage = new Image() 
  myImage.src = "avatars/" + img

  return myImage.height
}

// setHeight takes the image pointer and filename and resizes the image if necessary
// objImg = pointer to img tag
// strImg = filename string

function setHeight(objImg,strImg)
{
  if(getHeight(strImg) > 100)
  {
    objImg.height = 100;
    setAlt(objImg);
    setTitle(objImg);
  }
}

// setAlt changes the alt tag of objImg

function setAlt(objImg)
{
  objImg.alt = objImg.alt + ' (resized)';
}

// setTitle changes the title tag of objImg

function setTitle(objImg)
{
  objImg.title = objImg.title + ' (resized)';
}

</script>

</head>
<body>
<%
'''objImgFile becomes FileSystemObject
  Set objImgFile = Server.CreateObject("Scripting.FileSystemObject")
'''Folder becomes the folder "avatars"
  Set Folder = objImgFile.GetFolder(Server.MapPath("avatars"))

'''loop through all files in Folder and make the image accessable from objImg
  FOR EACH objImg in Folder.Files
%> 
<a href="javascript:window.close(self);" onClick="window.opener.document.frmEdit.strAvatar.value='<% =objImg.Name %>'">
<img src="avatars/<% =objImg.Name %>" border="0" alt="<% =objImg.Name %>" title="<% =objImg.Name %>" onLoad="setHeight(this,'<% =objImg.Name %>')"></a>
<%
  NEXT

'''cleanup
  Set objImgFile = Nothing
  Set Folder = Nothing
%> 
</body>
</html>
