<%@ Language = "JScript" %>

<%
  function fileListToHTML( fldr )
{
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  if( fso.FolderExists(Server.MapPath(fldr)) )
  {
    var folder;
    var folderCollection;
    folder = fso.GetFolder(Server.MapPath(fldr));
    folderCollection = new Enumerator(folder.files);

    var s = '<table>\n';

    var row = 0;
    while( !folderCollection.atEnd() )
    {
      s += '  <tr>\n';

      for( i = 0; i < imgsPerRow && !folderCollection.atEnd(); i++)
      {
        var size = sizeToString( folderCollection.item().Size );
        var dateCreated = new Date(folderCollection.item().DateCreated).toLocaleString();

        var rowFlag = 'Odd';
        if((row % 2) == 0)
          rowFlag = 'Even';

        s += 
          '    <td class="imgRow' + rowFlag + '">\n' +
          '      <a href="' + fldr + '/' + folderCollection.item().Name + '">' +
          '<img src="' + fldr + '/' + folderCollection.item().Name + '"></a><br>\n' +
          '      ' + dateCreated + '<br>\n' +
          '      ' + size + '\n' +
          '    </td>\n';
        folderCollection.moveNext();
      }
      s += '  </tr>\n';
      row++;
    }

    s += '</table>';
    return(s);
  }
  else
    return "Folder (" + fldr + ") does not exist.";
}


function sizeToString( size )
{
        if( size > 1048576 ){
          size /= 1048576;
          size = size.toFixed(2);
          size += ' MB';
        }
        else if( size > 1024 ){
          size /= 1024;
          size = size.toFixed(2);
          size += ' KB';
        }
        else
          size += ' B';
        return size;
}

var imgsPerRow = 4;
if(Request('imgsperrow') > '')
  imgsPerRow = Request('imgsperrow');

var htmlImages = fileListToHTML( "avatars" );

%>


<head>
<title>RnD PhotoServer</title>

<style type="text/css">
body{
}
img{
  border: 0;
}
table{
  width: 100%;
  font-family: verdana;
  font-size: 80%;
}
td.imgRowEven{
  background-color: #d0d0d0;
  text-align: center;
}
td.imgRowOdd{
  background-color: #f0f0f0;
  text-align: center;
}

</style>
  
</head>
<body>


<form method="get">
  Images Per Row:
  <input type="text" name="imgsperrow" value="<%= imgsPerRow %>" size="2">
  <input type="submit" value="Apply">
</form>

<br>

<%= htmlImages %>

</body>
</html>
