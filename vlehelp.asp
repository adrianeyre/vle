<%
' ****************************************************
' *                    vlehelp.asp                   *
' *                                                  *
' *            Coded by : Adrian Eyre                *
' *                Date : 13/09/2012                 *
' *             Version : 1.0.0                      *
' *                                                  *
' ****************************************************

response.cookies("link")="index.asp>vlehelp.asp"
response.cookies("linktext")="Home>VLE"


 dim folderlocation, currentfolder
 dim fs,fo,x
 
 folderlocation = "\subjects\"

 
 'Open the file and folder connection
 set fs=Server.CreateObject("Scripting.FileSystemObject")
 currentfolder = server.MapPath(folderlocation)
 set fo=fs.GetFolder(currentfolder)
%>

<style type="text/css">
<!--
a:link {
	color: #FFFFFF;
}
a:visited {
	color: #FFFFFF;
}
a:hover {
	color: #FFFF00;
}
a:active {
	color: #FF0000;
}
-->
</style></head>
<head>
	<link rel="stylesheet" type="text/css" href="css/main.css">
</head>
<body>
<TABLE width="100%">
  <TBODY>
    <TR>
      <TD width="100%" height="40" background="images/backdefault.png" bgcolor="#192F68"><div align="center"><span class="MenuTitle">VLE (Extended Learning) </span></div></TD>
    </TR>
    <TR>
      <TD height="40"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;</td>
          <td width="350"><p class="style6">&nbsp;</p></td>
        </tr>
        <tr>
          <td height="5"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="15">&nbsp;</td>
              <td bgcolor="#192F68"><p align="center" class="Title16 WhiteText">VLE Subjects </p>
                <p>
                  <%
 	for each x in fo.SubFolders
		if lcase(x.name) <> "feedback" then
%>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="40">&nbsp;</td>
                    <td width="245"><a href="<%response.write("vle.asp?Folder=\"+ x.name)%>" target="_self">
                      <%response.write(x.name)%>
                    </a></td>
                  </tr>
                </table>
                <%
 		'response.write("- " + x.name)
		'response.write("<br>")
		end if
 	next
%>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="70">&nbsp;</td>
                    <td width="245"><br>
                      </p></td>
                  </tr>
                </table></td>
              <td width="5">&nbsp;</td>
            </tr>
          </table></td>
          <td valign="top"><p class="style6"><strong>Logging into the VLE</strong></p>
            <p class="style2">1. Click on the links to the left</p>
            <p class="style2">2. If the login prompt appears then</p>
            <p class="style2">3 Enter your full username ie <span class="RedText">13student.a@school</span></p>
            <p class="style2">4. Enter your school password</p>
            <p class="style2">5. Check the &quot;Remember my credentials&quot; box</p>
            <p class="style2">6. Click the &quot;OK&quot; button</p></td>
        </tr>
        <tr>
          <td height="5">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td height="5" colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="15">&nbsp;</td>
              <td bgcolor="#192F68"><p align="center" class="YellowText"><br>
                  <strong>Log into your Microsoft Office 365 Account</strong></p>
                <p align="center" class="YellowText"><a href="https://login.microsoftonline.com/" target="_blank"><img src="images/Office365.png" width="559" height="131" alt="Microsoft Office 365" border=0></a></p>
                <p align="center" class="YellowText">Username = &lt;username&gt;@allsaints.notts.sch.uk<br>                
                  Password = Your network password<br>
                </p><br></td>
              <td width="15">&nbsp;</td>
            </tr>
          </table></td>
          </tr>
        </table></TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>

<%
set fs=nothing
set fo=nothing
%>
