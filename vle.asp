<%
' ****************************************************
' *                    vle.asp                       *
' *                                                  *
' *            Coded by : Adrian Eyre                *
' *                Date : 27/06/2012                 *
' *             Version : 1.0.0                      *
' *                                                  *
' ****************************************************

response.cookies("link")="index.asp>vle.asp"
response.cookies("linktext")="Home>VLE"
%><head>
	<link rel="stylesheet" type="text/css" href="css/main.css">
</head>
  <%
  
 Dim FullFileName(1000), DisplayFileName(1000), Extension(1000), iconname(1000)
 Dim Amount, lastlocation, feedback
 dim folderlocation, currentfolder, previousfolder, SubjectName
 dim fs,fo,x
 Dim arrUser, Username
 arrUser = Split(Request.ServerVariables("LOGON_USER"), "\")
 Username= arrUser(1)
 'Username = Request.ServerVariables("LOGON_USER")
 
 StaffUser = false
 if (asc(left(Username,1)) > 47 and asc(left(Username,1)) < 58) or lcase(username) = "parent" then
 	StaffUser = false
 else
 	StaffUser = true
 end if
 
 %>
 <!-- #include file = "config.asp" -->	
 <%
 
 folderlocation = request.querystring("folder")
 strMessage = Request.QueryString ("msg")
 
 
 if folderlocation = "" then
 	folderlocation = "\subjects"
 else
	folderlocation = "\subjects" + folderlocation
 end if
 
 'Open the file and folder connection
 set fs=Server.CreateObject("Scripting.FileSystemObject")
 currentfolder = server.MapPath(folderlocation)
 set fo=fs.GetFolder(currentfolder)

  if lcase(mid(folderlocation,1, 18)) = "\subjects\feedback" then
  	folderlocation = "\subjects\feedback"
	folderlocation = folderlocation + "\" + Username
	currentfolder = currentfolder + "\" + Username
	
	set fs2=Server.CreateObject("Scripting.FileSystemObject")
 	if fs2.FolderExists(currentfolder)=false then
		set f=fs2.CreateFolder(currentfolder)
 	end if
 	set fs2=nothing
	set fo=nothing
	set fo=fs.GetFolder(currentfolder)
	lastlocation =  request.Cookies("lastlocation")
	feedback = true
 end if

 TempSubjectName = mid(folderlocation,11,len(folderlocation))
' response.write(TempSubjectName)
 SubjectName = ""
 num = 0
 
 'Get the subject name from the breadcumbs feed
 for n = 1 to len(TempSubjectName)
 	if num = 0 and mid(TempSubjectName,n,1) <> "\" then
		SubjectName = SubjectName + mid(TempSubjectName,n,1)
	else
		num = 1
	end if
 next
 
 'Display the subject name
 if SubjectName = "" then 
 	DisplaySubject = "All Saints' VLE Subjects"
 else
 	DisplaySubject = SubjectName
 end if
 %>
 
<table width="715" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="40" background="images/backdefault.png" bgcolor="#192F68"><p align='center' class='MenuTitle'><%response.write(DisplaySubject)%></p></td>
  </tr>
</table>
 <%
 
 linkname = ""
 displaylink = ""
 subjectlink = ""
 u = 0
 
 'Display the breadcrumbs feed
  response.write(">")
 subjectlink = subjectlink + "\" + linkname
 response.write("<a href='vle.asp?folder=' target='_self' > Subjects </a>")
 for n = 11 to len(folderlocation)
	if mid(folderlocation, n, 1) = "\" then
		response.write(">")
		u = u + 1
		if u = 1 then subjectlink = subjectlink + linkname
		if u <> 1 then subjectlink = subjectlink + "\" + linkname
		response.write("<a href=vle.asp?folder=" & subjectlink & " target='_self' >" & displaylink & "</a>")
		linkname = ""
		displaylink = ""
	else
		if mid(folderlocation, n, 1) = " " then
			linkname = linkname + "%20"
			displaylink = displaylink + " "
		else
			linkname = linkname + mid(folderlocation, n, 1)
			displaylink = displaylink + mid(folderlocation, n, 1)
		end if 
	end if
 next
 'Add the last breadcrumbs feed
 response.write(">")
 if u > 0 then
	 subjectlink = subjectlink + "\" + linkname
 else
 	subjectlink = subjectlink + linkname
 end if
 response.write("<a href=vle.asp?folder=" & subjectlink & " target='_self' >" & displaylink & "</a>")
 
 'Replace spaces with %20 in the internet links
 filename = ""
 for n = 10 to len(folderlocation)
 	if mid(folderlocation, n, 1) = " " then
		filename = filename + "%20"
	else
		filename = filename + mid(folderlocation, n, 1)
	end if
 next
 folderlocation = filename
 response.cookies("lastlocation")=folderlocation
 filename = ""
 num = 0
 %>
 <table width="700" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="40"><img src="../images/icons/feedback.jpg" alt="Feedback" width="28" height="28" /></td>
    <td width="650"><a href="vle.asp?folder=\Feedback">Feedback</a></td>
  </tr>
 </table>
 <%
 
 
 'Get the last folder
 for n = len(folderlocation) to 1 step -1
 	if num = 0 and mid(folderlocation,n,1) = "\" then num = n
 next
 for n = 1 to (num - 1)
 	filename = filename + mid(folderlocation, n, 1)
 next
 previousfolder = filename
 
 if feedback = true then previousfolder = lastlocation 'mid(lastlocation,10,len(lastlocation))
 
 if SubjectName <> "" then
%>
<table width="700" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="40"><% response.write("<a href=vle.asp?folder=" & previousfolder & "><img src='images/icons/backfolder.png' alt='Previous Folder' width='30' height='30' border='0' /></a>") %></td>
    <td width="650"><% response.write("<a href=vle.asp?folder=" & previousfolder & "> BACK </a><br />") %></td>
  </tr>
</table>
<%
end if

'Display the folders
for each x in fo.SubFolders
   filename = ""
   for n = 1 to len(x.name)
   		if mid(x.name, n, 1) = " " then
			filename = filename + "%20"
		else
			filename = filename + mid(x.name, n, 1)
		end if
   next
   if left(x.name,1) <> "@" or StaffUser = true then
   	if lcase(filename) <> "feedback" then
%>   
   <table width="700" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="40"> <% response.write("<a href=vle.asp?folder=" & folderlocation & "\" & filename & "><img src='images/icons/folder.png' alt='Folder' width='30' height='30' border='0' /></a>") %></td>
   <td width="650"> <% Response.write("<a href=vle.asp?folder=" & folderlocation & "\" & filename & ">" & x.Name & "</a><br />") %></td>
<%
	 end if
	end if
next
%>
   </tr>
</table>

<%
DisplayUpload = false

for each x in fo.files
	Amount = Amount + 1
	FullFileName(Amount) = ""
	DisplayFileName(Amount) = ""
	Extension(Amount) = ""
	
	FullFileName(Amount) = x.name
	'response.write(FullFileName(Amount))
	num = 0
	for a = 1 to len(x.name)
		if mid(x.name,a,1) = "." then num = a
	next
	for a = num + 1 to len(x.name)
		Extension(Amount) = Extension(Amount) + mid(x.name,a,1)
	next
	for a = 1 to num - 1
		DisplayFileName(Amount) = DisplayFileName(Amount) + mid(x.name,a,1)
	next
	select case lcase(Extension(Amount))
		case "jpg","gif","tif","png":iconname(Amount) = "image"
		case "bmp":iconname(Amount) = "image.jpg"
		case "doc","docx","docm","txt","rtf","xps": iconname(Amount) = "word.jpg"
		case "xlsx", "xls": iconname(Amount) = "excel.jpg"
		case "pptx", "ppt", "pps", "pot", "pptm": iconname(Amount) = "powerpoint.jpg"
		case "pub": iconname(Amount) = "publisher.jpg"
		case "mdb": iconname(Amount) = "access.jpg"
		case "pdf": iconname(Amount) = "pdf.png"
		case "zip": iconname(Amount) = "zip.png"
		case "notebook": iconname(Amount) = "notebook.png"
		case "url", "lnk", "mht": iconname(Amount) = "url.png"
		case "wmv", "mp4", "mov", "avi", "mpg": iconname(Amount) = "video.png"
		case "mp3", "wma", "wav": iconname(Amount) = "audio.jpg"
		case "txt": iconname(Amount) = "text"
		case "htm","html": iconname(Amount) = "webpage.png"
		case "swf", "flv": iconname(Amount) = "flash.jpg"
		case else: iconname(Amount) = "blank.jpg"
	end select

	reset = false
	if lcase(FullFileName(Amount)) = "upload.txt" Then
		DisplayUpload = true
		reset = true
	end if
	if lcase(left(FullFileName(Amount),8)) = "shortcut" Then
		'response.write(currentfolder + "\" + FullFileName(Amount))

		Set wfile = fs.OpenTextFile(currentfolder + "\" + FullFileName(Amount))
			do while not wfile.AtEndOfStream
				shortcutname = wfile.Readline
				shortcutlink = wfile.Readline
				'response.write(filetext)
				'response.write("<br>")
			loop
		wfile.close
		%>
		<table width="700" border="0" cellspacing="0" cellpadding="0">
 		  <tr>
    		<td width="40"><a href="<%response.write(shortcutlink)%>" target="_self"><img src="../images/icons/shortcut.png" alt="Shortcut" width="32" height="32" border="0" /></a></td>
    		<td width="650"><a href="<%response.write(shortcutlink)%>" target="_self"><%=shortcutname%></a></td>
  		  </tr>
 		</table>
 	<%
		reset = true
	end if
	if lcase(FullFileName(Amount)) = "counter.txt" Then
		set FSO = CreateObject("scripting.FileSystemObject")

		if FSO.FileExists("D:\wwwroot\subjects\" & DisplaySubject & "\counter.csv") then
			set csvfile = FSO.OpenTextFile("D:\wwwroot\subjects\" & DisplaySubject & "\counter.csv",8, true, -1)
		else
			set csvfile = FSO.CreateTextFile("D:\wwwroot\subjects\" & DisplaySubject & "\counter.csv",true, true)
		end if
		if StaffUser = false then csvfile.writeLine(folderlocation & vbTab & date & vbTab & time & vbTab  & Username)
		csvfile.close
		reset = true
	end if
	if lcase(FullFileName(Amount)) = "counter.csv" Then
		reset = true
	end if
	if left(FullFileName(Amount),1) = "@" and StaffUser = false then
		reset = true
	end if
	if reset = true then
		FullFileName(Amount) = ""
		DisplayFileName(Amount) = ""
		Extension(Amount) = ""
		Amount = Amount - 1
	end if
	if Extension(Amount) = "txt" or Extension(Amount) = "db" or left(DisplayFileName(Amount),1) = "~" then 
		'response.write("here")
		FullFileName(Amount) = ""
		DisplayFileName(Amount) = ""
		Extension(Amount) = ""
		Amount = Amount - 1
	end if
next

if DisplayUpload = true then
	if strMessage <> "" then
	%>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
  		<tr>
	    <td width="40">&nbsp;</td>
    	<td>
			<table width="635" border="0" cellpadding="2" cellspacing="0" background="images/backcaritas.png">
			<tr>
				<td class="MiniTitle"><b><center><%=strMessage%></center></b></td>
			</tr>
		  </table>		
		</td>
	  </tr>
	</table>
	<br>
	<%
	end if
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="40">&nbsp;</td>
    <td>
	<table border="0" width="635" bgcolor="#192F68" cellspacing="0" cellpadding="2">
		<tr>
		<td height="20" colspan="2" background="images/backdefault.png" class="WhiteText">
		<center>
		  <span class="Title16 YellowText">File Upload</span><br>
		</center></td>
		</tr>
		<tr>
		  <td width="20" class="WhiteText"><br />
<br /></td>
	      <td width="617" class="WhiteText"><b>NOTE:</b> File name must contain your name and the subject of your document <br />
            <br />
for example <span class="WhiteText"> Joe Bloggs - Maths Coursework</span><br />
<br />
<%
		if strIncludes <> "" then
			Response.Write("File types which can be uploaded = ") & " "
			tempArray = Split(strIncludes,";")
			%>
<span class="Title16">
<%
			For i = 0 to UBOUND(tempArray)
				Response.Write (tempArray(i)) & " "
			Next
			%>
</span>
<%
		end if
		%><br><br></td>
		</tr>
	</table>
	</td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="40">	</td>
    <td>
		<form action="uploadfile.asp?home=vle.asp&folder=<%response.write(folderlocation)%>" method="post" enctype="multipart/form-data">

		<table border="0" width="635" align="left" bgcolor="#192F68" cellspacing="0" cellpadding="2">
		<tr>
			<td colspan="2" class="Title16 YellowText"><center>Select the file to upload</center></td>		
		</tr>
		<tr>
			<td bgcolor="#192F68" class="text"><center>
				<b><span class="WhiteText">File:</span></b>
				<input name="file1" type="file" size="60" />
				</center></td>
		</tr>

		
		<tr>
			<td align="center" bgcolor="#192F68">
            
			<input type="submit" value="Upload" name="submit"></td>
		</tr>		
	</table>
	</form>	
	</td>
  </tr>
</table>
<%
end if

for a = 1 to Amount
	num = 0
	if iconname(a) = "image" then num = 1
	if left(lcase(displaylink),7) = "gallery" then num = 2
	shortcut = "\subjects" + folderlocation + "/" + FullFileName(a)
	if iconname(a) = "url.png" then
		currentfilename = currentfolder + "\" + FullFileName(a)
		Set wfile = fs.OpenTextFile(currentfilename)
		temptext = ""
		for b = 1 to 5
			temptext = wfile.ReadLine
			if lcase(left(temptext,3)) = "url" then
				urltexttemp = temptext 'mid(temptext,5,len(temptext)-5)
				youtube = false
				for z = 1 to len(urltexttemp) - 11
					if lcase(mid(urltexttemp,z,11)) = "youtube.com" then youtube = true
				next
				if youtube = true then 
					youtube = false
					for z = 1 to len(urltexttemp)
						if lcase(mid(urltexttemp,z,1)) = "v" and youtube = false then
							youtube = true
							youtubeclip = mid(urltexttemp,z+2,11)
						end if
					next
				end if
			end if
		next
		shortcut = right(urltexttemp, len(urltexttemp)-4)
		wfile.close 
	end if
	
	videofile = false
	if iconname(a) = "flash.jpg" then videofile = true
	
	if fs.FileExists(currentfolder + "\" + DisplayFileName(a) + ".txt") then
		currentfilename = currentfolder + "\" + DisplayFileName(a) + ".txt"
		Set wfile = fs.OpenTextFile(currentfilename)
			do while not wfile.AtEndOfStream
				filetext = wfile.Readline
				response.write(filetext)
				response.write("<br>")
			loop
		wfile.close 
	end if
	if Extension(Amount) = "db" then num = 99
	
	
	select case num
		case 0
			if youtube = true then iconname(a) = "youtube.png"
%>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="40" height="35"><a href="<%response.write(shortcut)%>" target="_blank"><img src="images/icons/<%response.write(iconname(a))%>" alt="Icon" width="30" height="30" border="0" /></a></td>
    <td width="675" height="35"><a href="<%response.write(shortcut)%>" target="_blank"><%response.write(DisplayFileName(a))%></a></td>
  </tr>
</table>
<%
		if youtube = true then
%>
		<div align="center">
			<iframe width="640" height="360" src="http://www.youtube.com/embed/<%response.write(youtubeclip)%>" frameborder="0" allowfullscreen></iframe>
		</div>
<%
		end if
		youtube = false
		if videofile = true then
			videoname = ""
			for x = 1 to len(shortcut)
				if mid(shortcut,x,1) = "\" then
					videoname = videoname + "/"
				else
					videoname = videoname + mid(shortcut,x,1)
				end if
			next
			'response.write(videoname)
			%>
				<div align="center">
				<object id=0 type="application/x-shockwave-flash" data=player_flv_maxi.swf width=640 height=360>
	      		<noscript>
          		</noscript>
	      		<param name="movie" value=/documents/apps/player_flv_maxi.swf />
	      		<param name="wmode" value="opaque" />
	      		<param name="allowFullScreen" value="true" />
	      		<param name="allowScriptAccess" value="sameDomain" />
	      		<param name="quality" value="high" />
	      		<param name="menu" value="true" />
	      		<param name="autoplay" value="false" />
	      		<param name="autoload" value="false" />
	      		<param name="FlashVars" value="flv=<%response.write(videoname)%>&width=640&height=360&autoplay=0&autoload=0&buffer=5&buffermessage=&playercolor=464646&loadingcolor=999898&buttoncolor=ffffff&buttonovercolor=dddcdc&slidercolor=ffffff&sliderovercolor=dddcdc&showvolume=1&showfullscreen=1&playeralpha=100&title=&margin=0&buffershowbg=0" />
        		</object>
				</div>
			<%
			videofile = false
		end if
		case 1
%>
			<img src="<%response.write("\subjects" + folderlocation + "/" + FullFileName(a))%>" alt="Image" /><br>
<%
		case 2
%>
			<table width="715" border="0" cellspacing="0" cellpadding="0">
			  <tr>
<%
			if FullFileName(a) <> "" then
%>
			    <td width="205" height="205"><a href="<%response.write("\subjects" + folderlocation + "/" + FullFileName(a))%>" target="_blank"><img src="<%response.write("\subjects" + folderlocation + "/" + FullFileName(a))%>" alt="Image" width="200" height="200" border="0" /></a></td>
			    <td width="152" height="205"><a href="<%response.write("\subjects" + folderlocation + "/" + FullFileName(a))%>" target="_blank"><%response.write(DisplayFileName(a))%></a></td>
<%
			end if
			if FullFileName(a+1) <> "" then
%>
			    <td width="205" height="205"><a href="<%response.write("\subjects" + folderlocation + "/" + FullFileName(a+1))%>" target="_blank"><img src="<%response.write("\subjects" + folderlocation + "/" + FullFileName(a+1))%>" alt="Image" width="200" height="200" border="0" /></a></td>
			    <td width="153" height="205"><a href="<%response.write("\subjects" + folderlocation + "/" + FullFileName(a+1))%>" target="_blank"><%response.write(DisplayFileName(a+1))%></a></td>
<%
			else
%>
				<td width="205" height="205"></td>
			    <td width="153" height="205"></td>
<%	
			end if
%>
			  </tr>
			</table>
<%
			a = a + 1
	end select
next

set fo=nothing
 set fs=nothing
 set wfile=nothing
 %>
</p>