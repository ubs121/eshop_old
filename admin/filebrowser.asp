<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file="common.asp" -->
<%
'Control if logged in
if len(session("admin_user_id")) = 0 or session("admin_user_id") = "" then response.redirect("login.asp?p=login")

folder = request.querystring("folder")
if folder <> "category" AND folder <> "products" THEN
	folder = ""
end if

if len(folder) > 0 then
	path = server.mappath(strVirtualPath) & "\images\" & folder & "\"

	set fs = server.createobject("scripting.filesystemobject")
	set fo = fs.getfolder(path)
	
	if len(request.querystring("delfiles")) > 0 then
		arrFiles = split(request.querystring("delfiles"), ";")
		
		x = 0
		for x = 0 to ubound(arrFiles)
			set delfile = fs.getfile(path &  arrFiles(x))
			delfile.delete()
			
			set delfile = nothing
		next
	end if
	
	total_pics = 0
	for each picture in fo.files
		total_pics = total_pics + 1
	next
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Filebrowser v1.0</title>
<script>
var totalPictures = <%=total_pics%>;

function getPicstring(){
	var picString  = "";
	var stringUsed = false;
	for(x=1; x <= totalPictures; x++){
		slPic = document.getElementById("slPic_" + x);
		txtPic = document.getElementById("txtPic_" + x);
		
		if(slPic.checked == true){
			if(stringUsed){
				picString = picString + ";" + txtPic.value
			} else {
				picString  = txtPic.value;
				stringUsed = true;
			}
		}
	}
	return picString;
}

function attachFiles(){
	var picString = getPicstring();
	
	window.opener.document.getElementById("product_image").value = picString;
	window.close();
}

function deleteFiles(){
	var picString = getPicstring();
	document.location = "filebrowser.asp?folder=<%=request.querystring("folder")%>&delfiles=" + picString;
}
</script>
<style>
body{
	margin: 0px;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	background: #F5EB18
}
#spacer_middle{
	height: 300px;
	width: 6px;
	background: url("images/picbrowser/line_between.jpg");
	float: left;
}

#leftcontent{
	width: 150px;
	float: left;
	height: 300px;
	background: #FFFFFF;
}

#leftcontent h2{
	height: 25px;
	text-align: center;
	background: url("images/picbrowser/bar_bg.jpg");
	font-size: 14px;
	color: #FFFFFF;
	padding: 4 4 4 4;
	margin: 0px;
}

#rightcontent{
	float: left;
	margin: 0px;
	height: 300px;
	width: 400px;
	background: #FFFFFF;
}

#rightcontent h2{
	height: 25px;
	background: url("images/picbrowser/bar_bg.jpg");
	padding: 0px;
	margin: 0px;
}

#pictures{
	width: 390px;
	height: 260px;
	overflow: auto;
	margin: 4px;
	font-size: 10px;
}

img{
	border: 0px;
	margin: 0px;
	padding: 0px;
}

#header{
	height: 25px;
	color: #AABBCC;
	font-weight: bold;
	padding: 4px;	
}
</style>
</head>
<body>
<div id="header">FILE BROWSER v1.0</div>
<div id="leftcontent">
  <h2>Folders</h2>
  <img src="images/picbrowser/dossier.gif" width="20" height="20" align="absmiddle" /><a href="?folder=category">Category</a><br />
  <img src="images/picbrowser/dossier.gif" width="20" height="20" align="absmiddle" /><a href="?folder=products">Products</a></div>
<div id="spacer_middle">&nbsp;</div>
<div id="rightcontent">
  <h2>
    <a href="javascript:attachFiles();"><img src="images/picbrowser/attach_files.jpg" border="0" /></a>
	<a href="javascript:deleteFiles();"><img src="images/picbrowser/delete_files.jpg" border="0" /></a> </h2>
  <div id="pictures">
  	<% if len(folder) > 0 then %>
    <form name="frmFiles" id="frmFiles" action="" method="post">
    <table width="95%" cellspacing="0" cellpadding="2" style="border: 0;">
	  <tr>
	    <td width="20">&nbsp;</td>
		<td>Name</td>
		<td width="80">Size</td>
	  </tr>
	  <%
	  counter = 0
	  for each picture in fo.files
	  	counter = counter + 1
		response.write "<tr>" & chr(10) & _
			"  <td width=""20"">" & chr(10) & _
			"    <input type=""checkbox"" name=""slPic_" & counter & """ id=""slPic_" & counter & """ value=""1"" />" & chr(10) & _
			"  </td>" & chr(10) & _
			"  <td><input type=""hidden"" name=""txtPic_" & counter & """ id=""txtPic_" & counter & """ value=""" & picture.name & """ />"& picture.name & "</td>" & chr(10) & _
			"  <td>" & round(picture.size / 1024) & "kb</td>" & chr(10) & _
			"</tr>" & chr(10)
	  next
	  %>
	</table>
	</form>
	<%
	set fo = nothing
	set fs = nothing
	%>
	<% else %>
	Select a folder on the left
	<% end if %>
  </div>
</div>
</body>
</html>
