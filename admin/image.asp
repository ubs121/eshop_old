<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<style>
<!-- 
body
{
	margin: 0px;
	text-align: center;
}

img{
	border: 0;
}
-->
</style>
<%
img_src = request.querystring("img_src")
if instr(img_src, ";") > 0 then
	'there is more then 1 image
	imgArr = split(img_src, ";")
	image_string = ""
	counter = 0
	for counter = 0 to ubound(imgArr)
		if len(image_string) > 0 then
			image_string = image_string & "<br />" & chr(10) & "<img src=""../images/products/" & imgArr(counter) & """ />"
		else
			image_string = "<img src=""../images/products/" & imgArr(counter) & """ />"
		end if
	next
else
	'there is only 1 image
	image_string = "<img src=""../images/products/" & img_src & """ />"
end if
%>
<title>Image preview</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body onLoad="javascript:resizeImg();">
  <a href="javascript:window.close();"><%=image_string%></a>
</body>
</html>
