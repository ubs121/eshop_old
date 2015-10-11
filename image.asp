<!-- #include file="common.asp" -->
<%
strType = request.querystring("type")
id = request.querystring("id")

if IsNumeric(id) then id = cint(id)

select case strType
	case "product":
		set rsImage = server.createobject("ADODB.recordset")
		rsImage.cursortype = 3
		
		strSQL = "SELECT product_image, product_name FROM products WHERE product_id = " & id
		rsImage.open strSQL, adoCon
		
		if not rsImage.eof then
			image_url = rsImage("product_image")
			image_name = rsImage("product_name")
		end if
		
		rsImage.close
		set rsImage = nothing
		
		if instr(image_url, ";") > 0 then
			'There are multiple images
			arrImages = Split(image_url, ";")
			place     = cint(request.querystring("place"))
					
			if place > ubound(arrImages) then
				image_url = "images/products/" & arrImages(ubound(arrImages))
			elseif place < 0 then
				image_url = "images/products/" & arrImages(0)
			else
				image_url = "images/products/" & arrImages(place)
			end if
		else
			image_url = "images/products/" & image_url
		end if
end select
%>
<html>
<head>
<style>
<!-- 
body
{
	margin: 0px;
}
-->
</style>
<title><%=image_name%></title>
<script language="JavaScript" type="text/javascript">
<!--
function ResizeWindow(){
	imageWidth = document.resizeImage.width + 50;
	imageHeight = document.resizeImage.height + 60;
	winW = screen.width - 100;
	winH = screen.height - 100;
	
	if(winW < imageWidth || winH < imageHeight){
		window.resizeTo(winW, winH);
	} else {
		window.resizeTo(imageWidth, imageHeight);
	}
}
-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body onLoad="javascript:ResizeWindow();">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><a href="javascript:window.close();"><img src="<%=image_url%>" alt="<%=image_name%>" name="resizeImage" border="0" /></a></td>
  </tr>
</table>
</body>
</html>
