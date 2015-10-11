<!-- #include file="includes/functions/pure_asp_upload.asp" -->
<!-- #include file="common.asp" -->
<!-- #include file="includes/functions/upload.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <style type="text/css">
        fieldset
        {
            background: #EEEEEE;
        }
        legend
        {
            font-weight: bold;
        }
        .center
        {
            text-align: center;
            margin-top: 4px;
        }
    </style>

    <script>
var msg = "";
<%
upload_path = server.mappath(strVirtualPath)
upload_path = upload_path & "\images\" & request.querystring("type") & "\"

if len(request.querystring("first")) = 0 AND len(request.querystring("type")) > 0 then
	call uploadFile()
end if

%>
if (msg != ""){
	msg = "UPLOAD COMPLETE \n------------------------\n" + msg;
	choice = confirm(msg);
	if(choice != true){
		<% if request.querystring("type") = "products" then %>
		window.opener.document.getElementById("product_image").value = file_name;
		<% else %>
		window.opener.document.location.reload();
		<% end if %>
		window.close();
	}
}
    </script>

    <title>Upload</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
    <fieldset>
        <legend>upload</legend>
        <% if CheckLogin() = 0 then %>
        <p align="center">
            <b>You have to login to be able to upload</b></p>
        <% else %>
        <form method="post" action="upload.asp?type=<%=request.querystring("type")%>" enctype="multipart/form-data"
        name="frmUpload">
        <input type="file" name="File1" id="file1" size="26" />
        <input type="submit" name="cmdUpload" value="Upload" />
        <div class="center">
            <input type="button" name="btnClose" value="Close" onclick="javascript:window.close();" /></div>
        </form>
        <% end if %>
    </fieldset>
</body>
</html>
