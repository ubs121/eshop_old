<% @ Language=VBScript %>
<%
'Set the buffer to true
Response.Buffer = True

'Declare variables
Dim objADOStream		'Holds the stream object
Dim intImageNumber		'Holds which image number to show
Dim strImagePath		'Holds the image path and image name

'Initiliase variables
strVirtualPath = left(request.ServerVariables("SCRIPT_NAME"), len(request.ServerVariables("SCRIPT_NAME")) - 18)
strImagePath = "images/"

'Get the image number
intImageNumber = CInt(Request.QueryString("I"))

'Get the image security code number to display
intImageNumber = Mid(Session("SecurityCode"), intImageNumber, 1)

'Get the path and the image to show
strImagePath = server.mappath(strVirtualPath & strImagePath & intImageNumber & ".gif")

'Set the stream object
Set objADOStream = server.createobject("ADODB.Stream")

'Open the streem oject
objADOStream.Open

'Set the stream object type to binary
objADOStream.Type = 1

'Load in the image gif
objADOStream.LoadFromFile strImagePath



'Set the right response content type for the image
Response.ContentType = "image/gif"

'Display image
Response.BinaryWrite objADOStream.Read

'Flush the response object
Response.Flush

'Clean up
objADOStream.Close
Set objADOStream=Nothing
%>