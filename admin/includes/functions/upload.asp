<%
goodFiles = "|jpeg|jpg|gif|bmp|png|"
public function uploadFile()
	select case upload_method
		case "dundas":
			call uploadDundas()
		case "pureasp":
			call uploadPureAsp()
		case "fileUp":
			call uploadFileUp()
		case "aspUpload":
			call aspUpload()
		case else
			response.write("msg = ""no upload method selected"";")
	end select
end function

function getExtension(filename)
	extension = mid(filename, instr(filename, ".") + 1, len(filename))
	getExtension = extension
end function

public sub aspUpload()
	Set objUpload = Server.CreateObject("Persits.Upload.1")
	objUpload.save(upload_path)
	
	for each File in objUpload.files
		origName = File.fileName
		if instr(goodFiles, "|" & getExtension(origName) & "|") > 0 then
			response.write("msg = ""The file " & origName & " has been uploaded. \nDo you want to upload another file?"";" & chr(10))
			response.write("file_name = """ & origName & """;")
		else
			file.delete
		end if
	next
	
	set objUpload = nothing
end sub

public sub uploadDundas()
	Set objUpload = Server.CreateObject("Dundas.Upload.2")
	
	objUpload.UseUniqueNames = False
	
	objUpload.save(upload_path)
	origName = objUpload.GetFileName(objUpload.Files(0).OriginalPath)	
	
	if instr(goodFiles, "|" & getExtension(origName) & "|") > 0 then	
		if objUpload.Files.count > 0 then
			response.write("msg = ""The file " & origName & " has been uploaded. \nDo you want to upload another file?"";" & chr(10))
			response.write("file_name = """ & origName & """;")
		end if
	
	else
		For Each objUploadedFile in objUpload.Files
			objUploadedFile.delete
		next
	end if
	
	set objUpload = nothing
end sub

public sub uploadPureAsp()
	Dim MyUploader
	set MyUploader = new FileUploader
		
	MyUploader.Upload()	
		
	Dim File
  	For Each File In MyUploader.Files.Items
		if instr(goodFiles, "|" & lcase(getExtension(MyUploader.files("file1").filename)) & "|") > 0 then
			File.SaveToDisk upload_path
			response.write("msg = ""The file " & MyUploader.Files("file1").FileName & " has been uploaded. \nDo you want to upload another file?"";" & chr(10))
			response.write("file_name = """ & MyUploader.Files("file1").FileName & """;")
		end if
  	Next
end sub

public sub uploadFileUp()
	Dim oFileUp
	Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")
	
	oFileUp.Path = upload_path
	
	If Not oFileUp.Form("file1").IsEmpty Then
		origName = oFileUp.Form("file1").ShortFilename
		if instr(goodFiles, "|" & getExtension(origName) & "|") > 0 then
			oFileUp.Form("file1").Save
			response.write("msg = ""The file " & origName & " has been uploaded. \nDo you want to upload another file?"";" & chr(10))
			response.write("file_name = """ & origName & """;")
		end if
	end if
end sub
%>