<%
response.buffer = true

page = request.ServerVariables("SCRIPT_NAME")
do while instr(page,"/") > 0
	page = right(page, len(page) - instr(page,"/"))
loop

'Determine the comma-separator of the server
tempNumber = cstr(1 / 2)
if instr(tempNumber, ",") > 0 then
	strServerComma = ","
else
	strServerComma = "."
end if

strPath = server.MapPath("database/pdbaspshop.mdb")
'Make connection to the database
	set adoCon = server.createobject("ADODB.connection")
	driver     = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & strPath

	adoCon.connectionstring = driver
	adoCon.open

'Common functions
private function getConfig(config_name)
	rsConfig.filter = "config_name = '" & config_name & "'"
	if not rsConfig.eof then
		getConfig = rsConfig("config_value")
	end if
end function

set rsConfig = server.createobject("ADODB.recordset")
rsConfig.cursortype = 3

strSQL = "SELECT config_name, config_value FROM config"
rsConfig.open strSQL, adoCon
	'get config values
	strVirtualPath   = getConfig("virtual_path")
	auto_update_lang = getConfig("auto_update_lang")
	upload_method    = getConfig("upload_method")
	shop_currency    = getConfig("shop_currency")
	currVersion      = Replace ( getConfig("version"), ",", "." )
	mail_noreply     = getConfig("mail_noreply")
	strMailMethod    = getConfig("mail_method")
	strMailOut       = getConfig("mail_out")
	stock_autoUpdate = cint(getConfig("stock_autoupdate"))
	strWeightSign       = getConfig("weightSign")
	convertChars        = cint(getConfig("convertchars"))
rsConfig.close
set rsConfig = nothing

if right(strVirtualPath,1) <> "/" AND right(strVirtualPath,1) <> "\" then
	strVirtualPath = strVirtualPath & "/"
end if

set rsDefaultlang = server.createobject("ADODB.recordset")
rsDefaultlang.cursortype = 3

strSQL = "SELECT language_id, language_folder FROM lang WHERE language_default = -1"
rsDefaultlang.open strSQL, adocon

strDefaultLang = rsDefaultlang("language_folder")
default_lang_id = rsDefaultlang("language_id")

rsDefaultlang.close
set rsDefaultlang = nothing

function makeChars(text)
  if convertChars = 1 then
	  intI    = 0
	  newText = ""
	  for intI = 1 to len(text)
		newText = newText & "&#" & ascw(mid(text, intI, 1)) & ";"
	  next
	  
	  response.write newText
	  
	  makeChars = newText
  else
  	  makeChars = text
  end if
end function

Function RoundNumber(intNumber)
	intNumber = csng(Replace(intNumber, ".", strServerComma))
	
	if len(strRoundNumber) > 0 then
		roundtemp = Round(intNumber, strRoundNumber)
		comma_place = instr(roundtemp, strServerComma)
		
		if comma_place > 0 then
			x = 0
			for x = comma_place + 1 to (comma_place + strRoundNumber)
				if NOT isnumeric(mid(roundtemp, x, 1)) then
					roundtemp = roundtemp & "0"
				end if
			next
		else
			roundtemp = roundtemp & ","
			x = 0
			for x = 1 to strRoundNumber
				roundtemp = roundtemp & "0"
			next
		end if
	else
		roundtemp = Round(intNumber, 0)
	end if
	RoundNumber = roundtemp
	roundtemp = Replace(RoundNumber, strServerComma, strDecimalSign)
	
	if len(strSeparator) > 0 AND instr(roundtemp, strDecimalSign) > 0 then
		before_comma = getComma(left(roundtemp, instr(roundtemp, strDecimalSign) - 1))
		roundtemp = before_comma & mid(roundtemp, instr(roundtemp, strDecimalSign))
	end if
	RoundNumber = roundtemp
end function

function getComma(intNumber)
	if len(intNumber) > 0 and isnumeric(intNumber) then
		length = len(intNumber)
		if length / 3 = length \ 3 then
			total_comma = (length \ 3) - 1
		else
			total_comma = length \ 3
		end if
			
		y = 0
		if total_comma > 0 then
			for y = 1 to total_comma
				left_side = left(intNumber, length - ( y * 3))
				right_side = right(intNumber, ((y * 3) + (y - 1)))
				intNumber = left_side & strSeparator & right_side
			next
		end if		
		getComma = intNumber
	end if
end function
%>
<!-- #include file="includes/functions/userrights.asp" -->
<% if page="login.asp" then %>
<!-- #include file="includes/functions/hash1way.asp" -->
<!-- #include file="includes/functions/login.asp" -->
<% end if %>
<% if page="config.asp" then %>
<!-- #include file="includes/functions/config.asp" -->
<% end if %>