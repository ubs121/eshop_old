<%
response.buffer = true
startTimer = Timer

'General settings!!!!
	'Path to the database
	strPath = server.mapPath("admin/database/pdbaspshop.mdb")

function killChars(strWords) 
	dim badChars 
	dim newChars 
	
	badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_") 
	newChars = strWords 
	
	for i = 0 to uBound(badChars) 
	newChars = replace(newChars, badChars(i), "") 
	next 
	
	killChars = Replace(newChars,"'","''")
end function 

module = killChars(request.querystring("mod"))
submodule = killChars(request.querystring("sub"))
action = killChars(request.querystring("action"))
if module="product" or module = "cat" then 
	cat_id = killChars(request.querystring("cat_id"))
	parent_id = killChars(request.querystring("parent_id"))
end if

if len(action) = 0 OR action = "" then
	action = "view"
end if

if len(module) = 0 OR module = "" then
	module = "home"
end if

'Make connection to the database
	set adoCon = server.createobject("ADODB.connection")
	driver     = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & strPath

	adoCon.connectionstring = driver
	adoCon.open

private function GetConfig(value_name)
	rsConfig.filter = "config_name = '" & value_name & "'"
	GetConfig = rsConfig("config_value")
end function

'Determine the comma-separator of the server
tempNumber = cstr(1 / 2)
if instr(tempNumber, ",") > 0 then
	strServerComma = ","
else
	strServerComma = "."
end if


if len(cat_id) > 0 then
	arrCats = Split(cat_id, ",")
end if
'Get config-values from the database
	set rsConfig = server.createobject("ADODB.recordset")
	rsConfig.cursortype = 3
	
	strSQL = "SELECT config_name, config_value FROM config"
	rsConfig.open strSQL, adoCon
	
	strShopName          = GetConfig("shop_name")
	strPageTitle         = GetConfig("shop_title")
	strShopOwner         = GetConfig("shop_owner")
	strDefaultStylesheet = GetConfig("default_stylesheet")
	strCurrency          = GetConfig("shop_currency")
	strMailMethod        = GetConfig("mail_method")
	strMailOut           = GetConfig("mail_out")
	strMailInfo          = GetConfig("mail_info")
	strMailNoReply       = GetConfig("mail_noreply")
	strMailOrders        = GetConfig("mail_orders")
	strShopLink          = GetConfig("shop_link")
	write_complete_url   = cint(GetConfig("write_complete_url"))
	strShowLangModule    = cint(GetConfig("show_language_module"))
	strShowSearchModule  = cint(GetConfig("show_search"))
	strShowNewsletter    = cint(GetConfig("show_newsletter"))
	strShowManuModule    = cint(GetConfig("show_manufacturers"))
	strShowNews          = cint(GetConfig("show_news"))
	strShowPromotions    = cint(GetConfig("show_promotions"))
	strShowPopular       = cint(GetConfig("show_popular"))
	strShowMcart         = cint(GetConfig("show_mcart"))
	strShowLastproducts  = cint(GetConfig("show_lastproducts"))
	strVirtualPath       = GetConfig("virtual_path")
	strRoundNumber       = GetConfig("round_number")
	strProductsPerPage   = GetConfig("products_per_page")
	intShowLoadTime      = cint(GetConfig("show_loadtime"))
	strOffline           = cint(GetConfig("shop_offline"))
	strOfflineMessage    = GetConfig("offline_message")
	show_error           = cint(GetConfig("show_error"))
	strDecimalSign       = GetConfig("decimal_sign")
	strVersionNumber     = Replace(GetConfig("version"), ",", ".")
	strVirtualPath       = Getconfig("virtual_path")
	strSeparator         = getConfig("comma_separator")
	strWeightSign        = getConfig("weightSign")
	showStock            = cint(getConfig("showStock"))
	strMailOrderConf     = cint(getConfig("mail_sendOrderconfirmed"))
	
	rsConfig.close
	set rsConfig = nothing
	
	'Check if the round number is correct
	if IsNumeric(strRoundNumber) then
		strRoundNumber = cint(strRoundNumber)
	else
		strRoundNumber = 0
	end if
	
'Check if the shop is offline
	if strOffline = 1 then
		module = "offline"
	end if
'Get the current page
	if write_complete_url = 0 then
		strCurrPage = ""
		strCurrFile = ""
	else
		strCurrPage = "default.asp?" & request.querystring()
		strCurrFile = "default.asp"
	end if
	
'Select the proper stylesheet
	if len(request.querystring("stylesheet")) > 0 AND request.querystring("stylesheet") <> "" then
		session("stylesheet") = request.querystring("stylesheet")
	else
		if len(session("stylesheet")) = 0 then
			session("stylesheet") = strDefaultStyleSheet
		end if
	end if
	
	strStyleSheet = session("stylesheet")
	
'Select the proper language
if len(session("language")) = 0 OR session("language") = "" then
	set rsLanguage = server.createobject("ADODB.recordset")
	rsLanguage.cursortype = 3
	
	strSQL = "SELECT language_id, language_folder FROM lang WHERE language_default = -1"
	rsLanguage.open strSQL, adoCon
	
	session("language_id") = rsLanguage("language_id")
	session("language")    = rsLanguage("language_folder")
	
	rsLanguage.close
	set rsLanguage = nothing
end if

function getLink(menu_id)
	set rsTemp = server.createobject("ADODB.recordset")
	rsTemp.cursortype = 3
	
	strSQL = "SELECT menu_id, menu_parent_id, menu_name FrOM menu WHERE menu_lang_id = " & session("language_id")
	rsTemp.open strSQL, adoCon
	
	rsTemp.filter = "menu_id = " & cstr(menu_id)
	if not rsTemp.eof then
		parent_id = cint(rsTemp("menu_parent_id"))
	else
		parent_id = 0
	end if
	
	temp = menu_id
	do until parent_id = 0
		rsTemp.filter = "menu_id = " & parent_id
		parent_id = cint(rsTemp("menu_parent_id"))
		
		temp = rsTemp("menu_id") & "," & temp
	loop
	
	rsTemp.close
	set rsTemp = nothing
	
	getLink = temp
end function


sub WriteCopy()
	endTimer = timer
	strVersion = " v" & strVersionNumber
	if intShowLoadTime = 1 then
		response.write(Replace(strPageLoadedIn,"[seconds]", Round((endTimer - startTimer), 3)) & "<br />")
	end if
	response.write("Copyright &copy; " & year(now()) & "</a><br />")
	response.write("Зохиогч <a href=""http://www.izis.mn"" target=""blank""> Zoljargal " &  strVersion & "</a>")
	response.write("</a>")
end sub

function getFileContent(filePath)
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set f=fs.OpenTextFile(Server.MapPath(filePath), 1)
	
	getFileContent = f.readAll
	
	set fs = nothing
	set f  = nothing
end function
%>
<!-- #include file="includes/functions/common.asp" -->
<!-- #include file="languages.asp" -->