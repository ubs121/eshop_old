<%
private function CheckLogin()
	if len(session("admin_user_id")) > 0 AND Isnumeric(session("admin_user_id")) then
		if session("session_id") = session.SessionID then
			CheckLogin = 1
		else
			CheckLogin = 0
		end if
	else
		CheckLogin = 0
	end if
end function

private function DoLogin(login, pass)
	set rsLogin = server.createobject("ADODB.recordset")
	rsLogin.cursortype = 3
	
	strSQL = "SELECT admin_id, admin_salt, admin_password FROM admin_users WHERE admin_login = '" & login & "'"
	rsLogin.open strSQL, adoCon
	
	if not rsLogin.eof then
		correctPass = rsLogin("admin_password")
		controlPass = hashEncode(pass & rsLogin("admin_salt"))
		if correctPass = controlPass then
			DoLogin = 1
			session("admin_user_id") = rsLogin("admin_id")
			session("session_id") = session.SessionID
		else
			DoLogin = 0
		end if
	else
		DoLogin = 0
	end if
	
	rsLogin.close
	set rsLogin = nothing
end function

private function AddLogin()
	admin_login     = request.form("admin_login")
	password1       = request.form("password1")
	password2       = request.form("password2")
	intTotalModules = cint(request.form("intTotalModules"))
	
	redim arrModules(intTotalModules,2)
	
	for x = 1 to intTotalModules
		arrModules(x,1) = request.form("module_id_" & x)
		arrModules(x,2) = request.form("module_" & x)
	next
	
	intError = 0
	
	if len(admin_login) = 0 then
		strError = strError & "<li>You have to give a login</li>"
		intError = 1
	end if
	if len(password1) = 0 then
		strError = strError & "<li>You have to give a password</li>"
		intError = 1
	end if
	if password1 <> password2 then
		strError = strError & "<li>The passwords don't match</li>"
		intError = 1
	end if
	
	if intError = 0 then
		set rsAdmin = server.createobject("ADODB.recordset")
		rsAdmin.Cursortype = 3
		
		strSQL = "SELECT admin_id FROM admin_users WHERE admin_login = '" & admin_login & "'"
		rsAdmin.open strSQL, adoCon
		
		if not rsAdmin.eof then
			intError = 1
			strError = strError & "<li>This login is allready in use</li>"		
		end if
		
		rsAdmin.close
		set rsAdmin = nothing
	end if
	
	if intError = 0 then
		strSalt = getSalt(len(admin_login))
		strSecret = hashEncode(password1 & strSalt)
		
		set rsAdmin = server.createobject("ADODB.recordset")
		rsAdmin.open "admin_users", adoCon, 2, 2
		
		rsAdmin.addnew()
			rsAdmin("admin_login")    = admin_login
			rsAdmin("admin_salt")     = strSalt
			rsAdmin("admin_password") = strSecret
		rsAdmin.update()
		admin_id = rsAdmin("admin_id")
		
		rsAdmin.close
		set rsAdmin = nothing
		
		for x = 1 to intTotalModules
			strSQL = "INSERT INTO admin_rights (admin_id, module_id, module_right) VALUES(" & admin_id & "," & arrModules(x,1) & "," & arrModules(x,2) & ");"
			adoCon.execute(strSQL)
		next
		strError = "<li>The administrator has been added</li>"
	end if
	AddLogin = strError
end function

private function AdjustLogin()
	intError    = 0
	admin_login = request.form("Login")
	password1   = request.form("password1")
	password2   = request.form("password2")
	
	if len(password1) = 0 then
		intError   = 1
		strMessage = "You have to give a password in order to update your login"
	end if
	
	if password1 <> password2 then
		intError   = 1
		strMessage = "The 2 passwords you typed are not equal"
	end if
	
	if len(admin_login) = 0 then
		intError   = 1
		strMessage = "You have to give a login"
	end if

	if intError = 0 then
		strSalt = getSalt(len(admin_login))
		strSecret = hashEncode(password1 & strSalt)
		
		strSQL = "UPDATE admin_users SET admin_login = '" & admin_login & "', admin_salt = '" & strSalt & "', admin_password = '" & strSecret & "' WHERE admin_id = " & session("admin_user_id")
		adoCon.execute(strSQL)
		
		strMessage = "Your login / password has been updated with succes"
	end if
	
	AdjustLogin = strMessage
end function

function UpdateAdmin()
	IntTotalModules = cint(request.form("intTotalModules"))
	for x = 1 to intTotalModules
		right_id = request.form("module_right_id_" & x)
		module_right = request.form("module_" & x)
		if right_id <> 0 then
			strSQL = "UPDATE admin_rights SET module_right = " & module_right & " WHERE module_right_id = " & right_id		
		end if
		adoCon.execute(strSQL)
	next
end function
%>