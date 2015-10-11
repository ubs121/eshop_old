<%
private function updateConfig(config_name)
	config_value = request.form(config_name)
	strSQL = "UPDATE config SET config_value = '" & config_value & "' WHERE config_name = '" & config_name & "'"
	adoCon.execute(strSQL)
end function
%>