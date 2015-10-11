<%
set rsDLang = server.createobject("ADODB.recordset")
rsDLang.cursortype = 3

strSQL = "SELECT language_id FROM lang WHERE language_default = -1"
rsDLang.open strSQL, adoCon

default_lang_id = rsDLang("language_id")

rsDLang.close
set rsDLang = nothing

action = request.querystring("action")

select case action
	case "edit":
		%><!-- #include file="categories/edit.asp" --><%
	case "add":
		%><!-- #include file="categories/add.asp" --><%
	case "delete":
		%><!-- #include file="categories/delete.asp" --><%
	case else
		%><!-- #include file="categories/home.asp" --><%
end select
%>