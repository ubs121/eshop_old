<%
step = request.querystring("step")
if len(step) = 0 then
	step = 1
else
	step = cint(step)
end if

select case step
	case "1":
		%><!-- #include file="add/1.asp" --><%
	case "2":
	%><!-- #include file="add/2.asp" --><%
end select
%>