<%
if session("customer_id") = "" or len(session("customer_id")) = 0 then
	response.redirect("?mod=myaccount&sub=login&red=checkout")
end if

TotalProducts = request.cookies("total_products" & session.SessionID)
if len(TotalProducts) = 0 OR NOT isnumeric(TotalProducts) then
	TotalProducts = 0
end if

if TotalProducts = 0 then
	if request.querystring("p") <> "4" then
		response.redirect("?mod=cart")
	end if
end if
%>
<p class="pageheader"><%=strCheckout%></p>
<%
step = request.querystring("p")
select case step
	case "4":
		%><!-- #include file="step4/index.asp" --><%
	case "3":
		%><!-- #include file="step3/index.asp" --><%
	case "2":
		%><!-- #include file="step2/index.asp" --><%
	case "agreement":
		%><!-- #include file="agreement/index.asp" --><%
	case else
		%><!-- #include file="step1/index.asp" --><%
end select
%>