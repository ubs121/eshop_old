<%
p = request.querystring("p")
if len(p) > 0 then
	p = getPage(p)
end if

select case p
	case "categories":
		%><!-- #include file="categories.asp" --><%
	case "products":
		%><!-- #include file="products.asp" --><%
	case "reviews":
		%><!-- #include file="reviews.asp" --><%
	case "manufacturers":
		%><!-- #include file="manufacturers.asp" --><%
	case "newsletter":
		%><!-- #include file="newsletter.asp" --><%
	case "accessDenied":
		%><!-- #include file="accessdenied.asp" --><%
	case "stock":
		%><!-- #include file="stock.asp" --><%
end select
%>