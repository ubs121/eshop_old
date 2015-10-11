<%
select case module
	case "offline":
		%><!-- #include file="offline/index.asp" --><%
	case "newsletter":
		%><!-- #include file="newsletter/index.asp" --><%
	case "setlanguage":
		%><!-- #include file="setlanguage.asp" --><%
	case "myaccount":
		%><!-- #include file="myaccount/index.asp" --><%
	case "cat":
		%><!-- #include file="products/cat.asp" --><%
	case "product":
		%><!-- #include file="products/product_detail.asp" --><%
	case "redirect":
		%><!-- #include file="redirect.asp" --><%
	case "cart":
		%><!-- #include file="cart/index.asp" --><%
	case "checkout":
		%><!-- #include file="checkout/index.asp" --><%
	case "search":
		%><!-- #include file="search/results.asp" --><%
	case "confirm":
		%><!-- #include file="confirm/index.asp" --><%
	case "cpages":
		%><!-- #include file="custom_pages/index.asp" --><%
	case "contact":
		%><!-- #include file="contact/index.asp" --><%
	case "news":
		%><!-- #include file="news/index.asp" --><%
	case "cancel":
		%><!-- #include file="cancel/index.asp" --><%
	case else
	 	%><!-- #include file="home/index.asp " --><%
end select
%>