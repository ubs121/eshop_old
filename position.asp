<div id="position">
<%
strPosition = "<a href=""?mod=home"">" & strShop & "</a>"
cat_id = request.querystring("cat_id")

if module = "product" then
	set rsCat = server.createobject("ADODB.recordset")
	rsCat.cursortype = 3
	
	strSQL = "SELECT menu_id, menu_name FROM menu WHERE menu_lang_id = " & session("language_id")
	rsCat.open strSQL, adoCon
	
	x = 0
	catPos = ""
	temp   = ""
	for x = 0 to ubound(arrCats)
		if len(catPos) > 0 then
			catPos = catPos & "," & arrCats(x)
		else
			catPos = arrCats(x)
		end if
		rsCat.filter = "menu_id = " & arrCats(x)
		if not rsCat.eof then
			strProductPosition = strProductPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=cat&amp;cat_id=" & catPos & """>" & rsCat("menu_name") & "</a>"
		end if
	next
	
	rsCat.close
	set rsCat = nothing
	
	strPosition = strPosition & strProductPosition
	
	product_id = request.querystring("product_id")
	
	set rsProduct = server.Createobject("ADODB.recordset")
	rsProduct.cursortype = 3
	
	strSQL = "SELECT product_cat_id, product_name FROM products WHERE product_id = " & product_id
	rsProduct.open strSQL, adoCon
	
	product_name = rsProduct("product_name")
	product_parent_id = rsProduct("product_cat_id")
	
	rsProduct.close
	set rsProduct = nothing
	
	strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""#"">" & product_name & "</a>"
elseif module = "cat" then

	set rsCat = server.createobject("ADODB.recordset")
	rsCat.cursortype = 3
	
	strSQL = "SELECT menu_id, menu_name FROM menu WHERE menu_lang_id = " & session("language_id")
	rsCat.open strSQL, adoCon
	
	x = 0
	catPos = ""
	
	for x = 0 to ubound(arrCats)
		if len(catPos) > 0 then
			catPos = catPos & "," & arrCats(x)
		else
			catPos = arrCats(x)
		end if
		rsCat.filter = "menu_id = " & arrCats(x)
		strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=cat&amp;cat_id=" & catPos & """>" & rsCat("menu_name") & "</a>"
	next
	
	rsCat.Close
	set rsCAt = nothing
elseif module = "search" then
	if request.querystring("type") = "advanced" then
		strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=search&amp;type=advanced"">" & strAdvanced & "&nbsp;" & strSearch & "</a>"
	else
		strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=search&amp;type=simple"">" & strSimple & "&nbsp;" & strSearch & "</a>"
	end if
else
	select case module
		case "home":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=home"">" & strHome & "</a>"
		case "myaccount":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=myaccount"">" & strMyAccount & "</a>"
			select case request.querystring("sub")
				case "addressbook":
					strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=myaccount&amp;sub=addressbook"">" & strAddressbook & "</a>"
				case "details":
					strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=myaccount&amp;sub=details"">" & strPersonalInformation & "</a>"
				case "password":
					strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=myaccount&amp;sub=password"">" & strChangePassword & "</a>"
				case "register":
					strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=myaccount&amp;sub=register"">" & strRegister & "</a>"
				case "login":
					strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=myaccount&amp;sub=login"">" & strLogin & "</a>"
				case "lostpass":
					strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=myaccount&amp;sub=lostpass"">" & strLostPass & "</a>"
				case "newsletter":
					strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=myaccount&amp;sub=newsletter"">" & strNewsLetter & "</a>"
				case "orders_history":
					strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=myaccount&amp;sub=orders_history"">" & strOrderhistory & "</a>"
				case else
					strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=myaccount&amp;sub="">" & strOverview & "</a>"
			end select
		case "checkout":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=checkout"">" & strCheckOut & "</a>"
		case "cart":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=cart"">" & strShoppingCart & "</a>"
		case "conditions":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=conditions"">" & strMenuConditions & "</a>"
		case "about":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=about"">" & strMenuAbout & "</a>"
		case "contact":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=contact"">" & strContact & "</a>"
		case "confirm":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=confirm"">" & strConfirm & "</a>"
		case "newsletter":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=newsletter"">" & strNewsLetter & "</a>"
		case "offline":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=offline"">" & strOffline & "</a>"
		case "news":
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=news"">" & strNews & "</a>"
		case "cpages":
			set rsPage = server.createobject("ADODB.recordset")
			rsPage.cursortype = 3
			
			strSQL = "SELECT page_name FROM custom_pages WHERE page_id = " & request.querystring("page_id") & " AND page_lang_id = " & session("language_id")
			rsPage.open strSQL, adoCon
			
			if not rsPage.eof then
				page_name = rsPage("page_name")
			end if
			
			rsPage.close
			set rsPage = nothing
			
			strPosition = strPosition & "&nbsp;&raquo;&nbsp;<a href=""?mod=cpages&amp;page_id=" & request.querystring("page_id") & """>" & page_name & "</a>"
	end select
end if
%>
<p>
  <%=strPosition%>
</p>
</div>