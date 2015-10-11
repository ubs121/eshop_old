<script>
function switchDiscount(dType){
	dBlock = document.getElementById("discount_values");
	vInput = document.getElementById("discount_valuta");
	pInput = document.getElementById("discount_percentage");
	helpBlock = document.getElementById("discount_help")
	if(dType == "0"){
		dBlock.style.display = "none";
		vInput.value = 0;
		pInput.value = "0.00";
	} else if(dType == "1") {
		dBlock.style.display = "block";
		vInput.style.display = "block";
		pInput.style.display = "none";
		helpBlock.innerHTML = "&euro; (e.g. 10)";
	} else if(dType == "2") {
		dBlock.style.display = "block";
		vInput.style.display = "none";
		pInput.style.display = "block";
		helpBlock.innerHTML = " (e.g. 0.10 = 10% discount)";
	}
}
</script>
<%
'Initial values
	product_stock = 0

if len(request.form()) > 0 then
	product_name   = makeChars(Replace(request.form("product_name"), "'", "''"))
	product_price  = Replace(request.form("product_price"), ",", ".")
	product_img    = request.form("product_image")
	product_cat_id = request.form("product_cat_id")
	product_url    = Replace(request.form("product_url"), "'", "''")
	product_man_id = request.form("product_man_id")
	error_id       = 0
	product_discount_type = cint(request.form("discount_type"))
	product_weight        = replace(request.form("product_weight"), ",", ".")
	product_stock         = request.form("product_stock")
	
	if len(product_stock) > 0 and isnumeric(product_stock) then
		product_stock = cint(product_stock)
	else
		product_stock = 0
	end if
	
	if product_discount_type = 1 then
		product_discount = replace(request.form("discount_valuta"), ",", ".")
		
		newPrice = product_price - csng(replace(product_discount, ".", strServerComma))
	elseif product_discount_type = 2 then
		product_discount = replace(request.form("discount_percentage"), ",", ".")
		
		newPrice = product_price - (product_price * csng(replace(product_discount, ".", strServerComma)))
	else
		product_discount = 0
		newPrice         = product_price
	end if
	
	if len(product_name) = 0 then
		error_id = 1
	end if
	if len(product_price) = 0 or not isnumeric(product_price) then
		error_id = 1
	end if
	if len(product_man_id) = 0 then
		error_id = 1
	end if
	
	if error_id = 0 then
		set rsProduct = server.createobject("ADODB.recordset")
		rsProduct.open "products", adoCon, 2, 2
			rsProduct.addnew()
				rsProduct("product_name")            = product_name
				rsProduct("product_cat_id")          = product_cat_id
				rsProduct("product_manufacturer_id") = product_man_id
				rsProduct("product_price")           = product_price
				rsProduct("product_image")           = product_img
				rsProduct("product_link")            = product_url
				rsProduct("weight")                  = product_weight
				rsProduct("discount_type")           = product_discount_type
				rsProduct("discount")                = product_discount
				rsProduct("product_stock")           = product_stock
				rsProduct("product_date_added")      = now()
				rsProduct("newPrice")                = newPrice
			rsProduct.update()
			pid = rsProduct("product_id")
		rsProduct.close
		set rsProduct = nothing
		
		response.redirect("catalog.asp?p=" & request.querystring("p") & "&action=add&step=2&pid=" & pid)
	end if	
end if
%>
<form name="frmAddProduct" method="post" action="">
  <table width="500" align="center" cellpadding="2" cellspacing="2">
    <tr> 
      <td colspan="2"><b>Step 1: General information</b></td>
    </tr>
    <tr> 
      <td width="120">Productname:</td>
      <td><input name="product_name" type="text" id="product_name" value="<%=product_name%>"></td>
    </tr>
    <tr> 
      <td>Productprice:</td>
      <td><input name="product_price" type="text" id="product_price" value="<%=product_price%>"></td>
    </tr>
    <tr>
      <td>Weight:</td>
      <td><input name="product_weight" type="text" id="product_weight" value="<%=product_weight%>" /></td>
    </tr>
    <tr>
      <td align="left" valign="top">Stock:</td>
      <td><input name="product_stock" type="text" value="<%=product_stock%>" size="4" /></td>
    </tr>
    <tr>
      <td align="left" valign="top">Discount:</td>
      <td><label>
        <select name="discount_type" id="discount_type" onchange="switchDiscount(this.value);">
          <option value="0">None</option>
          <option value="1">Valuta</option>
          <option value="2">Percentage</option>
        </select>
		<div id="discount_values" style="display:none;">
		  <input name="discount_valuta" type="text" id="discount_valuta" value="0" size="6" />
		  <input name="discount_percentage" type="text" id="discount_percentage" value="0.00" size="6" />
		  <span id="discount_help"></span>		</div>
      </label></td>
    </tr>
    <tr> 
      <td>Productimage:</td>
      <td><input name="product_image" type="text" id="product_image" value="<%=product_img%>"> <input name="btnCheckImg" type="button" id="btnCheckImg" value="Check" onclick="javascript:checkImg();">
        <input name="btnUpload" type="button" id="btnUpload" value="Upload image" onclick="javascript:doUpload('products');" />
        <input name="filebrowser" type="button" id="filebrowser" value="Filebrowser" onclick="showFilebrowser('products');" /></td>
    </tr>
    <tr>
      <td>Product-url:</td>
      <td><input name="product_url" type="text" id="product_url" value="<%=product_url%>"></td>
    </tr>
    <tr> 
      <td>Manufacturer:</td>
      <td>
	    <select name="product_man_id" id="product_man_id">
		<%
		set rsMan = server.createobject("ADODB.recordset")
		rsMan.cursortype = 3
		
		strSQL = "SELECT manufacturer_id, manufacturer_name FROM manufacturer"
		rsMan.open strSQL, adoCon
		
		do while not rsMan.eof
			response.write("<option value=""" & rsMan("manufacturer_id") & """>" & rsMan("manufacturer_name") & "</option>" & chr(13))
			rsMan.movenext
		loop
		%>
        </select>	  </td>
    </tr>
    <tr> 
      <td>Productcategory:</td>
      <td>
<%
set rsMain = server.createobject("ADODB.recordset")
rsMain.cursortype = 3

set rsSubmenu = server.createobject("ADODB.recordset")
rsSubmenu.cursortype = 3

strSQL = "SELECT menu_id, menu_parent_id, menu_name FROM menu WHERE menu_lang_id = " & default_lang_id & " ORDER BY menu_name ASC;"

rsMain.open strSQL, adoCon
rsSubmenu.open strSQL, adoCon

function getSubs(parent_id)
	temp = ""
	rsSubMenu.filter = "menu_parent_id = " & parent_id
	do while not rsSubMenu.eof
		if len(temp) > 0 then
			temp = temp & ";" & rsSubMenu("menu_id")
		else
			temp = rsSubMenu("menu_id")
		end if
		rsSubMenu.movenext
	loop
	getSubs = temp
end function

function getName(menu_id)
	rsSubMenu.filter = "menu_id = " & menu_id
	getName = rsSubMenu("menu_name")
end function

function countSubs(menu_id)
	countSubs = 0
	
	rsSubmenu.filter = "menu_id = " & menu_id
	if not rsSubmenu.eof then
		parent_id = rsSubmenu("menu_parent_id")
	else
		parent_id = 0
	end if
	
	do until cint(parent_id) = 0
		countSubs = countSubs + 1
		
		rsSubMenu.filter = "menu_id = " & parent_id
		if not rsSubmenu.eof then
			parent_id = rsSubMenu("menu_parent_ID")
		else
			parent_id = 0
		end if	
	loop
end function

function writeSubMenusDrop(menu_ID)
	allMenus   = getSubs(menu_id)
	tempBefore = ""
	tempAfter  = ""
	usedID     = 0
	
	if len(allMenus) > 0 then
		arrSubMenus = Split(allMenus, ";")
		
		x = 0
		
		for x = 0 to ubound(arrSubMenus)
			spaces = 3 * countSubs(arrSubMenus(x))
			y = 0
			tempSpaces = ""
			
			for y = 1 to spaces
				tempSpaces = tempSpaces & "&nbsp;"
			next
			tempSpaces = tempSpaces & "|--&nbsp;"
			
			tempBefore = tempBefore & "<option value=""" & arrSubMenus(x) & """"
			if cint(arrSubMenus(x)) = product_cat_id then
				tempBefore = tempBefore & " selected=""selected"""
			end if
			tempBefore = tempBefore & ">" & tempSpaces & getName(arrSubMenus(x)) & "</option>" & chr(10) & writeSubmenusDrop(arrSubmenus(x))
		next
	end if
	writeSubMenusDrop = tempBefore
end function

response.write("<select name=""product_cat_id"" id=""product_cat_id"">" & chr(10))

rsMain.filter = "menu_parent_id = 0"
do while not rsMain.eof
	response.write("<option value=""" & rsMain("menu_id") & """")
	if cint(rsMain("menu_id")) = product_cat_id then
		response.write(" selected=""selected""")
	end if
	response.write(">" & getName(rsMain("menu_ID")) & "</option>" & chr(10))
	
	response.write writeSubmenusDrop(rsMain("menu_ID"))
	
	rsMain.movenext
loop
response.write("</select>" & chr(10))

rsMain.close
set rsMain = nothing

rsSubmenu.close
set rsSubmenu = nothing
%>	  </td>
    </tr>
    <tr align="center"> 
      <td colspan="2"> <%=BuildSubmitter("submit","Next step", request.querystring("p"))%> <input type="button" name="Cancel" value="Cancel" onclick="document.location='?p=<%=request.querystring("p")%>';">      </td>
    </tr>
  </table>
</form>