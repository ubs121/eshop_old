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
pid = request.querystring("pid")
if len(pid) = 0 or not isnumeric(pid) then response.redirect("?p=" & request.querystring("p"))
if len(request.form()) > 0 then
	product_name          = makeChars(request.form("product_name"))
	product_price         = Replace(request.form("product_price"), ",", ".")
	product_img           = request.form("product_image")
	product_cat_id        = cint(request.form("product_cat_id"))
	product_url           = request.form("product_url")
	product_man_id        = cint(request.form("product_man_id"))
	error_id              = 0
	product_discount_type = cint(request.form("discount_type"))
	product_weight        = replace(request.form("product_weight"), ",", ".")
	product_stock         = request.form("product_stock")
	
	product_price = Replace(product_price, ",", ".")
	product_price = csng(replace(product_price, ".", strServerComma))
	
	if len(product_stock) > 0 AND isnumeric(product_stock) then
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
	
	product_price = Replace(product_price, ",", ".")
	newPrice      = replace(newPrice, ",", ".")
	
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
		strSQL = "SELECT * FROM products WHERE product_id = " & pid
		
		rsProduct.open strSQL, adoCon, 2, 2
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
				rsProduct("newPrice")                = newPrice
			rsProduct.update()

		rsProduct.close
		set rsProduct = nothing
		
		response.redirect("catalog.asp?p=" & request.querystring("p") & "&action=edit&step=2&pid=" & pid)
	end if	
else
	set rsProduct = server.createobject("ADODB.recordset")
	rsProduct.cursortype = 3
	
	strSQL = "SELECT * FROM products WHERE product_id = " & pid
	rsProduct.open strSQL, adoCon
	
	product_name = Replace(rsProduct("product_name"), "'", "''")
	product_price = rsProduct("product_price")
	product_img   = rsProduct("product_image")
	product_url   = Replace(rsProduct("product_link"), "'", "''")
	product_man_id = rsProduct("product_manufacturer_id")
	product_cat_id = cint(rsProduct("product_cat_id"))
	product_weight = rsProduct("weight")
	product_discount_type = rsProduct("discount_type")
	product_discount      = rsProduct("discount")
	product_stock         = rsProduct("product_stock")
	
	rsProduct.close
	set rsProduct = nothing
end if
%>
<form name="frmEditProduct" method="post" action="">
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
      <td><input name="product_stock" type="text" id="product_stock" value="<%=product_stock%>" size="4" /></td>
    </tr>
    <tr>
      <td align="left" valign="top">Discount:</td>
      <td><label>
        <select name="discount_type" id="discount_type" onchange="switchDiscount(this.value);">
          <option value="0">None</option>
          <option value="1"<% if product_discount_type = 1 then %> selected="selected"<% end if %>>Valuta</option>
          <option value="2"<% if product_discount_type = 2 then %> selected="selected"<% end if %>>Percentage</option>
        </select>
		<div id="discount_values" style="display:none;">
		  <input name="discount_valuta" type="text" id="discount_valuta" value="<%=product_discount%>" size="6" />
		  <input name="discount_percentage" type="text" id="discount_percentage" value="<%=product_discount%>" size="6" />
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
			response.write("<option value=""" & rsMan("manufacturer_id") & """")
			if cint(rsMan("manufacturer_id")) = product_man_id then
				response.write(" selected=""selected""")
			end if
			response.write(">" & rsMan("manufacturer_name") & "</option>" & chr(13))
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
<script>
switchDiscount(<%=product_discount_type%>)
</script>