<%
did = request.querystring("did")

if len(did) > 0 and isnumeric(did) then
	did = cint(did)
else
	response.redirect("?p=" & request.querystring("p"))
end if

if len(request.form()) > 0 then
	totalLang = request.form("totalLang")
	a         = request.form("a")
	b         = request.form("b")
	x         = 0
	hasName   = false
	
	for x = 1 to totalLang
		if len(request.form("d_name_" & x)) > 0 then
			hasName = true
			exit for
		end if
	next
	
	if a <> "1" then
		tCats = request.form("tCats")
		
		x = 0
		a = ""
		b = ""
		for x = 1 to tCats
			weight = request.form("a_" & x)
			price  = request.form("b_" & x)
			
			if len(a) > 0 then
				a = a & ";" & weight
				b = b & ";" & price
			else
				a = weight
				b = price
			end if
		next
		
		'Order the products
		arrConditions = split(a, ";")
		arrPrices     = split(b, ";")
		
		x = 0
		for x = 0 to ubound(arrConditions) - 1
			for y = x + 1 to ubound(arrConditions) - 1
				weight = csng(Replace(arrConditions(x), ".", strServerComma))
				nWeight = csng(Replace(arrConditions(y), ".", strServerComma))
				price   = arrPrices(x)
				nPrice  = arrPrices(y)
			
				if weight < nWeight then
					arrConditions(x) = nWeight
					arrPrices(x)     = nPrice
					arrConditions(y) = weight
					arrPrices(y)     = Price
				end if					
			next
		next
		
		x = 0
		a = ""
		b = ""
		for x = 0 to ubound(arrConditions)
			if x = 0 AND x = ubound(arrConditions) then
				a = ">=" & arrConditions(x)
				b = arrPrices(x)
			elseif x = 0 then
				a = "<" & arrConditions(x)
				b = arrPrices(x)
			elseif x = ubound(arrConditions) then
				a = a & ";>=" & arrConditions(x)
				b = b & ";" & arrPrices(x)
			else
				a = a & ";<" & arrConditions(x)
				b = b & ";" & arrPrices(x)
			end if
		next
	end if
	
	if hasName then
		set rsDeliv = server.createobject("ADODB.recordset")
		
		strSQL = "SELECT * FROM delivery WHERE delivery_ID = " & did
		rsDeliv.open strSQL, adoCon, 2, 2
		
		x = 0
		for x = 1 to totalLang
			duid    = request.form("duid_" & x)
			d_name  = request.form("d_name_" & x)
			lang_id = request.form("lang_id_" & x)
			
			rsDeliv.filter = "delivery_unique_id = " & duid
			
			if not rsDeliv.eof then
				rsDeliv("delivery_name")  = d_name
				rsDeliv("a")              = replace(a, ",", ".")
				rsDeliv("b")              = replace(b, ",", ".")
				rsDeliv("lang_id")        = lang_id
				
				rsDeliv.update()
			else
				rsDeliv.addnew()
					rsDeliv("delivery_id")    = did
					rsDeliv("lang_id")        = lang_id
					rsDeliv("delivery_name")  = d_name
					rsDeliv("a")              = replace(a, ",", ".")
					rsDeliv("b")              = replace(b, ",", ".")
				rsDeliv.update()
			end if
		next
		
		rsDeliv.close
		set rsDeliv = nothing
	end if
end if

if len(did) > 0 and isnumeric(did) then
	did = cint(did)
else
	response.redirect("?p=" & request.querystring("p"))
end if

set rsDelivery = server.createobject("ADODB.recordset")
rsDelivery.cursortype = 3

strSQL = "SELECT delivery_unique_id, lang_id, delivery_name, a, b FROM delivery WHERE delivery_id = " & did
rsDelivery.open strSQL, adoCon

if not rsDelivery.eof then
	dPrice = rsDelivery("b")
else
	dPrice = 0
end if

set rsLang = server.createobject("ADODB.recordset")
rsLang.cursortype = 3

strSQL = "SELECT language_name, language_id FROM lang WHERE language_show = -1"
rsLang.open strSQL, adoCon
%>
<% if request.querystring("opt") = "f" then %>
<form name="frmUpdateDelivery" method="post" action="">
  <input type="hidden" name="a" value="1" />
  <table width="500" align="center" cellpadding="2" cellspacing="2" style="border: solid 1px #000000;">
    <tr> 
      <td colspan="2" bgcolor="#666666"><strong><font color="#FFFFFF">&nbsp;Update 
        deliverymethod</font></strong></td>
    </tr>
    <tr> 
      <td width="100">Price:</td>
      <td><%=shop_currency%> <input name="b" type="text" id="b" value="<%=dPrice%>" size="6"></td>
    </tr>
    <tr> 
      <td colspan="2"><strong>Translations:</strong></td>
    </tr>
	<%
	x = 0
	do while not rsLang.eof
		x = x + 1
		
		rsDelivery.filter = "lang_id = " & rsLang("language_id")
		if not rsDelivery.eof then
			duid   = rsDelivery("delivery_unique_id")
			d_name = rsDelivery("delivery_name")
		else
			duid   = 0
			d_name = ""
		end if
	%>
    <tr> 
      <td>&nbsp;<%=rsLang("language_name")%></td>
      <td>
	  	<input type="hidden" name="lang_id_<%=x%>" value="<%=rsLang("language_id")%>" />
		<input name="d_name_<%=x%>" type="text" id="d_name_<%=x%>" value="<%=d_name%>" size="40" />
		<input type="hidden" name="duid_<%=x%>" value="<%=duid%>" />
	  </td>
    </tr>
	<%
		rsLang.movenext
	loop
	%>
    <tr align="center"> 
      <td colspan="2">
	  	<input type="hidden" name="totalLang" value="<%=rsLang.recordcount%>" />
	  	<%=buildSubmitter("cmdUpdate","Update deliverymethod", request.querystring("p"))%>&nbsp;
        <input name="btnBack" type="button" id="btnBack" value="Back" onclick="document.location='?p=<%=request.querystring("p")%>';"></td>
    </tr>
  </table>
</form>
<% elseif request.querystring("opt") = "v" then %>
<script>
hWeight = 0;
tCats   = 1
function setWeight(weight, cats){
	hWeight = weight;
	tCats   = cats;
}

function changeCats(cats){
	document.location = "?p=<%=request.querystring("p")%>&action=edit&did=<%=did%>&opt=v&cats=" + cats;
}

function updateWeight(weight){
	weight = parseFloat(weight);
	if(weight > hWeight){
		hWeight = weight;
		document.getElementById("hWeight").value = hWeight;
		document.getElementById("a_" + tCats).value = hWeight;
	}
}
</script>
<form name="frmUpdateDelivery" method="post" action="">
  <%
  if not rsDelivery.eof then
  	a = rsDelivery("a")
	b = rsDelivery("b")
  end if
  arrConditions = split(a, ";")
  arrPrices     = split(b, ";")
  
  cats = request.querystring("cats")
  if len(cats) > 0 and isnumeric(cats) then
  	cats = cint(cats)
  else
  	cats = ubound(arrConditions) + 1
  end if
  %>
  <table width="500" align="center" cellpadding="2" cellspacing="2" style="border: solid 1px #000000;">
    <tr> 
      <td colspan="2" bgcolor="#666666"><strong><font color="#FFFFFF">&nbsp;Update 
        deliverymethod</font></strong></td>
	<tr>
	  <td colspan="2"><b>Variable prices:
	    <select name="cats" id="cats" onchange="changeCats(this.value);">
		<%
		x = 0
		for x = 1 to 10
			response.write "<option value=""" & x & """"
			if x = cats then
				response.write " selected=""selected"""
			end if
			response.write ">" & x & " Categories</option>" & chr(10)
		next
		%>
        </select>
	  </b></td>
    </tr>
	<tr>
	  <td colspan="2"><table width="300" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td width="20" bgcolor="#999999">&nbsp;</td>
          <td bgcolor="#999999"><strong>Weight</strong></td>
          <td width="20" align="center" bgcolor="#999999"><strong>=</strong></td>
          <td bgcolor="#999999"><strong>Price</strong></td>
        </tr>
		<%
		x = 0
		hWeight = 0
		for x = 1 to cats - 1
			if ubound(arrConditions) >= x then
				price = arrPrices(x - 1)
				weight = arrConditions(x - 1)
				weight = right(arrConditions(x - 1), len(arrConditions(x - 1)) - 1)
			else
				price  = 0
				weight = 0
			end if		
			if csng(replace(weight, ".", strServerCommma)) > hWeight then hWeight = weight
		%>
        <tr>
          <td>&lt;</td>
          <td><input name="a_<%=x%>" type="text" id="a_<%=x%>" value="<%=weight%>" size="10" onchange="updateWeight(this.value);" />
            <%=strWeightSign%></td>
          <td align="center">=</td>
          <td><%=shop_currency%><input name="b_<%=x%>" type="text" id="b_<%=x%>" value="<%=price%>" size="10" /></td>
        </tr>
		<%
		next
		
		if cats - 1 = ubound(arrPrices) then
			price = arrPrices(cats - 1)
		else
			price = 0
		end if
		%>
        <tr>
          <td>&gt;=</td>
          <td><input name="hWeight" type="text" id="hWeight" value="<%=hWeight%>" size="10" disabled="disabled" />
            <%=strWeightSign%></td>
          <td align="center">=</td>
          <td><%=shop_currency%><input name="b_<%=cats%>" type="text" id="b_<%=cats%>" value="<%=price%>" size="10" /></td>
        </tr>
      </table>
	  <input type="hidden" name="a_<%=cats%>" id="a_<%=cats%>" value="<%=hWeight%>" />
	  <input type="hidden" name="tCats" id="tCats" value="<%=cats%>" />
	  <script>setWeight(<%=hWeight%>,<%=cats%>);</script>
	  </td>
    </tr>
	<tr> 
      <td colspan="2"><strong>Translations:</strong></td>
    </tr>
    </tr>
	<%
	x = 0
	do while not rsLang.eof
		x = x + 1
		
		rsDelivery.filter = "lang_id = " & rsLang("language_id")
		if not rsDelivery.eof then
			duid   = rsDelivery("delivery_unique_id")
			d_name = rsDelivery("delivery_name")
		else
			duid   = 0
			d_name = ""
		end if
	%>
    <tr> 
      <td>&nbsp;<%=rsLang("language_name")%></td>
      <td>
	  	<input type="hidden" name="lang_id_<%=x%>" value="<%=rsLang("language_id")%>" />
		<input name="d_name_<%=x%>" type="text" id="d_name_<%=x%>" value="<%=d_name%>" size="40" />
		<input type="hidden" name="duid_<%=x%>" value="<%=duid%>" />	  </td>
    </tr>
	<%
		rsLang.movenext
	loop
	%>
    <tr align="center"> 
      <td colspan="2">
	  	<input type="hidden" name="totalLang" value="<%=rsLang.recordcount%>" />
	  	<%=buildSubmitter("cmdUpdate","Update deliverymethod", request.querystring("p"))%>&nbsp;
        <input name="btnBack" type="button" id="btnBack" value="Back" onclick="document.location='?p=<%=request.querystring("p")%>';"></td>
    </tr>
  </table>
</form>
<% end if %>
<%
rsDelivery.close
set rsDelivery = nothing

rsLang.close
set rsLang = nothing
%>