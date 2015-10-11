<script>
function submitForm(){
	document.getElementById("frmPay").submit();
}
</script>
<%
set rsPaypal = server.createobject("ADODB.recordset")
rsPaypal.cursortype = 3

strSQL = "SELECT payment_options FROM payment WHERE payment_ID = 2"
rsPaypal.open strSQL, adoCon

if not rsPaypal.eof then
	arrOptions = Split(rsPaypal("payment_options"), ";")
end if

rsPaypal.close
set rsPaypal = nothing

set rsOrder = server.createobject("ADODB.recordset")
rsOrder.cursortype = 3

strSQL = "SELECT user_ID, address_ID, total_price FROM orders WHERE order_id = " & order_id
rsOrder.open strSQL, adoCon

p_user_id    = rsOrder("user_ID")
p_address_id = rsOrder("address_ID")
total_price   = rsOrder("total_price")

rsOrder.close
set rsOrder = nothing

set rsUser = server.createobject("ADODB.recordset")
rsUser.cursortype = 3

strSQL = "SELECT user_email, user_telephone, user_street, user_postcode, user_city, user_province, user_address.user_firstname, user_address.user_lastname FROM users INNER JOIN user_address ON users.user_ID = user_address.user_ID WHERE user_address_id = " & p_address_id & " AND users.user_ID = " & p_user_ID
rsUser.open strSQL, adoCon

if not rsUser.eof then
	u_email = rsUser("user_email")
	u_zip   = rsUser("user_postcode")
	u_firstname = rsUser("user_firstname")
	u_lastname  = rsUser("user_lastname")
	u_address   = rsUser("user_street")
	u_city      = rsUser("user_city")
	u_state     = rsUser("user_province")
	u_telephone = rsUser("user_telephone")
end if

rsUser.close
set rsUser = nothing
%>
<div class="box_large">
  <h2>&nbsp;<%=strPaypalPayment%></h2>
  <p>
  	<%=strSelectedPaypalPayment%>
  </p>
  <br />
  <form action="https://www.paypal.com/cgi-bin/webscr" method="post" name="frmPay" id="frmPay">
  	<input type="hidden" name="cmd" value="_ext-enter" />
	<input type="hidden" name="redirect_cmd" value="_xclick" />
	<input type="hidden" name="business" value="<%=arrOptions(1)%>" />
	<input type="hidden" name="item_name" value="<%=strShopName & " OID-" & order_id %>" />
	<input type="hidden" name="currency_code" value="<%=arrOptions(0)%>" />
	<input type="hidden" name="amount" value="<%=total_price%>" />
	<input type="hidden" name="image" value="languages/<%=session("language")%>/images/pay_paypal.jpg" />
	<input type="hidden" name="no_note" value="1" />
	<input type="hidden" name="return" value="<%=strShopLink%>/default.asp?mod=confirm&amp;type=payment" />
	<input type="hidden" name="cancel_return" value="<%=strShopLink%>/default.asp?mod=cancel&amp;type=payment" />
	<input type="hidden" name="email" value="<%=u_email%>" />
	<input type="hidden" name="first_name" value="<%=u_firstname%>" />
	<input type="hidden" name="last_name" value="<%=u_lastname%>" />
	<input type="hidden" name="address1" value="<%=u_address%>" />
	<input type="hidden" name="city" value="<%=u_city%>" />
	<input type="hidden" name="state" value="<%=u_state%>" />
	<input type="hidden" name="zip" value="<%=u_zip%>" />
	<input type="hidden" name="day_phone_a" value="<%=u_telephone%>" />
  <p align="center">
  	<a href="javascript:submitForm();"><img src="languages/<%=session("language")%>/images/pay_paypal.jpg" alt="<%=strPaypalpayment%>" title="<%=strPaypalPayment%>" /></a>	
  </p>
  </form>
</div>