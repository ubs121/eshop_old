<%
if len(request.form()) > 0 then
	session("checkout_order_comment") = request.form("order_comment")
	session("checkout_payment_method") = request.form("paymentMethod")
end if

if len(session("checkout_payment_method")) = 0 then
	response.redirect("?mod=checkout")
end if
%>
<p><b><%=strDeliveryMethod%></b></p>
<%
set rsDelivery = server.createobject("ADODB.recordset")
rsDelivery.cursortype = 3

strSQL = "SELECT delivery_id, a, b, delivery_name FROM delivery WHERE lang_id = " & session("language_id") & " ORDER BY b ASC;"
rsDelivery.open strSQL, adoCon
%>
<form name="frmStep2" method="post" action="<%=strCurrFile%>?mod=checkout&amp;p=3">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="31" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="left" valign="top" class="content">
			<table width="400" border="0" cellspacing="0" cellpadding="0">
			<% do while not rsDelivery.eof %>
			<%
			if rsDelivery("a") = "1" then
				delivery_price = roundNumber(Replace(rsDelivery("b"), ".", strServerComma))
			else
				arrPrices = split(rsDelivery("b"), ";")
				arrConditions = split(rsDelivery("a"), ";")
				
				x = 0
				for x = 0 to ubound(arrConditions)
					if instr(arrConditions(x), ">=") > 0 then
						condition = csng(Replace(right(arrConditions(x), len(arrConditions(x)) - 2), ".", strServerComma))
						if csng(session("totalWeight")) >= condition then
							delivery_price = arrPrices(x) & " (" & arrConditions(x) & strWeightsign & ")"
						end if
					elseif instr(arrConditions(x), ">") > 0 then
						condition = csng(Replace(right(arrConditions(x), len(arrConditions(x)) - 1), ".", strServerComma))
						if csng(session("totalWeight")) > condition then
							delivery_price = arrPrices(x) & " (" & arrConditions(x) & strWeightsign & ")"
						end if
					else
						if csng(session("totalWeight")) < csng(right(arrConditions(x), len(arrConditions(x)) - 1)) then
							delivery_price = arrPrices(x) & " (" & arrConditions(x) & strWeightsign & ")"
						end if
					end if					
				next
			end if
			%>
                <tr>
                  <td width="20" class="content">
					<input type="radio" name="delivery_method" value="<%=rsDelivery("delivery_id")%>"<% if cint(session("checkout_delivery_method")) = cint(rsDelivery("delivery_id")) then %> checked="checked"<% end if %>></td>
                  <td width="50" class="content"><%=strCurrency%><%=delivery_price%></td>
                  <td width="340" class="content">&nbsp;<%=rsDelivery("delivery_name")%></td>
			   </tr>
			   <%
			   		rsDelivery.movenext
			   loop
			   %>
			</table>

		  </td>
        </tr>
      </table></td>
    </tr>
  </table>
<%
rsDelivery.close
set rsDelivery = nothing
%>
<br />
<p><b><%=strOrderComments%></b></p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="31" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td align="center" valign="middle" class="content"><textarea name="order_comment" cols="100" id="order_comment"><%=session("checkout_order_comment")%></textarea></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="50%" class="content"><a href="?mod=checkout"><img src="languages/<%=session("language")%>/images/button_back.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a></td>
    <td width="50%" align="right" class="content"><a href="javascript:document.frmStep2.submit();"><img src="languages/<%=session("language")%>/images/button_continue.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a></td>
  </tr>
</table>