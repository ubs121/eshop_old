<%
if len(request.form("Submit")) > 0 then
	updateConfig("shop_currency")
	updateConfig("shop_link")
	updateConfig("shop_name")
	updateConfig("shop_owner")
	updateConfig("shop_title")
	updateConfig("virtual_path")
	updateConfig("auto_update_lang")
	updateConfig("write_complete_url")
	updateConfig("show_language_module")
	updateConfig("round_number")
	updateConfig("decimal_sign")
	updateConfig("show_loadtime")
	updateConfig("show_search")
	updateConfig("show_newsletter")
	updateConfig("show_manufacturers")
	updateConfig("shop_offline")
	updateConfig("offline_message")
	updateConfig("upload_method")
	updateConfig("show_news")
	updateConfig("show_promotions")
	updateConfig("products_per_page")
	updateConfig("comma_separator")
	updateConfig("showStock")
	updateConfig("show_mcart")
	updateConfig("show_lastproducts")
	updateconfig("show_popular")
	updateConfig("weightSign")
	updateconfig("stock_autoupdate")
	updateConfig("convertChars")
	response.redirect("?p=" & request.querystring("p"))
end if

set rsConfig = server.createobject("ADODB.recordset")
rsConfig.cursortype = 3

strSQL = "SELECT config_name, config_value FROM config"
rsConfig.open strSQL, adoCon
%>
<form action="" method="post" name="frmMyStore" id="frmMyStore">
<center><br><br>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Store Setup</font></b></td>
<td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right"><a href="index.asp"><img border="0" src="images/toolbar_home.gif" width="18" height="19" alt="Main Admin Page"></a><img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

<tr><td>

  <table width="500" align="center" cellpadding="2" cellspacing="2">
              <tr> 
                <td colspan="2" align="center"><input type="button" value=">> Edit template <<" name="btnEditTemplate" onclick="document.location='?p=17';" /></td>
              </tr>
              <tr> 
                <td colspan="2" height="10"></td>
              </tr>
              <tr> 
                <td>Numbers behind decimal</td>
                <td><input name="round_number" type="text" id="round_number" value="<%=getConfig("round_number")%>" size="10"></td>
              </tr>
              <tr> 
                <td>Decimal sign</td>
                <td><input name="decimal_sign" type="text" id="decimal_sign" value="<%=getConfig("decimal_sign")%>" size="10"></td>
              </tr>
              <tr> 
                <td>Comma separator</td>
                <%
				comma_separator = getConfig("comma_separator")
				%>
                <td> <select name="comma_separator" id="comma_separator">
                    <option value="">None</option>
                    <option value=","<% if comma_separator = "," then %> selected="selected"<% end if %>>,</option>
                    <option value="&amp;nbsp;"<% if comma_separator = "&nbsp;" then %> selected="selected"<% end if %>>Space</option>
                  </select> </td>
              </tr>
              <tr>
                <td>Weight sign</td>
                <td><input name="weightSign" type="text" id="weightSign" value="<%=getConfig("weightSign")%>" size="10"></td>
              </tr>
              <tr> 
                <td>Currency</td>
                <%
				shop_currency = getConfig("shop_currency")
				shop_currency = Replace(shop_currency, "&", "&amp;")
				%>
                <td><input name="shop_currency" type="text" id="shop_currency" value="<%=shop_currency%>" size="10"></td>
              </tr>
              <tr> 
                <td colspan="2" height="10"></td>
              </tr>
              <tr> 
                <td>Show stock:</td>
                <td> <% showStock = getConfig("showStock") %> <select name="showStock" id="showStock">
                    <option value="1"<% if showStock = "1" then %> selected="selected"<% end if %>>Yes</option>
                    <option value="0"<% if showStock = "0" then %> selected="selected"<% end if %>>No</option>
                  </select></td>
              </tr>
              <tr>
                <td>Auto update stock: </td>
                <td>
				  <% autostock = getConfig("stock_autoupdate") %>
				  <select name="stock_autoupdate" id="stock_autoupdate">
				    <option value="1"<% if autostock = "1" then %> selected="selected"<% end if %>>Yes</option>
				    <option value="0"<% if autostock = "0" then %> selected="selected"<% end if %>>No</option>
			      </select>			    </td>
              </tr>
              <tr> 
                <td>Products per page</td>
                <td><input name="products_per_page" type="text" id="products_per_page" value="<%=getConfig("products_per_page")%>" size="10"></td>
              </tr>
              <tr> 
                <td colspan="2" height="10"></td>
              </tr>
              <tr>
                <td>Convert characters (only use this option when characters are shown incorrect) </td>
                <td>
				<% convertChars = cint(getConfig("convertChars")) %>
				<select name="convertChars" id="convertChars">
                  <option value="1"<% if convertChars = 1 then %> selected="selected"<% end if %>>Yes</option>
                  <option value="0"<% if convertChars = 0 then %> selected="selected"<% end if %>>No</option>
                </select>
                </td>
              </tr>
              <tr> 
                <td>Write complete url (including default.asp)</td>
                <td> <%
		write_complete_url = cint(GetConfig("write_complete_url"))
		%> <select name="write_complete_url" id="write_complete_url">
                    <option value="1"<% if write_complete_url = 1 then %> selected="selected"<% end if %>>yes</option>
                    <option value="0"<% if write_complete_url = 0 then %> selected="selected"<% end if %>>no</option>
                  </select></td>
              </tr>
              <tr> 
                <td colspan="2" height="10"></td>
              </tr>
              <tr> 
                <td>Show loadtime at end of page</td>
                <% show_loadtime = cint(getConfig("show_loadtime"))%>
                <td> <select name="show_loadtime" id="show_loadtime">
                    <option value="1"<% if show_loadtime = 1 then %> selected="selected"<% end if %>>Yes</option>
                    <option value="0"<% if show_loadtime = 0 then %> selected="selected"<% end if %>>No</option>
                  </select></td>
              </tr>
              <tr> 
                <td colspan="2"><strong>Leftmenu</strong></td>
              </tr>
              <tr> 
                <td>Show newsletter</td>
                <% show_newsletter = cint(getConfig("show_newsletter"))%>
                <td><select name="show_newsletter" id="show_newsletter">
                    <option value="1"<% if show_newsletter = 1 then %> selected="selected"<% end if %>>Yes</option>
                    <option value="0"<% if show_newsletter = 0 then %> selected="selected"<% end if %>>No</option>
                  </select></td>
              </tr>
              <tr> 
                <td>Show manufacturers-module</td>
                <% show_manufacturers = cint(getConfig("show_manufacturers")) %>
                <td><select name="show_manufacturers" id="show_manufacturers">
                    <option value="1"<% if show_manufacturers = 1 then %> selected="selected"<% end if %>>Yes</option>
                    <option value="0"<% if show_manufacturers = 0 then %> selected="selected"<% end if %>>No</option>
                  </select></td>
              </tr>
              <tr> 
                <td>Show search module</td>
                <% show_search = cint(getConfig("show_search")) %>
                <td><select name="show_search" id="show_search">
                    <option value="1"<% if show_search = 1 then %> selected="selected"<% end if %>>Yes</option>
                    <option value="0"<% if show_search = 0 then %> selected="selected"<% end if %>>No</option>
                  </select></td>
              </tr>
              <tr> 
                <td>Show language module</td>
                <td> <% show_lang_module = cint(GetConfig("show_language_module")) %> <select name="show_language_module" id="show_language_module">
                    <option value="1"<% if show_lang_module = 1 then %> selected="selected"<% end if %>>yes</option>
                    <option value="0"<% if show_lang_module = 0 then %> selected="selected"<% end if %>>no</option>
                  </select></td>
              </tr>
              <tr> 
                <td colspan="2"><strong>Rightmenu</strong></td>
              </tr>
              <tr> 
                <td>Show mini shoppingcart </td>
                <td> <% mcart = getConfig("show_mcart") %> <select name="show_mcart">
                    <option value="1"<% if mcart="1" then %> selected="selected"<% end if %>>yes</option>
                    <option value="0"<% if mcart="0" then %> selected="selected"<% end if %>>no</option>
                  </select></td>
              </tr>
              <tr> 
                <td>Show latest products </td>
                <td> <% last = getConfig("show_lastproducts") %> <select name="show_lastproducts">
                    <option value="1"<% if last="1" then %> selected="selected"<% end if %>>yes</option>
                    <option value="0"<% if last="0" then %> selected="selected"<% end if %>>no</option>
                  </select></td>
              </tr>
              <tr> 
                <td>Show popular products </td>
                <td> <% popular = getConfig("show_popular") %> <select name="show_popular">
                    <option value="1"<% if popular="1" then %> selected="selected"<% end if %>>yes</option>
                    <option value="0"<% if popular="0" then %> selected="selected"<% end if %>>no</option>
                  </select> </td>
              </tr>
              <tr> 
                <td colspan="2" height="10"></td>
              </tr>
              <tr> 
                <td>Show news</td>
                <td> <% show_news = cint( getConfig("show_news") ) %> <select name="show_news">
                    <option value="1"<% if show_news = 1 then %> selected="selected"<% end if %>>Yes</option>
                    <option value="0"<% if show_news = 0 then %> selected="selected"<% end if %>>No</option>
                  </select> </td>
              </tr>
              <tr> 
                <td>Show promotions</td>
                <td> <% show_promotions = cint( getConfig("show_promotions") ) %> <select name="show_promotions">
                    <option value="1"<% if show_promotions = 1 then %> selected="selected"<% end if %>>Yes</option>
                    <option value="0"<% if show_promotions = 0 then %> selected="selected"<% end if %>>No</option>
                  </select> </td>
              </tr>
              <tr> 
                <td colspan="2" height="10"></td>
              </tr>
              <tr> 
                <td>Upload method:</td>
                <td> <% upload_method = getConfig("upload_method") %> <select name="upload_method" id="upload_method">
                    <option>none</option>
                    <option value="dundas"<% if upload_method = "dundas" then %> selected="selected"<% end if %>>dundas</option>
                    <option value="pureasp"<% if upload_method = "pureasp" then %> selected="selected"<% end if %>>Pure 
                    Asp</option>
                    <option value="fileUp"<% if upload_method = "fileUp" then %> selected="selected"<% end if %>>File 
                    Up</option>
                    <option value="aspUpload"<% if upload_method = "aspUpload" then %> selected="selected"<% end if %>>aspUpload</option>
                  </select></td>
              </tr>
              <tr> 
                <td>Update method language-file:</td>
                <td><select name="auto_update_lang">
                    <option value="1"<% if getConfig("auto_update_lang") = "1" then %> selected="selected"<% end if %>>Auto-update</option>
                    <option value="2"<% if getConfig("auto_update_lang") = "2" then %> selected="selected"<% end if %>>Manually 
                    update</option>
                  </select> </td>
              </tr>
              <tr> 
                <td colspan="2" height="10"></td>
              </tr>
              <tr> 
                <td>Link to shop</td>
                <td><input name="shop_link" type="text" id="shop_link" value="<%=getConfig("shop_link")%>" size="40"></td>
              </tr>
              <tr> 
                <td>Virtual path</td>
                <td><input name="virtual_path" type="text" id="virtual_path" value="<%=strVirtualPath%>" size="40"></td>
              </tr>
              <tr> 
                <td>Shop name</td>
                <td><input name="shop_name" type="text" id="shop_name" value="<%=getConfig("shop_name")%>" size="40"></td>
              </tr>
              <tr> 
                <td>Shop owner</td>
                <td><input name="shop_owner" type="text" id="shop_owner" value="<%=getConfig("shop_owner")%>" size="40"></td>
              </tr>
              <tr> 
                <td colspan="2" height="10"></td>
              </tr>
              <tr> 
                <td>Page title</td>
                <td><input name="shop_title" type="text" id="shop_title" value="<%=getConfig("shop_title")%>" size="40"></td>
              </tr>
              <tr> 
                <td height="10" colspan="2" align="center"></td>
              </tr>
              <tr> 
                <td>Put shop offline</td>
                <%
				shop_offline = cint(getConfig("shop_offline"))
				%>
                <td> <select name="shop_offline" id="shop_offline">
                    <option value="1"<% if shop_offline = 1 then %> selected="selected"<% end if %>>Yes</option>
                    <option value="0"<% if shop_offline = 0 then %> selected="selected"<% end if %>>No</option>
                  </select></td>
              </tr>
              <tr> 
                <td align="left" valign="top">Offline message:</td>
                <td> <textarea name="offline_message" cols="37" rows="4" id="offline_message"><%=getConfig("offline_message")%></textarea></td>
              </tr>
              <tr> 
                <td colspan="2" align="center"><%=BuildSubmitter("submit","Update", request.querystring("p"))%> &nbsp;&nbsp; <input type="reset" name="Submit2" value="Reset"></td>
              </tr>
            </table>
</td></tr>
</table>
</center>
</form>
<%
rsConfig.close
%>