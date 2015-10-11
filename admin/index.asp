<!-- #include file="common.asp" -->
<% if CheckLogin() = 0 then response.redirect("login.asp?p=login") %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>

    <script language="javascript">
  var img_url = "";
  function swapImage(img_name, act){
    doc = document.getElementById("img_" + img_name);
	if(act == 1){
		doc.src = "images/panel/" + img_name + "_on.gif";
	} else {
		doc.src = "images/panel/" + img_name + "_off.gif";
	}
  }
    </script>

    <title>Admin Control Panel Index</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="includes/style.css" rel="stylesheet" type="text/css">
</head>
<body>
    <br>
    <br>
    <center>
        <table border="0" width="500" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
            <tr>
                <td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
                    <img border="0" src="images/gripblue.gif" width="15" height="19">
                </td>
                <td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle">
                    <b><font color="#FFFFFF" size="2" face="Verdana">Admin Panel</font></b>
                </td>
                <td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right">
                    <a href="aboutpanel.asp">
                        <img border="0" src="images/toolbar_info.gif" width="18" height="19"></a><img border="0"
                            src="images/downlevel.gif" width="25" height="19">
                </td>
            </tr>
        </table>
        <table border="1" width="500" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
            <tr>
                <td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC"
                    bordercolordark="#6699CC">
                    <table border="0" width="100%" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='config.asp?p=1';" onmouseover="javascript:swapImage('storesetup',1);"
                                    onmouseout="javascript:swapImage('storesetup',2);">
                                    <img src="images/panel/storesetup_off.gif" id="img_storesetup" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='config.asp?p=2';" onmouseover="javascript:swapImage('mailsetup',1);"
                                    onmouseout="javascript:swapImage('mailsetup',2);">
                                    <img src="images/panel/mailsetup_off.gif" id="img_mailsetup" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='config.asp?p=8';" onmouseover="javascript:swapImage('languages',1);"
                                    onmouseout="javascript:swapImage('languages',2);">
                                    <img src="images/panel/languages_off.gif" id="img_languages" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='config.asp?p=13';" onmouseover="javascript:swapImage('statistics',1);"
                                    onmouseout="javascript:swapImage('statistics',2);">
                                    <img src="images/panel/statistics_off.gif" id="img_statistics" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='catalog.asp?p=15';" onmouseover="javascript:swapImage('newsletter',1);"
                                    onmouseout="javascript:swapImage('newsletter',2);">
                                    <img src="images/panel/newsletter_off.gif" id="img_newsletter" width="66" height="61"><br>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='catalog.asp?p=3';" onmouseover="javascript:swapImage('categories',1);"
                                    onmouseout="javascript:swapImage('categories',2);">
                                    <img src="images/panel/categories_off.gif" id="img_categories" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='catalog.asp?p=4';" onmouseover="javascript:swapImage('products',1);"
                                    onmouseout="javascript:swapImage('products',2);">
                                    <img src="images/panel/products_off.gif" id="img_products" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='catalog.asp?p=5';" onmouseover="javascript:swapImage('reviews',1);"
                                    onmouseout="javascript:swapImage('reviews',2);">
                                    <img src="images/panel/reviews_off.gif" id="img_reviews" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='config.asp?p=14';" onmouseover="javascript:swapImage('custompage',1);"
                                    onmouseout="javascript:swapImage('custompage',2);">
                                    <img src="images/panel/custompage_off.gif" id="img_custompage" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='catalog.asp?p=16';" onmouseover="javascript:swapImage('manufactures',1);"
                                    onmouseout="javascript:swapImage('manufactures',2);">
                                    <img src="images/panel/manufactures_off.gif" id="img_manufactures" width="66" height="61"><br>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='customers.asp?p=6';" onmouseover="javascript:swapImage('customers',1);"
                                    onmouseout="javascript:swapImage('customers',2);">
                                    <img src="images/panel/customers_off.gif" id="img_customers" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='customers.asp?p=7';" onmouseover="javascript:swapImage('orders',1);"
                                    onmouseout="javascript:swapImage('orders',2);">
                                    <img src="images/panel/orders_off.gif" id="img_orders" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='customers.asp?p=9';" onmouseover="javascript:swapImage('orderstatus',1);"
                                    onmouseout="javascript:swapImage('orderstatus',2);">
                                    <img src="images/panel/orderstatus_off.gif" id="img_orderstatus" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='catalog.asp?p=22';" onmouseover="javascript:swapImage('stock',1);"
                                    onmouseout="javascript:swapImage('stock',2);">
                                    <img src="images/panel/stock_off.gif" id="img_stock" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='config.asp?p=19';" onmouseover="javascript:swapImage('news',1);"
                                    onmouseout="javascript:swapImage('news',2);">
                                    <img src="images/panel/news_off.gif" id="img_news" width="66" height="61"><br>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='config.asp?p=20';" onmouseover="javascript:swapImage('deliveries',1);"
                                    onmouseout="javascript:swapImage('deliveries',2);">
                                    <img src="images/panel/deliveries_off.gif" id="img_deliveries" width="66" height="61"><br>
                                </div>
                            </td>
                            <td valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='config.asp?p=21';" onmouseover="javascript:swapImage('payments',1);"
                                    onmouseout="javascript:swapImage('payments',2);">
                                    <img src="images/panel/payments_off.gif" id="img_payments" width="66" height="61"><br>
                                </div>
                            </td>
                            <td valign="Top" align="Center" bgcolor="#FFFFFF">
                                &nbsp;
                            </td>
                            <td valign="Top" align="Center" bgcolor="#FFFFFF">
                                &nbsp;
                            </td>
                            <td valign="Top" align="Center" bgcolor="#FFFFFF">
                                &nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='login.asp?p=10';" onmouseover="javascript:swapImage('passadmin',1);"
                                    onmouseout="javascript:swapImage('passadmin',2);">
                                    <img src="images/panel/passadmin_off.gif" id="img_passadmin" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='login.asp?p=12';" onmouseover="javascript:swapImage('editadmin',1);"
                                    onmouseout="javascript:swapImage('editadmin',2);">
                                    <img src="images/panel/editadmin_off.gif" id="img_editadmin" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='login.asp?p=11';" onmouseover="javascript:swapImage('addadmin',1);"
                                    onmouseout="javascript:swapImage('addadmin',2);">
                                    <img src="images/panel/addadmin_off.gif" id="img_addadmin" width="66" height="61"><br>
                                </div>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <img src="images/panel/empty.gif" width="66" height="61" /><br>
                            </td>
                            <td width="20%" valign="Top" align="Center" bgcolor="#FFFFFF">
                                <div class="body_image" onclick="document.location='../default.asp';" onmouseover="javascript:swapImage('storehome',1);"
                                    onmouseout="javascript:swapImage('storehome',2);">
                                    <img src="images/panel/storehome_off.gif" id="img_storehome" width="66" height="61"><br>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </center>
