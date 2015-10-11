<!-- #include file="common.lng" -->
<% if module="confirm" then %>
	<!-- #include file="confirm.lng" -->
<% end if %>

<% if module="cat" OR module="cart" then %>
	<!-- #include file="category.lng" -->
<% end if %>

<% if module="product" then %>
  <!-- #include file="product.lng" -->
<% end if %>

<% if module = "myaccount" then %>
  <!-- #include file="account.lng" -->
  <!-- #include file="mail_body.lng" -->
<% end if %>

<% if module="search" then %>
	<!-- #include file="search.lng" -->
<% end if %>

<% if module="checkout" then %>
	<!-- #include file="checkout.lng" -->
	<!-- #include file="mail_body.lng" -->
<% end if %>

<% if module = "newsletter" then %>
	<!-- #include file="newsletter.lng" -->
<% end if %>

<% if module = "contact" then %>
	<!-- #include file="contact.lng" -->
<% end if %>

<% if module = "news" then %>
	<!-- #include file="news.lng" -->
<% end if %>

<% if module = "cancel" then %>
	<!-- #include file="cancel.lng" -->
<% end if %>