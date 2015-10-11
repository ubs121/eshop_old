<!-- #include file="common.asp" -->
<% if submodule = "login" or submodule = "register" or (module="checkout" and p="4") or submodule = "lostpass" then %>
<!-- #include file="includes/functions/sendmail.asp" -->
<!-- #include file="includes/functions/functions_hash1way.asp" -->
<% end if %>
<% if module="contact" then %>
<!-- #include file="includes/functions/sendmail.asp" -->
<% end if %>