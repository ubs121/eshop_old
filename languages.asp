<%
select case session("language"):
 case "english":
%><!-- #include file="languages/english/index.inc.asp" --><%
case "mongol":
%><!-- #include file="languages/mongol/index.inc.asp" --><%
end select
%>
