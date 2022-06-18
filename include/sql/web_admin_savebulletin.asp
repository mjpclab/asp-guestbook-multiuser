<%
if IsAccess then
	sql_websavebulletin="UPDATE " &table_webmaster& " SET bulletinflag={0},[webbulletin]='{1}'"
elseif IsSqlServer then
	sql_websavebulletin="UPDATE " &table_webmaster& " SET bulletinflag={0},[webbulletin]=N'{1}'"
end if
%>
