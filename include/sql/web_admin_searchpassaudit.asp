<%
if IsAccess then
	sql_web_searchpassaudit=		"UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE (guestflag mod 32)\16<>0 AND id="
elseif IsSqlServer then
	sql_web_searchpassaudit=		"UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE guestflag & 16<>0 AND id="
end if
%>