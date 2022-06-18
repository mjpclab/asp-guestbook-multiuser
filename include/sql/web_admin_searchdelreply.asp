<%
sql_websearchdelreply_reply=		"DELETE FROM " &table_reply& " WHERE articleid="
if IsAccess then
	sql_websearchdelreply_unsetreply="UPDATE " &table_main& " SET replied=replied-(replied mod 2) WHERE id="
elseif IsSqlServer then
	sql_websearchdelreply_unsetreply="UPDATE " &table_main& " SET replied=replied-(replied & 1) WHERE id="
end if
%>
