<%
sql_admindelreply_delete="DELETE FROM " &table_reply& " WHERE articleid IN (SELECT id FROM " &table_main& " WHERE id={0} AND adminid={1})"
if IsAccess then
	sql_admindelreply_update="UPDATE " &table_main& " SET replied=replied-(replied mod 2) WHERE id={0} AND adminid={1}"
elseif IsSqlServer then
	sql_admindelreply_update="UPDATE " &table_main& " SET replied=replied-(replied & 1) WHERE id={0} AND adminid={1}"
end if
%>
