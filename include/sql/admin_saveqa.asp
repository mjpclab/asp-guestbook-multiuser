<%
sql_adminsaveqa_select="SELECT adminpass FROM " &table_supervisor& " WHERE adminid={0}"
if IsAccess then
	sql_adminsaveqa_update="UPDATE " &table_supervisor& " SET question='{0}',[key]='{1}' WHERE adminid={2}"
elseif IsSqlServer then
	sql_adminsaveqa_update="UPDATE " &table_supervisor& " SET question=N'{0}',[key]='{1}' WHERE adminid={2}"
end if
%>
