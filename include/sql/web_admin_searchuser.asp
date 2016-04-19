<%
if IsAccess then
	sql_websearchuser_condition="{0} LIKE '%{1}%'"
elseif IsSqlServer then
	sql_websearchuser_condition="{0} LIKE N'%{1}%'"
end if
sql_websearchuser_words_count="SELECT COUNT(adminname) FROM " &table_supervisor& " WHERE "
sql_websearchuser_words_query="SELECT adminname,name,regdate,lastlogin FROM " &table_supervisor& " WHERE "
%>
