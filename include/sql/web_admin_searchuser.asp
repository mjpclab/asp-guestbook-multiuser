<%
sql_websearchuser_condition="{0} LIKE '%{1}%'"
sql_websearchuser_words_count="SELECT COUNT(adminname) FROM " &table_supervisor& " WHERE "
sql_websearchuser_words_query="SELECT adminname,name,regdate,lastlogin FROM " &table_supervisor& " WHERE "
%>
