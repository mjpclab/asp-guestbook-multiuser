<%
sql_index_words_count="SELECT COUNT(1) FROM " &table_main& " WHERE parent_id<=0 AND adminid={0}"
sql_index_words_query="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id<=0 AND adminid={0}"
%>
