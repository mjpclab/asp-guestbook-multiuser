<%
sql_showword=		"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id={0} AND adminid={1}"
%>
