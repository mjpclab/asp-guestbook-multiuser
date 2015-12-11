<%
sql_showword_count=	"SELECT COUNT(1) FROM " &table_main& " WHERE adminid={0}"
sql_showword=		"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id={0} AND adminid={1}"
%>
