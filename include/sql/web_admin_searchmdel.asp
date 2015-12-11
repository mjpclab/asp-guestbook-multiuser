<%
sql_websearchmdel_reply="DELETE FROM " &table_reply& " WHERE articleid IN(SELECT id FROM " &table_main& " WHERE parent_id IN({0}) OR id IN({0}))"
sql_websearchmdel_main=	"DELETE FROM " &table_main& " WHERE parent_id IN({0}) OR id IN({0})"
%>
