<%
sql_websearchdel_reply=	"DELETE FROM " &table_reply& " WHERE articleid IN(SELECT id FROM " &table_main& " WHERE parent_id={0} OR id={0})"
sql_websearchdel_main=	"DELETE FROM " &table_main& " WHERE parent_id={0} OR id={0}"
%>
