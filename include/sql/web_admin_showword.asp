<%
if IsAccess then
	sql_web_showword="SELECT supervisor.adminname,main.*,reply.* FROM ((" &table_supervisor& " supervisor INNER JOIN " &table_main& " main ON supervisor.adminid=main.adminid) LEFT JOIN " &table_reply& " reply ON main.id=reply.articleid) WHERE id="
elseif IsSqlServer then
	sql_web_showword="SELECT supervisor.adminname,main.*,reply.* FROM " &table_supervisor& " supervisor INNER JOIN " &table_main& " main ON supervisor.adminid=main.adminid LEFT JOIN " &table_reply& " reply ON main.id=reply.articleid WHERE id="
end if
%>
