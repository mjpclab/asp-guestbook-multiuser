<%
sql_websearch_condition_init=		"parent_id<=0"
sql_websearch_condition_adminname=	" AND adminname LIKE '%{0}%'"
sql_websearch_condition_name=		" AND main.name LIKE '%{0}%'"
sql_websearch_condition_title=		" AND title LIKE '%{0}%'"
sql_websearch_condition_article=	" AND article LIKE '%{0}%'"
sql_websearch_condition_email=		" AND main.email LIKE '%{0}%'"
sql_websearch_condition_qqid=		" AND main.qqid LIKE '%{0}%'"
sql_websearch_condition_msnid=		" AND main.msnid LIKE '%{0}%'"
sql_websearch_condition_homepage=	" AND main.homepage LIKE '%{0}%'"
sql_websearch_condition_ipv4addr=	" AND ipv4addr LIKE '%{0}%'"
sql_websearch_condition_originalipv4=" AND originalipv4 LIKE '%{0}%'"
sql_websearch_condition_ipv6addr=	" AND ipv6addr LIKE '%{0}%'"
sql_websearch_condition_originalipv6=" AND originalipv6 LIKE '%{0}%'"
sql_websearch_condition_reply=		" AND id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%')"
if IsAccess then
	sql_websearch_count_inner=		"SELECT DISTINCT root_id FROM " &table_main& " main LEFT JOIN " &table_supervisor& " supervisor ON main.adminid=supervisor.adminid WHERE "
	sql_websearch_count=			"SELECT Count(*) FROM ({0})"
	sql_websearch_full_inner=		"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_websearch_full=				"SELECT supervisor.adminname,main.*,reply.* FROM ((" &table_supervisor& " AS supervisor INNER JOIN " &table_main& " AS main ON supervisor.adminid=main.adminid) LEFT JOIN " &table_reply& " AS reply ON main.id=reply.articleid) WHERE main.id IN({0})"
elseif IsSqlServer then
	sql_websearch_count_inner=		"SELECT DISTINCT root_id FROM " &table_main& " main LEFT JOIN " &table_supervisor& " supervisor ON main.adminid=supervisor.adminid WHERE "
	sql_websearch_count=			"SELECT Count(*) FROM ({0}) AS __temp1"
	sql_websearch_full_inner=		"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_websearch_full=				"SELECT supervisor.adminname,main.*,reply.* FROM " &table_supervisor& " AS supervisor INNER JOIN " &table_main& " AS main ON supervisor.adminid=main.adminid LEFT JOIN " &table_reply& " AS reply ON main.id=reply.articleid WHERE main.id IN({0})"
end if
sql_websearch_admininfo=			"SELECT name,faceid,faceurl,email,qqid,msnid,homepage FROM " &table_supervisor& " WHERE adminid={0}"
%>