<%
sql_webuserinfo=				"SELECT faceid,adminname,adminid,name,email,qqid,msnid,homepage FROM " &table_supervisor& " WHERE adminname='{0}'"
sql_webuserinfo_status=			"SELECT status FROM " &table_config& " WHERE adminid={0}"
sql_webuserinfo_count_view=		"SELECT [view] FROM " &table_stats& " WHERE adminid={0}"
sql_webuserinfo_count_words=	"SELECT COUNT(1) FROM " &table_main& " WHERE adminid={0}"
sql_webuserinfo_count_reply=	"SELECT COUNT(articleid) FROM " &table_main& " INNER JOIN " &table_reply& " ON (" &table_main& ".id=" &table_reply& ".articleid AND adminid={0})"
sql_webuserinfo_count_ipv4config="SELECT COUNT(1) FROM " &table_ipv4config& " WHERE adminid={0}"
sql_webuserinfo_count_ipv6config="SELECT COUNT(1) FROM " &table_ipv6config& " WHERE adminid={0}"
sql_webuserinfo_count_filterconfig="SELECT COUNT(filterid) FROM " &table_filterconfig& " WHERE adminid={0}"
sql_webuserinfo_reginfo=		"SELECT regdate,lastlogin FROM " &table_supervisor& " WHERE adminid={0}"
sql_webuserinfo_logininfo=		"SELECT login,loginfailed FROM " &table_stats& " WHERE adminid={0}"
%>
