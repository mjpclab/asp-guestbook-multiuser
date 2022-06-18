<%
sql_submitunreg_checkuser=				"SELECT adminid,adminpass FROM " &table_supervisor& " WHERE adminname='{0}' AND adminname<>'{1}'"
sql_submitunreg_delete_filterconfig=	"DELETE FROM " &table_filterconfig& " WHERE adminid={0}"
sql_submitunreg_delete_ipv4config=		"DELETE FROM " &table_ipv4config& " WHERE adminid={0}"
sql_submitunreg_delete_ipv6config=		"DELETE FROM " &table_ipv6config& " WHERE adminid={0}"
sql_submitunreg_delete_floodconfig=		"DELETE FROM " &table_floodconfig& " WHERE adminid={0}"
sql_submitunreg_delete_stats=			"DELETE FROM " &table_stats& " WHERE adminid={0}"
sql_submitunreg_delete_stats_clientinfo="DELETE FROM " &table_stats_clientinfo& " WHERE adminid={0}"
sql_submitunreg_delete_reply=			"DELETE FROM " &table_reply& " WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid={0})"
sql_submitunreg_delete_main=			"DELETE FROM " &table_main& " WHERE adminid={0}"
sql_submitunreg_delete_config=			"DELETE FROM " &table_config& " WHERE adminid={0}"
sql_submitunreg_delete_supervisor=		"DELETE FROM " &table_supervisor& " WHERE adminid={0}"
%>
