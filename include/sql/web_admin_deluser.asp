<%
sql_webdeluser_filterconfig=	"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_ipv4config=		"DELETE FROM " &table_ipv4config&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_ipv6config=		"DELETE FROM " &table_ipv6config&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_stats=			"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_stats_clientinfo="DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_reply=			"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0})))"
sql_webdeluser_main=			"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_config=			"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_supervisor=		"DELETE FROM " &table_supervisor&		" WHERE adminname IN ({0})"
%>
