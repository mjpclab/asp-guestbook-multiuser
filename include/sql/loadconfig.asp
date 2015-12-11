<%
sql_loadconfig_config=		"SELECT * FROM " &table_config& " WHERE adminid={0}"
sql_loadconfig_floodconfig=	"SELECT * FROM " &table_floodconfig& " WHERE adminid={0}"
sql_loadconfig_style=		"SELECT * FROM " &table_style& " WHERE styleid={0}"
sql_loadconfig_top1style=	"SELECT Top 1 * FROM " &table_style
%>
