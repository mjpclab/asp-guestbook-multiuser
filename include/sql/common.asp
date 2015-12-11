<%
sql_common_isbanipv4=	"SELECT TOP 1 1 FROM " &table_ipv4config& " WHERE cfgtype={0} And '{1}'>=ipfrom And '{1}'<=ipto AND adminid={2}"
sql_common_isbanipv6=	"SELECT TOP 1 1 FROM " &table_ipv6config& " WHERE cfgtype={0} And '{1}'>=ipfrom And '{1}'<=ipto AND adminid={2}"
sql_common_getstat=		"SELECT startdate,[{0}] FROM " &table_stats& " WHERE adminid={1}"
sql_common_initstat=	"INSERT INTO " &table_stats& "(startdate,adminid) VALUES('{0}',{1})"
sql_common_updatetime=	"UPDATE " &table_stats& " SET startdate='{0}' WHERE adminid={1}"
sql_common_addstat=		"UPDATE " &table_stats& " SET [{0}]=[{0}]+1 WHERE adminid={1}"

sql_common_checkuser=	"SELECT adminid FROM " &table_supervisor& " WHERE adminname='{0}'"
if IsAccess then
	'sql_common_dividepage=	"SELECT TOP {0} * FROM (" & _
	'							"SELECT * FROM (" & _
	'								"SELECT TOP {0} * FROM (" & _
	'									"SELECT TOP {1} {2} ORDER BY {3}" & _
	'								") ORDER BY {4}" & _
	'							") ORDER BY {3}" & _
	'						")"
	sql_common_dividepage=	"{6} {7} {8} IN(" & _
								"SELECT TOP {0} {8} FROM (" & _
									"SELECT {8} FROM (" & _
										"SELECT TOP {0} {2} FROM (" & _
											"SELECT TOP {1} {2} {3} ORDER BY {4}" & _
										") ORDER BY {5}" & _
									") ORDER BY {4}" & _
								")" & _
							") ORDER BY {4}"
elseif IsSqlServer then
	'sql_common_dividepage=	"SELECT * FROM (" & _
	'							"SELECT TOP {0} * FROM (" & _
	'								"SELECT TOP {1} {2} ORDER BY {3}" & _
	'							") AS __temp1 ORDER BY {4}" & _
	'						") AS __temp2 ORDER BY {3}"
	sql_common_dividepage=	"{6} {7} {8} IN(" & _
								"SELECT TOP {0} {8} FROM (" & _
									"SELECT TOP {1} {2} {3} ORDER BY {4}" & _
								") AS __temp1 ORDER BY {5}" & _
							") ORDER BY {4}"
end if
%>
