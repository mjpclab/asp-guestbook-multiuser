<%
if IsAccess then
	sql_webdoadvdel_beforedate_main="DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated<=#{0}#"
	sql_webdoadvdel_afterdate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>=#{0}#"
elseif IsSqlServer then
	sql_webdoadvdel_beforedate_main="DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated<='{0}'"
	sql_webdoadvdel_afterdate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>='{0}'"
end if
sql_webdoadvdel_firstn_main=				"DELETE FROM " &table_main& " WHERE id IN (SELECT top {0} id FROM " &table_main& " WHERE parent_id<=0 ORDER BY lastupdated ASC)"
sql_webdoadvdel_lastn_main=					"DELETE FROM " &table_main& " WHERE id IN (SELECT top {0} id FROM " &table_main& " WHERE parent_id<=0 ORDER BY lastupdated DESC)"
if IsAccess then
	sql_webdoadvdel_name_main=					"DELETE FROM " &table_main& " WHERE name LIKE '%{0}%'"
	sql_webdoadvdel_title_main=					"DELETE FROM " &table_main& " WHERE title LIKE '%{0}%'"
	sql_webdoadvdel_article_main=				"DELETE FROM " &table_main& " WHERE article LIKE '%{0}%'"
	sql_webdoadvdel_reply_main=					"DELETE FROM " &table_main& " WHERE id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%')"
	sql_webdoadvdel_reply_reply=				"DELETE FROM " &table_reply& " WHERE reinfo LIKE '%{0}%'"
	sql_webdoadvdel_reply_unsetreply=			"UPDATE " &table_main& " SET replied=0 WHERE id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%')"
elseif IsSqlServer then
	sql_webdoadvdel_name_main=					"DELETE FROM " &table_main& " WHERE name LIKE N'%{0}%'"
	sql_webdoadvdel_title_main=					"DELETE FROM " &table_main& " WHERE title LIKE N'%{0}%'"
	sql_webdoadvdel_article_main=				"DELETE FROM " &table_main& " WHERE article LIKE N'%{0}%'"
	sql_webdoadvdel_reply_main=					"DELETE FROM " &table_main& " WHERE id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE N'%{0}%')"
	sql_webdoadvdel_reply_reply=				"DELETE FROM " &table_reply& " WHERE reinfo LIKE N'%{0}%'"
	sql_webdoadvdel_reply_unsetreply=			"UPDATE " &table_main& " SET replied=0 WHERE id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE N'%{0}%')"
end if
sql_webdoadvdel_main=							"DELETE FROM " &table_main
sql_webdoadvdel_unsetreply=						"UPDATE " &table_main& " SET replied=0"
sql_webdoadvdel_reply=							"DELETE FROM " &table_reply
if IsAccess then
	sql_webdoadvdel_deldeclare=					"UPDATE " &table_supervisor& " SET [declare]='' WHERE [declare] LIKE '%{0}%'"
	sql_webdoadvdel_cleardeclare=					"UPDATE " &table_supervisor& " SET [declare]='' WHERE Len([declare])>0"
elseif IsSqlServer then
	sql_webdoadvdel_deldeclare=					"UPDATE " &table_supervisor& " SET [declare]='' WHERE [declare] LIKE N'%{0}%'"
	sql_webdoadvdel_cleardeclare=					"UPDATE " &table_supervisor& " SET [declare]='' WHERE [declare] LIKE N'_%'"
end if
if IsAccess then
	sql_webdoadvdel_regdate_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_ipv4config=			"DELETE FROM " &table_ipv4config&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_ipv6config=			"DELETE FROM " &table_ipv6config&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_stats=				"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_reply=				"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#))"
	sql_webdoadvdel_regdate_main=				"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_config=				"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_supervisor=			"DELETE FROM " &table_supervisor&		" WHERE regdate<=#{0}#"

	sql_webdoadvdel_lastlogin_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_ipv4config=		"DELETE FROM " &table_ipv4config&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_ipv6config=		"DELETE FROM " &table_ipv6config&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_stats=			"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_reply=			"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#))"
	sql_webdoadvdel_lastlogin_main=				"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_config=			"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_supervisor=		"DELETE FROM " &table_supervisor&		" WHERE lastlogin<=#{0}#"

	sql_webdoadvdel_logdate_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#) AND adminid<>{1}"
	sql_webdoadvdel_logdate_ipv4config=			"DELETE FROM " &table_ipv4config&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#) AND adminid<>{1}"
	sql_webdoadvdel_logdate_ipv6config=			"DELETE FROM " &table_ipv6config&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#) AND adminid<>{1}"
	sql_webdoadvdel_logdate_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#) AND adminid<>{1}"
	sql_webdoadvdel_logdate_stats=				"DELETE FROM " &table_stats&			" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#)"
	sql_webdoadvdel_logdate_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo& " WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#)"
	sql_webdoadvdel_logdate_reply=				"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#))"
	sql_webdoadvdel_logdate_main=				"DELETE FROM " &table_main&				" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#)"
	sql_webdoadvdel_logdate_config=				"DELETE FROM " &table_config&			" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#) AND adminid<>{1}"
	sql_webdoadvdel_logdate_supervisor=			"DELETE FROM " &table_supervisor&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#)"
elseif IsSqlServer then
	sql_webdoadvdel_regdate_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_ipv4config=			"DELETE FROM " &table_ipv4config&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_ipv6config=			"DELETE FROM " &table_ipv6config&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_stats=				"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_reply=				"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}'))"
	sql_webdoadvdel_regdate_main=				"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_config=				"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_supervisor=			"DELETE FROM " &table_supervisor&		" WHERE regdate<='{0}'"

	sql_webdoadvdel_lastlogin_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_ipv4config=		"DELETE FROM " &table_ipv4config&	    " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_ipv6config=		"DELETE FROM " &table_ipv6config&	    " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_stats=			"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_reply=			"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}'))"
	sql_webdoadvdel_lastlogin_main=				"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_config=			"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_supervisor=		"DELETE FROM " &table_supervisor&		" WHERE lastlogin<='{0}'"

	sql_webdoadvdel_logdate_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}') AND adminid<>{1}"
	sql_webdoadvdel_logdate_ipv4config=			"DELETE FROM " &table_ipv4config&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}') AND adminid<>{1}"
	sql_webdoadvdel_logdate_ipv6config=			"DELETE FROM " &table_ipv6config&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}') AND adminid<>{1}"
	sql_webdoadvdel_logdate_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}') AND adminid<>{1}"
	sql_webdoadvdel_logdate_stats=				"DELETE FROM " &table_stats&			" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}')"
	sql_webdoadvdel_logdate_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo& " WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}')"
	sql_webdoadvdel_logdate_reply=				"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}'))"
	sql_webdoadvdel_logdate_main=				"DELETE FROM " &table_main&				" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}')"
	sql_webdoadvdel_logdate_config=				"DELETE FROM " &table_config&			" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}') AND adminid<>{1}"
	sql_webdoadvdel_logdate_supervisor=			"DELETE FROM " &table_supervisor&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}')"
end if
sql_webdoadvdel_neverlogin_filterconfig=	"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate) AND adminid<>{0}"
sql_webdoadvdel_neverlogin_ipv4config=		"DELETE FROM " &table_ipv4config&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate) AND adminid<>{0}"
sql_webdoadvdel_neverlogin_ipv6config=		"DELETE FROM " &table_ipv6config&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate) AND adminid<>{0}"
sql_webdoadvdel_neverlogin_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate) AND adminid<>{0}"
sql_webdoadvdel_neverlogin_stats=			"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate)"
sql_webdoadvdel_neverlogin_stats_clientinfo="DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate)"
sql_webdoadvdel_neverlogin_reply=			"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate))"
sql_webdoadvdel_neverlogin_main=			"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate)"
sql_webdoadvdel_neverlogin_config=			"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate) AND adminid<>{0}"
sql_webdoadvdel_neverlogin_supervisor=		"DELETE FROM " &table_supervisor&		" WHERE lastlogin=regdate"

sql_webdoadvdel_all_filterconfig=	"DELETE FROM " &table_filterconfig& " WHERE adminid<>{0}"
sql_webdoadvdel_all_ipv4config=		"DELETE FROM " &table_ipv4config& " WHERE adminid<>{0}"
sql_webdoadvdel_all_ipv6config=		"DELETE FROM " &table_ipv6config& " WHERE adminid<>{0}"
sql_webdoadvdel_all_floodconfig=	"DELETE FROM " &table_floodconfig& " WHERE adminid<>{0}"
sql_webdoadvdel_all_stats=			"DELETE FROM " &table_stats
sql_webdoadvdel_all_stats_clientinfo="DELETE FROM " &table_stats_clientinfo
sql_webdoadvdel_all_reply=			"DELETE FROM " &table_reply
sql_webdoadvdel_all_main=			"DELETE FROM " &table_main
sql_webdoadvdel_all_config=			"DELETE FROM " &table_config& " WHERE adminid<>{0}"
sql_webdoadvdel_all_supervisor=		"DELETE FROM " &table_supervisor

sql_webdoadvdel_clearfragment_main=	"DELETE FROM " &table_main& " WHERE parent_id>0 AND parent_id NOT IN (SELECT DISTINCT id FROM " &table_main& " WHERE parent_id<=0)"
sql_webdoadvdel_clearfragment_reply="DELETE FROM " &table_reply& " WHERE articleid NOT IN(SELECT id FROM " &table_main& ")"
if IsAccess then
	sql_webdoadvdel_adjustadminreply_flag="UPDATE " &table_main& " SET replied=replied-(replied MOD 2) WHERE parent_id<=0 AND id NOT IN(" & _
												"SELECT articleid FROM " &table_reply & _
											")"
	sql_webdoadvdel_adjustguestreply_flag="UPDATE " &table_main& " SET replied=replied-((replied MOD 4)\2)*2 WHERE parent_id<=0 AND id NOT IN(" & _
												"SELECT DISTINCT parent_id FROM " &table_main & _
											")"
elseif IsSqlServer then
	sql_webdoadvdel_adjustadminreply_flag="UPDATE " &table_main& " SET replied=replied-(replied&1) WHERE parent_id<=0 AND id NOT IN(" & _
												"SELECT articleid FROM " &table_reply & _
											")"
	sql_webdoadvdel_adjustguestreply_flag="UPDATE " &table_main& " SET replied=replied-(replied&2) WHERE parent_id<=0 AND id NOT IN(" & _
												"SELECT DISTINCT parent_id FROM " &table_main & _
											")"
end if
%>
