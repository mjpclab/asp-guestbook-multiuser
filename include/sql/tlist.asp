<%
'sql_tlist_itemsperpage=	"SELECT itemsperpage FROM " &table_config& " WHERE adminid={0}"
sql_tlist_maxtime=		"SELECT TOP 1 lastupdated FROM " &table_main& " WHERE parent_id<=0 AND adminid={0} {1} ORDER BY lastupdated DESC"
if IsAccess then
	sql_tlist_mintime=	"SELECT TOP 1 lastupdated FROM (SELECT TOP {0} parent_id,lastupdated FROM " &table_main& " WHERE parent_id<=0 AND (guestflag mod 32)\16=0 AND (guestflag mod 64)\32=0 AND adminid={1} {2} ORDER BY parent_id ASC,lastupdated DESC) ORDER BY lastupdated ASC"
	sql_tlist=			"SELECT guestflag,title FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>=#{0}# AND lastupdated<=#{1}# AND adminid={2} {3} ORDER BY parent_id ASC,lastupdated DESC"
elseif IsSqlServer then
	sql_tlist_mintime=	"SELECT TOP 1 lastupdated FROM (SELECT TOP {0} parent_id,lastupdated FROM " &table_main& " WHERE parent_id<=0 AND guestflag & 48=0 AND adminid={1} {2} ORDER BY parent_id ASC,lastupdated DESC) AS __temp1 ORDER BY lastupdated ASC"
	sql_tlist=			"SELECT guestflag,title FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>='{0}' AND lastupdated<='{1}' AND adminid={2} {3} ORDER BY parent_id ASC,lastupdated DESC"
end if
%>
