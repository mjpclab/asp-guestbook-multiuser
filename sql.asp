<%
IsAccess=false
IsSqlServer=false

if dbtype>=1 and dbtype<=3 then
	IsAccess=true
    table_config			="[" & prefix & "config]"
    table_filterconfig		="[" & prefix & "filterconfig]"
    table_floodconfig		="[" & prefix & "floodconfig]"
    table_ipconfig			="[" & prefix & "ipconfig]"
    table_main				="[" & prefix & "main]"
    table_reply				="[" & prefix & "reply]"
    table_stats				="[" & prefix & "stats]"
    table_stats_clientinfo	="[" & prefix & "stats_clientinfo]"
    table_style				="[" & prefix & "style]"
    table_supervisor		="[" & prefix & "supervisor]"
    table_webmaster			="[" & prefix & "webmaster]"
elseif dbtype=10 then
	IsSqlServer=true
	if dbschema<>"" then
		schema="[" &dbschema& "]."
	else
		schema=""
	end if
    table_config			=schema & "[" & prefix & "config]"
    table_filterconfig		=schema & "[" & prefix & "filterconfig]"
    table_floodconfig		=schema & "[" & prefix & "floodconfig]"
    table_ipconfig			=schema & "[" & prefix & "ipconfig]"
    table_main				=schema & "[" & prefix & "main]"
    table_reply				=schema & "[" & prefix & "reply]"
    table_stats				=schema & "[" & prefix & "stats]"
    table_stats_clientinfo	=schema & "[" & prefix & "stats_clientinfo]"
    table_style				=schema & "[" & prefix & "style]"
    table_supervisor		=schema & "[" & prefix & "supervisor]"
    table_webmaster			=schema & "[" & prefix & "webmaster]"
end if

'PKs
sql_pk_main="id"
sql_pksearch_main="root_id"
sql_pk_supervisor="adminname"

'GLOBALs
if IsAccess then
	sql_global_noguestreply_flag=	"UPDATE " &table_main& " SET replied=replied-((replied MOD 4)\2)*2 WHERE parent_id<=0 AND id NOT IN({0}) AND id IN(" & _
										"SELECT parent_id FROM " &table_main& " WHERE id IN({0}) AND adminid={1}" & _
									") AND NOT EXISTS (" & _
										"SELECT 1 FROM " &table_main& " AS temp WHERE temp.parent_id=" &table_main& ".id AND temp.id NOT IN({0}) AND adminid={1}" & _
									")"
	sql_global_webnoguestreply_flag="UPDATE " &table_main& " SET replied=replied-((replied MOD 4)\2)*2 WHERE parent_id<=0 AND id NOT IN({0}) AND id IN(" & _
										"SELECT parent_id FROM " &table_main& " WHERE id IN({0})" & _
									") AND NOT EXISTS (" & _
										"SELECT 1 FROM " &table_main& " AS temp WHERE temp.parent_id=" &table_main& ".id AND temp.id NOT IN({0})" & _
									")"
elseif IsSqlServer then
	sql_global_noguestreply_flag=	"UPDATE " &table_main& " SET replied=replied-(replied&2) WHERE parent_id<=0 AND id NOT IN({0}) AND id IN(" & _
										"SELECT parent_id FROM " &table_main& " WHERE id IN({0}) AND adminid={1}" & _
									") AND NOT EXISTS (" & _
										"SELECT 1 FROM " &table_main& " AS temp WHERE temp.parent_id=" &table_main& ".id AND temp.id NOT IN({0}) AND adminid={1}" & _
									")"
	sql_global_webnoguestreply_flag="UPDATE " &table_main& " SET replied=replied-(replied&2) WHERE parent_id<=0 AND id NOT IN({0}) AND id IN(" & _
										"SELECT parent_id FROM " &table_main& " WHERE id IN({0})" & _
									") AND NOT EXISTS (" & _
										"SELECT 1 FROM " &table_main& " AS temp WHERE temp.parent_id=" &table_main& ".id AND temp.id NOT IN({0})" & _
									")"
end if

'common
sql_common_isbanip=		"SELECT TOP 1 1 FROM " &table_ipconfig& " WHERE ipconstatus={0} And '{1}'>=startip And '{1}'<=endip AND adminid={2}"
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

'common2
sql_common2_getadmininfo=	"SELECT name,faceid,faceurl,email,qqid,msnid,homepage FROM " &table_supervisor& " WHERE adminid={0}"
'sql_common2_reply=			"SELECT replydate,htmlflag,reinfo FROM " &table_reply& " WHERE articleid="
sql_common2_guestreply=		"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id={0}{1} AND adminid={2} ORDER BY id ASC"
sql_common2_replyinform=	"SELECT parent_id,email,guestflag,logdate,article FROM " &table_main& " WHERE id={0} AND adminid={1}"

'loadconfig/webconfig
sql_loadconfig_config=		"SELECT * FROM " &table_config& " WHERE adminid={0}"
sql_loadconfig_floodconfig=	"SELECT * FROM " &table_floodconfig& " WHERE adminid={0}"
sql_loadconfig_style=		"SELECT * FROM " &table_style& " WHERE styleid={0}"
sql_loadconfig_top1style=	"SELECT Top 1 * FROM " &table_style

'index/search/listword_guest/common2-outerword() condition
if IsAccess then
	sql_condition_hidehidden=	" AND (((guestflag mod 16)\8)=0 OR replied<>0)"
	sql_condition_hideaudit=	" AND (((guestflag mod 32)\16)=0 OR replied<>0)"
	sql_condition_hidewhisper=	" AND (((guestflag mod 64)\32)=0 OR replied<>0)"
elseif IsSqlServer then
	sql_condition_hidehidden=	" AND (guestflag & 8=0 OR replied<>0)"
	sql_condition_hideaudit=	" AND (guestflag & 16=0 OR replied<>0)"
	sql_condition_hidewhisper=	" AND (guestflag & 32=0 OR replied<>0)"
end if

'index
sql_index_words_count="SELECT COUNT(1) FROM " &table_main& " WHERE parent_id<=0 AND adminid={0}"
sql_index_words_query="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id<=0 AND adminid={0}"

'search
if IsAccess then
	sql_search_condition_reply=		"((guestflag mod 32)\16)=0 AND (((guestflag mod 64)\32)=0 or ((guestflag mod 128)\64)=0) AND id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%') AND adminid={1}"
	sql_search_condition_name=		"((guestflag mod 32)\16)=0 AND name LIKE '%{0}%' AND adminid={1}"
	sql_search_condition_title=		"((guestflag mod 32)\16)=0 AND ((guestflag mod 64)\32)=0 AND title LIKE '%{0}%' AND adminid={1}"
	sql_search_condition_article=	"((guestflag mod 16)\8)=0 AND ((guestflag mod 32)\16)=0 AND ((guestflag mod 64)\32)=0 AND article LIKE '%{0}%' AND adminid={1}"
	sql_search_count_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_search_count=				"SELECT Count(*) FROM ({0})"
	sql_search_full_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_search_full=				"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id IN({0})"
elseif IsSqlServer then
	sql_search_condition_reply=		"guestflag & 16=0 AND (guestflag & 32=0 OR guestflag & 64=0) AND id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%') AND adminid={1}"
	sql_search_condition_name=		"guestflag & 16=0 AND name LIKE '%{0}%' AND adminid={1}"
	sql_search_condition_title=		"guestflag & 48=0 AND title LIKE '%{0}%' AND adminid={1}"
	sql_search_condition_article=	"guestflag & 56=0 AND article LIKE '%{0}%' AND adminid={1}"
	sql_search_count_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_search_count=				"SELECT Count(*) FROM ({0}) AS __temp1"
	sql_search_full_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_search_full=				"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id IN({0})"
end if

'write
sql_write_filter=					"SELECT regexp,filtermode,replacestr FROM " &table_filterconfig& " WHERE adminid={0} ORDER BY filtersort ASC"
sql_write_flood_ids=				"SELECT DISTINCT id FROM " &table_main& " WHERE id="
sql_write_flood_idnull=				"SELECT DISTINCT id FROM " &table_main& " WHERE id IS NULL"
sql_write_flood_head=				"SELECT TOP 1 id FROM " &table_main& " WHERE ("
sql_write_flood_titlelike=			" title LIKE '%{0}%'"
sql_write_flood_titleequal=			" title='{0}'"
if IsAccess then
	sql_write_flood_articlelike=	" article LIKE '%{0}%'"
	sql_write_flood_articleequal=	" article LIKE '{0}'"
elseif IsSqlServer then
	sql_write_flood_articlelike=	" article LIKE CAST('%{0}%' AS NTEXT)"
	sql_write_flood_articleequal=	" article LIKE CAST('{0}' AS NTEXT)"
end if
sql_write_flood_wordstail=			") AND id IN (SELECT TOP {0} id FROM " &table_main& " WHERE parent_id<=0 AND adminid={1} ORDER BY id DESC)"
sql_write_flood_replytail=			") AND id IN (SELECT TOP {0} id FROM " &table_main& " WHERE (id={1} OR parent_id={1}) AND adminid={2} ORDER BY id DESC)"
sql_write_idnull=					"SELECT * FROM " &table_main& " WHERE id IS NULL"
if IsAccess then
	sql_write_verify_repliable=		"SELECT id FROM " &table_main& " WHERE parent_id<=0 AND (guestflag mod 1024)\512=0 AND id={0} AND adminid={1}"
	sql_write_updateparentflag=		"UPDATE " &table_main& " SET replied=(replied MOD 2) + 2 WHERE id={0} AND adminid={1}"
elseif IsSqlServer then
	sql_write_verify_repliable=		"SELECT id FROM " &table_main& " WHERE parent_id<=0 AND guestflag & 512=0 AND id={0} AND adminid={1}"
	sql_write_updateparentflag=		"UPDATE " &table_main& " SET replied=replied|2 WHERE id={0} AND adminid={1}"
end if
sql_write_updatelastupdated=		"UPDATE " &table_main& " SET lastupdated='{0}' WHERE id={1} AND adminid={2}"

'topbulletin
sql_topbulletin="SELECT declareflag,[declare] FROM " &table_supervisor& " WHERE adminid={0}"

'showword
sql_showword_count=	"SELECT COUNT(1) FROM " &table_main& " WHERE adminid={0}"
sql_showword=		"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id={0} AND adminid={1}"

'tlist
'sql_tlist_itemsperpage=	"SELECT itemsperpage FROM " &table_config& " WHERE adminid={0}"
sql_tlist_maxtime=		"SELECT TOP 1 lastupdated FROM " &table_main& " WHERE parent_id<=0 AND adminid={0} {1} ORDER BY lastupdated DESC"
if IsAccess then
	sql_tlist_mintime=	"SELECT TOP 1 lastupdated FROM (SELECT TOP {0} parent_id,lastupdated FROM " &table_main& " WHERE parent_id<=0 AND (guestflag mod 32)\16=0 AND (guestflag mod 64)\32=0 AND adminid={1} {2} ORDER BY parent_id ASC,lastupdated DESC) ORDER BY lastupdated ASC"
	sql_tlist=			"SELECT guestflag,title FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>=#{0}# AND lastupdated<=#{1}# AND adminid={2} {3} ORDER BY parent_id ASC,lastupdated DESC"
elseif IsSqlServer then
	sql_tlist_mintime=	"SELECT TOP 1 lastupdated FROM (SELECT TOP {0} parent_id,lastupdated FROM " &table_main& " WHERE parent_id<=0 AND guestflag & 48=0 AND adminid={1} {2} ORDER BY parent_id ASC,lastupdated DESC) AS __temp1 ORDER BY lastupdated ASC"
	sql_tlist=			"SELECT guestflag,title FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>='{0}' AND lastupdated<='{1}' AND adminid={2} {3} ORDER BY parent_id ASC,lastupdated DESC"
end if

'saveclientinfo
sql_saveclientinfo="INSERT INTO " &table_stats_clientinfo& "(os,browser,screenwidth,screenheight,timesect,sourceaddr,fullsource,adminid) VALUES('{0}','{1}',{2},{3},'{4}','{5}','{6}',{7})"

'listword_guest
sql_listword_guest="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id={0}{1} AND adminid={2} ORDER BY id ASC"

'listword_admin
sql_listword_admin="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id={0} AND adminid={1} ORDER BY id ASC"

'listword_web
sql_listword_web="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id={0} AND adminid={1} ORDER BY id ASC"

'admin
sql_admin_words_count="SELECT COUNT(1) FROM " &table_main& " WHERE parent_id<=0 AND adminid={0}"
sql_admin_words_query="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id<=0 AND adminid={0}"

'admin_showword
sql_admin_showword="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id={0} AND adminid={1}"

'admin_search
if IsAccess then
	sql_adminsearch_condition_audit=	"((guestflag mod 32)\16)={0} AND adminid={1}"
	sql_adminsearch_condition_reply=	"id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%') AND adminid={1}"
	sql_adminsearch_condition_else=		"{0} LIKE '%{1}%' AND adminid={2}"
	sql_adminsearch_count_inner=		"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_adminsearch_count=				"SELECT Count(*) FROM ({0})"
	sql_adminsearch_full_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_adminsearch_full=				"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id IN({0})"
elseif IsSqlServer then
	sql_adminsearch_condition_audit=	"guestflag % 32 / 16={0} AND adminid={1}"
	sql_adminsearch_condition_reply=	"id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%') AND adminid={1}"
	sql_adminsearch_condition_else=		"{0} LIKE '%{1}%' AND adminid={2}"
	sql_adminsearch_count_inner=		"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_adminsearch_count=				"SELECT Count(*) FROM ({0}) AS __temp1"
	sql_adminsearch_full_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_adminsearch_full=				"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id IN({0})"
end if

'admin_edit
sql_adminedit="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id={0} AND adminid={1}"

'admin_saveedit
sql_adminsaveedit_open="SELECT guestflag,title,article FROM " &table_main& " WHERE id={0} AND adminid={1}"

'admin_reply
sql_adminreply_reply="SELECT htmlflag,reinfo FROM " &table_reply& " WHERE articleid="
sql_adminreply_words="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id={0} AND adminid={1}"

'admin_delreply
sql_admindelreply_delete="DELETE FROM " &table_reply& " WHERE articleid IN (SELECT id FROM " &table_main& " WHERE id={0} AND adminid={1})"
if IsAccess then
	sql_admindelreply_update="UPDATE " &table_main& " SET replied=replied-(replied mod 2) WHERE id={0} AND adminid={1}"
elseif IsSqlServer then
	sql_admindelreply_update="UPDATE " &table_main& " SET replied=replied-(replied & 1) WHERE id={0} AND adminid={1}"
end if

'admin_savereply
sql_adminsavereply_main=			"SELECT guestflag,replied FROM " &table_main& " WHERE id={0} AND adminid={1}"
sql_adminsavereply_reply=			"SELECT articleid FROM " &table_reply& " WHERE articleid="
sql_adminsavereply_insert=			"INSERT INTO " &table_reply& "(articleid,replydate,htmlflag,reinfo) VALUES('{0}','{1}',{2},'{3}')"
sql_adminsavereply_update=			"UPDATE " &table_reply& " SET replydate='{0}',htmlflag={1},reinfo='{2}' WHERE articleid={3}"
if IsAccess then
	sql_adminsavereply_lock2top=	"UPDATE " &table_main& " SET parent_id=-1 WHERE parent_id=0 AND id=(SELECT root_id FROM " &table_main& " WHERE id={0} AND adminid={1})"
	sql_adminsavereply_bring2top=	"UPDATE " &table_main& " SET lastupdated='{0}' WHERE parent_id<=0 AND id=(SELECT root_id FROM " &table_main& " WHERE id={1} AND adminid={2})"
elseif IsSqlServer then
	sql_adminsavereply_lock2top=	"UPDATE " &table_main& " SET parent_id=-1 WHERE parent_id=0 AND id=(SELECT root_id FROM " &table_main& " WHERE id={0} AND adminid={1})"
	sql_adminsavereply_bring2top=	"UPDATE " &table_main& " SET lastupdated='{0}' WHERE parent_id<=0 AND id=(SELECT root_id FROM " &table_main& " WHERE id={1} AND adminid={2})"
end if

'admin_[action]
if IsAccess then
	sql_adminhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag+256 WHERE (guestflag mod 512)\256=0 AND id={0} AND adminid={1}"
	sql_adminunhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag-256 WHERE (guestflag mod 512)\256<>0 AND id={0} AND adminid={1}"
	sql_adminhideword=		"UPDATE " &table_main& " SET guestflag=guestflag+8 WHERE (guestflag mod 16)\8=0 AND id={0} AND adminid={1}"
	sql_adminunhideword=	"UPDATE " &table_main& " SET guestflag=guestflag-8 WHERE (guestflag mod 16)\8<>0 AND id={0} AND adminid={1}"
	sql_adminlockreply=		"UPDATE " &table_main& " SET guestflag=guestflag+512 WHERE (guestflag mod 1024)\512=0 AND id={0} AND adminid={1}"
	sql_adminunlockreply=	"UPDATE " &table_main& " SET guestflag=guestflag-512 WHERE (guestflag mod 1024)\512<>0 AND id={0} AND adminid={1}"
	sql_adminpassaudit=		"UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE (guestflag mod 32)\16<>0 AND id={0} AND adminid={1}"
	sql_adminpubwhisper=	"UPDATE " &table_main& " SET guestflag=guestflag-32 WHERE (guestflag mod 64)\32<>0 AND id={0} AND adminid={1}"
elseif IsSqlServer then
	sql_adminhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag+256 WHERE guestflag & 256=0 AND id={0} AND adminid={1}"
	sql_adminunhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag-256 WHERE guestflag & 256<>0 AND id={0} AND adminid={1}"
	sql_adminhideword=		"UPDATE " &table_main& " SET guestflag=guestflag+8 WHERE guestflag & 8=0 AND id={0} AND adminid={1}"
	sql_adminunhideword=	"UPDATE " &table_main& " SET guestflag=guestflag-8 WHERE guestflag & 8<>0 AND id={0} AND adminid={1}"
	sql_adminlockreply=		"UPDATE " &table_main& " SET guestflag=guestflag+512 WHERE guestflag & 512=0 AND id={0} AND adminid={1}"
	sql_adminunlockreply=	"UPDATE " &table_main& " SET guestflag=guestflag-512 WHERE guestflag & 512<>0 AND id={0} AND adminid={1}"
	sql_adminpassaudit=		"UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE guestflag & 16<>0 AND id={0} AND adminid={1}"
	sql_adminpubwhisper=	"UPDATE " &table_main& " SET guestflag=guestflag-32 WHERE guestflag & 32<>0 AND id={0} AND adminid={1}"
end if
sql_admindel_reply=			"DELETE FROM " &table_reply& " WHERE articleid IN(SELECT id FROM " &table_main& " WHERE (parent_id={0} OR id={0}) AND adminid={1})"
sql_admindel_main=			"DELETE FROM " &table_main& " WHERE (parent_id={0} OR id={0}) AND adminid={1}"
sql_adminlock2top=			"UPDATE " &table_main& " SET parent_id=-1 WHERE parent_id=0 AND id={0} AND adminid={1}"
sql_adminunlock2top=		"UPDATE " &table_main& " SET parent_id=0 WHERE parent_id<0 AND id={0} AND adminid={1}"
sql_adminbring2top=			"UPDATE " &table_main& " SET lastupdated='{0}' WHERE parent_id<=0 AND id={1} AND adminid={2}"
sql_adminreorder=			"UPDATE " &table_main& " SET parent_id=0,lastupdated=logdate WHERE parent_id<=0 AND id={0} AND adminid={1}"

'admin_mdel
sql_adminmdel_reply="DELETE FROM " &table_reply& " WHERE articleid IN(SELECT id FROM " &table_main& " WHERE (parent_id IN({0}) OR id IN({0})) AND adminid={1})"
sql_adminmdel_main=	"DELETE FROM " &table_main& " WHERE (parent_id IN({0}) OR id IN({0})) AND adminid={1}"

'admin_mpass
if IsAccess then
	sql_adminmpass="UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE ((guestflag mod 32)\16)<>0 AND id IN ({0}) AND adminid={1}"
elseif IsSqlServer then
	sql_adminmpass="UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE guestflag & 16<>0 AND id IN ({0}) AND adminid={1}"
end if

'admin_doadvdel
if IsAccess then
	sql_admindoadvdel_beforedate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated<=#{0}# AND adminid={1}"
	sql_admindoadvdel_afterdate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>=#{0}# AND adminid={1}"
elseif IsSqlServer then
	sql_admindoadvdel_beforedate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated<='{0}' AND adminid={1}"
	sql_admindoadvdel_afterdate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>='{0}' AND adminid={1}"
end if
sql_admindoadvdel_firstn_main=			"DELETE FROM " &table_main& " WHERE id IN (SELECT TOP {0} id FROM " &table_main& " WHERE parent_id<=0 AND adminid={1} ORDER BY lastupdated ASC)"
sql_admindoadvdel_lastn_main=			"DELETE FROM " &table_main& " WHERE id IN (SELECT top {0} id FROM " &table_main& " WHERE parent_id<=0 AND adminid={1} ORDER BY lastupdated DESC)"
sql_admindoadvdel_name_main=			"DELETE FROM " &table_main& " WHERE name LIKE '%{0}%' AND adminid={1}"
sql_admindoadvdel_title_main=			"DELETE FROM " &table_main& " WHERE title LIKE '%{0}%' AND adminid={1}"
sql_admindoadvdel_article_main=			"DELETE FROM " &table_main& " WHERE article LIKE '%{0}%' AND adminid={1}"
sql_admindoadvdel_main=					"DELETE FROM " &table_main& " WHERE adminid={0}"
sql_admindoadvdel_clearfragment_main=	"DELETE FROM " &table_main& " WHERE parent_id>0 AND parent_id NOT IN (SELECT DISTINCT id FROM " &table_main& " WHERE parent_id<=0)"
sql_admindoadvdel_clearfragment_reply=	"DELETE FROM " &table_reply& " WHERE articleid NOT IN(SELECT id FROM " &table_main& ")"
if IsAccess then
	sql_admindoadvdel_adjustguestreply_flag="UPDATE " &table_main& " SET replied=replied-((replied MOD 4)\2)*2 WHERE parent_id<=0 AND id NOT IN(" & _
												"SELECT DISTINCT parent_id FROM " &table_main & _
											") AND adminid={0}"
elseif IsSqlServer then
	sql_admindoadvdel_adjustguestreply_flag="UPDATE " &table_main& " SET replied=replied-(replied&2) WHERE parent_id<=0 AND id NOT IN(" & _
												"SELECT DISTINCT parent_id FROM " &table_main & " WHERE adminid={0}" & _
											") AND adminid={0}"
end if

'admin_setbulletin
sql_adminsetbulletin="SELECT declareflag,[declare] FROM " &table_supervisor& " WHERE adminid={0}"

'admin_savebulletin
sql_adminsavebulletin="UPDATE " &table_supervisor& " SET declareflag={0},[declare]='{1}' WHERE adminid={2}"

'admin_config/web_config
sql_adminconfig_config=	"SELECT * FROM " &table_config& " WHERE adminid={0}"
sql_adminconfig_style=	"SELECT styleid,stylename FROM " &table_style

'admin_saveconfig/web_saveconfig
sql_adminsaveconfig="SELECT * FROM " &table_config& " WHERE adminid={0}"

'admin_ipconfig/web_ipconfig
sql_adminipconfig_status=	"SELECT ipconstatus FROM " &table_config& " WHERE adminid={0}"
sql_adminipconfig_status1=	"SELECT listid,startip,endip FROM " &table_ipconfig& " WHERE ipconstatus=1 AND adminid={0}"
sql_adminipconfig_status2=	"SELECT listid,startip,endip FROM " &table_ipconfig& " WHERE ipconstatus=2 AND adminid={0}"

'admin_saveipconfig/web_saveipconfig
sql_adminsaveipconfig_delete1=	"DELETE FROM " &table_ipconfig& " WHERE listid IN ({0}) AND adminid={1}"
sql_adminsaveipconfig_delete2=	"DELETE FROM " &table_ipconfig& " WHERE listid IN ({0}) AND adminid={1}"
sql_adminsaveipconfig_insert1=	"INSERT INTO " &table_ipconfig& "(ipconstatus,startip,endip,adminid) VALUES (1,'{0}','{1}',{2})"
sql_adminsaveipconfig_insert2=	"INSERT INTO " &table_ipconfig& "(ipconstatus,startip,endip,adminid) VALUES (2,'{0}','{1}',{2})"
sql_adminsaveipconfig_update=	"UPDATE " &table_config& " SET ipconstatus={0} WHERE adminid={1}"

'admin_filter/web_filter
sql_adminfilter="SELECT * FROM " &table_filterconfig& " WHERE adminid={0} ORDER BY filtersort ASC"

'admin_appendfilter/web_appendfilter
sql_adminfilter_null=	"SELECT * FROM " &table_filterconfig& " WHERE filterid IS NULL"
sql_adminfilter_max=	"SELECT MAX(filterid) FROM " &table_filterconfig& " WHERE adminid={0}"
sql_adminfilter_update=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filterid={0} AND adminid={1}"

'admin_updatefilter/web_updatefilter
sql_adminupdatefilter="SELECT regexp,filtermode,replacestr,[memo] FROM " &table_filterconfig& " WHERE filterid={0} AND adminid={1}"

'admin_delfilter/web_delfilter
sql_admindelfilter="DELETE FROM " &table_filterconfig& " WHERE filterid={0} AND adminid={1}"

'admin_movefilter/web_movefilter
sql_adminmovefilter_sort1=			"SELECT filtersort FROM " &table_filterconfig& " WHERE filterid={0} AND adminid={1}"
sql_adminmovefilter_sort2_up=		"SELECT MAX(filtersort) FROM " &table_filterconfig& " WHERE filtersort<{0} AND adminid={1}"
sql_adminmovefilter_sort2_down=		"SELECT MIN(filtersort) FROM " &table_filterconfig& " WHERE filtersort>{0} AND adminid={1}"
sql_adminmovefilter_sort2_top=		"SELECT MIN(filtersort) FROM " &table_filterconfig& " WHERE adminid={0}"
sql_adminmovefilter_sort2_bottom=	"SELECT MAX(filtersort) FROM " &table_filterconfig& " WHERE adminid={0}"
sql_adminmovefilter_id2=			"SELECT filterid FROM " &table_filterconfig& " WHERE filtersort={0} AND adminid={1}"
sql_adminmovefilter_update_updown1=	"UPDATE " &table_filterconfig& " SET filtersort=-1 WHERE filterid={0} AND adminid={1}"
sql_adminmovefilter_update_updown2=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filtersort={1} AND adminid={2}"
sql_adminmovefilter_update_updown3=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filtersort=-1 AND adminid={1}"
sql_adminmovefilter_update_top1=	"UPDATE " &table_filterconfig& " SET filtersort=filtersort+1 WHERE filtersort<{0} AND adminid={1}"
sql_adminmovefilter_update_top2=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filterid={1} AND adminid={2}"
sql_adminmovefilter_update_bottom1=	"UPDATE " &table_filterconfig& " SET filtersort=filtersort-1 WHERE filtersort>{0} AND adminid={1}"
sql_adminmovefilter_update_bottom2=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filterid={1} AND adminid={2}"

'admin_floodconfig/web_floodconfig
sql_adminfloodconfig="SELECT * FROM " &table_floodconfig& " WHERE adminid={0}"

'admin_savefloodconfig/web_savefloodconfig
sql_adminsavefloodconfig="UPDATE " &table_floodconfig& " SET minwait={0},searchrange={1},searchflag={2} WHERE adminid={3}"

'admin_stats
sql_adminstats_startdate=		"SELECT startdate FROM " &table_stats& " WHERE adminid={0}"
sql_adminstats_insert=			"INSERT INTO " &table_stats& "(startdate,adminid) VALUES('{0}',{1})"
sql_adminstats=					"SELECT * FROM " &table_stats& " WHERE adminid={0}"
sql_adminstats_client_count=	"SELECT Count(*) FROM " &table_stats_clientinfo& " WHERE adminid={0}"
if IsAccess then
	sql_adminstats_client_os=		"SELECT TOP 30 * FROM (SELECT os,Count(os) AS count_os FROM " &table_stats_clientinfo& " WHERE adminid={0} GROUP BY os ORDER BY Count(os) DESC) ORDER BY count_os DESC"
	sql_adminstats_client_browser=	"SELECT TOP 30 * FROM (SELECT browser,Count(browser) AS count_browser FROM " &table_stats_clientinfo& " WHERE adminid={0} GROUP BY browser ORDER BY Count(browser) DESC) ORDER BY count_browser DESC"
	sql_adminstats_client_screen=	"SELECT TOP 30 * FROM (SELECT screenwh,Count(screenwh) AS count_screenwh FROM (SELECT screenwidth &'*'& screenheight AS screenwh FROM " &table_stats_clientinfo& " WHERE adminid={0}) GROUP BY screenwh ORDER BY Count(screenwh) DESC) ORDER BY count_screenwh DESC"
	sql_adminstats_client_timesect=	"SELECT hsect,Count(hsect) FROM (SELECT Hour(timesect) AS hsect FROM " &table_stats_clientinfo& " WHERE adminid={0}) GROUP BY hsect ORDER BY hsect ASC"
	sql_adminstats_client_week=		"SELECT weekno,Count(weekno) FROM (SELECT Weekday(timesect,1) AS weekno FROM " &table_stats_clientinfo& " WHERE adminid={0}) GROUP BY weekno ORDER BY weekno ASC"
	sql_adminstats_client_30day=	"SELECT * FROM (SELECT TOP 30 * FROM (SELECT datesect,Count(datesect) FROM (SELECT Year(timesect) & '-' & Month(timesect) & '-' & Day(timesect) AS datesect FROM " &table_stats_clientinfo& " WHERE adminid={0}) GROUP BY datesect ORDER BY CDate(datesect) DESC)) ORDER BY CDate(datesect) ASC"
	sql_adminstats_client_source=	"SELECT TOP 30 * FROM (SELECT sourceaddr,Count(sourceaddr) AS count_sourceaddr FROM " &table_stats_clientinfo& " WHERE adminid={0} GROUP BY sourceaddr ORDER BY Count(sourceaddr) DESC) ORDER BY count_sourceaddr DESC"
elseif IsSqlServer then
	sql_adminstats_client_os=		"SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT os,Count(os) AS count_os FROM " &table_stats_clientinfo& " WHERE adminid={0} GROUP BY os ORDER BY count_os DESC) AS __temp1 ORDER BY count_os DESC"
	sql_adminstats_client_browser=	"SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT browser,Count(browser) AS count_browser FROM " &table_stats_clientinfo& " WHERE adminid={0} GROUP BY browser ORDER BY count_browser DESC) AS __temp1 ORDER BY count_browser DESC"
	sql_adminstats_client_screen=	"SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT screenwh,Count(screenwh) AS count_screenwh FROM (SELECT CAST(screenwidth AS varchar) +'*'+ CAST(screenheight AS varchar) AS screenwh FROM " &table_stats_clientinfo& " WHERE adminid={0}) AS __temp1 GROUP BY screenwh ORDER BY count_screenwh DESC) AS __temp2 ORDER BY count_screenwh DESC"
	sql_adminstats_client_timesect=	"SELECT hsect,Count(hsect) AS count_hsect FROM (SELECT DATEPART(hh,timesect) AS hsect FROM " &table_stats_clientinfo& " WHERE adminid={0}) AS __temp1 GROUP BY hsect ORDER BY hsect ASC"
	sql_adminstats_client_week=		"SET DATEFIRST 7; SELECT weekno,Count(weekno) FROM (SELECT DATEPART(dw,timesect) AS weekno FROM " &table_stats_clientinfo& " WHERE adminid={0}) AS __temp1 GROUP BY weekno ORDER BY weekno ASC"
	sql_adminstats_client_30day=	"SELECT * FROM (SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT datesect,Count(datesect) AS count_datesect FROM (SELECT CAST(Year(timesect) AS varchar) + '-' + CAST(Month(timesect) AS varchar) + '-' + CAST(Day(timesect) AS varchar) AS datesect FROM " &table_stats_clientinfo& " WHERE adminid={0}) AS __temp1 GROUP BY datesect ORDER BY CAST(datesect AS datetime) DESC) AS __temp2) AS __temp3 ORDER BY CAST(datesect AS datetime) ASC"
	sql_adminstats_client_source=	"SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT sourceaddr,Count(sourceaddr) AS count_sourceaddr FROM " &table_stats_clientinfo& " WHERE adminid={0} GROUP BY sourceaddr ORDER BY count_sourceaddr DESC) AS __temp1 ORDER BY count_sourceaddr DESC"
end if

'admin_clearstats
sql_adminclearstats_startdate=	"UPDATE " &table_stats& " SET startdate='{0}',[view]=0,[search]=0,[leaveword]=0,[written]=0,[filtered]=0,[banned]=0,[login]=0,[loginfailed]=0 WHERE adminid={1}"
sql_adminclearstats_client=		"DELETE FROM " &table_stats_clientinfo& " WHERE adminid={0}"

'admin_setinfo
sql_adminsetinfo="SELECT name,faceid,faceurl,email,qqid,msnid,homepage FROM " &table_supervisor& " WHERE adminid={0}"

'admin_saveinfo
sql_adminsaveinfo="SELECT name,faceid,faceurl,email,qqid,msnid,homepage FROM " &table_supervisor& " WHERE adminid={0}"

'admin_chpass
sql_adminchpass_question="SELECT question FROM " &table_supervisor& " WHERE adminid={0}"

'admin_savepass
sql_adminsavepass_select="SELECT adminpass FROM " &table_supervisor& " WHERE adminid={0}"
sql_adminsavepass_update="UPDATE " &table_supervisor& " SET adminpass='{0}' WHERE adminid={1}"

'admin_saveqa
sql_adminsaveqa_select="SELECT adminpass FROM " &table_supervisor& " WHERE adminid={0}"
sql_adminsaveqa_update="UPDATE " &table_supervisor& " SET question='{0}',[key]='{1}' WHERE adminid={2}"

'login_verify
sql_loginverify=	"SELECT adminpass FROM " &table_supervisor& " WHERE adminid={0}"
sql_updatelastlogin="UPDATE " &table_supervisor& " SET lastlogin='{0}' WHERE adminid={1}"

'admin_verify
sql_adminverify="SELECT adminpass FROM " &table_supervisor& " WHERE adminid={0}"


'================================
'================================
'================================
'================================
'================================


'sql_sysbulletin
sql_sysbulletin="SELECT TOP 1 bulletinflag,webbulletin FROM " &table_webmaster

'checkuser
sql_checkuser="SELECT adminname FROM " &table_supervisor& " WHERE adminname='{0}'"

'submitreg
sql_submitreg_checkuser="SELECT 1 FROM " &table_supervisor& " WHERE adminname='{0}'"
sql_submitreg_init1=	"INSERT INTO " &table_supervisor& "(adminname,adminpass,name,faceid,regdate,lastlogin,question,[key],[declare]) VALUES('{0}','{1}','{2}',{3},'{4}','{5}','{6}','{7}','{8}')"
sql_submitreg_init2=	"INSERT INTO " &table_config& "(adminid,status,adminhtml,guesthtml,admintimeout,visualflag,ubbflag,styleid) SELECT TOP 1 adminid,{1},{2},{3},{4},{5},{6},'{7}' FROM " &table_supervisor& " WHERE adminname='{0}'"
sql_submitreg_init3=	"INSERT INTO " &table_floodconfig& "(adminid) SELECT TOP 1 adminid FROM " &table_supervisor& " WHERE adminname='{0}'"

'findkey2
sql_findkey2_checkuser="SELECT adminname,question FROM " &table_supervisor& " WHERE adminname='{0}'"

'findkey3
sql_findkey3_checkuser="SELECT adminname,[key] FROM " &table_supervisor& " WHERE adminname='{0}'"

'findkey4
sql_findkey4_checkuser="SELECT adminname,[key] FROM " &table_supervisor& " WHERE adminname='{0}'"
sql_findkey4_resetpass="UPDATE " &table_supervisor& " SET adminpass='{0}' WHERE adminname='{1}'"

'submitunreg
sql_submitunreg_checkuser=				"SELECT adminid,adminpass FROM " &table_supervisor& " WHERE adminname='{0}' AND adminname<>'{1}'"
sql_submitunreg_delete_filterconfig=	"DELETE FROM " &table_filterconfig& " WHERE adminid={0}"
sql_submitunreg_delete_ipconfig=		"DELETE FROM " &table_ipconfig& " WHERE adminid={0}"
sql_submitunreg_delete_floodconfig=		"DELETE FROM " &table_floodconfig& " WHERE adminid={0}"
sql_submitunreg_delete_stats=			"DELETE FROM " &table_stats& " WHERE adminid={0}"
sql_submitunreg_delete_stats_clientinfo="DELETE FROM " &table_stats_clientinfo& " WHERE adminid={0}"
sql_submitunreg_delete_reply=			"DELETE FROM " &table_reply& " WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid={0})"
sql_submitunreg_delete_main=			"DELETE FROM " &table_main& " WHERE adminid={0}"
sql_submitunreg_delete_config=			"DELETE FROM " &table_config& " WHERE adminid={0}"
sql_submitunreg_delete_supervisor=		"DELETE FROM " &table_supervisor& " WHERE adminid={0}"

'web_login_verify
sql_web_login_verify="SELECT TOP 1 webpass FROM " &table_webmaster

'web_admin_verify
sql_web_admin_verify="SELECT TOP 1 webpass FROM " &table_webmaster

'web_admin
sql_webadmin_words_count="SELECT COUNT(adminname) FROM " &table_supervisor
sql_webadmin_words_query="SELECT adminname,name,regdate,lastlogin FROM " &table_supervisor

'web_searchuser
sql_websearchuser_condition="{0} LIKE '%{1}%'"
sql_websearchuser_words_count="SELECT COUNT(adminname) FROM " &table_supervisor& " WHERE "
sql_websearchuser_words_query="SELECT adminname,name,regdate,lastlogin FROM " &table_supervisor& " WHERE "

'web_search
sql_websearch_condition_init=		"parent_id<=0"
sql_websearch_condition_adminname=	" AND adminname LIKE '%{0}%'"
sql_websearch_condition_name=		" AND main.name LIKE '%{0}%'"
sql_websearch_condition_title=		" AND title LIKE '%{0}%'"
sql_websearch_condition_article=	" AND article LIKE '%{0}%'"
sql_websearch_condition_email=		" AND email LIKE '%{0}%'"
sql_websearch_condition_qqid=		" AND qqid LIKE '%{0}%'"
sql_websearch_condition_msnid=		" AND msnid LIKE '%{0}%'"
sql_websearch_condition_homepage=	" AND homepage LIKE '%{0}%'"
sql_websearch_condition_ipaddr=		" AND ipaddr LIKE '%{0}%'"
sql_websearch_condition_originalip=	" AND originalip LIKE '%{0}%'"
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

'web_showword
sql_web_showword="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id="

'web_searchdec
sql_websearchdec_words_count="SELECT COUNT(adminname) FROM " &table_supervisor& " WHERE adminname LIKE '%{0}%' AND [declare] LIKE '%{1}%' AND [declare] LIKE '_%'"
sql_websearchdec_words_query="SELECT adminname,declareflag,[declare] FROM " &table_supervisor& " WHERE adminname LIKE '%{0}%' AND [declare] LIKE '%{1}%' AND [declare] LIKE '_%'"

'web_doadvdel
if IsAccess then
	sql_webdoadvdel_beforedate_main="DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated<=#{0}#"
	sql_webdoadvdel_afterdate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>=#{0}#"
elseif IsSqlServer then
	sql_webdoadvdel_beforedate_main="DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated<='{0}'"
	sql_webdoadvdel_afterdate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>='{0}'"
end if
sql_webdoadvdel_firstn_main=				"DELETE FROM " &table_main& " WHERE id IN (SELECT top {0} id FROM " &table_main& " WHERE parent_id<=0 ORDER BY lastupdated ASC)"
sql_webdoadvdel_lastn_main=					"DELETE FROM " &table_main& " WHERE id IN (SELECT top {0} id FROM " &table_main& " WHERE parent_id<=0 ORDER BY lastupdated DESC)"	
sql_webdoadvdel_name_main=					"DELETE FROM " &table_main& " WHERE name LIKE '%{0}%'"
sql_webdoadvdel_title_main=					"DELETE FROM " &table_main& " WHERE title LIKE '%{0}%'"
sql_webdoadvdel_article_main=				"DELETE FROM " &table_main& " WHERE article LIKE '%{0}%'"
sql_webdoadvdel_reply_main=					"DELETE FROM " &table_main& " WHERE id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%')"
sql_webdoadvdel_reply_reply=				"DELETE FROM " &table_reply& " WHERE reinfo LIKE '%{0}%'"
sql_webdoadvdel_reply_unsetreply=			"UPDATE " &table_main& " SET replied=0 WHERE id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%')"
sql_webdoadvdel_main=						"DELETE FROM " &table_main
sql_webdoadvdel_unsetreply=					"UPDATE " &table_main& " SET replied=0"
sql_webdoadvdel_reply=						"DELETE FROM " &table_reply
sql_webdoadvdel_deldeclare=					"UPDATE " &table_supervisor& " SET [declare]='' WHERE [declare] LIKE '%{0}%'"
sql_webdoadvdel_cleardeclare=				"UPDATE " &table_supervisor& " SET [declare]='' WHERE [declare] LIKE '_%'"
if IsAccess then
	sql_webdoadvdel_regdate_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_ipconfig=			"DELETE FROM " &table_ipconfig&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_stats=				"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_reply=				"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#))"
	sql_webdoadvdel_regdate_main=				"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_config=				"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<=#{0}#)"
	sql_webdoadvdel_regdate_supervisor=			"DELETE FROM " &table_supervisor&		" WHERE regdate<=#{0}#"

	sql_webdoadvdel_lastlogin_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_ipconfig=			"DELETE FROM " &table_ipconfig&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_stats=			"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_reply=			"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#))"
	sql_webdoadvdel_lastlogin_main=				"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_config=			"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<=#{0}#)"
	sql_webdoadvdel_lastlogin_supervisor=		"DELETE FROM " &table_supervisor&		" WHERE lastlogin<=#{0}#"

	sql_webdoadvdel_logdate_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#) AND adminid<>'{1}'"
	sql_webdoadvdel_logdate_ipconfig=			"DELETE FROM " &table_ipconfig&			" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#) AND adminid<>'{1}'"
	sql_webdoadvdel_logdate_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#) AND adminid<>'{1}'"
	sql_webdoadvdel_logdate_stats=				"DELETE FROM " &table_stats&			" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#)"
	sql_webdoadvdel_logdate_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo& " WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#)"
	sql_webdoadvdel_logdate_reply=				"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#))"
	sql_webdoadvdel_logdate_main=				"DELETE FROM " &table_main&				" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#)"
	sql_webdoadvdel_logdate_config=				"DELETE FROM " &table_config&			" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#) AND adminid<>'{1}'"
	sql_webdoadvdel_logdate_supervisor=			"DELETE FROM " &table_supervisor&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>#{0}#)"
elseif IsSqlServer then
	sql_webdoadvdel_regdate_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_ipconfig=			"DELETE FROM " &table_ipconfig&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_stats=				"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_reply=				"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}'))"
	sql_webdoadvdel_regdate_main=				"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_config=				"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE regdate<='{0}')"
	sql_webdoadvdel_regdate_supervisor=			"DELETE FROM " &table_supervisor&		" WHERE regdate<='{0}'"

	sql_webdoadvdel_lastlogin_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_ipconfig=			"DELETE FROM " &table_ipconfig&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_stats=			"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_reply=			"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}'))"
	sql_webdoadvdel_lastlogin_main=				"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_config=			"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin<='{0}')"
	sql_webdoadvdel_lastlogin_supervisor=		"DELETE FROM " &table_supervisor&		" WHERE lastlogin<='{0}'"

	sql_webdoadvdel_logdate_filterconfig=		"DELETE FROM " &table_filterconfig&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}') AND adminid<>'{1}'"
	sql_webdoadvdel_logdate_ipconfig=			"DELETE FROM " &table_ipconfig&			" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}') AND adminid<>'{1}'"
	sql_webdoadvdel_logdate_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}') AND adminid<>'{1}'"
	sql_webdoadvdel_logdate_stats=				"DELETE FROM " &table_stats&			" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}')"
	sql_webdoadvdel_logdate_stats_clientinfo=	"DELETE FROM " &table_stats_clientinfo& " WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}')"
	sql_webdoadvdel_logdate_reply=				"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}'))"
	sql_webdoadvdel_logdate_main=				"DELETE FROM " &table_main&				" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}')"
	sql_webdoadvdel_logdate_config=				"DELETE FROM " &table_config&			" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}') AND adminid<>'{1}'"
	sql_webdoadvdel_logdate_supervisor=			"DELETE FROM " &table_supervisor&		" WHERE adminid NOT IN (SELECT DISTINCT adminid FROM " &table_main& " WHERE logdate>'{0}')"
end if
sql_webdoadvdel_neverlogin_filterconfig=	"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate)"
sql_webdoadvdel_neverlogin_ipconfig=		"DELETE FROM " &table_ipconfig&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate)"
sql_webdoadvdel_neverlogin_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate)"
sql_webdoadvdel_neverlogin_stats=			"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate)"
sql_webdoadvdel_neverlogin_stats_clientinfo="DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate)"
sql_webdoadvdel_neverlogin_reply=			"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate))"
sql_webdoadvdel_neverlogin_main=			"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate)"
sql_webdoadvdel_neverlogin_config=			"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE lastlogin=regdate)"
sql_webdoadvdel_neverlogin_supervisor=		"DELETE FROM " &table_supervisor&		" WHERE lastlogin=regdate"

sql_webdoadvdel_all_filterconfig=	"DELETE FROM " &table_filterconfig& " WHERE adminid<>{0}"
sql_webdoadvdel_all_ipconfig=		"DELETE FROM " &table_ipconfig& " WHERE adminid<>{0}"
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

'web_setbulletin
sql_websetbulletin="SELECT TOP 1 bulletinflag,webbulletin FROM " &table_webmaster

'web_savebulletin
sql_websavebulletin="UPDATE " &table_webmaster& " SET bulletinflag={0},[webbulletin]='{1}'"

'web_savepass
sql_websavepass_select="SELECT TOP 1 webpass FROM " &table_webmaster
sql_websavepass_update="UPDATE " &table_webmaster& " SET webpass='{0}'"

'web_deluser
sql_webdeluser_filterconfig=	"DELETE FROM " &table_filterconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_ipconfig=		"DELETE FROM " &table_ipconfig&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_floodconfig=		"DELETE FROM " &table_floodconfig&		" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_stats=			"DELETE FROM " &table_stats&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_stats_clientinfo="DELETE FROM " &table_stats_clientinfo&	" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_reply=			"DELETE FROM " &table_reply&			" WHERE articleid IN (SELECT id FROM " &table_main& " WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0})))"
sql_webdeluser_main=			"DELETE FROM " &table_main&				" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_config=			"DELETE FROM " &table_config&			" WHERE adminid IN (SELECT adminid FROM " &table_supervisor& " WHERE adminname IN({0}))"
sql_webdeluser_supervisor=		"DELETE FROM " &table_supervisor&		" WHERE adminname IN ({0})"

'web_userinfo
sql_webuserinfo=				"SELECT faceid,adminname,adminid,name,email,qqid,msnid,homepage FROM " &table_supervisor& " WHERE adminname='{0}'"
sql_webuserinfo_count_view=		"SELECT [view] FROM " &table_stats& " WHERE adminid={0}"
sql_webuserinfo_count_words=	"SELECT COUNT(1) FROM " &table_main& " WHERE adminid={0}"
sql_webuserinfo_count_reply=	"SELECT COUNT(articleid) FROM " &table_main& " INNER JOIN " &table_reply& " ON (" &table_main& ".id=" &table_reply& ".articleid AND adminid={0})"
sql_webuserinfo_count_ipconfig=	"SELECT COUNT(1) FROM " &table_ipconfig& " WHERE adminid={0}"
sql_webuserinfo_count_filterconfig="SELECT COUNT(filterid) FROM " &table_filterconfig& " WHERE adminid={0}"
sql_webuserinfo_reginfo=		"SELECT regdate,lastlogin FROM " &table_supervisor& " WHERE adminid={0}"
sql_webuserinfo_logininfo=		"SELECT login,loginfailed FROM " &table_stats& " WHERE adminid={0}"

'web_saveuserpass
sql_websaveuserpass="UPDATE " &table_supervisor& " SET adminpass='{0}' WHERE adminname='{1}'"

'web_searchdel
sql_websearchdel_reply=	"DELETE FROM " &table_reply& " WHERE articleid IN(SELECT id FROM " &table_main& " WHERE parent_id={0} OR id={0})"
sql_websearchdel_main=	"DELETE FROM " &table_main& " WHERE parent_id={0} OR id={0}"

'web_searchdelreply
sql_websearchdelreply_reply=		"DELETE FROM " &table_reply& " WHERE articleid="
sql_websearchdelreply_unsetreply=	"UPDATE " &table_main& " SET replied=0 WHERE id="

'web_searchmdel
sql_websearchmdel_reply="DELETE FROM " &table_reply& " WHERE articleid IN(SELECT id FROM " &table_main& " WHERE parent_id IN({0}) OR id IN({0}))"
sql_websearchmdel_main=	"DELETE FROM " &table_main& " WHERE parent_id IN({0}) OR id IN({0})"

'web_deldec
sql_webdeldec="UPDATE " &table_supervisor& " SET [declare]='' WHERE adminname='{0}'"

'web_mdeldec
sql_webmdeldec="UPDATE " &table_supervisor& " SET [declare]='' WHERE adminname IN ({0})"

'compact
if IsSqlServer then
	sql_compact_dbfile="DECLARE @curr_dbname NVARCHAR(255); SET @curr_dbname=DB_NAME(0); DBCC SHRINKDATABASE(@curr_dbname,8)"
	sql_compact_dblog="DECLARE @curr_dbname NVARCHAR(255); SET @curr_dbname=DB_NAME(0); BACKUP LOG @curr_dbname WITH NO_LOG"
end if
%>