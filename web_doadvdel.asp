<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_doadvdel.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1
Dim affected
affected=0

set cn1=server.CreateObject("ADODB.Connection")
set rs1=server.CreateObject("ADODB.Recordset")
CreateConn cn1,dbtype
set rs1=nothing

dim tparam
tparam=Request.Form("iparam")
select case Request.Form("option")
case "1"
	if isdate(tparam) then
		tparam=DateTimeStr(tparam)
		cn1.Execute Replace(sql_webdoadvdel_beforedate_main,"{0}",tparam),affected,1
		Call TipsPage("已删除" &affected& "条留言。","web_advdel.asp")
	else
		Call TipsPage("您输入的日期有误，请检查。","web_advdel.asp")
	end if
case "2"
	if isdate(tparam) then
		tparam=DateTimeStr(tparam)
		cn1.Execute Replace(sql_webdoadvdel_afterdate_main,"{0}",tparam),affected,1
		Call TipsPage("已删除" &affected& "条留言。","web_advdel.asp")
	else
		Call TipsPage("您输入的日期有误，请检查。","web_advdel.asp")
	end if
case "3"
	if isnumeric(tparam) then
		if clng(tparam)>0 then
			tparam=clng(tparam)
			cn1.Execute Replace(sql_webdoadvdel_firstn_main,"{0}",tparam),affected,1
			Call TipsPage("已删除" &affected& "条留言。","web_advdel.asp")
		else
			Call TipsPage("请输入正确的数值。","web_advdel.asp")
		end if
	else
		Call TipsPage("请输入正确的数值。","web_advdel.asp")
	end if
case "4"
	if isnumeric(tparam) then
		if clng(tparam)>0 then
			tparam=clng(tparam)
			cn1.Execute Replace(sql_webdoadvdel_lastn_main,"{0}",tparam),affected,1
			Call TipsPage("已删除" &affected& "条留言。","web_advdel.asp")
		else
			Call TipsPage("请输入正确的数值。","web_advdel.asp")
		end if
	else
		Call TipsPage("请输入正确的数值。","web_advdel.asp")
	end if
case "5"
	tparam=replace(tparam,"'","''")
	tparam=replace(tparam,"[","[\[]")
	while left(tparam,1)="%" or left(tparam,1)="_"
		tparam=right(tparam,len(tparam)-1)
	wend
	while right(tparam,1)="%" or right(tparam,1)="_"
		tparam=left(tparam,len(tparam)-1)
	wend
	if tparam<>"" then
		cn1.Execute Replace(sql_webdoadvdel_name_main,"{0}",tparam),affected,1
		Call TipsPage("已删除" &affected& "条留言。","web_advdel.asp")
	else
		Call TipsPage("不能输入空字符串或全部为通配符。","web_advdel.asp")
	end if
case "6"
	tparam=replace(tparam,"'","''")
	tparam=replace(tparam,"[","[\[]")
	while left(tparam,1)="%" or left(tparam,1)="_"
		tparam=right(tparam,len(tparam)-1)
	wend
	while right(tparam,1)="%" or right(tparam,1)="_"
		tparam=left(tparam,len(tparam)-1)
	wend
	if tparam<>"" then
		cn1.Execute Replace(sql_webdoadvdel_title_main,"{0}",tparam),affected,1
		Call TipsPage("已删除" &affected& "条留言。","web_advdel.asp")
	else
		Call TipsPage("不能输入空字符串或全部为通配符。","web_advdel.asp")
	end if
case "7"
	tparam=replace(tparam,"'","''")
	tparam=replace(tparam,"[","[\[]")
	while left(tparam,1)="%" or left(tparam,1)="_"
		tparam=right(tparam,len(tparam)-1)
	wend
	while right(tparam,1)="%" or right(tparam,1)="_"
		tparam=left(tparam,len(tparam)-1)
	wend
	if tparam<>"" then
		cn1.Execute Replace(sql_webdoadvdel_article_main,"{0}",tparam),affected,1
		Call TipsPage("已删除" &affected& "条留言。","web_advdel.asp")
	else
		Call TipsPage("不能输入空字符串或全部为通配符。","web_advdel.asp")
	end if
case "8"
	tparam=replace(tparam,"'","''")
	tparam=replace(tparam,"[","[\[]")
	while left(tparam,1)="%" or left(tparam,1)="_"
		tparam=right(tparam,len(tparam)-1)
	wend
	while right(tparam,1)="%" or right(tparam,1)="_"
		tparam=left(tparam,len(tparam)-1)
	wend
	if tparam<>"" then
		cn1.Execute Replace(sql_webdoadvdel_reply_main,"{0}",tparam),affected,1
		cn1.Execute Replace(sql_webdoadvdel_reply_reply,"{0}",tparam),,1
		Call TipsPage("已删除" &affected& "条留言。","web_advdel.asp")
	else
		Call TipsPage("不能输入空字符串或全部为通配符。","web_advdel.asp")
	end if
case "9"
	tparam=replace(tparam,"'","''")
	tparam=replace(tparam,"[","[\[]")
	while left(tparam,1)="%" or left(tparam,1)="_"
		tparam=right(tparam,len(tparam)-1)
	wend
	while right(tparam,1)="%" or right(tparam,1)="_"
		tparam=left(tparam,len(tparam)-1)
	wend
	if tparam<>"" then
		cn1.Execute Replace(sql_webdoadvdel_reply_unsetreply,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_reply_reply,"{0}",tparam),affected,1
		Call TipsPage("已删除" &affected& "条回复。","web_advdel.asp")
	else
		Call TipsPage("不能输入空字符串或全部为通配符。","web_advdel.asp")
	end if
case "10"
	cn1.execute sql_webdoadvdel_main,affected,1
	Call TipsPage("已删除" &affected& "条留言。","web_advdel.asp")
case "11"
	cn1.Execute sql_webdoadvdel_unsetreply,,1
	cn1.Execute sql_webdoadvdel_reply,affected,1
	Call TipsPage("已删除" &affected& "条回复。","web_advdel.asp")
case "12"
	tparam=replace(tparam,"'","''")
	tparam=replace(tparam,"[","[\[]")
	while left(tparam,1)="%" or left(tparam,1)="_"
		tparam=right(tparam,len(tparam)-1)
	wend
	while right(tparam,1)="%" or right(tparam,1)="_"
		tparam=left(tparam,len(tparam)-1)
	wend
	if tparam<>"" then
		cn1.Execute Replace(sql_webdoadvdel_deldeclare,"{0}",tparam),affected,1
		Call TipsPage("已删除" &affected& "条公告。","web_advdel.asp")
	else
		Call TipsPage("不能输入空字符串或全部为通配符。","web_advdel.asp")
	end if
case "13"
	cn1.Execute sql_webdoadvdel_cleardeclare,affected,1
	Call TipsPage("已删除" &affected& "条公告。","web_advdel.asp")
case "14"
	if isdate(tparam) then
		tparam=DateTimeStr(tparam)
		cn1.BeginTrans
		cn1.Execute Replace(sql_webdoadvdel_regdate_filterconfig,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_regdate_ipv4config,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_regdate_ipv6config,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_regdate_floodconfig,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_regdate_stats,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_regdate_stats_clientinfo,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_regdate_reply,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_regdate_main,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_regdate_config,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_regdate_supervisor,"{0}",tparam),affected,1
		cn1.CommitTrans
	
		Call TipsPage("已删除" &affected& "个用户。","web_advdel.asp")
	else
		Call TipsPage("您输入的日期有误，请检查。","web_advdel.asp")
	end if
case "15"
	if isdate(tparam) then
		tparam=DateTimeStr(tparam)
		cn1.BeginTrans
		cn1.Execute Replace(sql_webdoadvdel_lastlogin_filterconfig,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_lastlogin_ipv4config,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_lastlogin_ipv6config,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_lastlogin_floodconfig,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_lastlogin_stats,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_lastlogin_stats_clientinfo,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_lastlogin_reply,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_lastlogin_main,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_lastlogin_config,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_lastlogin_supervisor,"{0}",tparam),affected,1
		cn1.CommitTrans
	
		Call TipsPage("已删除" &affected& "个用户。","web_advdel.asp")
	else
		Call TipsPage("您输入的日期有误，请检查。","web_advdel.asp")
	end if
case "16"
	if isdate(tparam) then
		tparam=DateTimeStr(tparam)
		cn1.BeginTrans
		cn1.Execute Replace(Replace(sql_webdoadvdel_logdate_filterconfig,"{0}",tparam),"{1}",wm_name),,1
		cn1.Execute Replace(Replace(sql_webdoadvdel_logdate_ipv4config,"{0}",tparam),"{1}",wm_name),,1
		cn1.Execute Replace(Replace(sql_webdoadvdel_logdate_ipv6config,"{0}",tparam),"{1}",wm_name),,1
		cn1.Execute Replace(Replace(sql_webdoadvdel_logdate_floodconfig,"{0}",tparam),"{1}",wm_name),,1
		cn1.Execute Replace(sql_webdoadvdel_logdate_stats,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_logdate_stats_clientinfo,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_logdate_reply,"{0}",tparam),,1
		cn1.Execute Replace(sql_webdoadvdel_logdate_main,"{0}",tparam),,1
		cn1.Execute Replace(Replace(sql_webdoadvdel_logdate_config,"{0}",tparam),"{1}",wm_name),,1
		cn1.Execute Replace(sql_webdoadvdel_logdate_supervisor,"{0}",tparam),affected,1
		cn1.CommitTrans
	
		Call TipsPage("已删除" &affected& "个用户。","web_advdel.asp")
	else
		Call TipsPage("您输入的日期有误，请检查。","web_advdel.asp")
	end if
case "17"
		cn1.BeginTrans
		cn1.Execute sql_webdoadvdel_neverlogin_filterconfig,,1
		cn1.Execute sql_webdoadvdel_neverlogin_ipv4config,,1
		cn1.Execute sql_webdoadvdel_neverlogin_ipv6config,,1
		cn1.Execute sql_webdoadvdel_neverlogin_floodconfig,,1
		cn1.Execute sql_webdoadvdel_neverlogin_stats,,1
		cn1.Execute sql_webdoadvdel_neverlogin_stats_clientinfo,,1
		cn1.Execute sql_webdoadvdel_neverlogin_reply,,1
		cn1.Execute sql_webdoadvdel_neverlogin_main,,1
		cn1.Execute sql_webdoadvdel_neverlogin_config,,1
		cn1.Execute sql_webdoadvdel_neverlogin_supervisor,affected,1
		cn1.CommitTrans
	
		Call TipsPage("已删除" &affected& "个用户。","web_advdel.asp")
case "18"
	cn1.BeginTrans
	cn1.Execute Replace(sql_webdoadvdel_all_filterconfig,"{0}",wm_name),,1
	cn1.Execute Replace(sql_webdoadvdel_all_ipv4config,"{0}",wm_name),,1
	cn1.Execute Replace(sql_webdoadvdel_all_ipv6config,"{0}",wm_name),,1
	cn1.Execute Replace(sql_webdoadvdel_all_floodconfig,"{0}",wm_name),,1
	cn1.Execute sql_webdoadvdel_all_stats,,1
	cn1.Execute sql_webdoadvdel_all_stats_clientinfo,,1
	cn1.Execute sql_webdoadvdel_all_reply,,1
	cn1.Execute sql_webdoadvdel_all_main,,1
	cn1.Execute Replace(sql_webdoadvdel_all_config,"{0}",wm_name),,1
	cn1.Execute sql_webdoadvdel_all_supervisor,affected,1
	cn1.CommitTrans
	
	Call TipsPage("已删除" &affected& "个用户。","web_advdel.asp")
end select
cn1.Execute sql_webdoadvdel_clearfragment_main,,1
cn1.Execute sql_webdoadvdel_clearfragment_reply,,1
cn1.Execute sql_webdoadvdel_adjustadminreply_flag,,1
cn1.Execute sql_webdoadvdel_adjustguestreply_flag,,1
cn1.Close
set cn1=nothing
%>