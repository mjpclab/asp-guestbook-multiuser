<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_deluser.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" --
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1

dim arr,users,affected,i
affected=0

users=FilterSql(Request.Form("users"))
if Len(users)>0 then
	users="'" & Replace(users,",","','") & "'"

	set cn=server.CreateObject("ADODB.Connection")
	CreateConn cn,dbtype
	
	cn.BeginTrans
	cn.Execute Replace(sql_webdeluser_filterconfig,"{0}",users),,1
	cn.Execute Replace(sql_webdeluser_ipv4config,"{0}",users),,1
	cn.Execute Replace(sql_webdeluser_ipv6config,"{0}",users),,1
	cn.Execute Replace(sql_webdeluser_floodconfig,"{0}",users),,1
	cn.Execute Replace(sql_webdeluser_stats,"{0}",users),,1
	cn.Execute Replace(sql_webdeluser_stats_clientinfo,"{0}",users),,1
	cn.Execute Replace(sql_webdeluser_reply,"{0}",users),,1
	cn.Execute Replace(sql_webdeluser_main,"{0}",users),,1
	cn.Execute Replace(sql_webdeluser_config,"{0}",users),,1
	cn.Execute Replace(sql_webdeluser_supervisor,"{0}",users),affected,1
	cn.CommitTrans
	
	cn.Close
	set cn=nothing
end if

dim backurl
if Request.Form("source")=1 then
	backurl = "web_admin.asp"
else
	backurl = "web_searchuser.asp"
end if
if Request.Form("arguments")<>"" then backurl=backurl & "?" & Request.Form("arguments")

Call TipsPage("已删除" &affected& "个用户。",backurl)
%>