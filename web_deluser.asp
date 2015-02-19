<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1

dim arr,users,affected,i
affected=0

users=replace(Request.Form("users"),"'","")
users=replace(users,";","")
users=replace(users,"#","")
users=replace(users,"%","")
users=replace(users,"&","")
users=replace(users,"=","")
users=replace(users," ","")
users=replace(users,"[","")
users=replace(users,"]","")
users=replace(users,"\","")

arr=split(users,",")
users=""
for i=0 to ubound(arr)
	if i>0 then
		users=users & ",'" & replace(replace(arr(i),"'",""),";","") & "'"
	else
		users="'" & replace(replace(arr(i),"'",""),";","") & "'"
	end if
next

if users<>"" then
	set cn=server.CreateObject("ADODB.Connection")
	CreateConn cn,dbtype
	
	cn.BeginTrans
	cn.Execute Replace(sql_webdeluser_filterconfig,"{0}",users),,1
	cn.Execute Replace(sql_webdeluser_ipconfig,"{0}",users),,1
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

Call MessagePage("已删除" &affected& "个用户。",backurl)
%>