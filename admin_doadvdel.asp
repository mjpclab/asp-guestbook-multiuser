<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

Dim affected
affected=0

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,true
set rs=nothing

dim tparam
tparam=Request.Form("iparam")
select case Request.Form("option")
case "1"
	if isdate(tparam) then
		tparam=DateTimeStr(tparam)
		cn.execute Replace(Replace(sql_admindoadvdel_beforedate_main,"{0}",tparam),"{1}",adminid),affected,1
		Call MessagePage("已删除" &affected& "条留言。","admin_advdel.asp?user=" & ruser)
	else
		Call MessagePage("您输入的日期有误，请检查。","admin_advdel.asp?user=" & ruser)
	end if
case "2"
	if isdate(tparam) then
		tparam=DateTimeStr(tparam)
		cn.execute Replace(Replace(sql_admindoadvdel_afterdate_main,"{0}",tparam),"{1}",adminid),affected,1
		Call MessagePage("已删除" &affected& "条留言。","admin_advdel.asp?user=" & ruser)
	else
		Call MessagePage("您输入的日期有误，请检查。","admin_advdel.asp?user=" & ruser)
	end if
case "3"
	if isnumeric(tparam) then
		if clng(tparam)>0 then
			tparam=clng(tparam)
			cn.execute Replace(Replace(sql_admindoadvdel_firstn_main,"{0}",tparam),"{1}",adminid),affected,1
			Call MessagePage("已删除" &affected& "条留言。","admin_advdel.asp?user=" & ruser)
		else
			Call MessagePage("请输入正确的数值。","admin_advdel.asp?user=" & ruser)
		end if
	else
		Call MessagePage("请输入正确的数值。","admin_advdel.asp?user=" & ruser)
	end if
case "4"
	if isnumeric(tparam) then
		if clng(tparam)>0 then
			tparam=clng(tparam)
			cn.execute Replace(Replace(sql_admindoadvdel_lastn_main,"{0}",tparam),"{1}",adminid),affected,1
			Call MessagePage("已删除" &affected& "条留言。","admin_advdel.asp?user=" & ruser)
		else
			Call MessagePage("请输入正确的数值。","admin_advdel.asp?user=" & ruser)
		end if
	else
		Call MessagePage("请输入正确的数值。","admin_advdel.asp?user=" & ruser)
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
		cn.execute Replace(Replace(sql_admindoadvdel_name_main,"{0}",tparam),"{1}",adminid),affected,1
		Call MessagePage("已删除" &affected& "条留言。","admin_advdel.asp?user=" & ruser)
	else
		Call MessagePage("不能输入空字符串或全部为通配符。","admin_advdel.asp?user=" & ruser)
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
		cn.execute Replace(Replace(sql_admindoadvdel_title_main,"{0}",tparam),"{1}",adminid),affected,1
		Call MessagePage("已删除" &affected& "条留言。","admin_advdel.asp?user=" & ruser)
	else
		Call MessagePage("不能输入空字符串或全部为通配符。","admin_advdel.asp?user=" & ruser)
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
		cn.execute Replace(Replace(sql_admindoadvdel_article_main,"{0}",tparam),"{1}",adminid),affected,1
		Call MessagePage("已删除" &affected& "条留言。","admin_advdel.asp?user=" & ruser)
	else
		Call MessagePage("不能输入空字符串或全部为通配符。","admin_advdel.asp?user=" & ruser)
	end if
case "8"
	cn.Execute Replace(sql_admindoadvdel_main,"{0}",adminid),affected,1
	Call MessagePage("已删除" &affected& "条留言。","admin_advdel.asp?user=" & ruser)
end select

cn.Execute Replace(sql_admindoadvdel_adjustguestreply_flag,"{0}",adminid),,1
cn.Execute sql_admindoadvdel_clearfragment_main,,1
cn.Execute sql_admindoadvdel_clearfragment_reply,,1
cn.Close : set cn=nothing		
%>