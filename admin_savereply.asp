<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/common2.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_savereply.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/mail.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if isnumeric(request.form("mainid"))=false or request.form("mainid")="" then
	Response.Redirect "admin.asp?user=" &ruser
	Response.End
end if

if Not IsEmpty(Request.Form) then
	dim tlimit
	tlimit=0
	if Request.Form("html1")="1" then tlimit=tlimit+1
	if Request.Form("ubb1")="1" then tlimit=tlimit+2
	if Request.Form("newline1")="1" then tlimit=tlimit+4

	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	rs.Open Replace(Replace(sql_adminsavereply_main,"{0}",Request.Form("mainid")),"{1}",adminid),cn,0,3,1
	if Not rs.EOF then		'留言存在
		cn.BeginTrans
		rs("replied")=clng(rs("replied") OR 1)
		if clng(rs("guestflag") and 16)<>0 then rs("guestflag")=rs("guestflag")-16	'通过审核
		rs.Update
		rs.Close

		replydate1=DateTimeStr(ServerTimeToUTC(now()))
		content1=Replace(Replace(Request.Form("rcontent"),"'","''"),"<%","< %")
		rs.Open sql_adminsavereply_reply & Request.Form("mainid"),cn,0,1,1
		if rs.EOF then	'新回复
			rs.Close
			cn.Execute Replace(Replace(Replace(Replace(sql_adminsavereply_insert,"{0}",Request.Form("mainid")),"{1}",replydate1),"{2}",clng(web_adminlimit and tlimit)),"{3}",content1),,129
		else	'更新回复
			rs.Close
			cn.Execute Replace(Replace(Replace(Replace(sql_adminsavereply_update,"{0}",replydate1),"{1}",clng(web_adminlimit and tlimit)),"{2}",content1),"{3}",Request.Form("mainid")),,129
		end if
		cn.CommitTrans

		if Request.Form("lock2top")="1" then	'置顶
			cn.Execute Replace(Replace(sql_adminsavereply_lock2top,"{0}",Request.Form("mainid")),"{1}",adminid),,129
		end if

		if Request.Form("bring2top")="1" then	'提前
			cn.Execute Replace(Replace(Replace(sql_adminsavereply_bring2top,"{0}",replydate1),"{1}",Request.Form("mainid")),"{2}",adminid),,129
		end if

		cn.close
		set rs=nothing
		set cn=nothing

		if web_MailReplyInform and MailReplyInform then replyinform()
	else
		rs.close : cn.close : set rs=nothing : set cn=nothing
	end if
end if
%><!-- #include file="include/template/admin_traceback.inc" -->