<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/const.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/common2.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/sql/topbulletin.asp" -->
<!-- #include file="include/sql/admin_showword.asp" -->
<!-- #include file="include/sql/admin_listword.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="include/utility/message.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"

Sub GoBack
	if Request("type")<>"" and Request("searchtxt")<>"" then
		Response.Redirect "admin_search.asp?user=" &ruser& "&page=" & Request("page") & "&type=" & server.URLEncode(Request("type")) & "&searchtxt=" & server.URLEncode(Request("searchtxt"))
	else
		Response.Redirect "admin.asp?user=" &ruser& "&page=" & Request("page")
	end if
	Response.End
End Sub

dim id,ipage
ipage=Request("page")
id=FilterKeyword(Trim(Request.QueryString("id")))
if id="" Or Not Isnumeric(id) then
	Call GoBack
else
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	rs.Open Replace(Replace(sql_admin_showword,"{0}",id),"{1}",adminid),cn,0,1,1
	if rs.EOF then
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Call Goback
	end if
end if
%>

<!-- #include file="include/template/dtd.inc" -->
<html lang="zh-CN">
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 管理首页 <%=rs.Fields("title")%></title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
	<script type="text/javascript" src="asset/js/jquery-1.x-min.js"></script>
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">
<div id="outerborder" class="outerborder">

	<%if ShowTitle then%><%Call InitHeaderData("管理")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<%dim sys_bul_flag
	sys_bul_flag=128%>
	<!-- #include file="include/template/sysbulletin.inc" -->
	<!-- #include file="include/template/topbulletin.inc" -->

	<form method="post" action="admin_mdel.asp" name="form7">
		<!-- #include file="include/template/admin_func.inc" -->
		<%
			dim pagename, inAdminPage, inWebAdminPage
			pagename="admin_showword"
			inAdminPage=true
			inWebAdminPage=false
			%><!-- #include file="include/template/admin_listword.inc" --><%
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
		%>

		<input type="hidden" name="rootid" value="<%=request("id")%>" />
		<input type="hidden" name="user" value="<%=ruser%>" />
		<input type="hidden" name="page" value="<%=Request.QueryString("page")%>" />
		<input type="hidden" name="type" value="<%=Request.QueryString("type")%>" />
		<input type="hidden" name="searchtxt" value="<%=Request.QueryString("searchtxt")%>" />
		<!-- #include file="include/template/admin_func.inc" -->
	</form>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>