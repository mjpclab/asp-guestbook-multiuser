<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 管理首页</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
	<!-- #include file="style.asp" -->
	<!-- #include file="adminstyle.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<%
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

dim id,ipage
ipage=Request("page")
if isnumeric(Request.QueryString("id")) and Request.QueryString("id")<>"" then
	id=Request.QueryString("id")
else
	id=0
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
rs.Open Replace(Replace(sql_admin_showword,"{0}",id),"{1}",Request.QueryString("user")),cn,0,1,1
if rs.EOF then		'留言不存在，退回主界面
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	if Request("type")<>"" and Request("searchtxt")<>"" then
		Response.Redirect "admin_search.asp?user=" &Request("user")& "&page=" & Request("page") & "&type=" & server.URLEncode(Request("type")) & "&searchtxt=" & server.URLEncode(Request("searchtxt"))
	else
		Response.Redirect "admin.asp?user=" &Request("user")& "&page=" & Request("page")
	end if
	Response.End
end if
%>

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admincontrols.inc" -->
	<!-- #include file="topbulletin.inc" -->
	
	<%dim sys_bul_flag
	sys_bul_flag=128%>
	<!-- #include file="sysbulletin.inc" -->
	<!-- #include file="topbulletin.inc" -->
	<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"admin.asp","[留言分页]","center",""%>

	<form method="post" action="admin_mdel.asp" name="form7">
		<!-- #include file="func_admin.inc" -->
		<%
			dim pagename
			pagename="admin_showword"
			%><!-- #include file="listword_admin.inc" --><%
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
		%>

		<input type="hidden" name="rootid" value="<%=request("id")%>" />
		<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
		<input type="hidden" name="page" value="<%=Request.QueryString("page")%>" />
		<input type="hidden" name="type" value="<%=Request.QueryString("type")%>" />
		<input type="hidden" name="searchtxt" value="<%=Request.QueryString("searchtxt")%>" />
		<!-- #include file="func_admin.inc" -->
	</form>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>