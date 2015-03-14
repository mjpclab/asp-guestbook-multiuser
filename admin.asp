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
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<%
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,false

Dim WordsPerPage
if AdminDisplayMode()="book" then
	WordsPerPage=ItemsPerPage
elseif AdminDisplayMode()="forum" then
	WordsPerPage=TitlesPerPage
end if

Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
get_divided_page cn,rs,sql_pk_main,Replace(sql_admin_words_count,"{0}",Request.QueryString("user")),Replace(sql_admin_words_query,"{0}",Request.QueryString("user")),"parent_id INC,lastupdated DEC,id DEC",Request.QueryString("page"),WordsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
%>

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admincontrols.inc" -->
	
	<%dim sys_bul_flag
	sys_bul_flag=128%>
	<!-- #include file="sysbulletin.inc" -->
	<!-- #include file="topbulletin.inc" -->
	<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"admin.asp","[留言分页]","center",""%>

	<form method="post" action="admin_mdel.asp" name="form7">
		<%RPage="admin.asp"%><!-- #include file="func_admin.inc" -->
		<%
		if ItemsCount=0 then
			Response.Write "<br/><br/><div class=""centertext"">目前尚无留言。</div><br/><br/>"
		else
			dim pagename
			pagename="admin"
			if AdminDisplayMode()="book" then
				%><!-- #include file="listword_admin.inc" --><%
			elseif AdminDisplayMode()="forum" then
				%><!-- #include file="listtitle_admin.inc" --><%
			end if
			rs.Close
		end if
		cn.Close : set rs=nothing : set cn=nothing%>

		<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
		<input type="hidden" name="page" value="<%=Request.QueryString("page")%>" />
		<!-- #include file="func_admin.inc" -->
	</form>

	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"admin.asp","[留言分页]","center",""%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>