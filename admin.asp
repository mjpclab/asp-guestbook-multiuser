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
<!-- #include file="include/sql/admin.asp" -->
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
%>

<!-- #include file="include/template/dtd.inc" -->
<html lang="zh-CN">
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 管理首页</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
	<script type="text/javascript" src="asset/js/jquery-1.x-min.js"></script>
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

Dim WordsPerPage
if AdminDisplayMode()="book" then
	WordsPerPage=ItemsPerPage
elseif AdminDisplayMode()="forum" then
	WordsPerPage=TitlesPerPage
end if

Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
get_divided_page cn,rs,sql_pk_main,Replace(sql_admin_words_count,"{0}",adminid),Replace(sql_admin_words_query,"{0}",adminid),"parent_id INC,lastupdated DEC,id DEC",Request.QueryString("page"),WordsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
%>

<div id="outerborder" class="outerborder">

	<%if ShowTitle then%><%Call InitHeaderData("管理")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<%dim sys_bul_flag
	sys_bul_flag=128%>
	<!-- #include file="include/template/sysbulletin.inc" -->
	<!-- #include file="include/template/topbulletin.inc" -->
	<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"[留言分页]",""%>

	<form method="post" action="admin_mdel.asp" name="form7">
		<%RPage="admin.asp"%><!-- #include file="include/template/admin_func.inc" -->
		<%
		if ItemsCount=0 then
			Response.Write "<br/><br/><div class=""centertext"">目前尚无留言。</div><br/><br/>"
		else
			dim pagename, inAdminPage, inWebAdminPage
			pagename="admin"
			inAdminPage=true
			inWebAdminPage=false
			if AdminDisplayMode()="book" then
				%><!-- #include file="include/template/admin_listword.inc" --><%
			elseif AdminDisplayMode()="forum" then
				%><!-- #include file="include/template/admin_listtitle.inc" --><%
			end if
			rs.Close
		end if
		cn.Close : set rs=nothing : set cn=nothing%>

		<input type="hidden" name="user" value="<%=ruser%>" />
		<input type="hidden" name="page" value="<%=Request.QueryString("page")%>" />
		<!-- #include file="include/template/admin_func.inc" -->
	</form>

	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"[留言分页]",""%>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>