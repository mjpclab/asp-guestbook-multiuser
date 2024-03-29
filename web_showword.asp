<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/web.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/const.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/common2.asp" -->
<!-- #include file="include/sql/web_admin_search.asp" -->
<!-- #include file="include/sql/web_admin_showword.asp" -->
<!-- #include file="include/sql/messagecondition.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/message.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"

dim id,ipage
ipage=Request("page")
id=FilterKeyword(Trim(Request.QueryString("id")))
if id="" Or Not Isnumeric(id) then
	Response.Redirect "web_search.asp" & get_param_str()
else
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	rs.Open sql_web_showword & id,cn,0,1,1
	if rs.EOF then		'留言不存在，退回主界面
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Response.Redirect "web_search.asp" & get_param_str()
	end if
end if
%>

<!-- #include file="include/template/dtd.inc" -->
<html lang="zh-CN">
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster管理中心 查看留言 <%=rs.Fields("title")%></title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>
<div id="outerborder" class="outerborder">

	<%Call WebInitHeaderData("","Webmaster管理中心","","")%><!-- #include file="include/template/header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->
	<form method="post" action="web_searchmdel.asp" name="form7">
		<!-- #include file="include/template/web_admin_func_search.inc" -->
		<%
			dim pagename, inAdminPage, inWebAdminPage
			pagename="web_showword"
			inAdminPage=true
			inWebAdminPage=true
			%><!-- #include file="include/template/web_admin_listword.inc" --><%
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
		%>

		<input type="hidden" name="rootid" value="<%=request("id")%>" />
		<input type="hidden" name="s_adminname" value="<%=request("s_adminname")%>" />
		<input type="hidden" name="s_name" value="<%=request("s_name")%>" />
		<input type="hidden" name="s_title" value="<%=request("s_title")%>" />
		<input type="hidden" name="s_article" value="<%=request("s_article")%>" />
		<input type="hidden" name="s_email" value="<%=request("s_email")%>" />
		<input type="hidden" name="s_qqid" value="<%=request("s_qqid")%>" />
		<input type="hidden" name="s_msnid" value="<%=request("s_msnid")%>" />
		<input type="hidden" name="s_homepage" value="<%=request("s_homepage")%>" />
		<input type="hidden" name="s_ipaddr" value="<%=request("s_ipaddr")%>" />
		<input type="hidden" name="s_originalip" value="<%=request("s_originalip")%>" />
		<input type="hidden" name="s_reply" value="<%=request("s_reply")%>" />
		<input type="hidden" name="page" value="<%=request("page")%>" />
		
		<!-- #include file="include/template/web_admin_func_search.inc" -->
	</form>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>