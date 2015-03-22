<!-- #include file="webconfig.asp" -->
<!-- #include file="web_common.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster管理中心 搜索留言</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<link rel="stylesheet" type="text/css" href="css/web_adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
	<!-- #include file="css/web_adminstyle.asp" -->
</head>

<body>

<%
'<ShowArticle>
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

'==================================================
'连接Where子句

sql_condition=""
param_str=""

if request("s_adminname")<>"" then	'用户名
	sql_condition=sql_condition & Replace(sql_websearch_condition_adminname,"{0}",FilterAdminLike(request("s_adminname")))
	param_str=param_str & "&s_adminname=" & server.URLEncode(request("s_adminname"))
end if

if request("s_name")<>"" then	'访客名
	sql_condition=sql_condition & Replace(sql_websearch_condition_name,"{0}",FilterAdminLike(request("s_name")))
	param_str=param_str & "&s_name=" & server.URLEncode(request("s_name"))
end if

if request("s_title")<>"" then	'标题
	sql_condition=sql_condition & Replace(sql_websearch_condition_title,"{0}",FilterAdminLike(request("s_title")))
	param_str=param_str & "&s_title=" & server.URLEncode(request("s_title"))
end if

if request("s_article")<>"" then	'留言内容
	sql_condition=sql_condition & Replace(sql_websearch_condition_article,"{0}",FilterAdminLike(request("s_article")))
	param_str=param_str & "&s_article=" & server.URLEncode(request("s_article"))
end if

if request("s_email")<>"" then	'邮件
	sql_condition=sql_condition & Replace(sql_websearch_condition_email,"{0}",FilterAdminLike(request("s_email")))
	param_str=param_str & "&s_email=" & server.URLEncode(request("s_email"))
end if

if request("s_qqid")<>"" then	'QQ号码
	sql_condition=sql_condition & Replace(sql_websearch_condition_qqid,"{0}",FilterAdminLike(request("s_qqid")))
	param_str=param_str & "&s_qqid=" & server.URLEncode(request("s_qqid"))
end if

if request("s_msnid")<>"" then	'Skype号
	sql_condition=sql_condition & Replace(sql_websearch_condition_msnid,"{0}",FilterAdminLike(request("s_msnid")))
	param_str=param_str & "&s_msnid=" & server.URLEncode(request("s_msnid"))
end if

if request("s_homepage")<>"" then	'主页
	sql_condition=sql_condition & Replace(sql_websearch_condition_homepage,"{0}",FilterAdminLike(request("s_homepage")))
	param_str=param_str & "&s_homepage=" & server.URLEncode(request("s_homepage"))
end if

if request("s_ipaddr")<>"" then	'IP
	sql_condition=sql_condition & Replace(sql_websearch_condition_ipaddr,"{0}",FilterAdminLike(request("s_ipaddr")))
	param_str=param_str & "&s_ipaddr=" & server.URLEncode(request("s_ipaddr"))
end if

if request("s_originalip")<>"" then	'原始IP
	sql_condition=sql_condition & Replace(sql_websearch_condition_originalip,"{0}",FilterAdminLike(request("s_originalip")))
	param_str=param_str & "&s_originalip=" & server.URLEncode(request("s_originalip"))
end if

if request("s_reply")<>"" then	'回复内容
	sql_condition=sql_condition & Replace(sql_websearch_condition_reply,"{0}",FilterAdminLike(request("s_reply")))
	param_str=param_str & "&s_reply=" & server.URLEncode(request("s_reply"))
end if


'====================================
sql_condition=sql_websearch_condition_init & sql_condition

	'=============分页控制=============

	Dim WordsPerPage
	if WebDisplayMode()="book" then
		WordsPerPage=ItemsPerPage
	elseif WebDisplayMode()="forum" then
		WordsPerPage=TitlesPerPage
	end if

	Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
	ItemsCount=0:PagesCount=0

	sql_count=sql_websearch_count_inner & sql_condition
	sql_count=Replace(sql_websearch_count,"{0}",sql_count)
	sql_full=sql_websearch_full_inner & sql_condition
	sql_full=Replace(sql_websearch_full,"{0}",sql_full)

	get_divided_page cn,rs,sql_pksearch_main,sql_count,sql_full,"lastupdated DEC,id DEC",Request.QueryString("page"),WordsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
%>

<div id="outerborder" class="outerborder">
	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admincontrols.inc" -->

<div class="region region-search-message">
	<h3 class="title">搜索留言</h3>
	<div class="content">
		<form method="post" action="web_search.asp" name="form1">
		<p>搜索：("%"代表任意个字符，"_"代表一个字符)</p>
		<div class="field">
			<span class="label">用户名：</span>
			<span class="value"><input type="text" name="s_adminname" value="<%=request("s_adminname")%>" size="<%=SearchTextWidth%>" maxlength="32" /></span>
		</div>
		<div class="field">
			<span class="label">访客姓名：</span>
			<span class="value"><input type="text" name="s_name" value="<%=request("s_name")%>" size="<%=SearchTextWidth%>" maxlength="64" /></span>
		</div>
		<div class="field">
			<span class="label">留言标题：</span>
			<span class="value"><input type="text" name="s_title" value="<%=request("s_title")%>" size="<%=SearchTextWidth%>" maxlength="64" /></span>
		</div>
		<div class="field">
			<span class="label">留言内容：</span>
			<span class="value"><input type="text" name="s_article" value="<%=request("s_article")%>" size="<%=SearchTextWidth%>" /></span>
		</div>
		<div class="field">
			<span class="label">访客邮件地址：</span>
			<span class="value"><input type="text" name="s_email" value="<%=request("s_email")%>" size="<%=SearchTextWidth%>" maxlength="48" /></span>
		</div>
		<div class="field">
			<span class="label">访客QQ号码：</span>
			<span class="value"><input type="text" name="s_qqid" value="<%=request("s_qqid")%>" size="<%=SearchTextWidth%>" maxlength="16" /></span>
		</div>
		<div class="field">
			<span class="label">访客Skype地址：</span>
			<span class="value"><input type="text" name="s_msnid" value="<%=request("s_msnid")%>" size="<%=SearchTextWidth%>" maxlength="48" /></span>
		</div>
		<div class="field">
			<span class="label">访客主页地址：</span>
			<span class="value"><input type="text" name="s_homepage" value="<%=request("s_homepage")%>" size="<%=SearchTextWidth%>" maxlength="255" /></span>
		</div>
		<div class="field">
			<span class="label">访客IP地址：</span>
			<span class="value"><input type="text" name="s_ipaddr" value="<%=request("s_ipaddr")%>" size="<%=SearchTextWidth%>" maxlength="15" /></span>
		</div>
		<div class="field">
			<span class="label">访客原IP地址：</span>
			<span class="value"><input type="text" name="s_originalip" value="<%=request("s_originalip")%>" size="<%=SearchTextWidth%>" maxlength="15" /></span>
		</div>
		<div class="field">
			<span class="label">版主回复内容：</span>
			<span class="value"><input type="text" name="s_reply" value="<%=request("s_reply")%>" size="<%=SearchTextWidth%>" /></span>
		</div>
		<div class="command"><input type="submit" value="搜索留言" name="searchsubmit" /></div>
		</form>
	</div>
</div>

	<%
	if PagesCount>1 and ShowTopPageList then
		show_page_list ipage,PagesCount,"web_search.asp","[搜索结果分页]",param_str
	end if
	%>
	<form method="post" action="web_searchmdel.asp" name="form7">

	<%RPage="web_search.asp"%><!-- #include file="func_web_search.inc" -->
	
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

	<%
	if ItemsCount=0 then
		Response.Write "<br/><br/><div class=""centertext"">没有找到符合条件的留言。</div><br/><br/>"
	else
		dim pagename
		pagename="admin_web_search"
		if WebDisplayMode()="book" then
			%><!-- #include file="listword_web.inc" --><%
		elseif WebDisplayMode()="forum" then
			%><!-- #include file="listtitle_web.inc" --><%
		end if
		rs.Close
	end if
	%>

<!-- #include file="func_web_search.inc" -->
</form>

<%
if PagesCount>1 and ShowBottomPageList then
	show_page_list ipage,PagesCount,"web_search.asp","[搜索结果分页]",param_str
end if
%>

</div>

<% cn.Close : set rs=nothing : set cn=nothing %>

<!-- #include file="bottom.asp" -->
</body>
</html>