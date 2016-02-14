<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/const.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/common2.asp" -->
<!-- #include file="include/sql/web_admin_searchdec.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/message.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster管理中心 搜索置顶公告</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<%
function filterstr(byref str,byval delquote)
	if delquote then
		filterstr=replace(replace(replace(str,"'",""),"[","[\[]"),"""","")
	else
		filterstr=replace(replace(replace(str,"'","''"),"[","[\[]"),"""","")
	end if
end function


'<ShowArticle>
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
'====================================

	'=============分页控制=============
	Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
	ItemsCount=0:PagesCount=0
	
	dim count_sql,sql
	count_sql=Replace(Replace(sql_websearchdec_words_count,"{0}",filterstr(Request("adminname"),true)),"{1}",filterstr(Request("searchtxt"),false))
	sql=Replace(Replace(sql_websearchdec_words_query,"{0}",filterstr(Request("adminname"),true)),"{1}",filterstr(Request("searchtxt"),false))
	
	get_divided_page cn,rs,sql_pk_supervisor,count_sql,sql,"adminname INC",Request.QueryString("page"),ItemsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
	'==============================
%>

<div id="outerborder" class="outerborder">

	<%Call WebInitHeaderData("","Webmaster管理中心","","")%><!-- #include file="include/template/header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->
	<div class="region form-region">
		<h3 class="title">搜索置顶公告</h3>
		<div class="content">
			<form method="post" action="web_searchdec.asp" name="form1">
			<p>搜索：("%"代表任意个字符，"_"代表一个字符)</p>
			<div class="field">
				<span class="label">用户名：</span>
				<span class="value"><input type="text" name="adminname" value="<%=Request("adminname")%>" size="<%=SearchTextWidth%>" /></span>
			</div>
			<div class="field">
				<span class="label">公告内容：</span>
				<span class="value"><input type="text" name="searchtxt" value="<%=Request("searchtxt")%>" size="<%=SearchTextWidth%>" /></span>
			</div>
			<div class="command"><input type="submit" value="搜索公告" name="searchsubmit" /></div>
			</form>
		</div>
	</div>

<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"web_searchdec.asp","[搜索结果分页]","adminname=" &server.URLEncode(Request("adminname"))& "&searchtxt=" & server.URLEncode(Request("searchtxt")) end if%>
<form method="post" action="web_mdeldec.asp" name="form7">

<input type="hidden" name="adminname" value="<%=request("adminname")%>" />
<input type="hidden" name="searchtxt" value="<%=request("searchtxt")%>" />
<input type="hidden" name="page" value="<%if isnumeric(request("page")) and request("page")<>"" then response.write request("page")%>" />
<div class="guest-functions">
	<div class="main">
		<a class="function-multi-del" href="javascript:void 0;"><img src="asset/image/icon_mdel.gif" />删除选定公告</a>
	</div>
</div>


<%
dim rsuser,pub_flag,pub
if ItemsCount=0 then
	Response.Write "<br/><br/><div class=""centertext"">没有找到符合条件的公告。</div><br/><br/>"
else
	while Not rs.EOF
	rsuser=rs("adminname")
%>

	<div class="region region-bulletin">
		<h3 class="title">所属用户：<a href="web_userinfo.asp?user=<%=rsuser%>" target="_blank"><%=rsuser%></a></h3>
		<div class="content">
			<%
			pub_flag=rs("declareflag")
			pub="" &rs("declare")& ""

			if AdminViewCode then		'Only view HTML code
				pub=server.HTMLEncode(pub)
			else
				convertstr pub,pub_flag,2
			end if

			Response.Write pub
			%>
		</div>
		<div class="admin-message-tools">
			<input type="checkbox" name="users" class="users checkbox" id="c<%=rsuser%>" value="<%=rsuser%>"><label for="c<%=rsuser%>">(选定)</label>
			<a href="web_deldec.asp?user=<%=rsuser%><%if isnumeric(request("page")) and request("page")<>"" then response.write "&page=" & request("page")%>&adminname=<%=server.URLEncode(request("adminname"))%>&searchtxt=<%=server.URLEncode(request("searchtxt"))%>" title="删除公告"<%if DelDecTip then Response.Write " onclick=""return confirm('确实要删除公告吗？');"""%>><img border="0" src="asset/image/icon_del.gif" class="imgicon" />[删除公告]</a>
		</div>
	</div>
<%
	rs.MoveNext
	wend
end if	'对应for上面一行的if
%>
<div class="guest-functions">
	<div class="main">
		<a class="function-multi-del" href="javascript:void 0;"><img src="asset/image/icon_mdel.gif" />删除选定公告</a>
	</div>
</div>

</form>

<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"web_searchdec.asp","[搜索结果分页]","adminname=" &server.URLEncode(Request("adminname"))& "&searchtxt=" & server.URLEncode(Request("searchtxt")) end if%>
</div>

<!-- #include file="include/template/footer.inc" -->
</div>
<script type="text/javascript" src="asset/js/jquery-1.x-min.js"></script>
<script type="text/javascript">
	(function multiDelConfirm(){
		var $multiDelLink = $('.function-multi-del');
		$multiDelLink.click(function () {
			var delSelDecTip = <%=LCase(CStr(DelSelDecTip))%>;
			var $selected = $('input.users:checked');
			if (!$selected.length) {
				alert('请先选定要删除的公告。');
			} else if (!delSelDecTip || confirm('确实要删除选定公告吗？')) {
				form7.submit();
			}
		});
	}());
</script>
</body>
</html>
<%
	if ItemsCount<>0 then rs.Close
	cn.Close
	set rs=nothing
	set cn=nothing
%>