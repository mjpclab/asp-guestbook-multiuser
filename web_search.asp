<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="web_common.asp" -->
<!-- #include file="common2.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
	<title><%=web_BookName%> Webmaster�������� ��������</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<%
'<ShowArticle>
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

'==================================================
'����Where�Ӿ�

sql_condition=""
param_str=""

if Len(Request("s_adminname"))>0 then	'�û���
	sql_condition=sql_condition & Replace(sql_websearch_condition_adminname,"{0}",FilterAdminLike(Request("s_adminname")))
	param_str=param_str & "&s_adminname=" & server.URLEncode(Request("s_adminname"))
end if

if Len(Request("s_name"))>0 then	'�ÿ���
	sql_condition=sql_condition & Replace(sql_websearch_condition_name,"{0}",FilterAdminLike(Request("s_name")))
	param_str=param_str & "&s_name=" & server.URLEncode(Request("s_name"))
end if

if Len(Request("s_title"))>0 then	'����
	sql_condition=sql_condition & Replace(sql_websearch_condition_title,"{0}",FilterAdminLike(Request("s_title")))
	param_str=param_str & "&s_title=" & server.URLEncode(Request("s_title"))
end if

if Len(Request("s_article"))>0 then	'��������
	sql_condition=sql_condition & Replace(sql_websearch_condition_article,"{0}",FilterAdminLike(Request("s_article")))
	param_str=param_str & "&s_article=" & server.URLEncode(Request("s_article"))
end if

if Len(Request("s_email"))>0 then	'�ʼ�
	sql_condition=sql_condition & Replace(sql_websearch_condition_email,"{0}",FilterAdminLike(Request("s_email")))
	param_str=param_str & "&s_email=" & server.URLEncode(Request("s_email"))
end if

if Len(Request("s_qqid"))>0 then	'QQ����
	sql_condition=sql_condition & Replace(sql_websearch_condition_qqid,"{0}",FilterAdminLike(Request("s_qqid")))
	param_str=param_str & "&s_qqid=" & server.URLEncode(Request("s_qqid"))
end if

if Len(Request("s_msnid"))>0 then	'Skype��
	sql_condition=sql_condition & Replace(sql_websearch_condition_msnid,"{0}",FilterAdminLike(Request("s_msnid")))
	param_str=param_str & "&s_msnid=" & server.URLEncode(Request("s_msnid"))
end if

if Len(Request("s_homepage"))>0 then	'��ҳ
	sql_condition=sql_condition & Replace(sql_websearch_condition_homepage,"{0}",FilterAdminLike(Request("s_homepage")))
	param_str=param_str & "&s_homepage=" & server.URLEncode(Request("s_homepage"))
end if

if Len(Request("s_ipv4addr"))>0 then	'IPv4
	sql_condition=sql_condition & Replace(sql_websearch_condition_ipv4addr,"{0}",FilterAdminLike(Request("s_ipv4addr")))
	param_str=param_str & "&s_ipv4addr=" & server.URLEncode(Request("s_ipv4addr"))
end if

if Len(Request("s_originalipv4"))>0 then	'ԭʼIPv4
	sql_condition=sql_condition & Replace(sql_websearch_condition_originalipv4,"{0}",FilterAdminLike(Request("s_originalipv4")))
	param_str=param_str & "&s_originalipv4=" & server.URLEncode(Request("s_originalipv4"))
end if

if Len(Request("s_ipv6addr"))>0 then	'IPv6
	sql_condition=sql_condition & Replace(sql_websearch_condition_ipv6addr,"{0}",FilterAdminLike(Request("s_ipv6addr")))
	param_str=param_str & "&s_ipv6addr=" & server.URLEncode(Request("s_ipv6addr"))
end if

if Len(Request("s_originalipv6"))>0 then	'ԭʼIPv6
	sql_condition=sql_condition & Replace(sql_websearch_condition_originalipv6,"{0}",FilterAdminLike(Request("s_originalipv6")))
	param_str=param_str & "&s_originalipv6=" & server.URLEncode(Request("s_originalipv6"))
end if

if Len(Request("s_reply"))>0 then	'�ظ�����
	sql_condition=sql_condition & Replace(sql_websearch_condition_reply,"{0}",FilterAdminLike(Request("s_reply")))
	param_str=param_str & "&s_reply=" & server.URLEncode(Request("s_reply"))
end if


'====================================
sql_condition=sql_websearch_condition_init & sql_condition

	'=============��ҳ����=============

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
	<!-- #include file="include/web_admin_title.inc" -->
	<!-- #include file="include/web_admin_mainmenu.inc" -->

<div class="region region-longtext region-search-message">
	<h3 class="title">��������</h3>
	<div class="content">
		<form method="post" action="web_search.asp" name="form1">
		<p>������("%"����������ַ���"_"����һ���ַ�)</p>
		<div class="field">
			<span class="label">�û�����</span>
			<span class="value"><input type="text" name="s_adminname" value="<%=Request("s_adminname")%>" maxlength="32" /></span>
		</div>
		<div class="field">
			<span class="label">�ÿ�������</span>
			<span class="value"><input type="text" name="s_name" value="<%=Request("s_name")%>" maxlength="64" /></span>
		</div>
		<div class="field">
			<span class="label">���Ա��⣺</span>
			<span class="value"><input type="text" name="s_title" value="<%=Request("s_title")%>" maxlength="64" /></span>
		</div>
		<div class="field">
			<span class="label">�������ݣ�</span>
			<span class="value"><input type="text" name="s_article" value="<%=Request("s_article")%>" /></span>
		</div>
		<div class="field">
			<span class="label">�ÿ��ʼ���ַ��</span>
			<span class="value"><input type="text" name="s_email" value="<%=Request("s_email")%>" maxlength="48" /></span>
		</div>
		<div class="field">
			<span class="label">�ÿ�QQ���룺</span>
			<span class="value"><input type="text" name="s_qqid" value="<%=Request("s_qqid")%>" maxlength="16" /></span>
		</div>
		<div class="field">
			<span class="label">�ÿ�Skype��ַ��</span>
			<span class="value"><input type="text" name="s_msnid" value="<%=Request("s_msnid")%>" maxlength="48" /></span>
		</div>
		<div class="field">
			<span class="label">�ÿ���ҳ��ַ��</span>
			<span class="value"><input type="text" name="s_homepage" value="<%=Request("s_homepage")%>" maxlength="255" /></span>
		</div>
		<div class="field">
			<span class="label">�ÿ�IPv4��ַ��</span>
			<span class="value"><input type="text" name="s_ipv4addr" value="<%=Request("s_ipv4addr")%>" maxlength="15" /></span>
		</div>
		<div class="field">
			<span class="label">�ÿ�ԭIPv4��ַ��</span>
			<span class="value"><input type="text" name="s_originalipv4" value="<%=Request("s_originalipv4")%>" maxlength="15" /></span>
		</div>
		<div class="field">
			<span class="label">�ÿ�IPv6��ַ��</span>
			<span class="value"><input type="text" name="s_ipv6addr" value="<%=Request("s_ipv6addr")%>" maxlength="15" /></span>
		</div>
		<div class="field">
			<span class="label">�ÿ�ԭIPv6��ַ��</span>
			<span class="value"><input type="text" name="s_originalipv6" value="<%=Request("s_originalipv6")%>" maxlength="15" /></span>
		</div>
		<div class="field">
			<span class="label">�����ظ����ݣ�</span>
			<span class="value"><input type="text" name="s_reply" value="<%=request("s_reply")%>" /></span>
		</div>
		<div class="command"><input type="submit" value="��������" name="searchsubmit" /></div>
		</form>
	</div>
</div>

	<%
	if PagesCount>1 and ShowTopPageList then
		show_page_list ipage,PagesCount,"web_search.asp","[���������ҳ]",param_str
	end if
	%>
	<form method="post" action="web_searchmdel.asp" name="form7">

	<%RPage="web_search.asp"%><!-- #include file="include/web_admin_func_search.inc" -->
	
	<input type="hidden" name="s_adminname" value="<%=Request("s_adminname")%>" />
	<input type="hidden" name="s_name" value="<%=Request("s_name")%>" />
	<input type="hidden" name="s_title" value="<%=Request("s_title")%>" />
	<input type="hidden" name="s_article" value="<%=Request("s_article")%>" />
	<input type="hidden" name="s_email" value="<%=Request("s_email")%>" />
	<input type="hidden" name="s_qqid" value="<%=Request("s_qqid")%>" />
	<input type="hidden" name="s_msnid" value="<%=Request("s_msnid")%>" />
	<input type="hidden" name="s_homepage" value="<%=Request("s_homepage")%>" />
	<input type="hidden" name="s_ipv4addr" value="<%=Request("s_ipv4addr")%>" />
	<input type="hidden" name="s_originalipv4" value="<%=Request("s_originalipv4")%>" />
	<input type="hidden" name="s_ipv6addr" value="<%=Request("s_ipv6addr")%>" />
	<input type="hidden" name="s_originalipv6" value="<%=Request("s_originalipv6")%>" />
	<input type="hidden" name="s_reply" value="<%=Request("s_reply")%>" />
	<input type="hidden" name="page" value="<%=Request("page")%>" />

	<%
	if ItemsCount=0 then
		Response.Write "<br/><br/><div class=""centertext"">û���ҵ��������������ԡ�</div><br/><br/>"
	else
		dim pagename
		pagename="admin_web_search"
		if WebDisplayMode()="book" then
			%><!-- #include file="include/web_admin_listword.inc" --><%
		elseif WebDisplayMode()="forum" then
			%><!-- #include file="include/web_admin_listtitle.inc" --><%
		end if
		rs.Close
	end if
	%>

<!-- #include file="include/web_admin_func_search.inc" -->
</form>

<%
if PagesCount>1 and ShowBottomPageList then
	show_page_list ipage,PagesCount,"web_search.asp","[���������ҳ]",param_str
end if
%>

</div>

<% cn.Close : set rs=nothing : set cn=nothing %>

<!-- #include file="include/footer.inc" -->
</body>
</html>