<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster�������� ��������</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<link rel="stylesheet" type="text/css" href="css/web_adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
	<!-- #include file="css/web_adminstyle.asp" -->
</head>

<body>

<%
function filterstr(byref str,byval delquote)
	if delquote=true then
		filterstr=replace(replace(replace(str,"'",""),"[","[\[]"),chr(34),"")
	else
		filterstr=replace(replace(replace(str,"'","''"),"[","[\[]"),chr(34),"")
	end if
end function


'<ShowArticle>
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
'====================================

	'=============��ҳ����=============
	Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
	ItemsCount=0:PagesCount=0
	
	dim count_sql,sql
	count_sql=Replace(Replace(sql_websearchdec_words_count,"{0}",filterstr(Request("adminname"),true)),"{1}",filterstr(Request("searchtxt"),false))
	sql=Replace(Replace(sql_websearchdec_words_query,"{0}",filterstr(Request("adminname"),true)),"{1}",filterstr(Request("searchtxt"),false))
	
	get_divided_page cn,rs,sql_pk_supervisor,count_sql,sql,"adminname INC",Request.QueryString("page"),ItemsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
	'==============================
%>

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admincontrols.inc" -->

	<div class="region form-region">
		<h3 class="title">�����ö�����</h3>
		<div class="content">
			<form method="post" action="web_searchdec.asp" name="form1">
			<p>������("%"����������ַ���"_"����һ���ַ�)</p>
			<div class="field">
				<span class="label">�û�����</span>
				<span class="value"><input type="text" name="adminname" value="<%=Request("adminname")%>" size="<%=SearchTextWidth%>" /></span>
			</div>
			<div class="field">
				<span class="label">�������ݣ�</span>
				<span class="value"><input type="text" name="searchtxt" value="<%=Request("searchtxt")%>" size="<%=SearchTextWidth%>" /></span>
			</div>
			<div class="command"><input type="submit" value="��������" name="searchsubmit" /></div>
			</form>
		</div>
	</div>

<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"web_searchdec.asp","[���������ҳ]","center","adminname=" &server.URLEncode(Request("adminname"))& "&searchtxt=" & server.URLEncode(Request("searchtxt")) end if%>
<form method="post" action="web_mdeldec.asp" name="form7">

<input type="hidden" name="adminname" value="<%=request("adminname")%>" />
<input type="hidden" name="searchtxt" value="<%=request("searchtxt")%>" />
<input type="hidden" name="page" value="<%if isnumeric(request("page")) and request("page")<>"" then response.write request("page")%>" />
<div class="guest-functions">
	<div class="main">
		<a href="javascript:for(var i=0;i<=form7.elements.length-1;i++)if(form7.elements[i].name=='users' && form7.elements[i].checked){<%if DelSelDecTip=true then Response.Write "if (confirm('ȷʵҪɾ��ѡ��������')==true)"%>form7.submit();break;}else if(i==form7.elements.length-1)alert('����ѡ��Ҫɾ���Ĺ��档');" onmouseover="return true;"><img src="image/icon_mdel.gif" style="border-width:0px;" />ɾ��ѡ������</a>
	</div>
</div>


<%
dim pub_flag,pub
if ItemsCount=0 then
	Response.Write "<br/><br/><div class=""centertext"">û���ҵ����������Ĺ��档</div><br/><br/>"
else
	while rs.eof=false
%>

	<div class="region region-bulletin">
		<h3 class="title">�����û���<a href="web_userinfo.asp?user=<%=rs("adminname")%>" target="_blank"><%=rs("adminname")%></a></h3>
		<div class="content">
			<%
			pub_flag=rs("declareflag")
			pub="" &rs("declare")& ""

			if AdminViewCode=true then		'Only view HTML code
				pub=server.HTMLEncode(pub)
			else
				convertstr pub,pub_flag,2
			end if

			Response.Write pub
			%>
		</div>
		<div class="admin-tools">
			<input type="checkbox" name="users" id="c<%=i%>" value="<%=rs("adminname")%>"><label for="c<%=i%>">(ѡ��)</label>
			<a href="web_deldec.asp?user=<%=rs("adminname")%><%if isnumeric(request("page")) and request("page")<>"" then response.write "&page=" & request("page")%>&adminname=<%=server.URLEncode(request("adminname"))%>&searchtxt=<%=server.URLEncode(request("searchtxt"))%>" title="ɾ������"<%if DelDecTip=true then Response.Write " onclick=""return confirm('ȷʵҪɾ��������');"""%>><img border="0" src="image/icon_del.gif" class="imgicon" />[ɾ������]</a>
		</div>
	</div>
<%
	rs.MoveNext
	wend
end if	'��Ӧfor����һ�е�if
%>
<div class="guest-functions"><a href="javascript:for(var i=0;i<=form7.elements.length-1;i++)if(form7.elements[i].name=='users' && form7.elements[i].checked){<%if DelSelDecTip=true then Response.Write "if (confirm('ȷʵҪɾ��ѡ��������')==true)"%>form7.submit();break;}else if(i==form7.elements.length-1)alert('����ѡ��Ҫɾ���Ĺ��档');" onmouseover="return true;"><img src="image/icon_mdel.gif" style="border-width:0px;">ɾ��ѡ������</a></div>
</form>

<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"web_searchdec.asp","[���������ҳ]","center","adminname=" &server.URLEncode(Request("adminname"))& "&searchtxt=" & server.URLEncode(Request("searchtxt")) end if%>
</div>

<%
	if ItemsCount<>0 then rs.Close
	cn.Close
	set rs=nothing
	set cn=nothing
%>


<!-- #include file="bottom.asp" -->
</body>
</html>