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
	<title><%=web_BookName%> Webmaster��������</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admincontrols.inc" -->

	<h3>ע���û��б�</h3>
	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype

	Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
	get_divided_page cn,rs,sql_pk_supervisor,sql_webadmin_words_count,sql_webadmin_words_query,"regdate DEC",Request.QueryString("page"),TitlesPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage

	if ItemsCount>0 then
	%>
		<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"web_admin.asp","[�û��б��ҳ����" &PagesCount& "ҳ��" &ItemsCount& "���û�]",""%>
		<form method="post" action="web_deluser.asp" name="frm_user" onsubmit="for(var i=0;i<=elements.length-1;i++)if(elements[i].name=='users' && elements[i].checked){if(confirm('���棡ɾ�����û������ݽ����ָܻ���\nȷʵҪִ��ɾ��������'))return confirm('���ٴ�ȷ���Ƿ�Ҫɾ���û���');else return false;}alert('����ѡ��Ҫɾ�����û���');return false;">

			<input type="hidden" name="source" value="1">
			<%if isnumeric(Request.QueryString("page")) and Request.QueryString("page")<>"" then%><input type="hidden" name="arguments" value="page=<%=Request.QueryString("page")%>"><%end if%>

			<table id="titlelist" class="topic-list">
			<thead>
				<tr><th class="select"><input type="checkbox" name="users"/></th><th>�û���</th><th>�ǳ�</th><th>ע������</th><th>����¼����</th><th>�����Ա�</th></tr>
			</thead>
			<tbody>
			<%while not rs.EOF%>
				<tr><td class="select"><input type="checkbox" name="users" value="<%=rs("adminname")%>" /></td><td><a href="web_userinfo.asp?user=<%=rs("adminname")%>" target="_blank"><%=rs("adminname")%></a></td><td><%=rs("name")%></td><td><%=rs("regdate")%></td><td><%=rs("lastlogin")%></td><td><a href="index.asp?user=<%=rs("adminname")%>" target="_blank">�����Ա�</a></td></tr>
			<%rs.MoveNext : wend%>
			</tbody>
			</table>
			<script type="text/javascript" src="js/jquery-1.x-min.js"></script>
			<script type="text/javascript" src="js/table-select.js"></script>

			<%if ItemsCount>0 then%>
			<div class="guest-functions">
			<div class="main">
				<input type="submit" value="ɾ��ѡ���û�" />
			</div>
			</div>
			<%end if%>
		</form>
	<%
	rs.Close
	end  if
	cn.Close : set rs=nothing : set cn=nothing
	%>

	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"web_admin.asp","[�û��б��ҳ����" &PagesCount& "ҳ��" &ItemsCount& "���û�]",""%>

	<!-- #include file="searchuserbox_web.inc" -->		
</div>		

<!-- #include file="bottom.asp" -->
</body>
</html>