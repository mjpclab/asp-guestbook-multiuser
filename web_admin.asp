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
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admintool.inc" -->
		
	<span style="color:<%=TableContentColor%>;">
		
	<span class="generalwindow noborder" style="text-align:left; font-weight:bold; display:block;">ע���û��б�</span>
		<%
		set cn=server.CreateObject("ADODB.Connection")
		set rs=server.CreateObject("ADODB.Recordset")
		CreateConn cn,dbtype

		Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
		get_divided_page cn,rs,sql_pk_supervisor,sql_webadmin_words_count,sql_webadmin_words_query,"regdate DEC",Request.QueryString("page"),TitlesPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
	
		if ItemsCount>0 then
		%>
			<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"web_admin.asp","[�û��б��ҳ����" &PagesCount& "ҳ��" &ItemsCount& "���û�]","left",""%>
			<form method="post" action="web_deluser.asp" name="frm_user" onsubmit="for(var i=0;i<=elements.length-1;i++)if(elements[i].name=='users' && elements[i].checked){if(confirm('���棡ɾ�����û������ݽ����ָܻ���\nȷʵҪִ��ɾ��������'))return confirm('���ٴ�ȷ���Ƿ�Ҫɾ���û���');else return false;}alert('����ѡ��Ҫɾ�����û���');return false;">

				<input type="hidden" name="source" value="1">
				<%if isnumeric(Request.QueryString("page")) and Request.QueryString("page")<>"" then%><input type="hidden" name="arguments" value="page=<%=Request.QueryString("page")%>"><%end if%>
				
				<table id="titlelist" border="1" cellpadding="2" class="generalwindow grid">
				<tr class="header"><th style="width:30px">ѡ��</th><th>�û���</th><th>�ǳ�</th><th>ע������</th><th>����¼����</th><th>�����Ա�</th></tr>
				<%while not rs.EOF%>
					<tr><td style="text-align:center;"><input type="checkbox" name="users" value="<%=rs("adminname")%>" /></td><td><a href="web_userinfo.asp?user=<%=rs("adminname")%>" target="_blank"><%=rs("adminname")%></a></td><td><%=rs("name")%></td><td><%=rs("regdate")%></td><td><%=rs("lastlogin")%></td><td><a href="index.asp?user=<%=rs("adminname")%>" target="_blank">�����Ա�</a></td></tr>
				<%rs.MoveNext : wend%>
				</table>
				<!-- #include file="rowlight.inc" -->
			
				<%if ItemsCount>0 then%><span class="generalwindow noborder" style="text-align:left; display:block;"><input type="checkbox" onclick="for(var i=0;i<=elements.length-1;i++)if (elements[i].name=='users') elements[i].checked=this.checked;" id="selall" /><label for="selall">ȫѡ��ҳ�û�</label>    <input type="submit" value="ɾ��ѡ���û�" /></span><%end if%>
			</form>
		<%
		rs.Close
		end  if
		cn.Close : set rs=nothing : set cn=nothing
		%>
	</span>
		
	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"web_admin.asp","[�û��б��ҳ����" &PagesCount& "ҳ��" &ItemsCount& "���û�]","left",""%>

	<!-- #include file="searchuserbox_web.inc" -->		
</div>		

<!-- #include file="bottom.asp" -->
</body>
</html>