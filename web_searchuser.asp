<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="inc_dtd.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn">
<head>
	<title><%=web_BookName%> Webmaster�������� �����û�</title>
	
	<!-- #include file="style.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admintool.inc" -->
		
	<span style="color:<%=TableContentColor%>;">
		
		<span class="generalwindow noborder" style="text-align:left; font-weight:bold; display:block;">(�������)ע���û��б�</span>
			<%
			dim arg
			arg=""
			if isnumeric(Request("page")) and request("page")<>"" then arg="page=" & request("page")
			if request("type")<>"" then
				if arg<>"" then arg=arg & "&"
				arg=arg & "type=" & server.URLEncode(request("type"))
			end if
			if request("searchtxt")<>"" then
				if arg<>"" then arg=arg & "&"
				arg=arg & "searchtxt=" & server.URLEncode(request("searchtxt"))
			end if

			set cn=server.CreateObject("ADODB.Connection")
			set rs=server.CreateObject("ADODB.Recordset")
			CreateConn cn,dbtype
		
			if replace(replace(replace(Request("type"),"'",""),";","")," ","")<>"" and request("searchtxt")<>"" then
				search_condition=Replace(Replace(sql_websearchuser_condition,"{0}",replace(replace(replace(replace(replace(replace(Request("type"),"'",""),";","")," ",""),"[",""),"--",""),"/*","")),"{1}",replace(replace(Request("searchtxt"),"'","''"),"[","[\[]"))
			else
				search_condition=""
			end if

			Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
			get_divided_page cn,rs,sql_pk_supervisor,sql_websearchuser_words_count & search_condition,sql_websearchuser_words_query & search_condition,"regdate DEC",Request.QueryString("page"),TitlesPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage

			if ItemsCount>0 then
			%>
				<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"web_searchuser.asp","(�������)[�û��б��ҳ����" &PagesCount& "ҳ��" &ItemsCount& "���û�]","left","type=" &Request("type")& "&searchtxt=" &server.URLEncode(request("searchtxt"))%>
				<form method="post" action="web_deluser.asp" name="frm_user" onsubmit="for(var i=0;i<=elements.length-1;i++)if(elements[i].name=='users' && elements[i].checked){if(confirm('���棡ɾ�����û������ݽ����ָܻ���\nȷʵҪִ��ɾ��������'))return confirm('���ٴ�ȷ���Ƿ�Ҫɾ���û���');else return false;}alert('����ѡ��Ҫɾ�����û���');return false;">

					<input type="hidden" name="source" value="2" />
					<input type="hidden" name="arguments" value="<%=arg%>">
					<table id="titlelist" border="1" cellpadding="2" class="generalwindow grid" ID="Table1">
					<tr class="header"><th style="width:30px">ѡ��</th><th>�û���</th><th>�ǳ�</th><th>ע������</th><th>����¼����</th><th>�����Ա�</th></tr>
					<%while not rs.EOF%>
						<tr><td style="text-align:center;"><input type="checkbox" name="users" value="<%=rs("adminname")%>" ID="Checkbox1"/></td><td><a href="web_userinfo.asp?user=<%=rs("adminname")%>" target="_blank"><%=rs("adminname")%></a></td><td><%=rs("name")%></td><td><%=rs("regdate")%></td><td><%=rs("lastlogin")%></td><td><a href="index.asp?user=<%=rs("adminname")%>" target="_blank">�����Ա�</a></td></tr>
					<%rs.MoveNext : wend%>
					</table>
					<!-- #include file="rowlight.inc" -->

					<%if ItemsCount>0 then%><span style="width:100%; text-align:left;"><br/> <input type="checkbox" onclick="for(var i=0;i<=elements.length-1;i++)if (elements[i].name=='users') elements[i].checked=this.checked;" id="selall" /><label for="selall">ȫѡ��ҳ�û�</label>    <input type="submit" value="ɾ��ѡ���û�" /></span><%end if%>
				</form>
			<%
			rs.Close
			end  if
			cn.Close : set rs=nothing : set cn=nothing
			%>
	</span>
		
	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"web_searchuser.asp","(�������)[�û��б��ҳ����" &PagesCount& "ҳ��" &ItemsCount& "���û�]","left","type=" &Request("type")& "&searchtxt=" &server.URLEncode(request("searchtxt"))%>

	<!-- #include file="searchuserbox_web.inc" -->	
</div>		

<!-- #include file="bottom.asp" -->
</body>
</html>