<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster�������� �鿴�û���Ϣ</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,false

%>

<div id="outerborder" class="outerborder">

<!-- #include file="web_admintitle.inc" -->


<div class="region">
	<h3 class="title">�鿴�û���Ϣ</h3>
	<div class="content">
		<%
		rs.Open Replace(sql_webuserinfo,"{0}",ruser),cn,,,1
		if rs("faceid")>0 then Response.Write "<img src=""" & FacePath & cstr(rs("faceid")) & ".gif" & """ style=""border:0px;"">"
		%>
		<p class="field">
			<a href="index.asp?user=<%=rs("adminname")%>" target="_blank">[�����Ա�]</a>
			<a href="web_chuserpass.asp?user=<%=rs("adminname")%>">[�����û�����]</a>
		</p>
		<table cellpadding="5">
			<tr><th scope="row">�û�����</th><td><%=rs("adminname")%></td></tr>
			<tr><th scope="row">�ǳƣ�</th><td><%=rs("name") & ""%></td></tr>
			<tr><th scope="row">E-mail��</th><td><%=rs("email") & ""%></td></tr>
			<tr><th scope="row">QQ�ţ�</th><td><%=rs("qqid") & ""%></td></tr>
			<tr><th scope="row">Skype��</th><td><%=rs("msnid") & ""%></td></tr>
			<tr><th scope="row">��ҳ��</th><td><%=rs("homepage") & ""%></td></tr>

			<%
			adminid=rs.Fields("adminid")
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_view,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">���Բ鿴������</th><td><%if rs.EOF=false then Response.Write rs(0) else Response.Write "0"%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_words,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">��������</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_reply,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">�ظ�����</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_ipv4config,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">�Զ���IPv4���β���������</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_ipv6config,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">�Զ���IPv6���β���������</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_filterconfig,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">�Զ������ݹ��˲���������</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_reginfo,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">ע�����ڣ�</th><td><%=rs("regdate")%></td></tr>
			<tr><th scope="row">����¼���ڣ�</th><td><%=rs("lastlogin")%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_logininfo,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">��¼������</th><td><%if not rs.EOF then Response.Write rs(0) else Response.Write "0"%></td></tr>
			<tr><th scope="row">���е�¼ʧ�ܴ�����</th><td><%if not rs.EOF then Response.Write rs(1) else Response.Write "0"%></td></tr>
		</table>
	</div>
</div>

<%
rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>