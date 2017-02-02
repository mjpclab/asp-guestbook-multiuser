<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/web.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/web_admin_userinfo.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%Response.Expires=-1%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster�������� �鿴�û���Ϣ</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
checkuser cn,rs,false

%>

<div id="outerborder" class="outerborder">

<%Call WebInitHeaderData("","Webmaster��������","","")%><!-- #include file="include/template/header.inc" -->

<div id="mainborder" class="mainborder">
<div class="region">
	<h3 class="title">�鿴�û���Ϣ</h3>
	<div class="content">
		<%
		rs.Open Replace(sql_webuserinfo,"{0}",ruser),cn,,,1
		if rs("faceid")>0 then Response.Write "<img src=""asset/face/" & rs("faceid") & ".gif" & """>"
		%>
		<p class="field">
			<a href="index.asp?user=<%=ruser%>" target="_blank">[�����Ա�]</a>
			<a href="web_chuserpass.asp?user=<%=ruser%>">[�����û�����]</a>
			<a href="web_chuserstatus.asp?user=<%=ruser%>">[�����˺�״̬]</a>
		</p>
		<table cellpadding="5">
			<tr><th scope="row">�û���</th><td><%=rs("adminname")%></td></tr>
			<tr><th scope="row">�ǳ�</th><td><%=rs("name") & ""%></td></tr>
			<tr><th scope="row">E-mail</th><td><%=rs("email") & ""%></td></tr>
			<tr><th scope="row">QQ��</th><td><%=rs("qqid") & ""%></td></tr>
			<tr><th scope="row">Skype</th><td><%=rs("msnid") & ""%></td></tr>
			<tr><th scope="row">��ҳ</th><td><%=rs("homepage") & ""%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_status,"{0}",adminid),cn,,,1
			Dim userstatus
			userstatus=rs.Fields(0)
			%>
			<tr><th scope="row">�˺�״̬</th><td><%if CBool(userstatus AND 1073741824) then Response.Write "����" else Response.Write "����"%></td></tr>
			<tr><th scope="row">�˺Ź����½״̬</th><td><%if CBool(userstatus AND 536870912) then Response.Write "����" else Response.Write "����"%></td></tr>
			<tr><th scope="row">�˺�ǩд����״̬</th><td><%if CBool(userstatus AND 268435456) then Response.Write "����" else Response.Write "����"%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_view,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">���Բ鿴����</th><td><%if Not rs.EOF then Response.Write rs(0) else Response.Write "0"%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_words,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">������</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_reply,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">�ظ���</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_ipv4config,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">�Զ���IPv4���β�������</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_ipv6config,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">�Զ���IPv6���β�������</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_count_filterconfig,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">�Զ������ݹ��˲�������</th><td><%=rs(0)%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_reginfo,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">ע������</th><td><%=UTCToDisplayTime(rs("regdate"))%></td></tr>
			<tr><th scope="row">����¼����</th><td><%=UTCToDisplayTime(rs("lastlogin"))%></td></tr>

			<%
			rs.Close
			rs.Open Replace(sql_webuserinfo_logininfo,"{0}",adminid),cn,,,1
			%>
			<tr><th scope="row">��¼����</th><td><%if not rs.EOF then Response.Write rs(0) else Response.Write "0"%></td></tr>
			<tr><th scope="row">���е�¼ʧ�ܴ���</th><td><%if not rs.EOF then Response.Write rs(1) else Response.Write "0"%></td></tr>
		</table>
	</div>
</div>
</div>

<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>
