<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=web_BookName%> Webmaster�������� �鿴�û���Ϣ</title>
	
	<!-- #include file="style.asp" -->
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

<br/>
<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" style="width:100%; border-collapse:collapse;">
	<tr>
		<td class="centertitle">�鿴�û���Ϣ</td>
	</tr>
	<tr>
	<td style="width:100%; color:<%=TableContentColor%>; background-color:<%=TableContentBGC%>;">
		<br/>
		<table  cellpadding="0" cellspacing="0" style="width:100%; border-width:0px;">
			<tr>
			<td style="width:<%=TableLeftWidth%>px; text-align:center; vertical-align:top;">
				<%
				rs.Open Replace(sql_webuserinfo,"{0}",Request.QueryString("user")),cn,,,1
				if rs("faceid")>0 then Response.Write "<img src=""" & FacePath & cstr(rs("faceid")) & ".gif" & """ style=""border:0px;"">"
				%>
				<br/><br/>
				<a href="index.asp?user=<%=rs("adminname")%>" target="_blank">[�����Ա�]</a>
				<br/><br/>
				<a href="web_chuserpass.asp?user=<%=rs("adminname")%>">[�����û�����]</a>
			</td>
			<td style="vertical-align:top;">
				<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" cellspacing="0" style="width:100%; border-collapse:collapse;">
					<tr><td style="width:180px; font-weight:bold;">�û�����</td><td><%=rs("adminname")%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">�ǳƣ�</td><td><%=rs("name") & ""%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">E-mail��</td><td><%=rs("email") & ""%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">QQ�ţ�</td><td><%=rs("qqid") & ""%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">MSN��</td><td><%=rs("msnid") & ""%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">��ҳ��</td><td><%=rs("homepage") & ""%></td></tr>
					
					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_count_view,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">���Բ鿴������</td><td><%if rs.EOF=false then Response.Write rs(0) else Response.Write "0"%></td></tr>
					
					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_count_words,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">��������</td><td><%=rs(0)%></td></tr>
					
					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_count_reply,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">�ظ�����</td><td><%=rs(0)%></td></tr>
					
					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_count_ipconfig,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">�Զ���IP���β���������</td><td><%=rs(0)%></td></tr>

					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_count_filterconfig,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">�Զ������ݹ��˲���������</td><td><%=rs(0)%></td></tr>

					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_reginfo,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">ע�����ڣ�</td><td><%=rs("regdate")%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">����¼���ڣ�</td><td><%=rs("lastlogin")%></td></tr>

					<%
					rs.Close
					rs.Open Replace(sql_webuserinfo_logininfo,"{0}",Request.QueryString("user")),cn,,,1
					%>
					<tr><td style="width:180px; font-weight:bold;">��¼������</td><td><%if not rs.EOF then Response.Write rs(0) else Response.Write "0"%></td></tr>
					<tr><td style="width:180px; font-weight:bold;">���е�¼ʧ�ܴ�����</td><td><%if not rs.EOF then Response.Write rs(1) else Response.Write "0"%></td></tr>
					
				</table>
			</td>
			</tr>
		</table>
		<br/>
	</td>
	</tr>
</table>

<%
rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>