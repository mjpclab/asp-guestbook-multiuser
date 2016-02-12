<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_setinfo.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

Call CreateConn(cn)
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> ���Ա� �޸İ�������</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<%
	rs.Open Replace(sql_adminsetinfo,"{0}",adminid),cn,,,1
	tfaceid=rs("faceid")
	%>

	<div class="region form-region region-longtext">
		<h3 class="title">�޸İ�������</h3>
		<div class="content">
			<form method="post" action="admin_saveinfo.asp" name="form1" onsubmit="form1.submit1.disabled=true;">
			<input type="hidden" name="user" value="<%=ruser%>" />
			<div class="field">
				<span class="label">�ǳƣ�</span>
				<span class="value"><input type="text" name="aname" maxlength="20" value="<%="" & rs("name") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">�ʼ���</span>
				<span class="value"><input type="text" name="aemail" maxlength="50" value="<%="" & rs("email") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">QQ�ţ�</span>
				<span class="value"><input type="text" name="aqqid" maxlength="16" value="<%="" & rs("qqid") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">Skype��</span>
				<span class="value"><input type="text" name="amsnid" maxlength="50" value="<%="" & rs("msnid") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">��ҳ��</span>
				<span class="value"><input type="text" name="ahomepage" maxlength="127" value="<%="" & rs("homepage") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">ͷ���ţ�</span>
				<span class="value"><input type="text" name="afaceid" maxlength="3" value="<%=tfaceid%>" title="��дͷ����ʱURL�������" /></span>
			</div>
			<div class="field">
				<span class="label">��URL��</span>
				<span class="value"><input type="text" name="afaceurl" maxlength="127" value="<%="" & rs("faceurl") & ""%>" title="��дURLʱ����ͷ����" /></span>
			</div>
			<div class="field">
				<%rs.Close : cn.Close : set rs=nothing : set cn=nothing

				dim listfacecount,defaultindex
				listfacecount=FrequentFaceCount
				defaultindex=tfaceid%>
				<!-- #include file="include/template/listface.inc" -->
			</div>
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>