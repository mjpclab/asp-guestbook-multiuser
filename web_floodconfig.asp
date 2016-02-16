<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_floodconfig.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster�������� ���Ա�����</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<%Call WebInitHeaderData("","Webmaster��������","","")%><!-- #include file="include/template/header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->

	<%rs.Open Replace(sql_adminfloodconfig,"{0}",wm_id),cn,,,1%>

	<div class="region form-region">
		<h3 class="title">����ˮ����</h3>
		<div class="content">
			<form method="post" action="web_savefloodconfig.asp" name="configform" onsubmit="return check();">
			<input type="hidden" name="user" value="<%=ruser%>" />
			<p>ͬһ�û���С����ʱ������<input type="text" name="minwait" size="10" maxlength="10" value="<%=flood_minwait%>" />�� (0=����)</p>

			<p>����<input type="text" name="searchrange" size="10" maxlength="10" value="<%=flood_searchrange%>" />��(0=����)
			<input type="checkbox" name="flag_newword" id="flag_newword" value="1"<%=cked(flood_sfnewword)%> /><label for="flag_newword">������</label>
			<input type="checkbox" name="flag_newreply" id="flag_newreply" value="1"<%=cked(flood_sfnewreply)%> /><label for="flag_newreply">�ÿͻظ�</label>
			<br/>������
			<input type="radio" name="flag_include_equal" id="flag_include" value="1"<%=cked(flood_include)%> /><label for="flag_include">����</label>
			<input type="radio" name="flag_include_equal" id="flag_equal" value="2"<%=cked(flood_equal)%> /><label for="flag_equal">����</label>
			<br/>��ͬ��
			<input type="checkbox" name="flag_title" id="flag_title" value="1"<%=cked(flood_sititle)%> /><label for="flag_title">����</label>
			<input type="checkbox" name="flag_content" id="flag_content" value="1"<%=cked(flood_sicontent)%> /><label for="flag_content">����</label>
			</p>

			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>

<script type="text/javascript" defer="defer">
function check()
{
	document.configform.submit1.disabled=true;
	return true;
}
</script>
</body>
</html>
<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>