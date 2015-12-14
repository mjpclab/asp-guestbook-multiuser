<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_chpass.asp" -->
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
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
end if
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> ���Ա� �޸Ĺ���Ա����</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function checkpass(cobject)
	{
		if (cobject.ioldpass.value.length===0) {alert('������ԭ���롣'); cobject.ioldpass.focus(); return false;}
		if (cobject.inewpass1.value.length===0) {alert('�����������롣'); cobject.inewpass1.focus(); return false;}
		if (cobject.inewpass2.value.length===0) {alert('������ȷ�����롣'); cobject.inewpass2.focus(); return false;}
		if (cobject.inewpass1.value!==cobject.inewpass2.value) {alert('��������ȷ�����벻ͬ�����������롣'); cobject.inewpass1.focus(); return(false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	function checkqa(cobject)
	{
		if (cobject.ioldpass.value.length===0) {alert('���������롣'); cobject.ioldpass.focus(); return false;}
		if (cobject.question.value.length===0) {alert('�������һ��������⡣'); cobject.question.focus(); return false;}
		if (cobject.key.value.length===0) {alert('�������һ�����𰸡�'); cobject.key.focus(); return false;}
		cobject.submit1.disabled=true;
		return (true);
	}
	</script>
</head>

<body<%=bodylimit%> onload="form4.ioldpass.focus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="include/template/admin_mainmenu.inc" -->

	<div class="region form-region region-longtext">
		<h3 class="title">�޸�����</h3>
		<div class="content">
			<form method="post" action="admin_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			<input type="hidden" name="user" value="<%=ruser%>" />
			<div class="field">
				<span class="label">ԭ���룺</span>
				<span class="value"><input type="password" name="ioldpass" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">�����룺</span>
				<span class="value"><input type="password" name="inewpass1" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">ȷ�����룺</span>
				<span class="value"><input type="password" name="inewpass2" maxlength="32" /></span>
			</div>
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)

	rs.Open Replace(sql_adminchpass_question,"{0}",adminid),cn,,,1
	%>

	<div class="region form-region region-longtext">
		<h3 class="title">�޸��һ���������/��</h3>
		<div class="content">
			<form method="post" action="admin_saveqa.asp" onsubmit="return checkqa(this)" name="form5">
				<input type="hidden" name="user" value="<%=ruser%>" />
				<div class="field">
					<span class="label">���룺</span>
					<span class="value"><input type="password" name="ioldpass" maxlength="32" /></span>
				</div>
				<div class="field">
					<span class="label">���⣺</span>
					<span class="value"><input type="text" name="question" maxlength="32" value="<%=server.HTMLEncode("" & rs("question") & "")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�𰸣�</span>
					<span class="value"><input type="text" name="key" maxlength="32" /></span>
				</div>
				<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

	<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>
</div>

<!-- #include file="include/template/footer.inc" -->
</body>
</html>