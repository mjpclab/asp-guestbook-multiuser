<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� �޸Ĺ���Ա����</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->

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
	<!-- #include file="admincontrols.inc" -->

	<div class="region form-region">
		<h3 class="title">�޸�����</h3>
		<div class="content">
			<form method="post" action="admin_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			<input type="hidden" name="user" value="<%=ruser%>" />
			<div class="field">
				<span class="label">ԭ���룺</span>
				<span class="value"><input type="password" name="ioldpass" size="<%=SetInfoTextWidth%>" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">�����룺</span>
				<span class="value"><input type="password" name="inewpass1" size="<%=SetInfoTextWidth%>" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">ȷ�����룺</span>
				<span class="value"><input type="password" name="inewpass2" size="<%=SetInfoTextWidth%>" maxlength="32" /></span>
			</div>
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype
	checkuser cn,rs,false

	rs.Open Replace(sql_adminchpass_question,"{0}",adminid),cn,,,1
	%>

	<div class="region form-region">
		<h3 class="title">�޸��һ���������/��</h3>
		<div class="content">
			<form method="post" action="admin_saveqa.asp" onsubmit="return checkqa(this)" name="form5">
				<input type="hidden" name="user" value="<%=ruser%>" />
				<div class="field">
					<span class="label">���룺</span>
					<span class="value"><input type="password" name="ioldpass" size="<%=SetInfoTextWidth%>" maxlength="32" /></span>
				</div>
				<div class="field">
					<span class="label">���⣺</span>
					<span class="value"><input type="text" name="question" size="<%=SetInfoTextWidth%>" maxlength="32" value="<%=server.HTMLEncode("" & rs("question") & "")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�𰸣�</span>
					<span class="value"><input type="text" name="key" size="<%=SetInfoTextWidth%>" maxlength="32" /></span>
				</div>
				<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

	<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>