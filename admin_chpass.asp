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
	<title><%=HomeName%> 留言本 修改管理员密码</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->

	<script type="text/javascript">
	//<![CDATA[
	
	function checkpass(cobject)
	{
		if (cobject.ioldpass.value=="") {alert('请输入原密码。'); cobject.ioldpass.focus(); return false;}
		if (cobject.inewpass1.value=="") {alert('请输入新密码。'); cobject.inewpass1.focus(); return false;}
		if (cobject.inewpass2.value=="") {alert('请输入确认密码。'); cobject.inewpass2.focus(); return false;}
		if (cobject.inewpass1.value!=cobject.inewpass2.value) {alert('新密码与确认密码不同，请重新输入。'); cobject.inewpass1.focus(); return(false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	function checkqa(cobject)
	{
		if (cobject.ioldpass.value=="") {alert('请输入密码。'); cobject.ioldpass.focus(); return false;}
		if (cobject.question.value=="") {alert('请输入找回密码问题。'); cobject.question.focus(); return false;}
		if (cobject.key.value=="") {alert('请输入找回密码答案。'); cobject.key.focus(); return false;}
		cobject.submit1.disabled=true;
		return (true);
	}
	
	//]]>
	</script>
</head>

<body<%=bodylimit%> onload="form4.ioldpass.focus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admintool.inc" -->

	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">修改密码</td>
		</tr>
		<tr>
			<td class="wordscontent" style="text-align:center; padding:20px 0px;">
			<form method="post" action="admin_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
			　原密码：<input type="password" name="ioldpass" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
			　新密码：<input type="password" name="inewpass1" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
			确认密码：<input type="password" name="inewpass2" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/><br/>
			<input value="更新数据" type="submit" name="submit1" />
			</form>
			</td>
		</tr>
	</table>

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype
	checkuser cn,rs,false

	rs.Open Replace(sql_adminchpass_question,"{0}",Request.QueryString("user")),cn,,,1
	%>
	
	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">修改找回密码问题/答案</td>
		</tr>
		<tr>
			<td class="wordscontent" style="text-align:center; padding:20px 0px;">
			<form method="post" action="admin_saveqa.asp" onsubmit="return checkqa(this)" name="form5">
			<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
			密码：<input type="password" name="ioldpass" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
			问题：<input type="text" name="question" size="<%=SetInfoTextWidth%>" maxlength="32" value="<%=server.HTMLEncode("" & rs("question") & "")%>" /><br/>
			答案：<input type="text" name="key" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/><br/>
			<input value="更新数据" type="submit" name="submit1" />
			</form>
			</td>
		</tr>
	</table>
	<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>