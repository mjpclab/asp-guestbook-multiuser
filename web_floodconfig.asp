<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
%>

<!-- #include file="inc_dtd.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn">
<head>
	<title><%=web_BookName%> Webmaster管理中心 留言本参数</title>
	
	<!-- #include file="style.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admintool.inc" -->

	<%rs.Open Replace(sql_adminfloodconfig,"{0}",wm_name),cn,,,1%>

	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">防灌水策略</td>
		</tr>
		<tr>
			<td class="wordscontent" style="padding:20px 2px;">
			<form method="post" action="web_savefloodconfig.asp" name="configform" onsubmit="return check();">
			<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
			同一用户最小发言时间间隔：<input type="text" name="minwait" size="10" maxlength="10" value="<%=flood_minwait%>" />秒 (0=不限)<br/><br/>
			
			最新<input type="text" name="searchrange" size="10" maxlength="10" value="<%=flood_searchrange%>" />条(0=不限)
			<input type="checkbox" name="flag_newword" id="flag_newword" value="1"<%=cked(flood_sfnewword)%> /><label for="flag_newword">新留言</label>
			<input type="checkbox" name="flag_newreply" id="flag_newreply" value="1"<%=cked(flood_sfnewreply)%> /><label for="flag_newreply">访客回复</label>
			<br/>不允许
			<input type="radio" name="flag_include_equal" id="flag_include" value="1"<%=cked(flood_include)%> /><label for="flag_include">含有</label>
			<input type="radio" name="flag_include_equal" id="flag_equal" value="2"<%=cked(flood_equal)%> /><label for="flag_equal">具有</label>
			<br/>相同的
			<input type="checkbox" name="flag_title" id="flag_title" value="1"<%=cked(flood_sititle)%> /><label for="flag_title">标题</label>
			<input type="checkbox" name="flag_content" id="flag_content" value="1"<%=cked(flood_sicontent)%> /><label for="flag_content">内容</label>
			
			<p style="text-align:center;"><input value="更新数据" type="submit" name="submit1" /></p>
			</form>
		</td>
	</tr>
</table>
<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

</div>

<script type="text/javascript" defer="defer">
//<![CDATA[

function check()
{
	document.configform.submit1.disabled=true;
	return true;
}

//]]>
</script>

<!-- #include file="bottom.asp" -->
</body>
</html>