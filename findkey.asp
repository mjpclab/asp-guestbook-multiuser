<!-- #include file="webconfig.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusFindkey=false then
	Response.Redirect "web_err.asp?number=2"
	Response.End
end if
%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
	<title><%=web_BookName%> �һ����벽��1</title>
	<!-- #include file="inc_stylesheet.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.user.value.length===0) {alert('�������û�����'); frm.user.focus(); return false;}
		else if (/^\w+$/.test(frm.user.value)===false) {alert('�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�'); frm.user.select(); return false;}
		else {frm.submit1.disabled=true; return true;}
	}
	</script>
</head>

<body onload="findform.user.select();">

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">��ҳ</a> &gt;&gt; <a href="findkey.asp" style="color:<%=TitleColor%>">�һ�����</a> &gt;&gt; ����1</div>
</div>

<%
set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="include/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/web_guest_func.inc" -->

<div class="region form-region">
	<h3 class="title">�һ����� ����1</h3>
	<div class="content">
	<form name="findform" method="post" action="findkey2.asp" onsubmit="return submitcheck(this)">
		<div class="field">
			<span class="label">�û�����</span>
			<span class="value"><input type="text" name="user" maxlength="32" size="32" title="ֻ�ܰ���Ӣ�ġ����ֺ��»��ߣ����32λ" /></span>
		</div>
		<div class="command"><input type="submit" name="submit1" value="��һ��" />����<input type="reset" value="������д" /></div>
	</form>
	</div>
</div>

</div>

<!-- #include file="include/footer.inc" -->
</body>
</html>