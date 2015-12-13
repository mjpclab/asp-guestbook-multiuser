<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif StatusUnreg=false then
	Call WebErrorPage(5)
	Response.End
end if

if VcodeCount>0 then session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> ��ɾ�ʺ�</title>
	<!-- #include file="inc_stylesheet.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.user.value.length===0) {alert('�������û�����'); frm.user.focus(); return false;}
		else if (/^\w+$/.test(frm.user.value)===false) {alert('�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�'); frm.user.focus(); return false;}
		else if (frm.pass1.value.length===0) {alert('���������롣'); frm.pass1.focus(); return false;}
		else if (frm.vcode && frm.vcode.value.length===0) {alert('��������֤�롣'); frm.vcode.focus(); return false;}
		else return confirm('ɾ���û�����ͬʱ���ȫ�����ԡ��ظ������棬ȷʵҪ������');
	}
	</script>
	
</head>

<body onload="unregform.user.focus();">

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">��ҳ</a> &gt;&gt; ��ɾ�ʺ�</div>
</div>
<%
set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="include/template/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/template/web_guest_func.inc" -->

<div class="region form-region">
	<h3 class="title">��ɾ�ʺ�</h3>
	<div class="content">
		<form name="unregform" method="post" action="submitunreg.asp" onsubmit="return submitcheck(this)">

		<div class="field">
			<span class="label">�û�����</span>
			<span class="value"><input type="text" name="user" maxlength="32" size="32" title="ֻ�ܰ���Ӣ�ġ����ֺ��»��ߣ����32λ" /></span>
		</div>
		<div class="field">
			<span class="label">���룺</span>
			<span class="value"><input type="password" name="pass1" maxlength="32" size="32" /></span>
		</div>
		<%if VcodeCount>0 then%>
		<div class="field">
			<span class="label">��֤�룺</span>
			<span class="value"><input type="text" name="vcode" size="10" /><img id="captcha" class="captcha" src="show_vcode.asp?t=0"/></span>
		</div>
		<%end if%>
		<div class="command"><input type="submit" value="�ύ" />����<input type="reset" value="��д" /></div>

		</form>
	</div>
</div>

</div>
<!-- #include file="include/template/footer.inc" -->
<script type="text/javascript">
	<!-- #include file="asset/js/refresh-captcha.js" -->
</script>
</body>
</html>