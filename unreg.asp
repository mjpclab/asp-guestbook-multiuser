<!-- #include file="webconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusUnreg=false then
	Response.Redirect "web_err.asp?number=5"
	Response.End
end if

if VcodeCount>0 then session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> ��ɾ�ʺ�</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<!-- #include file="css/style.asp" -->

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
<!-- #include file="sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="func_web.inc" -->

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
			<span class="value"><input type="text" name="vcode" size="10" /><img class="captcha" src="web_show_vcode.asp"/></span>
		</div>
		<%end if%>
		<div class="command"><input type="submit" value="�ύ" />����<input type="reset" value="��д" /></div>

		</form>
	</div>
</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>