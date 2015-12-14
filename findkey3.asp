<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/sql/web_findkey3.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="tips.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
elseif StatusFindkey=false then
	Call WebErrorPage(2)
	Response.End
end if

if VcodeCount>0 and (Request.Form("vcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	Call TipsPage("��֤�����","findkey.asp")
	Response.End
elseif VcodeCount>0 then
	session("vcode")=getvcode(VcodeCount)
else
	session("vcode")=""
end if

'===============================��ʽ��֤
dim re
set re=new RegExp
re.Pattern="^\w+$"

ruser=Request.Form("user")
if ruser="" then
	Call TipsPage("�û�������Ϊ�ա�","findkey.asp")
	Response.End
elseif re.Test(ruser)=false then
	Call TipsPage("�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�","findkey.asp")
	Response.End
elseif Request.Form("key")="" then
	Call TipsPage("�һ�����𰸲���Ϊ�ա�","findkey.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

'=============================��������֤
rs.Open Replace(sql_findkey3_checkuser,"{0}",ruser),cn,,,1
if rs.EOF then
	Call TipsPage("�û��������ڡ�","findkey.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if

if md5(Request.Form("key"),32)<>rs("key") then
	Call TipsPage("�һ�����𰸲���ȷ��","findkey.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : set rs=nothing
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> �һ����벽��3</title>
	<!-- #include file="inc_stylesheet.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.pass1.value.length===0) {alert('���������롣'); frm.pass1.focus(); return false;}
		else if (frm.pass2.value.length===0) {alert('������ȷ�����롣'); frm.pass2.focus(); return false;}
		else if (frm.pass1.value!==frm.pass2.value) {alert('���벻һ�£����顣'); frm.pass2.select(); return false;}
		else if (frm.vcode && frm.vcode.value.length===0) {alert('��������֤�롣'); frm.vcode.focus(); return false;}
		else {frm.submit1.disabled=true; return true;}
	}
	</script>
</head>

<body onload="findform3.pass1.focus();">

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">��ҳ</a> &gt;&gt; <a href="findkey.asp" style="color:<%=TitleColor%>">�һ�����</a> &gt;&gt; ����3</div>
</div>

<%
dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="include/template/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/template/web_guest_func.inc" -->

<div class="region form-region">
	<h3 class="title">�һ����� ����3</h3>
	<div class="content">
		<form name="findform3" method="post" action="findkey4.asp" onsubmit="return submitcheck(this)">
			<input type="hidden" name="user" value="<%=ruser%>" />
			<input type="hidden" name="key" value="<%=Request.Form("key")%>" />

			<h4>������������������</h4>
			<div class="field">
				<span class="label">���룺</span>
				<span class="value"><input type="password" name="pass1" maxlength="32" size="32" /></span>
			</div>
			<div class="field">
				<span class="label">ȷ�����룺</span>
				<span class="value"><input type="password" name="pass2" maxlength="32" size="32" /></span>
			</div>
			<%if VcodeCount>0 then%>
			<div class="field">
				<span class="label">��֤�룺</span>
				<span class="value"><input type="text" name="vcode" size="10" autocomplete="off" /><img id="captcha" class="captcha" src="show_vcode.asp?t=0"/></span>
			</div>
			<%end if%>
			<div class="command"><input type="submit" name="submit1" value="�������" />����<input type="reset" value="������д" /></div>
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