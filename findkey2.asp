<!-- #include file="webconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusFindkey=false then
	Response.Redirect "web_err.asp?number=2"
	Response.End
end if

if VcodeCount>0 then session("vcode")=getvcode(VcodeCount)
'===============================合式验证
dim re
set re=new RegExp
re.Pattern="^\w+$"

ruser=Request.Form("user")
if ruser="" then
	Call MessagePage("用户名不能为空。","findkey.asp")
	Response.End
elseif re.Test(ruser)=false then
	Call MessagePage("用户名中只能包含英文字母、数字和下划线。","findkey.asp")
	Response.End
end if

dim cn,rs,question
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

'=============================存在性验证
rs.Open Replace(sql_findkey2_checkuser,"{0}",ruser),cn,,,1
if rs.EOF then
	Call MessagePage("用户名不存在。","findkey.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if

question="" & rs("question") & ""
rs.Close : set rs=nothing
%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
	<title><%=web_BookName%> 找回密码步骤2</title>
	<!-- #include file="inc_stylesheet.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.key.value.length===0) {alert('请输入找回密码答案。'); frm.key.focus(); return false;}
		else if (frm.vcode && frm.vcode.value.length===0) {alert('请输入验证码。'); frm.vcode.focus(); return false;}
		else {frm.submit1.disabled=true; return true;}
	}
	</script>
</head>

<body onload="findform2.key.focus();">

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">首页</a> &gt;&gt; <a href="findkey.asp" style="color:<%=TitleColor%>">找回密码</a> &gt;&gt; 步骤2</div>
</div>

<%
dim sys_bul_flag
sys_bul_flag=32
%>
<!-- #include file="include/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/web_guest_func.inc" -->

<div class="region form-region">
	<h3 class="title">找回密码 步骤2</h3>
	<div class="content">
		<form name="findform2" method="post" action="findkey3.asp" onsubmit="return submitcheck(this)">
			<input type="hidden" name="user" value="<%=ruser%>">

			<div class="field">
				<span class="label">问题：</span>
				<span class="value"><%=server.HTMLEncode(question)%></span>
			</div>
			<div class="field">
				<span class="label">答案：</span>
				<span class="value"><input type="text" name="key" maxlength="32" size="32" /></span>
			</div>
			<%if VcodeCount>0 then%>
			<div class="field">
				<span class="label">验证码：</span>
				<span class="value"><input type="text" name="vcode" size="10" autocomplete="off" /><img id="captcha" class="captcha" src="web_show_vcode.asp?t=0" /></span>
			</div>
			<%end if%>
			<div class="command"><input type="submit" name="submit1" value="下一步" />　　<input type="reset" value="重新填写" /></div>
		</form>
	</div>
</div>

</div>
<!-- #include file="include/footer.inc" -->
<script type="text/javascript">
	<!-- #include file="js/refresh-captcha.js" -->
</script>
</body>
</html>