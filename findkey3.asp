<!-- #include file="webconfig.asp" -->
<!-- #include file="md5.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusFindkey=false then
	Response.Redirect "web_err.asp?number=2"
	Response.End
end if

if VcodeCount>0 and (Request.Form("vcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	Call MessagePage("��֤�����","findkey.asp")
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

if Request.Form("user")="" then
	Call MessagePage("�û�������Ϊ�ա�","findkey.asp")
	Response.End
elseif re.Test(Request.Form("user"))=false then
	Call MessagePage("�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�","findkey.asp")
	Response.End
elseif Request.Form("key")="" then
	Call MessagePage("�һ�����𰸲���Ϊ�ա�","findkey.asp")
	Response.End
end if

dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

'=============================��������֤
rs.Open Replace(sql_findkey3_checkuser,"{0}",Request.Form("user")),cn,,,1
if rs.EOF then
	Call MessagePage("�û��������ڡ�","findkey.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if

if md5(Request.Form("key"),32)<>rs("key") then
	Call MessagePage("�һ�����𰸲���ȷ��","findkey.asp")
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : set rs=nothing
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> �һ����벽��3</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.pass1.value=='') {alert('���������롣'); frm.pass1.focus(); return false;}
		else if (frm.pass2.value=='') {alert('������ȷ�����롣'); frm.pass2.focus(); return false;}
		else if (frm.pass1.value!=frm.pass2.value) {alert('���벻һ�£����顣'); frm.pass2.select(); return false;}
		else if (frm.vcode && frm.vcode.value=='') {alert('��������֤�롣'); frm.vcode.focus(); return false;}
		else {frm.submit1.disabled=true; return true;}
	}
	</script>
</head>

<body onload="findform3.pass1.focus();">

<div id="outerborder" class="outerborder">

<table cellpadding="2" cellspacing="0" style="width:100%; border-width:0px; border-collapse:collapse;">
	<tr style="height:60px;">
		<td style="width:100%; color:<%=TitleColor%>; background-color:<%=TitleBGC%>; background-image:url(<%=TitlePic%>);" nowrap="nowrap">
			<%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">��ҳ</a> &gt;&gt; <a href="findkey.asp" style="color:<%=TitleColor%>">�һ�����</a> &gt;&gt; ����3
		</td>
	</tr>		
	<tr>
		<td>
			<%			
			dim sys_bul_flag
			sys_bul_flag=32
			%>
			<!-- #include file="sysbulletin.inc" -->
			<%cn.Close : set cn=nothing%>
		</td>
	</tr>
	<tr>
		<td style="width:100%">
		<!-- #include file="func_web.inc" -->
		</td>
	</tr>
		<td style="width:100%; text-align:center; padding:20px 0px; color:<%=TableContentColor%>;">
			<form name="findform3" method="post" action="findkey4.asp" onsubmit="return submitcheck(this)">
				<input type="hidden" name="user" value="<%=Request.Form("user")%>" />
				<input type="hidden" name="key" value="<%=Request.Form("key")%>" />

				<table cellpadding="5" cellspacing="0" style="margin-left:auto; margin-right:auto; border-width:0px;">
				<tr style="height:20px;"><td></td></tr>
				<tr>
					<td colspan="2" align="center">
						������������������
					</td>
				</tr>
				<tr style="height:5px;"><td></td></tr>
				<tr>
					<td>���룺</td>
					<td><input type="password" name="pass1" maxlength="32" size="32" /></td>
				</tr>
				<tr style="height:5px;"><td></td></tr>
				<tr>
					<td>ȷ�����룺</td>
					<td><input type="password" name="pass2" maxlength="32" size="32" /></td>
				</tr>
				<%if VcodeCount>0 then%>
				<tr style="height:5px;"><td></td></tr>
				<tr>
					<td>��֤�룺</td>
					<td><input type="text" name="vcode" size="10" /> <img src="web_show_vcode.asp" class="vcode" onclick="this.src=this.src" /></td>
				</tr>
				<%end if%>
				<tr style="height:20px;"><td></td></tr>
				<tr>
				<td colspan="2" style="text-align:center;" />
						<input type="submit" name="submit1" value="�������" />����<input type="reset" value="������д" />
				</td>
			</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>