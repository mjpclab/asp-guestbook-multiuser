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
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=web_BookName%> ��ɾ�ʺ�</title>

	<script type="text/javascript">
	//<![CDATA[
	
	function submitcheck(frm)
	{
		if (frm.user.value=='') {alert('�������û�����'); frm.user.focus(); return false;}
		else if (/^\w+$/.test(frm.user.value)==false) {alert('�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�'); frm.user.focus(); return false;}
		else if (frm.pass1.value=='') {alert('���������롣'); frm.pass1.focus(); return false;}
		else if (frm.vcode && frm.vcode.value=='') {alert('��������֤�롣'); frm.vcode.focus(); return false;}
		else return confirm('ɾ���û�����ͬʱ���ȫ�����ԡ��ظ������棬ȷʵҪ������');
	}
	
	//]]>
	</script>
	
	<!-- #include file="style.asp" -->
</head>

<body onload="unregform.user.focus();">

<div id="outerborder" class="outerborder">

<table cellpadding="2" cellspacing="0" style="width:100%; border-width:0px; border-collapse:collapse;">
	<tr style="height:60px;">
		<td style="width:100%; color:<%=TitleColor%>; background-color:<%=TitleBGC%>; background-image:url(<%=TitlePic%>);" nowrap="nowrap">
			<%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">��ҳ</a> &gt;&gt; ��ɾ�ʺ�
		</td>
	</tr>		
	<tr>
		<td>
			<%			
			set cn=server.CreateObject("ADODB.Connection")
			CreateConn cn,dbtype
			
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
			<form name="unregform" method="post" action="submitunreg.asp" onsubmit="return submitcheck(this)">

			<table cellpadding="5" cellspacing="0" style="margin-left:auto; margin-right:auto; border-width:0px;">
			<tr>
				<td>�û�����</td>
				<td><input type="text" name="user" maxlength="32" size="32" title="ֻ�ܰ���Ӣ�ġ����ֺ��»��ߣ����32λ" /></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>���룺</td>
				<td><input type="password" name="pass1" maxlength="32" size="32" /></td>
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
				<td colspan="2" style="text-align:center;">
					<input type="submit" value="�ύ" />����<input type="reset" value="��д" />
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