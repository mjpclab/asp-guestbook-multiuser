<!-- #include file="webconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif StatusReg=false then
	Response.Redirect "web_err.asp?number=1"
	Response.End
end if

if VcodeCount>0 then session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=web_BookName%> �������Ա�</title>

	<script type="text/javascript">
	//<![CDATA[
	
	function submitcheck(frm)
	{
		if (frm.user.value=='') {alert('�������û�����'); frm.user.focus(); return false;}
		else if (/^\w+$/.test(frm.user.value)==false) {alert('�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�'); frm.user.select(); return false;}
		else if (frm.pass1.value=='') {alert('���������롣'); frm.pass1.focus(); return false;}
		else if (frm.pass2.value=='') {alert('������ȷ�����롣'); frm.pass2.focus(); return false;}
		else if (frm.pass1.value!=frm.pass2.value) {alert('���벻һ�£����顣'); frm.pass2.select(); return false;}
		else if (frm.question.value=='') {alert('�������һ��������⡣'); frm.question.focus(); return false;}
		else if (frm.key.value=='') {alert('�������һ�����𰸡�'); frm.key.focus(); return false;}
		else if (frm.vcode && frm.vcode.value=='') {alert('��������֤�롣'); frm.vcode.focus(); return false;}
		else return true;
	}
	
	<!-- #include file="xmlhttp.inc" -->
	function checkuser(frm,showChecking)
	{
		/*
		if (frm.user.value=='') {alert('�������û�����'); frm.user.focus(); return;}
		else if (/^\w+$/.test(frm.user.value)==false) {alert('�û�����ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�'); frm.user.select(); return;}
		else window.open('checkuser.asp?user=' + frm.user.value,'win_checkuser','width=300,height=185,menubar=no,toolbar=no,status=no,scrollbars=no,resizable=no');
		*/
		var username=frm.user.value;
		var divCheckuser=document.getElementById('divCheckuser');
		if(divCheckuser)
		{
			while(divCheckuser.childNodes.length>0) divCheckuser.removeChild(divCheckuser.lastChild);
			if(username=='') setPureText(divCheckuser,'�������û�����');
			else if(!/^\w+$/.test(username)) setPureText(divCheckuser,'�û���ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߡ�');
			else
			{
				window.xmlHttp=createXmlHttp();
				if(xmlHttp)
				{
					if(showChecking)
						{setPureText(divCheckuser,'���ڼ���û��������Ժ򡭡�');}
					else
						{clearChildren(divCheckuser);}
						
					xmlHttp.abort();
					xmlHttp.onreadystatechange=checkuserComplete;
					xmlHttp.open('GET','checkuser.asp?user=' + username);
					xmlHttp.send(null);
				}
			}
		}
	}
	
	function checkuserComplete()
	{
		if(xmlHttp.readyState==4 && xmlHttp.status==200)
		{
			var divCheckuser=document.getElementById('divCheckuser');
			if(divCheckuser) setPureText(divCheckuser,xmlHttp.responseText);
			xmlHttp.abort();
		}
	}
	
	//]]>
	</script>
	
	<!-- #include file="style.asp" -->
</head>

<body onload="if(regform.user.value=='')regform.user.focus();">

<div id="outerborder" class="outerborder">

<table cellpadding="2" cellspacing="0" style="width:100%; border-width:0px; border-collapse:collapse;">
	<tr style="height:60px;">
		<td style="width:100%; color:<%=TitleColor%>; background-color:<%=TitleBGC%>; background-image:url(<%=TitlePic%>);" nowrap="nowrap">
			<%=web_BookName%> <a href="face.asp" style="color:<%=TitleColor%>">��ҳ</a> &gt;&gt; �������Ա�
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
		<td style="width:100%; text-align:center; color:<%=TableContentColor%>;">
			<form name="regform" method="post" action="submitreg.asp" onsubmit="return submitcheck(this)">

			<table cellpadding="5" cellspacing="0" style="margin:20px auto; border-style:none; border-width:0px;">
				<td>�û���*</td>
				<td>
					<input type="text" name="user" maxlength="32" size="18" title="ֻ�ܰ���Ӣ����ĸ�����ֺ��»��ߣ����32λ" onchange="checkuser(this.form,false)" />
					<input type="button" value="����ʺ�" onclick="checkuser(this.form,true)" />
				</td>
			</tr>
			<tr>
				<td></td>
				<td><div id="divCheckuser" style="width:230px; height:10ex;"></div></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>����*</td>
				<td><input type="password" name="pass1" maxlength="32" size="32" /></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>ȷ������*</td>
				<td><input type="password" name="pass2" maxlength="32" size="32" /></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>�����ǳ�</td>
				<td><input type="text" name="nick" maxlength="64" size="32" /></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>��ʾ����*</td>
				<td><input type="text" name="question" maxlength="32" size="32" title="�����һ�����" /></td>
			</tr>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>�����*</td>
				<td><input type="text" name="key" maxlength="32" size="32" /></td>
			</tr>
			<%if VcodeCount>0 then%>
			<tr style="height:5px;"><td></td></tr>
			<tr>
				<td>��֤��</td>
				<td><input type="text" name="vcode" size="10" /> <img src="web_show_vcode.asp" class="vcode" onclick="this.src=this.src" /></td>
			</tr>
			<%end if%>
			<tr style="height:20px;"><td></td></tr>
			<tr>
				<td colspan="2" style="text-align:center;">
					<input type="submit" value="�ύ" style="width:80px; height:25px;" />����<input type="reset" value="���" style="width:80px; height:25px;" />
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