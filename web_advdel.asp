<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster�������� �߼�ɾ��</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<link rel="stylesheet" type="text/css" href="css/web_adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
	<!-- #include file="css/web_adminstyle.asp" -->

	<script type="text/javascript">
	function del_all_user_warning(frm)
	{
		if(confirm('���棡ɾ�����û������ݽ����ָܻ���\nȷʵҪִ��ɾ��������'))
			if (confirm('���ٴ�ȷ���Ƿ�Ҫɾ���û���'))
				{frm.submit1.disabled=true;return true;}
			else
				return false;
		else
			return false;
	}
	</script>
</head>

<body>

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admincontrols.inc" -->

	<div class="region form-region">
		<h3 class="title">�߼�ɾ��</h3>
		<div class="content">
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ָ�����ں�ʱ��ǰ�����ԣ�����������/ʱ�䡣<br/>
				<input type="hidden" name="option" value="1" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />(ʱ��Ĭ��Ϊ0:0:0)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ָ�����ں�ʱ�������ԣ�����������/ʱ�䡣<br/>
				<input type="hidden" name="option" value="2" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />(ʱ��Ĭ��Ϊ0:0:0)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ���ǰ(����)��n�����ԣ�������n��ֵ��<br/>
				<input type="hidden" name="option" value="3" />
				<input type="text" name="iparam" maxlength="10" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ�����(����)��n�����ԣ�������n��ֵ��<br/>
				<input type="hidden" name="option" value="4" />
				<input type="text" name="iparam" maxlength="10" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ���ƺ��а����ض��ַ��������ԣ�<br/>
				<input type="hidden" name="option" value="5" />
				<input type="text" name="iparam" maxlength="64" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ�������а����ض��ַ��������ԣ�<br/>
				<input type="hidden" name="option" value="6" />
				<input type="text" name="iparam" maxlength="64" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ���ÿ������а����ض��ַ��������ԣ�<br/>
				<input type="hidden" name="option" value="7" />
				<input type="text" name="iparam" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ�������ظ��а����ض��ַ��������ԣ�<br/>
				<input type="hidden" name="option" value="8" />
				<input type="text" name="iparam" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ�������ظ��а����ض��ַ����Ļظ���<br/>
				<input type="hidden" name="option" value="9" />
				<input type="text" name="iparam" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ���ö������а����ض��ַ����Ĺ��棺<br/>
				<input type="hidden" name="option" value="12" />
				<input type="text" name="iparam" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ȫ�����ԣ�������Ա��������ʹ�ã�
				<input type="hidden" name="option" value="10" />
				<input type="hidden" name="iparam" value="DEL_ALL" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ȫ�������ظ��������ʹ�ã�
				<input type="hidden" name="option" value="11" />
				<input type="hidden" name="iparam" value="DEL_ALL_REPLY" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ȫ���ö����棬�����ʹ�ã�
				<input type="hidden" name="option" value="13" />
				<input type="hidden" name="iparam" value="DEL_ALL_DECLARE" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>

			<form method="post" action="web_doadvdel.asp" onsubmit="return del_all_user_warning(this)">
				ɾ��ע��������ָ�����ں�ʱ��ǰ���û�������������/ʱ�䣺<br/>
				<input type="hidden" name="option" value="14" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />(ʱ��Ĭ��Ϊ0:0:0)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="return del_all_user_warning(this)" id=form1 name=form1>
				ɾ������¼������ָ�����ں�ʱ��ǰ���û�������������/ʱ�䣺<br/>
				<input type="hidden" name="option" value="15" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />(ʱ��Ĭ��Ϊ0:0:0)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="return del_all_user_warning(this)" id=form2 name=form2>
				ɾ���������������ָ�����ں�ʱ��ǰ���û�������������/ʱ�䣺<br/>
				<input type="hidden" name="option" value="16" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />(ʱ��Ĭ��Ϊ0:0:0)
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="return del_all_user_warning(this)">
				ɾ����δ��¼�����û��������ʹ�ã�
				<input type="hidden" name="option" value="17" />
				<input type="hidden" name="iparam" value="DEL_NeverLogined_USER" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
			<form method="post" action="web_doadvdel.asp" onsubmit="return del_all_user_warning(this)">
				ɾ��ȫ��ע���û��������ʹ�ã�
				<input type="hidden" name="option" value="18" />
				<input type="hidden" name="iparam" value="DEL_ALL_USER" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
		</div>
	</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>