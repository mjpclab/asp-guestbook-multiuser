<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster�������� �޸Ĺ���Ա����</title>

	<script type="text/javascript">
	function checkpass(cobject)
	{
		if (cobject.ioldpass.value.length===0) {alert('������ԭ���롣'); cobject.ioldpass.focus(); return(false);}
		if (cobject.inewpass1.value.length===0) {alert('�����������롣'); cobject.inewpass1.focus(); return(false);}
		if (cobject.inewpass2.value.length===0) {alert('������ȷ�����롣'); cobject.inewpass2.focus(); return(false);}
		if (cobject.inewpass1.value!==cobject.inewpass2.value) {alert('��������ȷ�����벻ͬ�����������롣'); cobject.inewpass1.focus(); return(false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	</script>
	
</head>

<body onload="form4.ioldpass.focus();">

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admincontrols.inc" -->

	<div class="region form-region">
		<h3 class="title">�޸�����</h3>
		<div class="content">
			<form method="post" action="web_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			<div class="field">
				<span class="label">ԭ���룺</span>
				<span class="value"><input type="password" name="ioldpass" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">�����룺</span>
				<span class="value"><input type="password" name="inewpass1" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">ȷ�����룺</span>
				<span class="value"><input type="password" name="inewpass2" maxlength="32" /></span>
			</div>
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
    	</div>
	</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>