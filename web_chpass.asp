<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/web.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster�������� �޸Ĺ���Ա����</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->

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

	<%Call WebInitHeaderData("","Webmaster��������","","")%><!-- #include file="include/template/header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->
	<div class="region form-region region-longtext">
		<h3 class="title">�޸�����</h3>
		<div class="content">
			<form method="post" action="web_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			<div class="field">
				<span class="label">ԭ����</span>
				<span class="value"><input type="password" name="ioldpass" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">������</span>
				<span class="value"><input type="password" name="inewpass1" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">ȷ������</span>
				<span class="value"><input type="password" name="inewpass2" maxlength="32" /></span>
			</div>
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
    	</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>