<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster�������� ����ϵͳ����</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<link rel="stylesheet" type="text/css" href="css/web_adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
	<!-- #include file="css/web_adminstyle.asp" -->

	<script type="text/javascript">
	function sfocus()
	{
		if (!form6.abulletin.modified)
		{
			form6.abulletin.focus();
			form6.abulletin.select();
		}
	}
	</script>
</head>

<body onload="sfocus();">

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admincontrols.inc" -->

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	CreateConn cn,dbtype

	rs.Open sql_websetbulletin,cn,,,1
	dim tbul,tflag

	tflag=rs("bulletinflag")
	tbul="" & rs("webbulletin") & ""

	if isnull(tbul) then
		tbul=""
	else
		tbul=replace(server.HTMLEncode(tbul),chr(13)&chr(10),"&#13;&#10;")
	end if

	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	%>

	<div class="region">
		<h3 class="title">����ϵͳ����</h3>
		<div class="content">
			<form method="post" action="web_savebulletin.asp" name="form6" onsubmit="form6.submit1.disabled=true;">
			�������ݣ�<br/><textarea name="abulletin" id="abulletin" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" rows="<%=ReplyTextHeight%>"><%=tbul%></textarea>
			<!-- #include file="ubbtoolbar.inc" -->
			<%ShowUbbToolBar(1)%>
			<p>
			<input type="checkbox" name="html2" id="html2" value="1"<%if cint(tflag and 1)<>0 then Response.Write " checked=""checked"""%> /><label for="html2">֧��HTML���</label><br/>
			<input type="checkbox" name="ubb2" id="ubb2" value="1"<%if cint(tflag and 2)<>0 then Response.Write " checked=""checked"""%> /><label for="ubb2">֧��UBB���</label><br/>
			<input type="checkbox" name="newline2" id="newline2" value="1"<%if cint(tflag and 4)<>0 then Response.Write " checked=""checked"""%> /><label for="newline2">��֧��HTML��UBB���ʱ����س�����</label><br/><br/>

			<input type="checkbox" name="pub_at_face" id="pub_at_face" value="1"<%if cint(tflag and 16)<>0 then Response.Write " checked=""checked"""%> /><label for="pub_at_face">����ҳ����</label><br/>
			<input type="checkbox" name="pub_at_function" id="pub_at_function" value="1"<%if cint(tflag and 32)<>0 then Response.Write " checked=""checked"""%> /><label for="pub_at_function">�ڹ���ҳ����(ע�ᡢά������)</label><br/>
			<input type="checkbox" name="pub_at_index" id="pub_at_index" value="1"<%if cint(tflag and 64)<>0 then Response.Write " checked=""checked"""%> /><label for="pub_at_index">���û����ԡ�����ҳ����</label><br/>
			<input type="checkbox" name="pub_at_search" id="pub_at_search" value="1"<%if cint(tflag and 128)<>0 then Response.Write " checked=""checked"""%> /><label for="pub_at_search">���û�����Ա��ҳ����</label>
			</p>
			<div class="command"><input value="��������" type="submit" name="submit1" id="submit1" /></div>
			</form>
		</div>
	</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>