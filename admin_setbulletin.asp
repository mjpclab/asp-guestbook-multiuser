<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� �����ö�����</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->

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

<body<%=bodylimit%> onload="sfocus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admincontrols.inc" -->

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	CreateConn cn,dbtype
	checkuser cn,rs,false

	rs.Open Replace(sql_adminsetbulletin,"{0}",adminid),cn,,,1
	dim tbul,tflag
	tflag=rs("declareflag")
	tbul=rs("declare")

	if isnull(tbul) then
		tbul=""
	else
		tbul=replace(server.HTMLEncode(tbul),chr(13)&chr(10),"&#13;&#10;")
	end if
	if tbul="" then tflag=(web_adminlimit and adminlimit)

	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	%>

	<div class="region">
		<h3 class="title">�����ö�����</h3>
		<div class="content">
			<form method="post" action="admin_savebulletin.asp" name="form6" onsubmit="form6.submit1.disabled=true;">
				<input type="hidden" name="user" value="<%=ruser%>" />
				�������ݣ�<br/><textarea name="abulletin" id="abulletin" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" rows="<%=ReplyTextHeight%>"><%=tbul%></textarea>
				<!-- #include file="ubbtoolbar.inc" -->
				<%if web_AdminUBBSupport then ShowUbbToolBar(2)%>
				<p>
				<%if web_AdminHTMLSupport=true then%><input type="checkbox" name="html2" id="html2" value="1"<%if cint(tflag and 1)<>0 then Response.Write " checked=""checked"""%> /><label for="html2">֧��HTML���</label><br/><%end if%>
				<%if web_AdminUBBSupport=true then%><input type="checkbox" name="ubb2" id="ubb2" value="1"<%if cint(tflag and 2)<>0 then Response.Write " checked=""checked"""%> /><label for="ubb2">֧��UBB���</label><br/><%end if%>
				<%if web_AdminAllowNewLine=true then%><input type="checkbox" name="newline2" id="newline2" value="1"<%if cint(tflag and 4)<>0 then Response.Write " checked=""checked"""%> /><label for="newline2">��֧��HTML��UBB���ʱ����س�����</label><%end if%>
				</p>
				<div class="command"><input value="��������" type="submit" name="submit1" id="submit1" /></div>
			</form>
		</div>
	</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>