<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=HomeName%> ���Ա� �����ö�����</title>

	<script type="text/javascript">
	//<![CDATA[
	
	function sfocus()
	{
		if (!form6.abulletin.modified)
		{
			form6.abulletin.focus();
			form6.abulletin.select();
		}
	}
	
	//]]>
	</script>
	
	<!-- #include file="style.asp" -->
</head>

<body<%=bodylimit%> onload="sfocus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admintool.inc" -->

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	CreateConn cn,dbtype
	checkuser cn,rs,false

	rs.Open Replace(sql_adminsetbulletin,"{0}",Request.QueryString("user")),cn,,,1
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

	<table border="1" cellpadding="2" bordercolor="<%=TableBorderColor%>" class="generalwindow">
		<tr>
			<td class="centertitle">�����ö�����</td>
		</tr>
		<tr>
			<td class="wordscontent" style="text-align:center; padding:20px 2px;">
				<form method="post" action="admin_savebulletin.asp" name="form6" onsubmit="form6.submit1.disabled=true;">
				<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
				�������ݣ�<br/><textarea name="abulletin" id="abulletin" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" cols="<%=ReplyTextWidth%>" rows="<%=ReplyTextHeight%>"><%=tbul%></textarea>
				<!-- #include file="ubbtoolbar.inc" -->
				<%if web_AdminUBBSupport then ShowUbbToolBar(2)%>
				<p style="text-align:left;">
				<%if web_AdminHTMLSupport=true then%><input type="checkbox" name="html2" id="html2" value="1"<%if cint(tflag and 1)<>0 then Response.Write " checked=""checked"""%> /><label for="html2">֧��HTML���</label><br/><%end if%>
				<%if web_AdminUBBSupport=true then%><input type="checkbox" name="ubb2" id="ubb2" value="1"<%if cint(tflag and 2)<>0 then Response.Write " checked=""checked"""%> /><label for="ubb2">֧��UBB���</label><br/><%end if%>
				<%if web_AdminAllowNewLine=true then%><input type="checkbox" name="newline2" id="newline2" value="1"<%if cint(tflag and 4)<>0 then Response.Write " checked=""checked"""%> /><label for="newline2">��֧��HTML��UBB���ʱ����س�����</label><%end if%>
				</p>
				<input value="��������" type="submit" name="submit1" id="submit1" />
				</form>
			</td>
		</tr>
	</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>