<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "admin.asp?user=" &Request.QueryString("user")
	Response.End 
end if

'================================
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,false

dim t_html

rs.Open sql_adminreply_reply & Request("id"),cn,,,1
	
if rs.EOF=false then
	t_html=rs("htmlflag")
	c_old="" & rs("reinfo") & ""
	c_old=replace(server.HTMLEncode(c_old),chr(13)&chr(10),"&#13;&#10;")
else
	t_html=adminlimit
	c_old=""
end if

rs.Close
cn.close

%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 回复留言</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->

	<script type="text/javascript">
	//<![CDATA[
	
	function submitcheck(cobject)
	{
		if (cobject.rcontent.value=="") {alert('请输入回复内容。'); cobject.rcontent.focus(); return (false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	function sfocus()
	{
		if (!form3.rcontent.modified)
		{
			form3.rcontent.focus();
			form3.rcontent.select();
		}
	}
	
	//]]>
	</script>
</head>

<body<%=bodylimit%> onload="sfocus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

<%if ShowTitle=true then show_book_title 3,"管理"%>
<!-- #include file="admintool.inc" -->

<table cellpadding="2" class="generalwindow" ID="Table1">
	<tr>
		<td class="centertitle">回复留言</td>
	</tr>
	<tr>
		<td class="wordscontent" style="text-align:center; padding:20px 2px;">
			<form method="post" action="admin_savereply.asp" onsubmit="return submitcheck(this)" name="form3">
			回复内容：<br/>
			<textarea name="rcontent" id="rcontent" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" cols="<%=ReplyTextWidth%>" rows="<%=ReplyTextHeight%>"><%=c_old%></textarea>
			<!-- #include file="ubbtoolbar.inc" -->
			<%if web_AdminUBBSupport then ShowUbbToolBar(2)%>
			<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
			<input type="hidden" name="rootid" value="<%=request.QueryString("rootid")%>" />
			<input type="hidden" name="mainid" value="<%=request.QueryString("id")%>" />
			<input type="hidden" name="page" value="<%=request.QueryString("page")%>" />
			<input type="hidden" name="type" value="<%=request.QueryString("type")%>" />
			<input type="hidden" name="searchtxt" value="<%=request.QueryString("searchtxt")%>" ID="Hidden6"/>
			<p style="text-align:left;">
				<%if web_AdminHTMLSupport=true then%><input type="checkbox" name="html1" id="html1" value="1"<%if cint(t_html and 1)<>0 then Response.Write " checked=""checked""" %> /><label for="html1">支持HTML标记</label><br/><%end if%>
				<%if web_AdminUBBSupport=true then%><input type="checkbox" name="ubb1" id="ubb1" value="1"<%if cint(t_html and 2)<>0 then Response.Write " checked=""checked""" %> /><label for="ubb1">支持UBB标记</label><br/><%end if%>
				<%if web_AdminAllowNewLine=true then%><input type="checkbox" name="newline1" id="newline1" value="1"<%if cint(t_html and 4)<>0 then Response.Write " checked=""checked""" %> /><label for="newline1">不支持HTML和UBB标记时允许回车换行</label><br/><%end if%>
				<br/>
				<input type="checkbox" name="lock2top" id="lock2top" value="1" /><label for="lock2top">回复后置顶留言</label><br/>
				<input type="checkbox" name="bring2top" id="bring2top" value="1" /><label for="bring2top">回复后提前留言</label>
			</p>
			<input type="submit" value="发表回复" name="submit1" id="submit1" />
			</form>
		</td>
	</tr>
</table>

<%
CreateConn cn,dbtype
rs.Open Replace(Replace(sql_adminreply_words,"{0}",Request.QueryString("id")),"{1}",Request.QueryString("user")),cn,,,1
	
if rs.EOF=false then
	dim pagename
	pagename="admin_reply"
	%><!-- #include file="listword_admin.inc" --><%
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing	
%>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>