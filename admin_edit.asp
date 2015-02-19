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

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
checkuser cn,rs,false

rs.Open Replace(Replace(sql_adminedit,"{0}",Request.QueryString("id")),"{1}",Request.QueryString("user")),cn,,,1
	
if rs.EOF=false then
	guestflag=rs("guestflag")
	guest_txt="" & rs("article") & ""
	guest_txt=replace(server.htmlEncode(guest_txt),chr(13)&chr(10),"&#13;&#10;")
else
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "admin.asp?user=" &Request.QueryString("user")
	Response.End
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn">
<head>
	<title><%=HomeName%> 留言本 编辑留言</title>

	<script type="text/javascript">
	//<![CDATA[
	
	function submitcheck(cobject)
	{
		if (cobject.etitle.value=="") {alert('请输入标题。'); cobject.etitle.focus(); return false;}
		cobject.submit1.disabled=true;
		return (true);
	}
	function sfocus()
	{
		if (!form3.modified)
		{
			form3.econtent.focus();
			form3.econtent.select();
		}
	}
	
	//]]>
	</script>

	<!-- #include file="style.asp" -->
</head>

<body<%=bodylimit%> onload="sfocus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

<%if ShowTitle=true then show_book_title 3,"管理"%>
<!-- #include file="admintool.inc" -->

<table cellpadding="2" class="generalwindow" ID="Table1">
	<tr>
		<td class="centertitle">编辑留言</td>
	</tr>
	<tr>
		<td class="wordscontent" style="text-align:center; padding:20px 2px;">
			<form method="post" action="admin_saveedit.asp" onsubmit="return submitcheck(this)" name="form3">
			标题：<br/>
			<input type="text" name="etitle" onkeydown="if(!this.form.modified)this.form.modified=true;" size="<%=ReplyTextWidth%>" maxlength="30" value="<%=rs.Fields("title")%>"><br/><br/>
			内容：<br/>
			<textarea name="econtent" id="econtent" onkeydown="if(!this.form.modified)this.form.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" cols="<%=ReplyTextWidth%>" rows="<%=ReplyTextHeight%>"><%=guest_txt%></textarea>
			<!-- #include file="ubbtoolbar.inc" -->
			<%if web_AdminUBBSupport then ShowUbbToolBar(2)%>
			<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
			<input type="hidden" name="rootid" value="<%=request.QueryString("rootid")%>" />
			<input type="hidden" name="mainid" value="<%=request.QueryString("id")%>" />
			<input type="hidden" name="page" value="<%=request.QueryString("page")%>" />
			<input type="hidden" name="type" value="<%=request.QueryString("type")%>" />
			<input type="hidden" name="searchtxt" value="<%=request.QueryString("searchtxt")%>" />
			<p style="text-align:left;">
			<%if web_AdminHTMLSupport=true then%><input type="checkbox" name="html1" id="html1" value="1"<%if cint(guestflag and 1)<>0 then Response.Write " checked=""checked"""%> /><label for="html1">支持HTML标记</label><br/><%end if%>
			<%if web_AdminUBBSupport=true then%><input type="checkbox" name="ubb1" id="ubb1" value="1"<%if cint(guestflag and 2)<>0 then Response.Write " checked=""checked"""%> /><label for="ubb1">支持UBB标记</label><br/><%end if%>
			<%if web_AdminAllowNewLine=true then%><input type="checkbox" name="newline1" id="newline1" value="1"<%if cint(guestflag and 4)<>0 then Response.Write " checked=""checked"""%> /><label for="newline1">不支持HTML和UBB标记时允许回车换行</label><%end if%>
			</p>
			<input type="submit" value="保存留言" name="submit1" id="submit1" />
			</form>
		</td>
	</tr>
</table>

<%dim pagename
pagename="admin_edit"%>
<!-- #include file="listword_admin.inc" -->
<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>