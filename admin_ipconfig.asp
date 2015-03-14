<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
checkuser cn,rs,false
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 IP屏蔽策略</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admincontrols.inc" -->

	<%
	function hextoip(byref valip)
		dim vi,strip
		strip=""
		for vi=0 to 3
			strip=strip & cstr(cint("&H" & mid(valip,vi*2+1,2)))
			if vi<>3 then strip=strip & "."
		next
		hextoip=strip
	end function

	rs.Open Replace(sql_adminipconfig_status,"{0}",Request.QueryString("user")),cn,,,1
	tipconstatus=rs("ipconstatus")
	rs.Close
	%>
	
<div class="region form-region">
	<h3 class="title">IP屏蔽策略</h3>
	<div class="content">
		<form method="post" action="admin_saveipconfig.asp" name="ipconfigform" onsubmit="submit1.disabled=true;">
		<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
		<table cellpadding="10">
			<tr>
				<td colspan="2"><input type="radio" name="ipconstatus" value="0" id="r1"<%=cked(tipconstatus=0)%> /><label for="r1">不使用IP屏蔽策略</label></td>
			</tr>
			<tr>
				<td style="width:50%; vertical-align:top;">
					<p class="row"><input type="radio" name="ipconstatus" value="1" id="r2"<%=cked(tipconstatus=1)%> /><label for="r2">只屏蔽以下IP段，其余放行</label></p>
					<p class="row">添加新IP段,格式:"起始IP-终止IP"</p>
					<p class="row"><textarea name="txt1" rows="6" style="width:100%"></textarea></p>
					<p class="row">
					<%rs.Open Replace(sql_adminipconfig_status1,"{0}",Request.QueryString("user")),cn,,,1
					if rs.EOF=false then
						Response.Write "选择要删除的IP段："
						while rs.EOF=false
							tlistid=rs("listid")
							tstartip=rs("startip")
							tendip=rs("endip")
							Response.Write "<input type=""checkbox"" name=""savediplist1"" value=""" &tlistid& """ id=""" &tlistid& """ /><label for=""" &tlistid& """>" & hextoip(tstartip) & "-" & hextoip(tendip) &"</label>"
							rs.MoveNext
						wend
						Response.Write ""
					end if
					rs.Close
					%>
					</p>
				</td>
				<td style="width:50%; vertical-align:top;">
					<p class="row"><input type="radio" name="ipconstatus" value="2" id="r3"<%=cked(tipconstatus=2)%> /><label for="r3">只允许以下IP段，其余均不放行</label></p>
					<p class="row">添加新IP段,格式:"起始IP-终止IP"</p>
					<p class="row"><textarea name="txt2" rows="6" style="width:100%"></textarea></p>
					<p class="row">
					<%rs.Open Replace(sql_adminipconfig_status2,"{0}",Request.QueryString("user")),cn,,,1
					if rs.EOF=false then
						Response.Write "选择要删除的IP段："
						while rs.EOF=false
							tlistid=rs("listid")
							tstartip=rs("startip")
							tendip=rs("endip")
							Response.Write "<input type=""checkbox"" name=""savediplist2"" value=""" &tlistid& """ id=""" &tlistid& """ /><label for=""" &tlistid& """>" & hextoip(tstartip) & "-" & hextoip(tendip) &"</label>"
							rs.MoveNext
						wend
						Response.Write ""
					end if
					rs.Close
					%>
					</p>
				</td>
			</tr>
		</table>
		<div class="command"><input type="submit" name="submit1" value="更新数据" /></div>
		</form>
	</div>
</div>

<%cn.Close : set rs=nothing : set cn=nothing%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>