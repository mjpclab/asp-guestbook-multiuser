<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster管理中心 IP屏蔽策略</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<link rel="stylesheet" type="text/css" href="css/web_adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
	<!-- #include file="css/web_adminstyle.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admincontrols.inc" -->

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

rs.Open Replace(sql_adminipconfig_status,"{0}",wm_name),cn,,,1
tipconstatus=rs("ipconstatus")
rs.Close
%>

<div class="region form-region">
	<h3 class="title">IP屏蔽策略</h3>
	<div class="content">
		<form method="post" action="web_saveipconfig.asp" name="ipconfigform" onsubmit="submit1.disabled=true;">
		<input type="hidden" name="user" value="<%=request("user")%>" />
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
					<%rs.Open Replace(sql_adminipconfig_status1,"{0}",wm_name),cn,,,1
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
					<%rs.Open Replace(sql_adminipconfig_status2,"{0}",wm_name),cn,,,1
					if rs.EOF=false then
						Response.Write "选择要删除的IP段："
						while rs.EOF=false
							tlistid=rs("listid")
							tstartip=rs("startip")
							tendip=rs("endip")
							Response.Write "<input type=""checkbox"" name=""savediplist2"" value=""" &tlistid& """ id=""" &tlistid& """ /><label for=""" &tlistid& """>" & hextoip(tstartip) &"-"& hextoip(tendip) &"</label>"
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