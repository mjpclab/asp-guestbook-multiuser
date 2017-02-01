<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/web_admin_userinfo.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires = -1

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
checkuser cn,rs,false

Dim status, statusDisabled
rs.Open Replace(sql_webuserinfo_status,"{0}",adminid), cn, 0, 1, 1
status=rs.Fields(0)
statusDisabled=CBool(status AND 1073741824)
statusLoginDisabled=CBool(status AND 536870912)
statusLeavewordDisabled=CBool(status AND 268435456)

rs.Close
cn.Close
set rs=nothing
set cn=nothing

%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster�������� �����˺�״̬</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->

</head>

<body>

<div id="outerborder" class="outerborder">

	<%Call WebInitHeaderData("","Webmaster��������","","")%><!-- #include file="include/template/header.inc" -->
	<div id="mainborder" class="mainborder">
	<div class="region form-region region-longtext region-config">
		<h3 class="title">�����˺�״̬</h3>
		<div class="content">
			<form method="post" action="web_saveuserstatus.asp">
			<div class="field">
				<span class="label">�û�����</span>
				<span class="value"><input type="text" name="user" maxlength="32" value="<%=ruser%>" /></span>
			</div>
			<div class="field">
				<span class="label">�˺�״̬��</span>
				<span class="value">
					<input type="radio" name="statusDisabled" value="0" id="statusDisabled0"<%=cked(Not statusDisabled)%>/><label for="statusDisabled0">����</label>����<input type="radio" name="statusDisabled" value="1073741824" id="statusDisabled1"<%=cked(statusDisabled)%>/><label for="statusDisabled1">����</label>
				</span>
			</div>
			<div class="field">
				<span class="label">�˺Ź����½״̬��</span>
				<span class="value">
					<input type="radio" name="statusLoginDisabled" value="0" id="statusLoginDisabled0"<%=cked(Not statusLoginDisabled)%>/><label for="statusLoginDisabled0">����</label>����<input type="radio" name="statusLoginDisabled" value="536870912" id="statusLoginDisabled1"<%=cked(statusLoginDisabled)%>/><label for="statusLoginDisabled1">����</label>
				</span>
			</div>
			<div class="field">
				<span class="label">�˺�ǩд����״̬��</span>
				<span class="value">
					<input type="radio" name="statusLeavewordDisabled" value="0" id="statusLeavewordDisabled0"<%=cked(Not statusLeavewordDisabled)%>/><label for="statusLeavewordDisabled0">����</label>����<input type="radio" name="statusLeavewordDisabled" value="268435456" id="statusLeavewordDisabled1"<%=cked(statusLeavewordDisabled)%>/><label for="statusLeavewordDisabled1">����</label>
				</span>
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