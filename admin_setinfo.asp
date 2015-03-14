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
	<title><%=HomeName%> ���Ա� �޸İ�������</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admincontrols.inc" -->

	<%
	rs.Open Replace(sql_adminsetinfo,"{0}",Request.QueryString("user")),cn,,,1
	tfaceid=rs("faceid")
	%>

	<div class="region">
		<h3 class="title">�޸İ�������</h3>
		<div class="content">
			<form method="post" action="admin_saveinfo.asp" name="form1" onsubmit="form1.submit1.disabled=true;">
			<input type="hidden" name="user" value="<%=Request.QueryString("user")%>" />
			<div class="field">
				<span class="label">�ǳƣ�</span>
				<span class="value"><input type="text" name="aname" size="<%=SetInfoTextWidth%>" maxlength="20" value="<%="" & rs("name") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">�ʼ���</span>
				<span class="value"><input type="text" name="aemail" size="<%=SetInfoTextWidth%>" maxlength="50" value="<%="" & rs("email") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">QQ�ţ�</span>
				<span class="value"><input type="text" name="aqqid" size="<%=SetInfoTextWidth%>" maxlength="16" value="<%="" & rs("qqid") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">Skype��</span>
				<span class="value"><input type="text" name="amsnid" size="<%=SetInfoTextWidth%>" maxlength="50" value="<%="" & rs("msnid") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">��ҳ��</span>
				<span class="value"><input type="text" name="ahomepage" size="<%=SetInfoTextWidth%>" maxlength="127" value="<%="" & rs("homepage") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">ͷ���ţ�</span>
				<span class="value"><input type="text" name="afaceid" size="<%=SetInfoTextWidth%>" maxlength="3" value="<%=tfaceid%>" title="��дͷ����ʱURL�������" /></span>
			</div>
			<div class="field">
				<span class="label">��URL��</span>
				<span class="value"><input type="text" name="afaceurl" size="<%=SetInfoTextWidth%>" maxlength="127" value="<%="" & rs("faceurl") & ""%>" title="��дURLʱ����ͷ����" /></span>
			</div>
			<div class="field">
				<%rs.Close : cn.Close : set rs=nothing : set cn=nothing

				dim listfacecount,defaultindex
				listfacecount=FrequentFaceCount
				defaultindex=tfaceid%>
				<!-- #include file="listface.inc" -->
			</div>
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>