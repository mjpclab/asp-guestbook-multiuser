<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="include/utility/book.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> ���Ա� ����</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<div class="guest-functions">
		<div class="aside">
			<a class="function" href="reg.asp">�������Ա�</a>
			<a class="function" href="admin.asp?user=<%=ruser%>">����</a>
		</div>
	</div>

	<%
	dim errmsg
	select case Request("number")
	case "1"
		errmsg="�Բ������ķ����ѱ���ֹ��"
	case "2"
		errmsg="��Ǹ�����Ա��ѹرգ����Ժ����ԡ�"
	case "3"
		errmsg="��Ǹ������Ȩ���ѹرգ����Ժ����ԡ�"
	case "4"
		errmsg="�Բ������������к��н�ֹ���ֵ����ݡ�"
	case "5"
		errmsg="��Ǹ������Ȩ���ѹرգ����Ժ����ԡ�"
	case "6"
		errmsg="�Բ������ķ����ٶ�̫���ˣ�����Ϣһ�¡�"
	case "7"
		errmsg="�Բ����벻Ҫ�����ظ����ݡ�"
	case else
		errmsg="δ֪��������ϵ����Ա��"
	end select
	Response.Write "<br/><br/><span style=""width:100%; text-align:left;"">��������" &errmsg& "</span><br/><br/><br/>"
	%>
</div>

<!-- #include file="include/template/footer.inc" -->
</body>
</html>