<%
sub WebErrorPage(errorCode)
	Response.Status="403 Forbidden"

	dim errmsg
	select case errorCode
	case 1
		errmsg="�Բ������Ա����������ѹرգ����Ժ����ԡ�"
	case 2
		errmsg="�Բ����һ����빦���ѹرգ����Ժ����ԡ�"
	case 3
		errmsg="�Բ��𣬹���ά�������ѹرգ����Ժ����ԡ�"
	case 4
		errmsg="�Բ������ķ����ѱ���ֹ�����Ժ����ԡ�"
	case 5
		errmsg="�Բ�����ɾ�ʺŹ����ѹرգ����Ժ����ԡ�"
	case else
		errmsg="δ֪��������ϵ����Ա��"
	end select
	%>
	<!-- #include file="include/template/dtd.inc" -->
	<html>
	<head>
		<!-- #include file="include/template/metatag.inc" -->
		<title><%=web_BookName%> ����</title>
		<!-- #include file="inc_web_admin_stylesheet.asp" -->
	</head>

	<body>

	<%Call WebInitHeaderData("","����","","")%><!-- #include file="include/template/header.inc" -->
	<div id="outerborder" class="outerborder">
		<div id="mainborder" class="mainborder">
		<!-- #include file="include/template/web_guest_func.inc" -->
		<p style="margin-bottom: 3em;">��������<%=errmsg%></p>
		</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
	</body>
	</html>
	<%
end sub
%>
