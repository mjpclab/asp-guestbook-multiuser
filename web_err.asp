<!-- #include file="webconfig.asp" -->
<%Response.Expires=-1%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
	<title><%=web_BookName%> ����</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" class="booktitle">��ҳ</a> &gt;&gt; ����</div>
</div>

<!-- #include file="include/web_guest_func.inc" -->

<%
dim errmsg
select case Request("number")
case "1"
	errmsg="�Բ������Ա����������ѹرգ����Ժ����ԡ�"
case "2"
	errmsg="�Բ����һ����빦���ѹرգ����Ժ����ԡ�"
case "3"
	errmsg="�Բ��𣬹���ά�������ѹرգ����Ժ����ԡ�"
case "4"
	errmsg="�Բ������ķ����ѱ���ֹ�����Ժ����ԡ�"
case "5"
	errmsg="�Բ�����ɾ�ʺŹ����ѹرգ����Ժ����ԡ�"
case else
	errmsg="δ֪��������ϵ����Ա��"
end select
Response.Write "<br/><br/><p>��������" &errmsg& "</p><br/><br/><br/>"
%>
</div>

<!-- #include file="include/footer.inc" -->
</body>
</html>