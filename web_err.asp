<!-- #include file="webconfig.asp" -->
<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> ����</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<link rel="stylesheet" type="text/css" href="css/web_adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
	<!-- #include file="css/web_adminstyle.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%> <a href="face.asp" class="booktitle">��ҳ</a> &gt;&gt; ����</div>
</div>

<!-- #include file="func_web.inc" -->

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

<!-- #include file="bottom.asp" -->
</body>
</html>