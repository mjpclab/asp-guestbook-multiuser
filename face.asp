<!-- #include file="webconfig.asp" -->

<%
Response.Expires=-1
if web_isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or web_isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%></title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<div class="header">
	<div class="breadcrumb"><%=web_BookName%></div>
</div>

<%
set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

dim sys_bul_flag
sys_bul_flag=16
%>
<!-- #include file="sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="func_web.inc" -->


<div style="overflow: hidden; padding: 10px;">
	<img src="intro1.gif" style="float: left; margin-right: 10px;" />
	<div style="overflow: hidden;">
		<h3>���Խ��ܣ�</h3>

		<h4>����֧��UBB��HTML���ɰ�������������ر�</h4>
		<p>ͨ��UBB��HTML��Ϊ�������Ӳ���Ч������ʱ������ʹ��������ȷ�����м���HTML��������ķ����ԣ��������������HTML���ܣ���������Խϰ�ȫ��������ƫ�ڼ򵥵�UBB��ǡ�</p>

		<h4>����ʱ������رյķÿ�����</h4>
		<p>�ÿͲ�����Ϊ�Ҳ�����ǰ�����Զ������ͨ���ÿ����������⡢�������ݡ������ظ�������</p>

		<h4>������ɫ������ѡ��</h4>
		<p>����������ѡ���Ա���ɫ�������Ա����ҳɫ������һ�¡�</p>

		<h4>��֤����Ч��ֹ����ύ</h4>
		<p>ʹ����֤�����Ч��ֹ����ʹ��������⡢�ظ����ύ���ݡ�</p>

		<h4>IP���β��Խ�������ʿ��֮����</h4>
		<p>ͨ��IP���β��ԣ����Խ�ͼı����֮�˾�֮���⣬Ҳ����ֻ�����ض�IP�ķ��ʣ���Ҫ����ֻ�����ò��Բ�����IP�ζ��ѡ�</p>

		<h4>���ݹ��˲��ԣ������İ�ȫ��ʩ</h4>
		<p>���˵IP���β�����������ȫ���Ρ�������Ч�Ĺ̶�IP����ô���ݹ��˲��Խ����Ǹ����Ľ���������ò�������ʹ��������ʽ��ƥ������Ե��ַ��������Ұ������Ը������ѡ���滻�ַ�������ֱ�Ӿܾ����ԡ�</p>

		<h4>��������</h4>
		<p>������ˡ����Ļ����ʼ�֪ͨ�ȡ�</p>
	</div>
</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>