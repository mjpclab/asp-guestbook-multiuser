<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/sysbulletin.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->

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
	<title><%=web_BookName%></title>
	<!-- #include file="inc_stylesheet.asp" -->
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
<!-- #include file="include/template/sysbulletin.inc" -->
<%cn.Close : set cn=nothing%>

<!-- #include file="include/template/web_guest_func.inc" -->


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

<!-- #include file="include/template/footer.inc" -->
</body>
</html>