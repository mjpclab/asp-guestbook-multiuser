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
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=web_BookName%></title>
	
	<!-- #include file="style.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

<table cellpadding="2" style="width:100%; border-width:0px; border-collapse:collapse;">
	<tr style="height:60px">
		<td style="width:100%; color:<%=TitleColor%>; background-color:<%=TitleBGC%>; background-image:url(<%=TitlePic%>);" nowrap="nowrap">
			<%=web_BookName%>
		</td>
	</tr>
	<tr>
		<td>
			<%
			set cn=server.CreateObject("ADODB.Connection")
			CreateConn cn,dbtype
			
			dim sys_bul_flag
			sys_bul_flag=16
			%>
			<!-- #include file="sysbulletin.inc" -->
			<%cn.Close : set cn=nothing%>
		</td>
	</tr>
	<tr>
		<td style="width:100%;">
		<!-- #include file="func_web.inc" -->
		</td>
	</tr>
	<tr>
		<td style="width:100%; color:<%=TableContentColor%>;">

			<table border="0" width="100%" style="color:<%=TableContentColor%>">
				<tr>
					<td style="width:150px; vertical-align:top;">
						<img src="intro1.gif" style="width:142px; height:159px; border-width:0px;" />
					</td>
					<td style="vertical-align:top;">
						<p style="font-weight:bold">���Խ��ܣ�</p>
			
						<p><span style="font-weight:bold">����֧��UBB��HTML���ɰ�������������ر�</span><br/>
						����ͨ��UBB��HTML��Ϊ�������Ӳ���Ч������ʱ������ʹ��������ȷ�����м���HTML��������ķ����ԣ��������������HTML���ܣ���������Խϰ�ȫ��������ƫ�ڼ򵥵�UBB��ǡ�</p>
						<p><span style="font-weight:bold">����ʱ������رյķÿ�����</span><br/>
						�����ÿͲ�����Ϊ�Ҳ�����ǰ�����Զ������ͨ���ÿ����������⡢�������ݡ������ظ�������</p>
						<p><span style="font-weight:bold">������ɫ������ѡ��</span><br/>
						��������������ѡ���Ա���ɫ�������Ա����ҳɫ������һ�¡�</p>
						<p><span style="font-weight:bold">��֤����Ч��ֹ����ύ</span><br/>
						����ʹ����֤�����Ч��ֹ����ʹ��������⡢�ظ����ύ���ݡ�</p>
						<p><span style="font-weight:bold">IP���β��Խ�������ʿ��֮����</span><br/>
						����ͨ��IP���β��ԣ����Խ�ͼı����֮�˾�֮���⣬Ҳ����ֻ�����ض�IP�ķ��ʣ���Ҫ����ֻ�����ò��Բ�����IP�ζ��ѡ�</p>
						<p><span style="font-weight:bold">���ݹ��˲��ԣ������İ�ȫ��ʩ</span><br/>
						�������˵IP���β�����������ȫ���Ρ�������Ч�Ĺ̶�IP����ô���ݹ��˲��Խ����Ǹ����Ľ���������ò�������ʹ��������ʽ��ƥ������Ե��ַ��������Ұ������Ը������ѡ���滻�ַ�������ֱ�Ӿܾ����ԡ�</p>
						<p><span style="font-weight:bold">��������</span><br/>
						����������ˡ����Ļ����ʼ�֪ͨ�ȡ�</p>
						<br/><br/>
					</td>
				</tr>
			</table>		
		
		</td>
	</tr>
</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>