<!-- #include file="webconfig.asp" -->
<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> ����</title>
	
	<!-- #include file="style.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<table cellpadding="2" style="width:100%; border-width:0px; border-collapse:collapse;">
		<tr style="height:60px;">
			<td style="width:100%; color:<%=TitleColor%>; background-color:<%=TitleBGC%>; background-image:url(<%=TitlePic%>);">
				<%=web_BookName%> <a href="face.asp" class="booktitle">��ҳ</a> &gt;&gt; ����
			</td>
		</tr>
	</table>
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
	Response.Write "<br/><br/><span style=""width:100%; text-align:left;"">��������" &errmsg& "</span><br/><br/><br/>"
	%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>