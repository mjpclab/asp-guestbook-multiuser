<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
checkuser cn,rs,false

%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
	<title><%=HomeName%> ���Ա� ���Ա�����</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="include/admin_mainmenu.inc" -->
	
	<%rs.Open Replace(sql_adminconfig_config,"{0}",adminid),cn,,,1%>

	<div class="region region-config admin-tools">
		<h3 class="title">���Ա���������</h3>
		<div class="content flex-box">
			<ul>
				<li><a href="admin_config.asp?user=<%=ruser%>&amp;page=1">��������</a></li>
				<li><a href="admin_config.asp?user=<%=ruser%>&amp;page=2">�ʼ�֪ͨ</a></li>
				<li><a href="admin_config.asp?user=<%=ruser%>&amp;page=4">����ߴ�</a></li>
				<li><a href="admin_config.asp?user=<%=ruser%>&amp;page=8">��������</a></li>
				<li><a href="admin_config.asp?user=<%=ruser%>">ȫ������</a></li>
			</ul>

			<form method="post" action="admin_saveconfig.asp" name="configform" onsubmit="return check();">
			<%
			showpage=15
			if isnumeric(Request.QueryString("page")) and Request.QueryString("page")<>"" then showpage=clng(Request.QueryString("page"))

			if clng(showpage and 1)<> 0 then
			%>
				<h4>��������</h4>
				<%tstatus=rs("status")%>
				<div class="field">
					<span class="label">���Ա�״̬��</span>
					<span class="value"><input type="radio" name="status1" value="1" id="status11"<%=cked(StatusOpen)%><%=dised(web_StatusOpen=false)%> /><label for="status11"<%=dised(web_StatusOpen=false)%>>����</label>����<input type="radio" name="status1" value="0" id="status12"<%=cked(StatusOpen=false)%><%=dised(web_StatusOpen=false)%> /><label for="status12"<%=dised(web_StatusOpen=false)%>>�ر�</label> (�ر�ʱ"�ÿ�����Ȩ��"������Ч)</span>
				</div>
				<div class="field">
					<span class="label">�ÿ�����Ȩ�ޣ�</span>
					<span class="value"><input type="radio" name="status2" value="2" id="status21"<%=cked(StatusWrite)%><%=dised(web_StatusWrite=false)%> /><label for="status21"<%=dised(web_StatusWrite=false)%>>����</label>����<input type="radio" name="status2" value="0" id="status22"<%=cked(StatusWrite=false)%><%=dised(web_StatusWrite=false)%> /><label for="status22"<%=dised(web_StatusWrite=false)%>>�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ���������Ȩ�ޣ�</span>
					<span class="value"><input type="radio" name="status3" value="4" id="status31"<%=cked(StatusSearch)%><%=dised(web_StatusSearch=false)%> /><label for="status31"<%=dised(web_StatusSearch=false)%>>����</label>����<input type="radio" name="status3" value="0" id="status32"<%=cked(StatusSearch=false)%><%=dised(web_StatusSearch=false)%> /><label for="status32"<%=dised(web_StatusSearch=false)%>>�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�ͷ���ܣ�</span>
					<span class="value"><input type="radio" name="status7" value="8" id="status71"<%=cked(clng(tstatus and 8)<>0)%> /><label for="status71">����</label>����<input type="radio" name="status7" value="0" id="status72"<%=cked(clng(tstatus and 8)=0)%> /><label for="status72">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">������ʾǰ����ˣ�</span>
					<span class="value"><input type="radio" name="status4" value="16" id="status41"<%=cked(clng(tstatus and 16)<>0)%> /><label for="status41">���</label>����<input type="radio" name="status4" value="0" id="status42"<%=cked(clng(tstatus and 16)=0)%> /><label for="status42">�����</label></span>
				</div>
				<div class="field">
					<span class="label">���Ļ����ܣ�</span>
					<span class="value"><input type="radio" name="status5" value="32" id="status51"<%=cked(clng(tstatus and 32)<>0)%> /><label for="status51">����</label>����<input type="radio" name="status5" value="0" id="status52"<%=cked(clng(tstatus and 32)=0)%> /><label for="status52">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�Ϊ���Ļ����ܣ�</span>
					<span class="value"><input type="radio" name="status6" value="64" id="status61"<%=cked(clng(tstatus and 64)<>0)%> /><label for="status61">����</label>����<input type="radio" name="status6" value="0" id="status62"<%=cked(clng(tstatus and 64)=0)%> /><label for="status62">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">����ÿͻظ���</span>
					<span class="value"><input type="radio" name="status8" value="128" id="status81"<%=cked(clng(tstatus and 128)<>0)%> /><label for="status81">����</label>����<input type="radio" name="status8" value="0" id="status82"<%=cked(clng(tstatus and 128)=0)%> /><label for="status82">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">����ͳ�ƣ�</span>
					<span class="value"><input type="radio" name="status9" value="256" id="status91"<%=cked(clng(tstatus and 256)<>0)%> /><label for="status91">����</label>����<input type="radio" name="status9" value="0" id="status92"<%=cked(clng(tstatus and 256)=0)%> /><label for="status92">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">������վLogo��ַ��</span>
					<span class="value"><input type="text" class="longtext" maxlength="127" name="homelogo" value="<%=rs("homelogo")%>" /></span>
				</div>
				<div class="field">
					<span class="label">������վ���ƣ�</span>
					<span class="value"><input type="text" class="longtext" maxlength="30" name="homename" value="<%=rs("homename")%>" /></span>
				</div>
				<div class="field">
					<span class="label">������վ��ַ��</span>
					<span class="value"><input type="text" class="longtext" maxlength="127" name="homeaddr" value="<%=rs("homeaddr")%>" /></span>
				</div>
				<%adminlimit=rs("adminhtml")%>
				<div class="field">
					<span class="label">����Ա��ȫ�����ã�</span>
					<span class="value"><input type="checkbox" value="1" name="adminviewcode" id="adminviewcode"<%=cked(clng(adminlimit and 8)<>0)%> /><label for="adminviewcode">�ÿ�������ʾʵ��HTML��UBB����</label></span>
				</div>
				<div class="field">
					<span class="label">����ԱĬ��HTMLȨ�ޣ�</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="adminhtml" id="adminhtml"<%=cked(clng(web_adminlimit and adminlimit and 1)<>0)%><%=dised(web_AdminHTMLSupport=false)%> /><label for="adminhtml"<%=dised(web_AdminHTMLSupport=false)%>>����Ա�ظ�������Ĭ��֧��HTML</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminubb" id="adminubb"<%=cked(clng(web_adminlimit and adminlimit and 2)<>0)%><%=dised(web_AdminUBBSupport=false)%> /><label for="adminubb"<%=dised(web_AdminUBBSupport=false)%>>����Ա�ظ�������Ĭ��֧��UBB</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminertn" id="adminertn"<%=cked(clng(web_adminlimit and adminlimit and 4)<>0)%><%=dised(web_AdminAllowNewLine=false)%> /><label for="adminertn"<%=dised(web_AdminAllowNewLine=false)%>>����Ա��֧��HTML��UBBʱ��Ĭ������س�����</label></span>
					</span>
				</div>
				<%guestlimit=rs("guesthtml")%>
				<div class="field">
					<span class="label">�ÿ�HTMLȨ�ޣ�</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="guesthtml" id="guesthtml"<%=cked(clng(web_guestlimit and guestlimit and 1)<>0)%><%=dised(web_HTMLSupport=false)%> /><label for="guesthtml"<%=dised(web_HTMLSupport=false)%>>�ÿ�����֧��HTML</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestubb" id="guestubb"<%=cked(clng(web_guestlimit and guestlimit and 2)<>0)%><%=dised(web_UBBSupport=false)%> /><label for="guestubb"<%=dised(web_UBBSupport=false)%>>�ÿ�����֧��UBB</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestertn" id="guestertn"<%=cked(clng(web_guestlimit and guestlimit and 4)<>0)%><%=dised(web_AllowNewLine=false)%> /><label for="guestertn"<%=dised(web_AllowNewLine=false)%>>�ÿͲ�֧��HTML��UBBʱ������س�����</label></span>
					</span>
				</div>
				<div class="field">
					<span class="label">����Ա��¼��ʱ��</span>
					<span class="value"><input type="text" size="4" maxlength="4" name="admintimeout" value="<%=rs("admintimeout")%>" />�� (Ĭ��=20)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ�ÿ���ʾIPv4ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="showipv4" value="<%=ShowIPv4%>" />�ֽ� (��ѡֵ��0��4)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ�ÿ���ʾIPv6ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="showipv6" value="<%=ShowIPv6%>" />�� (��ѡֵ��0��8)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾIPv4ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshowipv4" value="<%=AdminShowIPv4%>" />�ֽ� (��ѡֵ��0��4)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾIPv6ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshowipv6" value="<%=AdminShowIPv6%>" />�� (��ѡֵ��0��8)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾԭIPv4ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshoworiginalipv4" value="<%=AdminShowOriginalIPv4%>" />�ֽ� (��ѡֵ��0��4,ʹ�ô��������ʱ������ʾԭʼIP)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾԭIPv6ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshoworiginalipv6" value="<%=AdminShowOriginalIPv6%>" />�� (��ѡֵ��0��8,ʹ�ô��������ʱ������ʾԭʼIP)</span>
				</div>
				<div class="field">
					<span class="label">��¼��֤�볤�ȣ�</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="vcodecount" value="<%=VcodeCount%>" />λ (��ѡֵ��<%=web_VcodeCount%>��10)</span>
				</div>
				<div class="field">
					<span class="label">������֤�볤�ȣ�</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="writevcodecount" value="<%=WriteVcodeCount%>" />λ (��ѡֵ��<%=web_WriteVcodeCount%>��10)</span>
				</div>
			<%
			end if

			if clng(showpage and 2)<> 0 then
			%>
				<%MailFlag=rs("mailflag")%>
				<h4>�ʼ�֪ͨ������ʹ����Ҫ�����Է�й�ܣ�</h4>
				<div class="field">
					<span class="label">�����Ե���֪ͨ������</span>
					<span class="value"><input type="checkbox" value="1" name="mailnewinform" id="mailnewinform"<%=cked(clng(web_MailFlag and MailFlag and 1)<>0)%><%=dised(web_MailNewInform=false)%> /><label for="mailnewinform"<%=dised(web_MailNewInform=false)%>>����</label></span>
				</div>
				<div class="field">
					<span class="label">�����ظ�֪ͨ�����ˣ�</span>
					<span class="value"><input type="checkbox" value="1" name="mailreplyinform" id="mailreplyinform"<%=cked(clng(web_MailFlag and MailFlag and 2)<>0)%><%=dised(web_MailReplyInform=false)%> /><label for="mailreplyinform"<%=dised(web_MailReplyInform=false)%>>����</label></span>
				</div>
				<div class="field">
					<span class="label">������֪ͨ���յ�ַ��</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="mailreceive" value="<%=rs("mailreceive")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����˵�ַ��</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="mailfrom" value="<%=rs("mailfrom")%>" /></span>
				</div>
				<div class="field">
					<span class="label">������SMTP��������ַ��</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="mailsmtpserver" value="<%=rs("mailsmtpserver")%>" /></span>
				</div>
				<div class="field">
					<span class="label">��¼�û���(����Ҫ)��</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="mailuserid" value="<%=rs("mailuserid")%>" /></span>
				</div>
				<div class="field">
					<span class="label">��¼����(����Ҫ)��</span>
					<span class="value"><input type="password" class="longtext" maxlength="48" name="mailuserpass" value="<%=rs("mailuserpass")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�ʼ������̶ȣ�</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="maillevel" value="<%=rs("maillevel")%>" /> (1��졫5����)</span>
				</div>
			<%end if

			if clng(showpage and 4)<> 0 then
			%>
				<h4>����ߴ�</h4>
				<div class="field">
					<span class="label">��������</span>
					<span class="value"><input type="text" size="10" maxlength="30" name="cssfontfamily" value="<%=rs("cssfontfamily")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����С��</span>
					<span class="value"><input type="text" size="10" maxlength="30" name="cssfontsize" value="<%=rs("cssfontsize")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����м�ࣺ</span>
					<span class="value"><input type="text" size="10" maxlength="30" name="csslineheight" value="<%=rs("csslineheight")%>" /></span>
				</div>
				<div class="field">
					<span class="label">���Ա�����ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="tablewidth" value="<%=rs("tablewidth")%>" /> (Ĭ��=630,���ðٷֱ�)</span>
				</div>
				<div class="field">
					<span class="label">���������ࣺ</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="windowspace" value="<%=rs("windowspace")%>" /> (Ĭ��=20,��λ:����)</span>
				</div>
				<div class="field">
					<span class="label">���Ա��󴰸��ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="tableleftwidth" value="<%=rs("tableleftwidth")%>" /> (Ĭ��=150,���ðٷֱ�)</span>
				</div>
				<div class="field">
					<span class="label">���������ݡ��ı��߶ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="leavecontentheight" value="<%=rs("leavecontentheight")%>" /> (Ĭ��=7,��λ:��ĸ�߶�)</span>
				</div>
				<div class="field">
					<span class="label">�������ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="searchtextwidth" value="<%=rs("searchtextwidth")%>" /> (Ĭ��=20,��λ:��ĸ���)</span>
				</div>
				<div class="field">
					<span class="label">�ظ�������༭��߶ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="replytextheight" value="<%=rs("replytextheight")%>" /> (Ĭ��=10,��λ:��ĸ�߶�)</span>
				</div>
				<div class="field">
					<span class="label">ÿҳ��ʾ����������</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="itemsperpage" value="<%=rs("itemsperpage")%>" /> (Ĭ��=5)</span>
				</div>
				<div class="field">
					<span class="label">ÿҳ��ʾ�ı�������</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="titlesperpage" value="<%=rs("titlesperpage")%>" /> (Ĭ��=20)</span>
				</div>
				<div class="field">
					<span class="label">ͷ��ÿ����ʾ����Ŀ��</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="picturesperrow" value="<%=rs("picturesperrow")%>" /> (Ĭ��=5)</span>
				</div>
				<div class="field">
					<span class="label">���������ͷ������</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="frequentfacecount" value="<%=rs("frequentfacecount")%>" /> (Ĭ��=15,����"����ͷ��"��ʾȫ��)</span>
				</div>
			<%
			end if

			if clng(showpage and 8)<> 0 then
			%>
				<h4>��������</h4>
				<%tvisualflag=rs("visualflag")%>
				<div class="field">
					<span class="label">Ĭ�ϰ���ģʽ��</span>
					<span class="value"><input type="radio" name="displaymode" value="1" id="displaymode1"<%=cked(clng(tvisualflag and 1024)<>0)%> /><label for="displaymode1">����ģʽ</label>����<input type="radio" name="displaymode" value="0" id="displaymode2"<%=cked(clng(tvisualflag and 1024)=0)%> /><label for="displaymode2">����ģʽ</label></span>
				</div>
				<div class="field">
					<span class="label">�ظ�������ʾλ�ã�</span>
					<span class="value"><input type="radio" name="replyinword" value="1" id="replyinword1"<%=cked(clng(tvisualflag and 1)<>0)%> /><label for="replyinword1">��Ƕ�ڷÿ�����</label>����<input type="radio" name="replyinword" value="0" id="replyinword2"<%=cked(clng(tvisualflag and 1)=0)%> /><label for="replyinword2">��ʾ�ڷÿ������·�</label></span>
				</div>
				<div class="field">
					<span class="label">�������ݱ����ص����ԣ�</span>
					<span class="value"><input type="radio" name="hidehidden" value="1" id="hidehidden1"<%=cked(clng(tvisualflag and 128)<>0)%> /><label for="hidehidden1" >����</label>����<input type="radio" name="hidehidden" value="0" id="hidehidden2"<%=cked(clng(tvisualflag and 128)=0)%> /><label for="hidehidden2" >��ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">���ش�������ԣ�</span>
					<span class="value"><input type="radio" name="hideaudit" value="1" id="hideaudit1"<%=cked(clng(tvisualflag and 256)<>0)%> /><label for="hideaudit1" >����</label>����<input type="radio" name="hideaudit" value="0" id="hideaudit2"<%=cked(clng(tvisualflag and 256)=0)%> /><label for="hideaudit2" >��ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">���ذ���δ�ظ����Ļ���</span>
					<span class="value"><input type="radio" name="hidewhisper" value="1" id="hidewhisper1"<%=cked(clng(tvisualflag and 512)<>0)%> /><label for="hidewhisper1" >����</label>����<input type="radio" name="hidewhisper" value="0" id="hidewhisper2"<%=cked(clng(tvisualflag and 512)=0)%> /><label for="hidewhisper2" >��ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">Logo��ʾģʽ��</span>
					<span class="value"><input type="radio" name="logobannermode" value="1" id="logobannermode1"<%=cked(clng(tvisualflag and 2048)<>0)%> /><label for="logobannermode1" >���ģʽ</label>����<input type="radio" name="logobannermode" value="0" id="logobannermode2"<%=cked(clng(tvisualflag and 2048)=0)%> /><label for="logobannermode2" >ͼƬģʽ</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ������ʾλ�ã�</span>
					<span class="value"><input type="radio" name="showpagelist" value="3" id="showpagelist3"<%=cked(clng(tvisualflag and 12)=12)%> /><label for="showpagelist3">���·�</label>��<input type="radio" name="showpagelist" value="1" id="showpagelist1"<%=cked(clng(tvisualflag and 12)=4)%> /><label for="showpagelist1">�Ϸ�</label>����<input type="radio" name="showpagelist" value="2" id="showpagelist2"<%=cked(clng(tvisualflag and 12)=8)%> /><label for="showpagelist2">�·�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�����������ʾλ�ã�</span>
					<span class="value"><input type="radio" name="showsearchbox" value="3" id="showsearchbox3"<%=cked(clng(tvisualflag and 48)=48)%> /><label for="showsearchbox3">���·�</label>��<input type="radio" name="showsearchbox" value="1" id="showsearchbox1"<%=cked(clng(tvisualflag and 48)=16)%> /><label for="showsearchbox1">�Ϸ�</label>����<input type="radio" name="showsearchbox" value="2" id="showsearchbox2"<%=cked(clng(tvisualflag and 48)=32)%> /><label for="showsearchbox2">�·�</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ�б�ģʽ��</span>
					<span class="value"><input type="radio" name="advpagelist" value="1" id="advpagelist1"<%=cked(clng(tvisualflag and 64)<>0)%> /><label for="advpagelist1">����ʽ</label>��<input type="radio" name="advpagelist" value="0" id="advpagelist2"<%=cked(clng(tvisualflag and 64)=0)%> /><label for="advpagelist2">ƽ��ʽ</label></span>
				</div>
				<div class="field">
					<span class="label">����ʽ��ҳ������</span>
					<span class="value"><input type="text" size="5" maxlength="3" name="advpagelistcount" value="<%=rs("advpagelistcount")%>" /> (Ĭ��=10)</span>
				</div>
				<div class="field">
					<span class="label">UBB��������</span>
					<span class="value"><input type="radio" name="showubbtool" value="1" id="showubbtool1"<%=cked(clng(tvisualflag and 2)<>0)%> /><label for="showubbtool1">��ʾ</label>����<input type="radio" name="showubbtool" value="0" id="showubbtool2"<%=cked(clng(tvisualflag and 2)=0)%> /><label for="showubbtool2">����</label></span>
				</div>
				<div class="field">
					<span class="label">UBB����(������UBB)��</span>
					<span class="value">
						<span class="row">
							<input type="checkbox" name="ubbflag_image" id="ubbflag_image" value="1"<%=cked(web_UbbFlag_image and UbbFlag_image)%><%=dised(web_UbbFlag_image=false)%> /><label for="ubbflag_image"<%=dised(web_UbbFlag_image=false)%>>ͼƬ</label>������
							<input type="checkbox" name="ubbflag_url" id="ubbflag_url" value="1"<%=cked(web_UbbFlag_url and UbbFlag_url)%><%=dised(web_UbbFlag_url=false)%> /><label for="ubbflag_url"<%=dised(web_UbbFlag_url=false)%>>URL��Email</label>
							<input type="checkbox" name="ubbflag_autourl" id="ubbflag_autourl" value="1"<%=cked(web_UbbFlag_autourl and UbbFlag_autourl)%><%=dised(web_UbbFlag_autourl=false)%> /><label for="ubbflag_autourl"<%=dised(web_UbbFlag_autourl=false)%>>�Զ�ʶ����ַ</label>
						</span>
						<span class="row">
							<input type="checkbox" name="ubbflag_player" id="ubbflag_player" value="1"<%=cked(web_UbbFlag_player and UbbFlag_player)%><%=dised(web_UbbFlag_player=false)%> /><label for="ubbflag_player"<%=dised(web_UbbFlag_player=false)%>>���ſؼ�</label>��
							<input type="checkbox" name="ubbflag_paragraph" id="ubbflag_paragraph" value="1"<%=cked(web_UbbFlag_paragraph and UbbFlag_paragraph)%><%=dised(web_UbbFlag_paragraph=false)%> /><label for="ubbflag_paragraph"<%=dised(web_UbbFlag_paragraph=false)%>>������ʽ</label>��
							<input type="checkbox" name="ubbflag_fontstyle" id="ubbflag_fontstyle" value="1"<%=cked(web_UbbFlag_fontstyle and UbbFlag_fontstyle)%><%=dised(web_UbbFlag_fontstyle=false)%> /><label for="ubbflag_fontstyle"<%=dised(web_UbbFlag_fontstyle=false)%>>������ʽ</label>
						</span>
						<span class="row">
							<input type="checkbox" name="ubbflag_fontcolor" id="ubbflag_fontcolor" value="1"<%=cked(web_UbbFlag_fontcolor and UbbFlag_fontcolor)%><%=dised(web_UbbFlag_fontcolor=false)%> /><label for="ubbflag_fontcolor"<%=dised(web_UbbFlag_fontcolor=false)%>>������ɫ</label>��
							<input type="checkbox" name="ubbflag_alignment" id="ubbflag_alignment" value="1"<%=cked(web_UbbFlag_alignment and UbbFlag_alignment)%><%=dised(web_UbbFlag_alignment=false)%> /><label for="ubbflag_alignment"<%=dised(web_UbbFlag_alignment=false)%>>���뷽ʽ</label>��
							<input type="checkbox" name="ubbflag_face" id="ubbflag_face" value="1"<%=cked(web_UbbFlag_face and UbbFlag_face)%><%=dised(web_UbbFlag_face=false)%> /><label for="ubbflag_face"<%=dised(web_UbbFlag_face=false)%>>����ͼ��</label>
						</span>
					</span>
				</div>
				<%talign=rs("tablealign")%>
				<div class="field">
					<span class="label">���Ա����뷽ʽ��</span>
					<span class="value"><input type="radio" name="tablealign" value="left" id="align1"<%=cked(talign<>"center" and talign<>"right")%> /><label for="align1">�����</label>��<input type="radio" name="tablealign" value="center" id="align2"<%=cked(talign="center")%> /><label for="align2">����</label>����<input type="radio" name="tablealign" value="right" id="align3"<%=cked(talign="right")%> /><label for="align3">�Ҷ���</label></span>
				</div>
				<%tpagecontrol=rs("pagecontrol")%>
				<div class="field">
					<span class="label">���Ա��߿��ߣ�</span>
					<span class="value"><input type="radio" name="showborder" value="1" id="showborder1"<%=cked(clng(tpagecontrol and 1)<>0)%> /><label for="showborder1">��ʾ</label>����<input type="radio" name="showborder" value="0" id="showborder2"<%=cked(clng(tpagecontrol and 1)=0)%> /><label for="showborder2">����</label></span>
				</div>
				<div class="field">
					<span class="label">���Ա��ܱ��⣺</span>
					<span class="value"><input type="radio" name="showtitle" value="1" id="showtitle1"<%=cked(clng(tpagecontrol and 2)<>0)%> /><label for="showtitle1">��ʾ</label>����<input type="radio" name="showtitle" value="0" id="showtitle2"<%=cked(clng(tpagecontrol and 2)=0)%> /><label for="showtitle2">����</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ�Ҽ��˵���</span>
					<span class="value"><input type="radio" name="showcontext" value="1" id="showcontext1"<%=cked(clng(tpagecontrol and 4)<>0)%> /><label for="showcontext1">����</label>����<input type="radio" name="showcontext" value="0" id="showcontext2"<%=cked(clng(tpagecontrol and 4)=0)%> /><label for="showcontext2">����</label></span>
				</div>
				<div class="field">
					<span class="label">ѡ����ҳ���ݣ�</span>
					<span class="value"><input type="radio" name="selectcontent" value="1" id="selectcontent1"<%=cked(clng(tpagecontrol and 8)<>0)%> /><label for="selectcontent1">����</label>����<input type="radio" name="selectcontent" value="0" id="selectcontent2"<%=cked(clng(tpagecontrol and 8)=0)%> /><label for="selectcontent2">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">������ҳ(Ctrl-C)��</span>
					<span class="value"><input type="radio" name="copycontent" value="1" id="copycontent1"<%=cked(clng(tpagecontrol and 16)<>0)%> /><label for="copycontent1">����</label>����<input type="radio" name="copycontent" value="0" id="copycontent2"<%=cked(clng(tpagecontrol and 16)=0)%> /><label for="copycontent2">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">��IFrame��Ƕ��</span>
					<span class="value"><input type="radio" name="beframed" value="1" id="beframed1"<%=cked(clng(tpagecontrol and 32)<>0)%> /><label for="beframed1">����</label>����<input type="radio" name="beframed" value="0" id="beframed2"<%=cked(clng(tpagecontrol and 32)=0)%> /><label for="beframed2">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">�����������ƣ�</span>
					<span class="value"><input type="text" size="10" maxlength="10" name="wordslimit" value="<%=rs("wordslimit")%>" /> (0=����,�س�����ռ2��,��UBB����ǰͳ��)</span>
				</div>
				<%tdelconfirm=rs("delconfirm")%>
				<div class="field">
					<span class="label">����ͨ�����ǰ��ʾ��</span>
					<span class="value"><input type="radio" name="passaudittip" value="1" id="passaudittip1"<%=cked(clng(tdelconfirm and 16)<>0)%> /><label for="passaudittip1">��ʾ</label>����<input type="radio" name="passaudittip" value="0" id="passaudittip2"<%=cked(clng(tdelconfirm and 16)=0)%> /><label for="passaudittip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ѡ������ͨ�������ʾ��</span>
					<span class="value"><input type="radio" name="passseltip" value="1" id="passseltip1"<%=cked(clng(tdelconfirm and 32)<>0)%> /><label for="passseltip1">��ʾ</label>����<input type="radio" name="passseltip" value="0" id="passseltip2"<%=cked(clng(tdelconfirm and 32)=0)%> /><label for="passseltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ������ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="deltip" value="1" id="deltip1"<%=cked(clng(tdelconfirm and 1)<>0)%> /><label for="deltip1">��ʾ</label>����<input type="radio" name="deltip" value="0" id="deltip2"<%=cked(clng(tdelconfirm and 1)=0)%> /><label for="deltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ���ظ�ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="delretip" value="1" id="delretip1"<%=cked(clng(tdelconfirm and 2)<>0)%> /><label for="delretip1">��ʾ</label>����<input type="radio" name="delretip" value="0" id="delretip2"<%=cked(clng(tdelconfirm and 2)=0)%> /><label for="delretip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ��ѡ������ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="delseltip" value="1" id="delseltip1"<%=cked(clng(tdelconfirm and 4)<>0)%> /><label for="delseltip1">��ʾ</label>����<input type="radio" name="delseltip" value="0" id="delseltip2"<%=cked(clng(tdelconfirm and 4)=0)%> /><label for="delseltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ִ�и߼�ɾ��ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="deladvtip" value="1" id="deladvtip1"<%=cked(clng(tdelconfirm and 8)<>0)%> /><label for="deladvtip1">��ʾ</label>����<input type="radio" name="deladvtip" value="0" id="deladvtip2"<%=cked(clng(tdelconfirm and 8)=0)%> /><label for="deladvtip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">�������Ļ�ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="pubwhispertip" value="1" id="pubwhispertip1"<%=cked(clng(tdelconfirm and 64)<>0)%> /><label for="pubwhispertip1">��ʾ</label>����<input type="radio" name="pubwhispertip" value="0" id="pubwhispertip2"<%=cked(clng(tdelconfirm and 64)=0)%> /><label for="pubwhispertip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">�ö�����ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="lock2toptip" value="1" id="lock2toptip1"<%=cked(clng(tdelconfirm and 256)<>0)%> /><label for="lock2toptip1">��ʾ</label>����<input type="radio" name="lock2toptip" value="0" id="lock2toptip2"<%=cked(clng(tdelconfirm and 256)=0)%> /><label for="lock2toptip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">��ǰ����ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="bring2toptip" value="1" id="bring2toptip1"<%=cked(clng(tdelconfirm and 128)<>0)%> /><label for="bring2toptip1">��ʾ</label>����<input type="radio" name="bring2toptip" value="0" id="bring2toptip2"<%=cked(clng(tdelconfirm and 128)=0)%> /><label for="bring2toptip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">��������˳��ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="reordertip" value="1" id="reordertip1"<%=cked(clng(tdelconfirm and 512)<>0)%> /><label for="reordertip1">��ʾ</label>����<input type="radio" name="reordertip" value="0" id="reordertip2"<%=cked(clng(tdelconfirm and 512)=0)%> /><label for="reordertip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">���Ա���ɫ������</span>
					<span class="value">
						<select name="style">
						<%
						styleid=rs("styleid")
						rs.Close
						rs.Open sql_adminconfig_style,cn,,,1

						dim onestyleid,onestylename
						while rs.EOF=false
							onestyleid=rs("styleid")
							onestylename=rs("stylename")
							Response.Write "<option value=" &chr(34)& onestyleid &chr(34)
							if onestyleid=styleid or onestyleid="" then Response.Write " selected=""selected"""
							Response.Write ">" &onestylename& "</option>"
							rs.MoveNext
						wend
						%>
						</select>
					</span>
				</div>
			<%end if%>

			<input type="hidden" name="page" value="<%=showpage%>" />
			<input type="hidden" name="user" value="<%=ruser%>" />
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>
</div>

<script type="text/javascript">
function check()
{
	var tv,showpage=<%=showpage%>;
	if ((showpage & 1) != 0)
	{
		if (isNaN(tv=Number(document.configform.admintimeout.value)))
			{alert('������Ա��¼��ʱ������Ϊ���֡�');document.configform.admintimeout.select();return false;}
		else if (tv<1 || tv>1440)
			{alert('������Ա��¼��ʱ��������1��1440�ķ�Χ�ڡ�');document.configform.admintimeout.select();return false;}

		if (isNaN(tv=Number(document.configform.showipv4.value)))
			{alert('��Ϊ�ÿ���ʾIPv4������Ϊ���֡�');document.configform.showipv4.select();return false;}
		else if (tv<0 || tv>4 || document.configform.showipv4.value==='')
			{alert('��Ϊ�ÿ���ʾIPv4��������0��4�ķ�Χ�ڡ�');document.configform.showipv4.select();return false;}

		if (isNaN(tv=Number(document.configform.showipv6.value)))
			{alert('��Ϊ�ÿ���ʾIPv6������Ϊ���֡�');document.configform.showipv6.select();return false;}
		else if (tv<0 || tv>8 || document.configform.showipv6.value==='')
			{alert('��Ϊ�ÿ���ʾIPv6��������0��8�ķ�Χ�ڡ�');document.configform.showipv6.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshowipv4.value)))
			{alert('��Ϊ����Ա��ʾIPv4������Ϊ���֡�');document.configform.adminshowipv4.select();return false;}
		else if (tv<0 || tv>4 || document.configform.adminshowipv4.value==='')
			{alert('��Ϊ����Ա��ʾIPv4��������0��4�ķ�Χ�ڡ�');document.configform.adminshowipv4.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshowipv6.value)))
			{alert('��Ϊ����Ա��ʾIPv6������Ϊ���֡�');document.configform.adminshowipv6.select();return false;}
		else if (tv<0 || tv>8 || document.configform.adminshowipv6.value==='')
			{alert('��Ϊ����Ա��ʾIPv6��������0��8�ķ�Χ�ڡ�');document.configform.adminshowipv6.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshoworiginalipv4.value)))
			{alert('��Ϊ����Ա��ʾԭIPv4������Ϊ����');document.configform.adminshoworiginalipv4.select();return false;}
		else if (tv<0 || tv>4 || document.configform.adminshoworiginalipv4.value==='')
			{alert('��Ϊ����Ա��ʾԭIPv4��������0��4�ķ�Χ�ڡ�');document.configform.adminshoworiginalipv4.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshoworiginalipv6.value)))
			{alert('��Ϊ����Ա��ʾԭIPv6������Ϊ����');document.configform.adminshoworiginalipv6.select();return false;}
		else if (tv<0 || tv>8 || document.configform.adminshoworiginalipv6.value==='')
			{alert('��Ϊ����Ա��ʾԭIPv6��������0��8�ķ�Χ�ڡ�');document.configform.adminshoworiginalipv6.select();return false;}

		if (isNaN(tv=Number(document.configform.vcodecount.value)))
			{alert('����¼��֤�볤�ȡ�����Ϊ���֡�');document.configform.vcodecount.select();return false;}
		else if (tv<<%=web_VcodeCount%> || tv>10)
			{alert('����¼��֤�볤�ȡ�������<%=web_VcodeCount%>��10�ķ�Χ�ڡ�');document.configform.vcodecount.select();return false;}

		if (isNaN(tv=Number(document.configform.writevcodecount.value)))
			{alert('��������֤�볤�ȡ�����Ϊ���֡�');document.configform.writevcodecount.select();return false;}
		else if (tv<<%=web_WriteVcodeCount%> || tv>10)
			{alert('��������֤�볤�ȡ�������<%=web_WriteVcodeCount%>��10�ķ�Χ�ڡ�');document.configform.writevcodecount.select();return false;}
	}
	if ((showpage & 2) != 0)
	{
		if (isNaN(tv=Number(document.configform.maillevel.value)))
			{alert('���ʼ������̶ȡ�����Ϊ���֡�');document.configform.maillevel.select();return false;}
		else if (tv<1 || tv>5)
			{alert('���ʼ������̶ȡ�������1��5�ķ�Χ�ڡ�');document.configform.maillevel.select();return false;}	
	}
	if ((showpage & 4) != 0)
	{
		if (isNaN(tv=Number(document.configform.tablewidth.value)))
			{if (/^\d+%$/.test(document.configform.tablewidth.value)==false) {alert('�����Ա�����ȡ�����Ϊ���������ٷֱȡ�');document.configform.tablewidth.select();return false;}}
		else if (tv<1)
			{alert('�����Ա�����ȡ���������㡣');document.configform.tablewidth.select();return false;}

		if (isNaN(tv=Number(document.configform.tableleftwidth.value)))
			{if (/^\d+%$/.test(document.configform.tableleftwidth.value)==false) {alert('�����Ա��󴰸��ȡ�����Ϊ���������ٷֱȡ�');document.configform.tableleftwidth.select();return false;}}
		else if (tv<1)
			{alert('�����Ա��󴰸��ȡ���������㡣');document.configform.tableleftwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.windowspace.value)))
			{alert('�����������ࡱ����Ϊ���֡�');document.configform.windowspace.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�����������ࡱ������1��255�ķ�Χ�ڡ�');document.configform.windowspace.select();return false;}

		if (isNaN(tv=Number(document.configform.leavecontentheight.value)))
			{alert('�����������ݡ��ı��߶ȡ�����Ϊ���֡�');document.configform.leavecontentheight.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�����������ݡ��ı��߶ȡ�������1��255�ķ�Χ�ڡ�');document.configform.leavecontentheight.select();return false;}

		if (isNaN(tv=Number(document.configform.searchtextwidth.value)))
			{alert('���������ȡ�����Ϊ���֡�');document.configform.searchtextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('���������ȡ�������1��255�ķ�Χ�ڡ�');document.configform.searchtextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.replytextheight.value)))
			{alert('���ظ�������༭��߶ȡ�����Ϊ���֡�');document.configform.replytextheight.select();return false;}
		else if (tv<1 || tv>255)
			{alert('���ظ�������༭��߶ȡ�������1��255�ķ�Χ�ڡ�');document.configform.replytextheight.select();return false;}

		if (isNaN(tv=Number(document.configform.itemsperpage.value)))
			{alert('��ÿҳ��ʾ��������������Ϊ���֡�');document.configform.itemsperpage.select();return false;}
		else if (tv<1 || tv>32767)
			{alert('��ÿҳ��ʾ����������������1��32767�ķ�Χ�ڡ�');document.configform.itemsperpage.select();return false;}

		if (isNaN(tv=Number(document.configform.titlesperpage.value)))
			{alert('��ÿҳ��ʾ�ı�����������Ϊ���֡�');document.configform.titlesperpage.select();return false;}
		else if (tv<1 || tv>32767)
			{alert('��ÿҳ��ʾ�ı�������������1��32767�ķ�Χ�ڡ�');document.configform.titlesperpage.select();return false;}

		if (isNaN(tv=Number(document.configform.picturesperrow.value)))
			{alert('��ͷ��ÿ����ʾ����Ŀ������Ϊ���֡�');document.configform.picturesperrow.select();return false;}
		else if (tv<1 || tv>255)
			{alert('��ͷ��ÿ����ʾ����Ŀ��������1��255�ķ�Χ�ڡ�');document.configform.picturesperrow.select();return false;}

		if (isNaN(tv=Number(document.configform.frequentfacecount.value)))
			{alert('�����������ͷ����������Ϊ���֡�');document.configform.frequentfacecount.select();return false;}
		else if (tv<0 || tv>255)
			{alert('�����������ͷ������������0��255�ķ�Χ�ڡ�');document.configform.frequentfacecount.select();return false;}
	}
	if ((showpage & 8) != 0)
	{
		if (isNaN(tv=Number(document.configform.advpagelistcount.value)))
			{alert('������ʽ��ҳ����������Ϊ���֡�');document.configform.advpagelistcount.select();return false;}
		else if (tv<1 || tv>255 || document.configform.wordslimit.value=='')
			{alert('������ʽ��ҳ������������1��255�ķ�Χ�ڡ�');document.configform.advpagelistcount.select();return false;}

		if (isNaN(tv=Number(document.configform.wordslimit.value)))
			{alert('�������������ơ�����Ϊ���֡�');document.configform.wordslimit.select();return false;}
		else if (tv<0 || tv>2147483647 || document.configform.wordslimit.value=='')
			{alert('�������������ơ�������0��2147483647�ķ�Χ�ڡ�');document.configform.wordslimit.select();return false;}
	}
	document.configform.submit1.disabled=true;
	return true;
}
</script>

<!-- #include file="include/footer.inc" -->
</body>
</html>