<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_config.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster�������� ���Ա�����</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<%Call WebInitHeaderData("","Webmaster��������","","")%><!-- #include file="include/template/header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->
	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	Call CreateConn(cn)
	rs.Open Replace(sql_adminconfig_config,"{0}",wm_id),cn,,,1
	tstatus=rs("status")
	adminlimit=rs("adminhtml")
	guestlimit=rs("guesthtml")
    MailFlag=rs("mailflag")
    tvisualflag=rs("visualflag")
    'tpagecontrol=rs("pagecontrol")
    tdelconfirm=rs("delconfirm")
	%>

	<div class="region region-config admin-tools">
		<h3 class="title">���Ա���������</h3>
		<div class="content">
			<form method="post" action="web_saveconfig.asp" name="configform" onsubmit="return check();">

			<div id="tabContainer"></div>

			<div id="tab-switch">
				<h4>����</h4>
				<div class="field">
					<span class="label">���Ա���������</span>
					<span class="value"><input type="radio" name="status4" value="1" id="status41"<%=cked(CBool(tstatus AND 8))%> /><label for="status41">����</label>����<input type="radio" name="status4" value="0" id="status42"<%=cked(Not CBool(tstatus AND 8))%> /><label for="status42">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�һ����빦��</span>
					<span class="value"><input type="radio" name="status5" value="1" id="status51"<%=cked(CBool(tstatus AND 16))%> /><label for="status51">����</label>����<input type="radio" name="status5" value="0" id="status52"<%=cked(Not CBool(tstatus AND 16))%> /><label for="status52">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�û���¼ά������</span>
					<span class="value"><input type="radio" name="status6" value="1" id="status61"<%=cked(CBool(tstatus AND 32))%> /><label for="status61">����</label>����<input type="radio" name="status6" value="0" id="status62"<%=cked(Not CBool(tstatus AND 32))%> /><label for="status62">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�û���ɾ�ʺŹ���</span>
					<span class="value"><input type="radio" name="status7" value="1" id="status71"<%=cked(CBool(tstatus AND 64))%> /><label for="status71">����</label>����<input type="radio" name="status7" value="0" id="status72"<%=cked(Not CBool(tstatus AND 64))%> /><label for="status72">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">���Ա�״̬</span>
					<span class="value"><input type="radio" name="status1" value="1" id="status11"<%=cked(CBool(tstatus AND 1))%> /><label for="status11">����</label>����<input type="radio" name="status1" value="0" id="status12"<%=cked(Not CBool(tstatus AND 1))%> /><label for="status12">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�����Ȩ��</span>
					<span class="value"><input type="radio" name="status2" value="1" id="status21"<%=cked(CBool(tstatus AND 2))%> /><label for="status21">����</label>����<input type="radio" name="status2" value="0" id="status22"<%=cked(Not CBool(tstatus AND 2))%> /><label for="status22">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ���������Ȩ��</span>
					<span class="value"><input type="radio" name="status3" value="1" id="status31"<%=cked(CBool(tstatus AND 4))%> /><label for="status31">����</label>����<input type="radio" name="status3" value="0" id="status32"<%=cked(Not CBool(tstatus AND 4))%> /><label for="status32">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">����ͳ��</span>
					<span class="value"><input type="radio" name="status9" value="1" id="status91"<%=cked(CBool(tstatus AND 256))%> /><label for="status91">����</label>����<input type="radio" name="status9" value="0" id="status92"<%=cked(Not CBool(tstatus AND 256))%> /><label for="status92">�ر�</label></span>
				</div>
			</div>

			<div id="tab-code">
				<h4>����</h4>
				<div class="field">
					<span class="label">�������İ�ȫ������</span>
					<span class="value"><input type="checkbox" value="1" name="adminviewcode" id="adminviewcode"<%=cked(CBool(adminlimit AND 8))%> /><label for="adminviewcode">�û����ÿ�������ʾʵ��HTML��UBB����</label></span>
				</div>
				<div class="field">
					<span class="label">�û�����ԱHTMLȨ��</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="adminhtml" id="adminhtml"<%=cked(CBool(adminlimit AND 1))%> /><label for="adminhtml">����HTMLȨ�ޣ����Ƽ���</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminubb" id="adminubb"<%=cked(CBool(adminlimit AND 2))%> /><label for="adminubb">����UBBȨ��</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminertn" id="adminertn"<%=cked(CBool(adminlimit AND 4))%> /><label for="adminertn">��֧��HTML��UBBʱ������س�����</label></span>
					</span>
				</div>
				<div class="field">
					<span class="label">�ÿ�HTMLȨ��</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="guesthtml" id="guesthtml"<%=cked(CBool(guestlimit AND 1))%> /><label for="guesthtml">����HTMLȨ�ޣ����Ƽ���</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestubb" id="guestubb"<%=cked(CBool(guestlimit AND 2))%> /><label for="guestubb">����UBBȨ��</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestertn" id="guestertn"<%=cked(CBool(guestlimit AND 4))%> /><label for="guestertn">��֧��HTML��UBBʱ������س�����</label></span>
					</span>
				</div>
				<div class="field">
					<span class="label">UBB����(������UBB)</span>
					<span class="value">
						<span class="row">
							<input type="checkbox" name="ubbflag_image" id="ubbflag_image" value="1"<%=cked(UbbFlag_image)%> /><label for="ubbflag_image">ͼƬ</label>
							<input type="checkbox" name="ubbflag_url" id="ubbflag_url" value="1"<%=cked(UbbFlag_url)%> /><label for="ubbflag_url">URL��Email</label>
							<input type="checkbox" name="ubbflag_autourl" id="ubbflag_autourl" value="1"<%=cked(UbbFlag_autourl)%> /><label for="ubbflag_autourl">�Զ�ʶ����ַ</label>
						</span>
						<span class="row">
							<input type="checkbox" name="ubbflag_player" id="ubbflag_player" value="1"<%=cked(UbbFlag_player)%> /><label for="ubbflag_player">���ſؼ�</label>
							<input type="checkbox" name="ubbflag_paragraph" id="ubbflag_paragraph" value="1"<%=cked(UbbFlag_paragraph)%> /><label for="ubbflag_paragraph">������ʽ</label>
							<input type="checkbox" name="ubbflag_fontstyle" id="ubbflag_fontstyle" value="1"<%=cked(UbbFlag_fontstyle)%> /><label for="ubbflag_fontstyle">������ʽ</label>
						</span>
						<span class="row">
							<input type="checkbox" name="ubbflag_fontcolor" id="ubbflag_fontcolor" value="1"<%=cked(UbbFlag_fontcolor)%> /><label for="ubbflag_fontcolor">������ɫ</label>
							<input type="checkbox" name="ubbflag_alignment" id="ubbflag_alignment" value="1"<%=cked(UbbFlag_alignment)%> /><label for="ubbflag_alignment">���뷽ʽ</label>
							<input type="checkbox" name="ubbflag_face" id="ubbflag_face" value="1"<%=cked(UbbFlag_face)%> /><label for="ubbflag_face">����ͼ��</label>
						</span>
						<span class="row">
							<input type="checkbox" name="ubbflag_markdown_paragraph" id="ubbflag_markdown_paragraph" value="1"<%=cked(UbbFlag_markdown_paragraph)%> /><label for="ubbflag_markdown_paragraph">Markdown������ʽ</label>
							<input type="checkbox" name="ubbflag_markdown_fontstyle" id="ubbflag_markdown_fontstyle" value="1"<%=cked(UbbFlag_markdown_fontstyle)%> /><label for="ubbflag_markdown_fontstyle">Markdown������ʽ</label>
						</span>
					</span>
				</div>
			</div>

			<div id="tab-security">
				<h4>��ȫ</h4>
				<div class="field">
					<span class="label">����Ա��¼��ʱ</span>
					<span class="value"><input type="text" size="4" maxlength="4" name="admintimeout" value="<%=rs("admintimeout")%>" />�� (Ĭ��=20)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ�ÿ���ʾIPv4ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="showipv4" value="<%=ShowIPv4%>" />�ֽ� (��ѡֵ��0��4)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ�ÿ���ʾIPv6ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="showipv6" value="<%=ShowIPv6%>" />�� (��ѡֵ��0��8)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾIPv4ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshowipv4" value="<%=AdminShowIPv4%>" />�ֽ� (��ѡֵ��0��4)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾIPv6ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshowipv6" value="<%=AdminShowIPv6%>" />�� (��ѡֵ��0��8)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾԭIPv4ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshoworiginalipv4" value="<%=AdminShowOriginalIPv4%>" />�ֽ� (��ѡֵ��0��4,ʹ�ô��������ʱ������ʾԭʼIP)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾԭIPv6ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshoworiginalipv6" value="<%=AdminShowOriginalIPv6%>" />�� (��ѡֵ��0��8,ʹ�ô��������ʱ������ʾԭʼIP)</span>
				</div>
				<div class="field">
					<span class="label">��¼��֤�볤��</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="vcodecount" value="<%=rs("vcodecount") AND &H0F%>" />λ (��ѡֵ��0��10)</span>
				</div>
				<div class="field">
					<span class="label">������֤�볤��</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="writevcodecount" value="<%=(rs("vcodecount") and &HF0) \ &H10%>" />λ (��ѡֵ��0��10)</span>
				</div>
			</div>

			<div id="tab-paging">
				<h4>��ҳ</h4>
				<div class="field">
					<span class="label">Ĭ�ϰ���ģʽ</span>
					<span class="value"><input type="radio" name="displaymode" value="0" id="displaymode2"<%=cked(Not CBool(tvisualflag AND 1024))%> /><label for="displaymode2">����ģʽ</label>����<input type="radio" name="displaymode" value="1" id="displaymode1"<%=cked(CBool(tvisualflag AND 1024))%> /><label for="displaymode1">����ģʽ</label></span>
				</div>
				<div class="field">
					<span class="label">ÿҳ��ʾ��������</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="itemsperpage" value="<%=rs("itemsperpage")%>" /> (Ĭ��=5)</span>
				</div>
				<div class="field">
					<span class="label">ÿҳ��ʾ�ı�����</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="titlesperpage" value="<%=rs("titlesperpage")%>" /> (Ĭ��=20)</span>
				</div>
				<div class="field">
					<span class="label">��ҳ������ʾλ��</span>
					<span class="value"><input type="radio" name="showpagelist" value="3" id="showpagelist3"<%=cked((tvisualflag and 12)=12)%> /><label for="showpagelist3">���·�</label>��<input type="radio" name="showpagelist" value="1" id="showpagelist1"<%=cked((tvisualflag and 12)=4)%> /><label for="showpagelist1">�Ϸ�</label>����<input type="radio" name="showpagelist" value="2" id="showpagelist2"<%=cked((tvisualflag and 12)=8)%> /><label for="showpagelist2">�·�</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ�б�ģʽ</span>
					<span class="value"><input type="radio" name="advpagelist" value="1" id="advpagelist1"<%=cked(CBool(tvisualflag AND 64))%> /><label for="advpagelist1">����ʽ</label>��<input type="radio" name="advpagelist" value="0" id="advpagelist2"<%=cked(Not CBool(tvisualflag AND 64))%> /><label for="advpagelist2">ƽ��ʽ</label></span>
				</div>
				<div class="field">
					<span class="label">����ʽ��ҳ����</span>
					<span class="value"><input type="text" size="5" maxlength="3" name="advpagelistcount" value="<%=rs("advpagelistcount")%>" /> (Ĭ��=10)</span>
				</div>
			</div>

			<div id="tab-ui">
				<h4>����</h4>
				<div class="field">
					<span class="label">�����б�(","�ָ�)</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="cssfontfamily" value="<%=rs("cssfontfamily")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����С</span>
					<span class="value"><input type="text" size="10" maxlength="8" name="cssfontsize" value="<%=rs("cssfontsize")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����м��</span>
					<span class="value"><input type="text" size="10" maxlength="8" name="csslineheight" value="<%=rs("csslineheight")%>" /></span>
				</div>
				<div class="field">
					<span class="label">���Ա������</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="tablewidth" value="<%=rs("tablewidth")%>" /> (Ĭ��=630,���ðٷֱ�)</span>
				</div>
				<div class="field">
					<span class="label">����������</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="windowspace" value="<%=rs("windowspace")%>" /> (Ĭ��=20,��λ:����)</span>
				</div>
				<div class="field">
					<span class="label">���Ա��󴰸���</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="tableleftwidth" value="<%=rs("tableleftwidth")%>" /> (Ĭ��=150,���ðٷֱ�)</span>
				</div>
				<div class="field">
					<span class="label">��������</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="searchtextwidth" value="<%=rs("searchtextwidth")%>" /> (Ĭ��=20,��λ:��ĸ���)</span>
				</div>
				<div class="field">
					<span class="label">����༭��߶�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="replytextheight" value="<%=rs("replytextheight")%>" /> (Ĭ��=10,��λ:��ĸ�߶�)</span>
				</div>
				<div class="field">
					<span class="label">���Ա���ɫ����</span>
					<span class="value">
						<select name="style">
						<%
						styleid=rs("styleid")

						set rs1=server.CreateObject("ADODB.Recordset")
						rs1.Open sql_adminconfig_style,cn,,,1

						dim onestyleid,onestylename
						while rs1.EOF=false
							onestyleid=rs1("styleid")
							onestylename=rs1("stylename")
							Response.Write "<option value=""" & onestyleid & """"
							Response.Write seled(onestyleid=styleid or onestyleid="")
							Response.Write ">" &onestylename& "</option>"
							rs1.MoveNext
						wend
						rs1.Close : set rs1=nothing
						%>
						</select>
					</span>
				</div>
			</div>

			<div id="tab-behavior">
				<h4>��Ϊ</h4>
				<div class="field">
					<span class="label">������ʱ��ƫ��</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="servertimezoneoffset" value="<%=rs("servertimezoneoffset")%>" /> (Ĭ��=0,��λ:����)</span>
				</div>
				<div class="field">
					<span class="label">��ʾʱ��ƫ��</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="displaytimezoneoffset" value="<%=rs("displaytimezoneoffset")%>" /> (Ĭ��=0,��λ:����)</span>
				</div>
				<div class="field">
					<span class="label">�ظ�������ʾλ��</span>
					<span class="value"><input type="radio" name="replyinword" value="1" id="replyinword1"<%=cked(CBool(tvisualflag AND 1))%> /><label for="replyinword1">��Ƕ�ڷÿ�����</label>����<input type="radio" name="replyinword" value="0" id="replyinword2"<%=cked(Not CBool(tvisualflag AND 1))%> /><label for="replyinword2">��ʾ�ڷÿ������·�</label></span>
				</div>
			</div>

			<div id="tab-mail">
				<h4>�ʼ�</h4>
				<div class="field">
					<span class="label">�����Ե���֪ͨ����</span>
					<span class="value"><input type="checkbox" value="1" name="mailnewinform" id="mailnewinform"<%=cked(CBool(MailFlag AND 1))%> /><label for="mailnewinform">����</label></span>
				</div>
				<div class="field">
					<span class="label">�����ظ�֪ͨ������</span>
					<span class="value"><input type="checkbox" value="1" name="mailreplyinform" id="mailreplyinform"<%=cked(CBool(MailFlag AND 2))%> /><label for="mailreplyinform">����</label></span>
				</div>
				<div class="field">
					<span class="label">�ʼ��������</span>
					<span class="value"><input type="radio" value="0" name="mailcomponent" id="mailcomponent0"<%=cked(Not CBool(MailFlag AND 4))%> /><label for="mailcomponent0">JMail</label>��<input type="radio" value="1" name="mailcomponent" id="mailcomponent1"<%=cked(CBool(MailFlag AND 4))%> /><label for="mailcomponent1">CDO</label></span>
				</div>
			</div>

			<div id="tab-confirm">
				<h4>ȷ��</h4>
				<div class="field">
					<span class="label">ͨ���������</span>
					<span class="value"><input type="radio" name="passaudittip" value="1" id="passaudittip1"<%=cked(CBool(tdelconfirm AND 16))%> /><label for="passaudittip1">��ʾȷ��</label>����<input type="radio" name="passaudittip" value="0" id="passaudittip2"<%=cked(Not CBool(tdelconfirm AND 16))%> /><label for="passaudittip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ͨ�����ѡ������</span>
					<span class="value"><input type="radio" name="passseltip" value="1" id="passseltip1"<%=cked(CBool(tdelconfirm AND 32))%> /><label for="passseltip1">��ʾȷ��</label>����<input type="radio" name="passseltip" value="0" id="passseltip2"<%=cked(Not CBool(tdelconfirm AND 32))%> /><label for="passseltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ������</span>
					<span class="value"><input type="radio" name="deltip" value="1" id="deltip1"<%=cked(CBool(tdelconfirm AND 1))%> /><label for="deltip1">��ʾȷ��</label>����<input type="radio" name="deltip" value="0" id="deltip2"<%=cked(Not CBool(tdelconfirm AND 1))%> /><label for="deltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ���ظ�</span>
					<span class="value"><input type="radio" name="delretip" value="1" id="delretip1"<%=cked(CBool(tdelconfirm AND 2))%> /><label for="delretip1">��ʾȷ��</label>����<input type="radio" name="delretip" value="0" id="delretip2"<%=cked(Not CBool(tdelconfirm AND 2))%> /><label for="delretip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ��ѡ������</span>
					<span class="value"><input type="radio" name="delseltip" value="1" id="delseltip1"<%=cked(CBool(tdelconfirm AND 4))%> /><label for="delseltip1">��ʾȷ��</label>����<input type="radio" name="delseltip" value="0" id="delseltip2"<%=cked(Not CBool(tdelconfirm AND 4))%> /><label for="delseltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ִ�и߼�ɾ��</span>
					<span class="value"><input type="radio" name="deladvtip" value="1" id="deladvtip1"<%=cked(CBool(tdelconfirm AND 8))%> /><label for="deladvtip1">��ʾȷ��</label>����<input type="radio" name="deladvtip" value="0" id="deladvtip2"<%=cked(Not CBool(tdelconfirm AND 8))%> /><label for="deladvtip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ���ö�����</span>
					<span class="value"><input type="radio" name="deldectip" value="1" id="deldectip1"<%=cked(CBool(tdelconfirm AND 64))%> /><label for="deldectip1">��ʾȷ��</label>����<input type="radio" name="deldectip" value="0" id="deldectip2"<%=cked(Not CBool(tdelconfirm AND 64))%> /><label for="deldectip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ��ѡ���ö�����</span>
					<span class="value"><input type="radio" name="delseldectip" value="1" id="delseldectip1"<%=cked(CBool(tdelconfirm AND 128))%> /><label for="delseldectip1">��ʾȷ��</label>����<input type="radio" name="delseldectip" value="0" id="delseldectip2"<%=cked(Not CBool(tdelconfirm AND 128))%> /><label for="delseldectip2">����ʾ</label></span>
				</div>
			</div>

			<input type="hidden" name="tabIndex" id="tabIndex" value="<%=Request.QueryString("tabIndex")%>" />
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>

<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

<script type="text/javascript" src="asset/js/tabcontrol.js"></script>
<script type="text/javascript">
var tab=new TabControl('tabContainer');

tab.addPage('tab-switch','����');
tab.addPage('tab-code','����');
tab.addPage('tab-security','��ȫ');
tab.addPage('tab-paging','��ҳ');
tab.addPage('tab-ui','����');
tab.addPage('tab-behavior','��Ϊ');
tab.addPage('tab-mail','�ʼ�');
tab.addPage('tab-confirm','ȷ��');

tab.restoreFromField('tabIndex');

function check()
{
	function checkRange(tabIndex, name, textbox, min, max) {
		var value;
		if(isNaN(value = parseInt(textbox.value, 10))) {
			tab.selectPage(tabIndex);
			textbox.select();
			alert(name + ' ����Ϊ���֡�');
			return false;
		} else if(value<min || value>max) {
			tab.selectPage(tabIndex);
			textbox.select();
			alert(name + ' ������' + min + '��' + max + '�ķ�Χ�ڡ�');
			return false;
		}

		return true;
	}
	function checkCssSize(tabIndex, name, textbox) {
		var value;
		if(isNaN(value = parseInt(textbox.value, 10))) {
			if (!/^\s*\d+%\s*$/.test(value)) {
				tab.selectPage(tabIndex);
				textbox.select();
				alert(name + ' ����Ϊ���������ٷֱȡ�');
				return false;
			}
		} else if (value <= 0) {
			tab.selectPage(tabIndex);
			textbox.select();
			alert(name + ' ��������㡣');
			return false;
		}

		return true;
	}

	var frm=document.configform;
	return checkRange(2, "����Ա��¼��ʱ", frm.admintimeout, 1, 1440) &&
		checkRange(2, "Ϊ�ÿ���ʾIPv4", frm.showipv4, 0, 4) &&
		checkRange(2, "Ϊ�ÿ���ʾIPv6", frm.showipv6, 0, 8) &&
		checkRange(2, "Ϊ����Ա��ʾIPv4", frm.adminshowipv4, 0, 4) &&
		checkRange(2, "Ϊ����Ա��ʾIPv6", frm.adminshowipv6, 0, 8) &&
		checkRange(2, "Ϊ����Ա��ʾԭIPv4", frm.adminshoworiginalipv4, 0, 4) &&
		checkRange(2, "Ϊ����Ա��ʾԭIPv6", frm.adminshoworiginalipv6, 0, 8) &&
		checkRange(2, "��¼��֤�볤��", frm.vcodecount, 0, 10) &&
		checkRange(2, "������֤�볤��", frm.writevcodecount, 0, 10) &&
		checkRange(3, "ÿҳ��ʾ��������", frm.itemsperpage, 1, 32767) &&
		checkRange(3, "ÿҳ��ʾ�ı�����", frm.titlesperpage, 1, 32767) &&
		checkRange(3, "����ʽ��ҳ����", frm.advpagelistcount, 1, 255) &&
		checkCssSize(4, "���Ա������", frm.tablewidth) &&
		checkRange(4, "����������", frm.windowspace, 1, 255) &&
		checkCssSize(4, "���Ա��󴰸���", frm.tableleftwidth) &&
		checkRange(4, "��������", frm.searchtextwidth, 1, 255) &&
		checkRange(4, "�ظ�������༭��߶�", frm.replytextheight, 1, 255) &&
		checkRange(5, "������ʱ��ƫ��", frm.servertimezoneoffset, -1440, 1440) &&
		checkRange(5, "��ʾʱ��ƫ��", frm.displaytimezoneoffset, -1440, 1440) &&
		(document.configform.submit1.disabled=true, true);
}
</script>
</body>
</html>