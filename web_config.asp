<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_config.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

Call CreateConn(cn)
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

	<!-- #include file="include/template/web_admin_header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->
	<%rs.Open Replace(sql_adminconfig_config,"{0}",wm_id),cn,,,1%>

	<div class="region region-config admin-tools">
		<h3 class="title">���Ա���������</h3>
		<div class="content flex-box">
			<ul>
				<li><a href="web_config.asp?page=1">��������</a></li>
				<li><a href="web_config.asp?page=2">�ʼ�֪ͨ</a></li>
				<li><a href="web_config.asp?page=4">������ߴ�</a></li>
				<li><a href="web_config.asp?page=8">��������</a></li>
				<li><a href="web_config.asp">ȫ������</a></li>
			</ul>

			<form method="post" action="web_saveconfig.asp" name="configform" onsubmit="return check();">
			<%
			showpage=15
			if isnumeric(Request.QueryString("page")) and len(Request.QueryString("page"))<=2 and Request.QueryString("page")<>"" then showpage=clng(Request.QueryString("page"))

			if CBool(showpage AND 1) then
			%>
				<h4>�������ã���ĳ���ܹر�ʱ���û�������Ч��IP��ʾ�������û����ã���֤�볤����Ϊ�û�������Сֵ����</h4>
				<%tstatus=rs("status")%>
				<div class="field">
					<span class="label">���Ա��������ܣ�</span>
					<span class="value"><input type="radio" name="status4" value="1" id="status41"<%=cked(CBool(tstatus AND 8))%> /><label for="status41">����</label>����<input type="radio" name="status4" value="0" id="status42"<%=cked(Not CBool(tstatus AND 8))%> /><label for="status42">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�һ����빦�ܣ�</span>
					<span class="value"><input type="radio" name="status5" value="1" id="status51"<%=cked(CBool(tstatus AND 16))%> /><label for="status51">����</label>����<input type="radio" name="status5" value="0" id="status52"<%=cked(Not CBool(tstatus AND 16))%> /><label for="status52">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�û���¼ά�����ܣ�</span>
					<span class="value"><input type="radio" name="status6" value="1" id="status61"<%=cked(CBool(tstatus AND 32))%> /><label for="status61">����</label>����<input type="radio" name="status6" value="0" id="status62"<%=cked(Not CBool(tstatus AND 32))%> /><label for="status62">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�û���ɾ�ʺŹ��ܣ�</span>
					<span class="value"><input type="radio" name="status7" value="1" id="status71"<%=cked(CBool(tstatus AND 64))%> /><label for="status71">����</label>����<input type="radio" name="status7" value="0" id="status72"<%=cked(Not CBool(tstatus AND 64))%> /><label for="status72">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">���Ա�״̬��</span>
					<span class="value"><input type="radio" name="status1" value="1" id="status11"<%=cked(CBool(tstatus AND 1))%> /><label for="status11">����</label>����<input type="radio" name="status1" value="0" id="status12"<%=cked(Not CBool(tstatus AND 1))%> /><label for="status12">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�����Ȩ�ޣ�</span>
					<span class="value"><input type="radio" name="status2" value="1" id="status21"<%=cked(CBool(tstatus AND 2))%> /><label for="status21">����</label>����<input type="radio" name="status2" value="0" id="status22"<%=cked(Not CBool(tstatus AND 2))%> /><label for="status22">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ���������Ȩ�ޣ�</span>
					<span class="value"><input type="radio" name="status3" value="1" id="status31"<%=cked(CBool(tstatus AND 4))%> /><label for="status31">����</label>����<input type="radio" name="status3" value="0" id="status32"<%=cked(Not CBool(tstatus AND 4))%> /><label for="status32">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">����ͳ�ƣ�</span>
					<span class="value"><input type="radio" name="status9" value="1" id="status91"<%=cked(CBool(tstatus AND 256))%> /><label for="status91">����</label>����<input type="radio" name="status9" value="0" id="status92"<%=cked(Not CBool(tstatus AND 256))%> /><label for="status92">�ر�</label></span>
				</div>
				<%adminlimit=rs("adminhtml")%>
				<div class="field">
					<span class="label">�������İ�ȫ�����ã�</span>
					<span class="value"><input type="checkbox" value="1" name="adminviewcode" id="adminviewcode"<%=cked(CBool(adminlimit AND 8))%> /><label for="adminviewcode">�û����ÿ�������ʾʵ��HTML��UBB����</label></span>
				</div>
				<div class="field">
					<span class="label">�û�����ԱHTMLȨ�ޣ�</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="adminhtml" id="adminhtml"<%=cked(CBool(adminlimit AND 1))%> /><label for="adminhtml">����HTMLȨ�ޣ����Ƽ���</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminubb" id="adminubb"<%=cked(CBool(adminlimit AND 2))%> /><label for="adminubb">����UBBȨ��</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminertn" id="adminertn"<%=cked(CBool(adminlimit AND 4))%> /><label for="adminertn">��֧��HTML��UBBʱ������س�����</label></span>
					</span>
				</div>
				<%guestlimit=rs("guesthtml")%>
				<div class="field">
					<span class="label">�ÿ�HTMLȨ�ޣ�</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="guesthtml" id="guesthtml"<%=cked(CBool(guestlimit AND 1))%> /><label for="guesthtml">����HTMLȨ�ޣ����Ƽ���</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestubb" id="guestubb"<%=cked(CBool(guestlimit AND 2))%> /><label for="guestubb">����UBBȨ��</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestertn" id="guestertn"<%=cked(CBool(guestlimit AND 4))%> /><label for="guestertn">��֧��HTML��UBBʱ������س�����</label></span>
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
					<span class="value"><input type="text" size="4" maxlength="2" name="vcodecount" value="<%=clng(rs("vcodecount") and &H0F)%>" />λ (��ѡֵ��0��10)</span>
				</div>
				<div class="field">
					<span class="label">������֤�볤�ȣ�</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="writevcodecount" value="<%=clng(rs("vcodecount") and &HF0) \ &H10%>" />λ (��ѡֵ��0��10)</span>
				</div>
			<%
			end if

			if CBool(showpage AND 2) then
			%>
				<%MailFlag=rs("mailflag")%>
				<h4>�ʼ�֪ͨ�������Ƿ���û����ţ���</h4>
				<div class="field">
					<span class="label">�����Ե���֪ͨ������</span>
					<span class="value"><input type="checkbox" value="1" name="mailnewinform" id="mailnewinform"<%=cked(CBool(MailFlag AND 1))%> /><label for="mailnewinform">����</label></span>
				</div>
				<div class="field">
					<span class="label">�����ظ�֪ͨ�����ˣ�</span>
					<span class="value"><input type="checkbox" value="1" name="mailreplyinform" id="mailreplyinform"<%=cked(CBool(MailFlag AND 2))%> /><label for="mailreplyinform">����</label></span>
				</div>
				<div class="field">
					<span class="label">�ʼ����������</span>
					<span class="value"><input type="radio" value="0" name="mailcomponent" id="mailcomponent0"<%=cked(Not CBool(MailFlag AND 4))%> /><label for="mailcomponent0">JMail</label>��<input type="radio" value="1" name="mailcomponent" id="mailcomponent1"<%=cked(CBool(MailFlag AND 4))%> /><label for="mailcomponent1">CDO</label></span>
				</div>
			<%end if

			if CBool(showpage AND 4) then
			%>
				<h4>������ߴ磨����ֻӰ�챾����ҳ�棩��</h4>
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
					<span class="label">�������ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="searchtextwidth" value="<%=rs("searchtextwidth")%>" /> (Ĭ��=20,��λ:��ĸ���)</span>
				</div>
				<div class="field">
					<span class="label">����༭��߶ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="replytextheight" value="<%=rs("replytextheight")%>" /> (Ĭ��=10,��λ:��ĸ�߶�)</span>
				</div>
				<div class="field">
					<span class="label">ÿҳ��ʾ����Ŀ����</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="itemsperpage" value="<%=rs("itemsperpage")%>" /> (Ĭ��=5)</span>
				</div>
				<div class="field">
					<span class="label">ÿҳ��ʾ�ı�������</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="titlesperpage" value="<%=rs("titlesperpage")%>" /> (Ĭ��=20)</span>
				</div>
				<div class="field">
					<span class="label">���Ա���ɫ������</span>
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
			<%
			end if

			if CBool(showpage AND 8) then
			%>
				<h4>�������ã�</h4>
				<%tvisualflag=rs("visualflag")%>
				<div class="field">
					<span class="label">Ĭ�ϰ���ģʽ��</span>
					<span class="value"><input type="radio" name="displaymode" value="1" id="displaymode1"<%=cked(CBool(tvisualflag AND 1024))%> /><label for="displaymode1">����ģʽ</label>����<input type="radio" name="displaymode" value="0" id="displaymode2"<%=cked(Not CBool(tvisualflag AND 1024))%> /><label for="displaymode2">����ģʽ</label></span>
				</div>
				<div class="field">
					<span class="label">�ظ�������ʾλ�ã�</span>
					<span class="value"><input type="radio" name="replyinword" value="1" id="replyinword1"<%=cked(CBool(tvisualflag AND 1))%> /><label for="replyinword1">��Ƕ�ڷÿ�����</label>����<input type="radio" name="replyinword" value="0" id="replyinword2"<%=cked(Not CBool(tvisualflag AND 1))%> /><label for="replyinword2">��ʾ�ڷÿ������·�</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ������ʾλ�ã�</span>
					<span class="value"><input type="radio" name="showpagelist" value="3" id="showpagelist3"<%=cked(clng(tvisualflag and 12)=12)%> /><label for="showpagelist3">���·�</label>��<input type="radio" name="showpagelist" value="1" id="showpagelist1"<%=cked(clng(tvisualflag and 12)=4)%> /><label for="showpagelist1">�Ϸ�</label>����<input type="radio" name="showpagelist" value="2" id="showpagelist2"<%=cked(clng(tvisualflag and 12)=8)%> /><label for="showpagelist2">�·�</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ�б�ģʽ��</span>
					<span class="value"><input type="radio" name="advpagelist" value="1" id="advpagelist1"<%=cked(CBool(tvisualflag AND 64))%> /><label for="advpagelist1">����ʽ</label>��<input type="radio" name="advpagelist" value="0" id="advpagelist2"<%=cked(Not CBool(tvisualflag AND 64))%> /><label for="advpagelist2">ƽ��ʽ</label></span>
				</div>
				<div class="field">
					<span class="label">����ʽ��ҳ������</span>
					<span class="value"><input type="text" size="5" maxlength="3" name="advpagelistcount" value="<%=rs("advpagelistcount")%>" /> (Ĭ��=10)</span>
				</div>
				<div class="field">
					<span class="label">UBB����(������UBB)��</span>
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
					</span>
				</div>
				<%tpagecontrol=rs("pagecontrol")%>
				<%tdelconfirm=rs("delconfirm")%>
				<div class="field">
					<span class="label">ɾ������ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="deltip" value="1" id="deltip1"<%=cked(CBool(tdelconfirm AND 1))%> /><label for="deltip1">��ʾ</label>����<input type="radio" name="deltip" value="0" id="deltip2"<%=cked(Not CBool(tdelconfirm AND 1))%> /><label for="deltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ���ظ�ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="delretip" value="1" id="delretip1"<%=cked(CBool(tdelconfirm AND 2))%> /><label for="delretip1">��ʾ</label>����<input type="radio" name="delretip" value="0" id="delretip2"<%=cked(Not CBool(tdelconfirm AND 2))%> /><label for="delretip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ��ѡ������ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="delseltip" value="1" id="delseltip1"<%=cked(CBool(tdelconfirm AND 4))%> /><label for="delseltip1">��ʾ</label>����<input type="radio" name="delseltip" value="0" id="delseltip2"<%=cked(Not CBool(tdelconfirm AND 4))%> /><label for="delseltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ִ�и߼�ɾ��ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="deladvtip" value="1" id="deladvtip1"<%=cked(CBool(tdelconfirm AND 8))%> /><label for="deladvtip1">��ʾ</label>����<input type="radio" name="deladvtip" value="0" id="deladvtip2"<%=cked(Not CBool(tdelconfirm AND 8))%> /><label for="deladvtip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ���ö�����ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="deldectip" value="1" id="deldectip1"<%=cked(CBool(tdelconfirm AND 16))%> /><label for="deldectip1">��ʾ</label>����<input type="radio" name="deldectip" value="0" id="deldectip2"<%=cked(Not CBool(tdelconfirm AND 16))%> /><label for="deldectip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ��ѡ���ö�����ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="delseldectip" value="1" id="delseldectip1"<%=cked(CBool(tdelconfirm AND 32))%> /><label for="delseldectip1">��ʾ</label>����<input type="radio" name="delseldectip" value="0" id="delseldectip2"<%=cked(Not CBool(tdelconfirm AND 32))%> /><label for="delseldectip2">����ʾ</label></span>
				</div>
			<%end if%>

			<input type="hidden" name="page" value="<%=showpage%>" />
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>

<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

<script type="text/javascript" defer="defer">
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
		else if (tv<0 || tv>10)
			{alert('����¼��֤�볤�ȡ�������0��10�ķ�Χ�ڡ�');document.configform.vcodecount.select();return false;}

		if (isNaN(tv=Number(document.configform.writevcodecount.value)))
			{alert('��������֤�볤�ȡ�����Ϊ���֡�');document.configform.writevcodecount.select();return false;}
		else if (tv<0 || tv>10)
			{alert('��������֤�볤�ȡ�������0��10�ķ�Χ�ڡ�');document.configform.writevcodecount.select();return false;}
	}
	if ((showpage & 4) != 0)
	{
		if (isNaN(tv=Number(document.configform.tablewidth.value)))
			{if (/^\d+%$/.test(document.configform.tablewidth.value)==false) {alert('�����Ա�����ȡ�����Ϊ���������ٷֱȡ�');document.configform.tablewidth.select();return false;}}
		else if (tv<1)
			{alert('�����Ա�����ȡ���������㡣');document.configform.tablewidth.select();return false;}

		if (isNaN(tv=Number(document.configform.windowspace.value)))
			{alert('�����������ࡱ����Ϊ���֡�');document.configform.windowspace.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�����������ࡱ������1��255�ķ�Χ�ڡ�');document.configform.windowspace.select();return false;}

		if (isNaN(tv=Number(document.configform.tableleftwidth.value)))
			{if (/^\d+%$/.test(document.configform.tableleftwidth.value)==false) {alert('�����Ա��󴰸��ȡ�����Ϊ���������ٷֱȡ�');document.configform.tableleftwidth.select();return false;}}
		else if (tv<1)
			{alert('�����Ա��󴰸��ȡ���������㡣');document.configform.tableleftwidth.select();return false;}

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
	}
	if ((showpage & 8) != 0)
	{
		if (isNaN(tv=Number(document.configform.advpagelistcount.value)))
			{alert('������ʽ��ҳ����������Ϊ���֡�');document.configform.advpagelistcount.select();return false;}
		else if (tv<1 || tv>255 || document.configform.advpagelistcount.value=='')
			{alert('������ʽ��ҳ������������1��255�ķ�Χ�ڡ�');document.configform.advpagelistcount.select();return false;}
	}
	document.configform.submit1.disabled=true;
	return true;
}
</script>
</body>
</html>