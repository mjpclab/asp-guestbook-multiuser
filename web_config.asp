<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
%>

<!-- #include file="inc_dtd.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn">
<head>
	<title><%=web_BookName%> Webmaster�������� ���Ա�����</title>
	
	<!-- #include file="style.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admintool.inc" -->

	<%rs.Open Replace(sql_adminconfig_config,"{0}",wm_name),cn,,,1%>

	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">���Ա���������</td>
		</tr>
		<tr>
		<td style="width:100%; text-align:left; vertical-align:top; color:<%=TableContentColor%>; background-color:<%=TableContentBGC%>;">
			<table border="1" bordercolor="<%=TableBorderColor%>" style="width:100%; text-align:center; border-collapse:collapse;">
			<tr style="cursor:hand;">
				<td style="width:20%;" onclick="window.location='web_config.asp?page=1'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="web_config.asp?page=1" onmouseover="return true;">[��������]</a></td>
				<td style="width:20%;" onclick="window.location='web_config.asp?page=2'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="web_config.asp?page=2" onmouseover="return true;">[�ʼ�֪ͨ]</a></td>
				<td style="width:20%;" onclick="window.location='web_config.asp?page=4'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="web_config.asp?page=4" onmouseover="return true;">[����ߴ�]</a></td>
				<td style="width:20%;" onclick="window.location='web_config.asp?page=8'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="web_config.asp?page=8" onmouseover="return true;">[��������]</a></td>
				<td style="width:20%;" onclick="window.location='web_config.asp'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="web_config.asp" onmouseover="return true;">[ȫ������]</a></td>
			</tr>
			</table>

		<br/>
		<form method="post" action="web_saveconfig.asp" name="configform" onsubmit="return check();">
		<%
		showpage=15
		if isnumeric(Request.QueryString("page")) and len(Request.QueryString("page"))<=2 and Request.QueryString("page")<>"" then showpage=clng(Request.QueryString("page"))
		
		if clng(showpage and 1)<> 0 then
		%>
		
			<span style="font-weight:bold">��������<br/>����ĳ���ܹر�ʱ���û�������Ч��IP��ʾ�������û����ã���֤�볤����Ϊ�û�������Сֵ����</span><br/><br/>
			<%tstatus=rs("status")%>
			���Ա��������ܣ�������<input type="radio" name="status4" value="8" id="status41"<%=cked(clng(tstatus and 8)<>0)%> /><label for="status41">����</label>����<input type="radio" name="status4" value="0" id="status42"<%=cked(clng(tstatus and 8)=0)%> /><label for="status42">�ر�</label><br/><br/>
			�һ����빦�ܣ���������<input type="radio" name="status5" value="16" id="status51"<%=cked(clng(tstatus and 16)<>0)%> /><label for="status51">����</label>����<input type="radio" name="status5" value="0" id="status52"<%=cked(clng(tstatus and 16)=0)%> /><label for="status52">�ر�</label><br/><br/>
			�û���¼ά�����ܣ�����<input type="radio" name="status6" value="32" id="status61"<%=cked(clng(tstatus and 32)<>0)%> /><label for="status61">����</label>����<input type="radio" name="status6" value="0" id="status62"<%=cked(clng(tstatus and 32)=0)%> /><label for="status62">�ر�</label><br/><br/>
			�û���ɾ�ʺŹ��ܣ�����<input type="radio" name="status7" value="64" id="status71"<%=cked(clng(tstatus and 64)<>0)%> /><label for="status71">����</label>����<input type="radio" name="status7" value="0" id="status72"<%=cked(clng(tstatus and 64)=0)%> /><label for="status72">�ر�</label><br/><br/><br/>

			���Ա�״̬������������<input type="radio" name="status1" value="1" id="status11"<%=cked(clng(tstatus and 1)<>0)%> /><label for="status11">����</label>����<input type="radio" name="status1" value="0" id="status12"<%=cked(clng(tstatus and 1)=0)%> /><label for="status12">�ر�</label><br/><br/>
			�ÿ�����Ȩ�ޣ���������<input type="radio" name="status2" value="2" id="status21"<%=cked(clng(tstatus and 2)<>0)%> /><label for="status21">����</label>����<input type="radio" name="status2" value="0" id="status22"<%=cked(clng(tstatus and 2)=0)%> /><label for="status22">�ر�</label><br/><br/>
			�ÿ���������Ȩ�ޣ�����<input type="radio" name="status3" value="4" id="status31"<%=cked(clng(tstatus and 4)<>0)%> /><label for="status31">����</label>����<input type="radio" name="status3" value="0" id="status32"<%=cked(clng(tstatus and 4)=0)%> /><label for="status32">�ر�</label><br/><br/><br/>

			<%adminlimit=rs("adminhtml")%>
			�������İ�ȫ�����ã���<input type="checkbox" value="1" name="adminviewcode" id="adminviewcode"<%=cked(clng(adminlimit and 8)<>0)%> /><label for="adminviewcode">�û����ÿ�������ʾʵ��HTML��UBB����</label><br/><br/>
			�û�����ԱHTMLȨ�ޣ���<input type="checkbox" value="1" name="adminhtml" id="adminhtml"<%=cked(clng(adminlimit and 1)<>0)%> /><label for="adminhtml">����HTMLȨ�ޣ����Ƽ���</label><br/>
			����������������������<input type="checkbox" value="1" name="adminubb" id="adminubb"<%=cked(clng(adminlimit and 2)<>0)%> /><label for="adminubb">����UBBȨ��</label><br/>
			����������������������<input type="checkbox" value="1" name="adminertn" id="adminertn"<%=cked(clng(adminlimit and 4)<>0)%> /><label for="adminertn">��֧��HTML��UBBʱ������س�����</label><br/><br/>

			<%guestlimit=rs("guesthtml")%>
			�ÿ�HTMLȨ�ޣ���������<input type="checkbox" value="1" name="guesthtml" id="guesthtml"<%=cked(clng(guestlimit and 1)<>0)%> /><label for="guesthtml">����HTMLȨ�ޣ����Ƽ���</label><br/>
			����������������������<input type="checkbox" value="1" name="guestubb" id="guestubb"<%=cked(clng(guestlimit and 2)<>0)%> /><label for="guestubb">����UBBȨ��</label><br/>
			����������������������<input type="checkbox" value="1" name="guestertn" id="guestertn"<%=cked(clng(guestlimit and 4)<>0)%> /><label for="guestertn">��֧��HTML��UBBʱ������س�����</label><br/><br/><br/>
		
			����Ա��¼��ʱ��������<input type="text" size="4" maxlength="4" name="admintimeout" value="<%=rs("admintimeout")%>" />�� (Ĭ��=20)<br/><br/>
			Ϊ�ÿ���ʾIPǰ��������<input type="text" size="4" maxlength="1" name="showip" value="<%=rs("showip")%>" />λ (��ѡֵ��0��4)<br/><br/>
			Ϊ����Ա��ʾIPǰ������<input type="text" size="4" maxlength="1" name="adminshowip" value="<%=rs("adminshowip")%>" />λ (��ѡֵ��0��4)<br/><br/>
			Ϊ����Ա��ʾԭIPǰ����<input type="text" size="4" maxlength="1" name="adminshoworiginalip" value="<%=rs("adminshoworiginalip")%>" />λ (��ѡֵ��0��4,ʹ�ô��������ʱ������ʾԭʼIP)<br/><br/>
			
			��¼��֤�볤�ȣ�������<input type="text" size="4" maxlength="2" name="vcodecount" value="<%=clng(rs("vcodecount") and &H0F)%>" />λ (��ѡֵ��0��10)<br/><br/>
			������֤�볤�ȣ�������<input type="text" size="4" maxlength="2" name="writevcodecount" value="<%=clng(rs("vcodecount") and &HF0) \ &H10%>" />λ (��ѡֵ��0��10)<br/><br/><br/>
		<%
		end if
		
		if clng(showpage and 2)<> 0 then
		%>
			<%MailFlag=rs("mailflag")%>
			<span style="font-weight:bold">�ʼ�֪ͨ�������Ƿ���û����ţ���</span><br/><br/>
			�����Ե���֪ͨ��������<input type="checkbox" value="1" name="mailnewinform" id="mailnewinform"<%=cked(clng(MailFlag and 1)<>0)%> /><label for="mailnewinform">����</label><br/><br/>
			�����ظ�֪ͨѡ�����<input type="checkbox" value="1" name="mailreplyinform" id="mailreplyinform"<%=cked(clng(MailFlag and 2)<>0)%> /><label for="mailreplyinform">����</label><br/><br/>
			�ʼ������������������<input type="radio" value="0" name="mailcomponent" id="mailcomponent0"<%=cked(clng(MailFlag and 4)=0)%> /><label for="mailcomponent0">JMail</label>��<input type="radio" value="1" name="mailcomponent" id="mailcomponent1"<%=cked(clng(MailFlag and 4)<>0)%> /><label for="mailcomponent1">CDO</label><br/><br/><br/>
		<%end if
		
		if clng(showpage and 4)<> 0 then
		%>
		
			<span style="font-weight:bold">����ߴ磨����ֻӰ�챾����ҳ�棩��</span><br/><br/>
		
		
			����������������������<input type="text" size="10" maxlength="30" name="cssfontfamily" value="<%=rs("cssfontfamily")%>" /><br/><br/>
			�����С��������������<input type="text" size="10" maxlength="30" name="cssfontsize" value="<%=rs("cssfontsize")%>" /><br/><br/>
			�����м�ࣺ����������<input type="text" size="10" maxlength="30" name="csslineheight" value="<%=rs("csslineheight")%>" /><br/><br/>			
			���Ա��ܿ�ȣ���������<input type="text" size="10" maxlength="5" name="tablewidth" value="<%=rs("tablewidth")%>" /> (Ĭ��=630,���ðٷֱ�)<br/><br/>
			���������ࣺ��������<input type="text" size="10" maxlength="3" name="windowspace" value="<%=rs("windowspace")%>" /> (Ĭ��=20,��λ:����)<br/><br/>
			���Ա��󴰸��ȣ�����<input type="text" size="10" maxlength="5" name="tableleftwidth" value="<%=rs("tableleftwidth")%>" /> (Ĭ��=150,���ðٷֱ�)<br/><br/>
			UBB ��������ȣ�������<input type="text" size="10" maxlength="4" name="ubbtoolwidth" value="<%=rs("ubbtoolwidth")%>" /> (Ĭ��=320,��λ:����)<br/><br/>
			UBB �������߶ȣ�������<input type="text" size="10" maxlength="4" name="ubbtoolheight" value="<%=rs("ubbtoolheight")%>" /> (Ĭ��=100,��λ:����)<br/><br/>
			�������ȣ�����������<input type="text" size="10" maxlength="3" name="searchtextwidth" value="<%=rs("searchtextwidth")%>" /> (Ĭ��=20,��λ:��ĸ���)<br/><br/>
			���߼�ɾ�������ı���<input type="text" size="10" maxlength="3" name="advdeltextwidth" value="<%=rs("advdeltextwidth")%>" /> (Ĭ��=20,��λ:��ĸ���)<br/><br/>
			���޸����롱���ı���<input type="text" size="10" maxlength="3" name="setinfotextwidth" value="<%=rs("setinfotextwidth")%>" /> (Ĭ��=40,��λ:��ĸ���)<br/><br/>
			�����ݹ��ˡ����ı���<input type="text" size="10" maxlength="3" name="filtertextwidth" value="<%=rs("filtertextwidth")%>" /> (Ĭ��=62,��λ:��ĸ���)<br/><br/>
			����༭���ȣ�������<input type="text" size="10" maxlength="3" name="replytextwidth" value="<%=rs("replytextwidth")%>" /> (Ĭ��=98,��λ:��ĸ���)<br/><br/>
			����༭��߶ȣ�������<input type="text" size="10" maxlength="3" name="replytextheight" value="<%=rs("replytextheight")%>" /> (Ĭ��=10,��λ:��ĸ�߶�)<br/><br/>

			ÿҳ��ʾ����Ŀ��������<input type="text" size="10" maxlength="5" name="itemsperpage" value="<%=rs("itemsperpage")%>" /> (Ĭ��=5)<br/><br/>
			ÿҳ��ʾ�ı�����������<input type="text" size="10" maxlength="5" name="titlesperpage" value="<%=rs("titlesperpage")%>" /> (Ĭ��=20)<br/><br/><br/>
		
		<%
		end if
		
		if clng(showpage and 8)<> 0 then
		%>	
		
			<span style="font-weight:bold">�������ã�</span><br/><br/>
			<%tvisualflag=rs("visualflag")%>
			Ĭ�ϰ���ģʽ����������<input type="radio" name="displaymode" value="1" id="displaymode1"<%=cked(clng(tvisualflag and 1024)<>0)%> /><label for="displaymode1">��̳�б���ʽ</label>����<input type="radio" name="displaymode" value="0" id="displaymode2"<%=cked(clng(tvisualflag and 1024)=0)%> /><label for="displaymode2">���Ա���ʽ</label><br/><br/>
			�ظ�������ʾλ�ã�����<input type="radio" name="replyinword" value="1" id="replyinword1"<%=cked(clng(tvisualflag and 1)<>0)%> /><label for="replyinword1">��Ƕ�ڷÿ�����</label>����<input type="radio" name="replyinword" value="0" id="replyinword2"<%=cked(clng(tvisualflag and 1)=0)%> /><label for="replyinword2">��ʾ�ڷÿ������·�</label><br/><br/>
			��ҳ������ʾλ�ã�����<input type="radio" name="showpagelist" value="3" id="showpagelist3"<%=cked(clng(tvisualflag and 12)=12)%> /><label for="showpagelist3">���·�</label>��<input type="radio" name="showpagelist" value="1" id="showpagelist1"<%=cked(clng(tvisualflag and 12)=4)%> /><label for="showpagelist1">�Ϸ�</label>����<input type="radio" name="showpagelist" value="2" id="showpagelist2"<%=cked(clng(tvisualflag and 12)=8)%> /><label for="showpagelist2">�·�</label><br/><br/>
			��ҳ�б�ģʽ����������<input type="radio" name="advpagelist" value="1" id="advpagelist1"<%=cked(clng(tvisualflag and 64)<>0)%> /><label for="advpagelist1">����ʽ</label>��<input type="radio" name="advpagelist" value="0" id="advpagelist2"<%=cked(clng(tvisualflag and 64)=0)%> /><label for="advpagelist2">ƽ��ʽ</label><br/><br/>
			
			����ʽ��ҳ������������<input type="text" size="5" maxlength="3" name="advpagelistcount" value="<%=rs("advpagelistcount")%>" /> (Ĭ��=10)<br/><br/>

			UBB����(������UBB)����<input type="checkbox" name="ubbflag_image" id="ubbflag_image" value="1"<%=cked(UbbFlag_image)%> /><label for="ubbflag_image">ͼƬ</label>������
								<input type="checkbox" name="ubbflag_url" id="ubbflag_url" value="1"<%=cked(UbbFlag_url)%> /><label for="ubbflag_url">URL��Email</label>
								<input type="checkbox" name="ubbflag_autourl" id="ubbflag_autourl" value="1"<%=cked(UbbFlag_autourl)%> /><label for="ubbflag_autourl">�Զ�ʶ����ַ</label><br/>

			����������������������<input type="checkbox" name="ubbflag_player" id="ubbflag_player" value="1"<%=cked(UbbFlag_player)%> /><label for="ubbflag_player">���ſؼ�</label>��
								<input type="checkbox" name="ubbflag_paragraph" id="ubbflag_paragraph" value="1"<%=cked(UbbFlag_paragraph)%> /><label for="ubbflag_paragraph">������ʽ</label>��
								<input type="checkbox" name="ubbflag_fontstyle" id="ubbflag_fontstyle" value="1"<%=cked(UbbFlag_fontstyle)%> /><label for="ubbflag_fontstyle">������ʽ</label><br/>

			����������������������<input type="checkbox" name="ubbflag_fontcolor" id="ubbflag_fontcolor" value="1"<%=cked(UbbFlag_fontcolor)%> /><label for="ubbflag_fontcolor">������ɫ</label>��
								<input type="checkbox" name="ubbflag_alignment" id="ubbflag_alignment" value="1"<%=cked(UbbFlag_alignment)%> /><label for="ubbflag_alignment">���뷽ʽ</label>��
								<input type="checkbox" name="ubbflag_movement" id="ubbflag_movement" value="1"<%=cked(UbbFlag_movement)%> /><label for="ubbflag_movement">�ƶ�Ч��(<span title="�����ƶ�Ч�����޷�ͨ��W3C XHTML 1.0 Traditional ��֤">����ʹ��</span>)</label><br/>

			����������������������<input type="checkbox" name="ubbflag_cssfilter" id="ubbflag_cssfilter" value="1"<%=cked(UbbFlag_cssfilter)%> /><label for="ubbflag_cssfilter">�˾�Ч��</label>��
								<input type="checkbox" name="ubbflag_face" id="ubbflag_face" value="1"<%=cked(UbbFlag_face)%> /><label for="ubbflag_face">����ͼ��</label><br/><br/>

			<%talign=rs("tablealign")%>
			���Ա����뷽ʽ��������<input type="radio" name="tablealign" value="left" id="align1"<%=cked(talign<>"center" and talign<>"right")%> /><label for="align1">�����</label>����<input type="radio" name="tablealign" value="center" id="align2"<%=cked(talign="center")%> /><label for="align2">����</label>����<input type="radio" name="tablealign" value="right" id="align3"<%=cked(talign="right")%> /><label for="align3">�Ҷ���</label><br/><br/>
			
			<%tdelconfirm=rs("delconfirm")%>
			ɾ������ʱ��ʾ��������<input type="radio" name="deltip" value="1" id="deltip1"<%=cked(clng(tdelconfirm and 1)<>0)%> /><label for="deltip1">��ʾ</label>����<input type="radio" name="deltip" value="0" id="deltip2"<%=cked(clng(tdelconfirm and 1)=0)%> /><label for="deltip2">����ʾ</label><br/><br/>
			ɾ���ظ�ʱ��ʾ��������<input type="radio" name="delretip" value="1" id="delretip1"<%=cked(clng(tdelconfirm and 2)<>0)%> /><label for="delretip1">��ʾ</label>����<input type="radio" name="delretip" value="0" id="delretip2"<%=cked(clng(tdelconfirm and 2)=0)%> /><label for="delretip2">����ʾ</label><br/><br/>
			ɾ��ѡ������ʱ��ʾ����<input type="radio" name="delseltip" value="1" id="delseltip1"<%=cked(clng(tdelconfirm and 4)<>0)%> /><label for="delseltip1">��ʾ</label>����<input type="radio" name="delseltip" value="0" id="delseltip2"<%=cked(clng(tdelconfirm and 4)=0)%> /><label for="delseltip2">����ʾ</label><br/><br/>
			ִ�и߼�ɾ��ʱ��ʾ����<input type="radio" name="deladvtip" value="1" id="deladvtip1"<%=cked(clng(tdelconfirm and 8)<>0)%> /><label for="deladvtip1">��ʾ</label>����<input type="radio" name="deladvtip" value="0" id="deladvtip2"<%=cked(clng(tdelconfirm and 8)=0)%> /><label for="deladvtip2">����ʾ</label><br/><br/>
			ɾ���ö�����ʱ��ʾ����<input type="radio" name="deldectip" value="1" id="deldectip1"<%=cked(clng(tdelconfirm and 16)<>0)%> /><label for="deldectip1">��ʾ</label>����<input type="radio" name="deldectip" value="0" id="deldectip2"<%=cked(clng(tdelconfirm and 16)=0)%> /><label for="deldectip2">����ʾ</label><br/><br/>
			ɾ��ѡ������ʱ��ʾ����<input type="radio" name="delseldectip" value="1" id="delseldectip1"<%=cked(clng(tdelconfirm and 32)<>0)%> /><label for="delseldectip1">��ʾ</label>����<input type="radio" name="delseldectip" value="0" id="delseldectip2"<%=cked(clng(tdelconfirm and 32)=0)%> /><label for="delseldectip2">����ʾ</label><br/><br/>

			���Ա���ɫ������������<select name="style">
			<%
			stylename=rs("stylename")
			rs.Close
			rs.Open sql_adminconfig_style,cn,,,1

			dim onestyle
			while rs.EOF=false
				onestyle=rs("stylename")
				Response.Write "<option value=" &chr(34)& onestyle &chr(34)
				if onestyle=stylename or stylename="" then Response.Write " selected=""selected"""
				Response.Write ">" &onestyle& "</option>"
				rs.MoveNext
			wend
			%>
			</select><br/><br/><br/>
		<%end if%>	
		<input type="hidden" name="page" value="<%=showpage%>" />
		<p style="text-align:center;"><input value="��������" type="submit" name="submit1" /></p>
		</form>
		<br/>
		</td>
		</tr>
	</table>
</div>

<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

<script type="text/javascript" defer="defer">
//<![CDATA[

function check()
{
	var tv,showpage=<%=showpage%>;
	if ((showpage & 1) != 0)
	{
		if (isNaN(tv=Number(document.configform.admintimeout.value)))
			{alert('������Ա��¼��ʱ������Ϊ���֡�');document.configform.admintimeout.select();return false;}
		else if (tv<1 || tv>1440)
			{alert('������Ա��¼��ʱ��������1��1440�ķ�Χ�ڡ�');document.configform.admintimeout.select();return false;}

		if (isNaN(tv=Number(document.configform.showip.value)))
			{alert('��Ϊ�ÿ���ʾIP������Ϊ���֡�');document.configform.showip.select();return false;}
		else if (tv<0 || tv>4 || document.configform.showip.value=='')
			{alert('��Ϊ�ÿ���ʾIP��������0��4�ķ�Χ�ڡ�');document.configform.showip.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshowip.value)))
			{alert('��Ϊ����Ա��ʾIP������Ϊ���֡�');document.configform.adminshowip.select();return false;}
		else if (tv<0 || tv>4 || document.configform.adminshowip.value=='')
			{alert('��Ϊ����Ա��ʾIP��������0��4�ķ�Χ�ڡ�');document.configform.adminshowip.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshoworiginalip.value)))
			{alert('��Ϊ����Ա��ʾԭIP������Ϊ����');document.configform.adminshoworiginalip.select();return false;}
		else if (tv<0 || tv>4 || document.configform.adminshoworiginalip.value=='')
			{alert('��Ϊ����Ա��ʾԭIP��������0��4�ķ�Χ�ڡ�');document.configform.adminshoworiginalip.select();return false;}

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
			{if (/^\d+%$/.test(document.configform.tablewidth.value)==false) {alert('�����Ա��ܿ�ȡ�����Ϊ���������ٷֱȡ�');document.configform.tablewidth.select();return false;}}
		else if (tv<1)
			{alert('�����Ա��ܿ�ȡ���������㡣');document.configform.tablewidth.select();return false;}

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

		if (isNaN(tv=Number(document.configform.ubbtoolwidth.value)))
			{alert('����д���ԡ�UBB���߿�����Ϊ���֡�');document.configform.ubbtoolwidth.select();return false;}
		else if (tv<1 || tv>9999)
			{alert('����д���ԡ�UBB���߿�������1��9999�ķ�Χ�ڡ�');document.configform.ubbtoolwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.ubbtoolheight.value)))
			{alert('����д���ԡ�UBB���߸ߡ�����Ϊ���֡�');document.configform.ubbtoolheight.select();return false;}
		else if (tv<1 || tv>9999)
			{alert('����д���ԡ�UBB���߸ߡ�������1��9999�ķ�Χ�ڡ�');document.configform.ubbtoolheight.select();return false;}

		if (isNaN(tv=Number(document.configform.advdeltextwidth.value)))
			{alert('�����߼�ɾ�������ı�������Ϊ���֡�');document.configform.advdeltextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�����߼�ɾ�������ı���������1��255�ķ�Χ�ڡ�');document.configform.advdeltextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.setinfotextwidth.value)))
			{alert('�����޸����ϡ����ı�������Ϊ���֡�');document.configform.setinfotextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�����޸����ϡ����ı���������1��255�ķ�Χ�ڡ�');document.configform.setinfotextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.filtertextwidth.value)))
			{alert('�������ݹ��ˡ����ı�������Ϊ���֡�');document.configform.filtertextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�������ݹ��ˡ����ı���������1��255�ķ�Χ�ڡ�');document.configform.filtertextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.replytextwidth.value)))
			{alert('���ظ�������༭���ȡ�����Ϊ���֡�');document.configform.replytextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('���ظ�������༭���ȡ�������1��255�ķ�Χ�ڡ�');document.configform.replytextwidth.select();return false;}

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

//]]>
</script>

<!-- #include file="bottom.asp" -->
</body>
</html>