<!-- #include file="common.asp" -->
<%
dim lcn,lrs
set lcn=server.CreateObject("ADODB.Connection")
set lrs=server.CreateObject("ADODB.Recordset")

CreateConn lcn,dbtype
checkuser lcn,lrs,false

'================================
'================================
'================================
'Load Web Master Global Settings
'================================
'================================
'================================
lrs.Open Replace(sql_loadconfig_config,"{0}",wm_id),lcn,0,1,1

status=lrs("status")
if clng(status and 1)<>0 then	'���Ա�����
	StatusOpen=true
else
	StatusOpen=false
end if
if clng(status and 2)<>0 then	'����Ȩ�޿���
	StatusWrite=true
else
	StatusWrite=false
end if
if clng(status and 4)<>0 then	'����Ȩ�޿���
	StatusSearch=true
else
	StatusSearch=false
end if
if clng(status and 32)<>0 then	'�û���¼Ȩ�޿���
	StatusLogin=true
else
	StatusLogin=false
end if
if clng(status and 256)<>0 then	'����ͳ��
	StatusStatistics=true
else
	StatusStatistics=false
end if

web_StatusOpen=StatusOpen
web_StatusWrite=StatusWrite
web_StatusSearch=StatusSearch
web_StatusStatistics=StatusStatistics

web_IPConStatus=lrs("ipconstatus")	'IP���β��ԣ���4λ����IPv4����4λ����IPv6
web_IPv4ConStatus=web_IPConStatus mod 16
if web_IPv4ConStatus<0 or web_IPv4ConStatus>2 then web_IPv4ConStatus=0
web_IPv6ConStatus=web_IPConStatus \ 16
if web_IPv6ConStatus<0 or web_IPv6ConStatus>2 then web_IPv6ConStatus=0


'========HTMLȨ���趨========
web_AdminHTMLSupport=false
web_AdminUBBSupport=false
web_AdminAllowNewLine=false

web_HTMLSupport=false
web_UBBSupport=false
web_AllowNewLine=false

dim web_adminlimit,web_guestlimit
web_adminlimit=lrs("adminhtml")
if clng(web_adminlimit and 1) <>0 then web_AdminHTMLSupport=true		'����Ա�ظ�������Ĭ���Ƿ�֧��HTML True:�� False:��
if clng(web_adminlimit and 2) <>0 then web_AdminUBBSupport=true		'����Ա�ظ�������Ĭ���Ƿ�֧��UBB���
if clng(web_adminlimit and 4) <>0 then web_AdminAllowNewLine=true	'����Ա��֧��HTML��UBBʱ��Ĭ���Ƿ�����س�����

web_guestlimit=lrs("guesthtml")
if clng(web_guestlimit and 1) <>0 then web_HTMLSupport=true			'�ÿ������Ƿ�֧��HTML
if clng(web_guestlimit and 2) <>0 then web_UBBSupport=true			'�ÿ������Ƿ�֧��UBB���
if clng(web_guestlimit and 4) <>0 then web_AllowNewLine=true			'�ÿͲ�֧��HTML��UBBʱ���Ƿ�����س�����


'========��ȫ������========
web_ShowIP=lrs("showip")			'������IP��ʾ����4λ����IPv4����4λ����IPv6
web_ShowIPv4=web_ShowIP mod 16
web_ShowIPv6=web_ShowIP \ 16
web_AdminShowIP=lrs("adminshowip")	'Ϊ����Ա��ʾIPλ������4λ����IPv4����4λ����IPv6
web_AdminShowIPv4=web_AdminShowIP mod 16
web_AdminShowIPv6=web_AdminShowIP \ 16
web_AdminShowOriginalIP=lrs("adminshoworiginalip")	'Ϊ����Ա��ʾԭʼIPλ������4λ����IPv4����4λ����IPv6
web_AdminShowOriginalIPv4=web_AdminShowOriginalIP mod 16
web_AdminShowOriginalIPv6=web_AdminShowOriginalIP \ 16

web_VcodeCount=clng(lrs("vcodecount") and &H0F)		'��¼��֤�볤��
web_WriteVcodeCount=clng(lrs("vcodecount") and &HF0) \ &H10		'������֤�볤��

'========�ʼ�����========
web_MailFlag=lrs("mailflag")
if clng(web_Mailflag and 1)<>0 then web_MailNewInform=true else web_MailNewInform=false		'������֪ͨ
if clng(web_Mailflag and 2)<>0 then web_MailReplyInform=true else web_MailReplyInform=false		'�ظ�֪ͨ
if clng(web_MailFlag and 4)<>0 then MailComponent="cdo" else MailComponent="jmail"

'========web UBB Flags========
web_UbbFlag=lrs("ubbflag")
if clng(web_UbbFlag and 1)<>0 then web_UbbFlag_image=true else web_UbbFlag_image=false
if clng(web_UbbFlag and 2)<>0 then web_UbbFlag_url=true else web_UbbFlag_url=false
if clng(web_UbbFlag and 4)<>0 then web_UbbFlag_autourl=true else web_UbbFlag_autourl=false
if clng(web_UbbFlag and 8)<>0 then web_UbbFlag_player=true else web_UbbFlag_player=false
if clng(web_UbbFlag and 16)<>0 then web_UbbFlag_paragraph=true else web_UbbFlag_paragraph=false
if clng(web_UbbFlag and 32)<>0 then web_UbbFlag_fontstyle=true else web_UbbFlag_fontstyle=false
if clng(web_UbbFlag and 64)<>0 then web_UbbFlag_fontcolor=true else web_UbbFlag_fontcolor=false
if clng(web_UbbFlag and 128)<>0 then web_UbbFlag_alignment=true else web_UbbFlag_alignment=false
'if clng(web_UbbFlag and 256)<>0 then web_UbbFlag_movement=true else web_UbbFlag_movement=false
'if clng(web_UbbFlag and 512)<>0 then web_UbbFlag_cssfilter=true else web_UbbFlag_cssfilter=false
if clng(web_UbbFlag and 1024)<>0 then web_UbbFlag_face=true else web_UbbFlag_face=false

'========================
lrs.Close
lrs.Open Replace(sql_loadconfig_floodconfig,"{0}",wm_id),lcn,0,1,1
web_flood_minwait=abs(lrs.Fields("minwait"))
web_flood_searchrange=abs(lrs.Fields("searchrange"))
web_flood_searchflag=lrs.Fields("searchflag")
if clng(web_flood_searchflag and 1)<>0 then web_flood_sfnewword=true else web_flood_sfnewword=false
if clng(web_flood_searchflag and 2)<>0 then web_flood_sfnewreply=true else web_flood_sfnewreply=false
if clng(web_flood_searchflag and 16)<>0 then web_flood_include=true else web_flood_include=false
if clng(web_flood_searchflag and 32)<>0 then web_flood_equal=true else web_flood_equal=false
if clng(web_flood_searchflag and 256)<>0 then web_flood_sititle=true else web_flood_sititle=false
if clng(web_flood_searchflag and 512)<>0 then web_flood_sicontent=true else web_flood_sicontent=false

lrs.Close

'================================
'================================
'================================
'Load User Settings
'================================
'================================
'================================
lrs.Open Replace(sql_loadconfig_config,"{0}",adminid),lcn,0,1,1

status=lrs("status")
if StatusOpen=true then
	if clng(status and 1)<>0 then	'���Ա�����
		StatusOpen=true
	else
		StatusOpen=false
	end if
end if
if StatusWrite=true then
	if clng(status and 2)<>0 then	'����Ȩ�޿���
		StatusWrite=true
	else
		StatusWrite=false
	end if
end if
if StatusSearch=true then
	if clng(status and 4)<>0 then	'����Ȩ�޿���
		StatusSearch=true
	else
		StatusSearch=false
	end if
end if
if clng(status and 8)<>0 then	'�ÿ�ͷ����ʾ����
	StatusShowHead=true
else
	StatusShowHead=false
end if
if clng(status and 16)<>0 then	'������Ҫ���
	StatusNeedAudit=true
else
	StatusNeedAudit=false
end if
if clng(status and 32)<>0 then	'�������Ļ�
	StatusWhisper=true
else
	StatusWhisper=false
end if
if clng(status and 64)<>0 then	'����ÿͼ������Ļ��ظ�
	StatusEncryptWhisper=true
else
	StatusEncryptWhisper=false
end if
if clng(status and 128)<>0 then	'����ÿͻظ�
	StatusGuestReply=true
else
	StatusGuestReply=false
end if
if StatusStatistics=true then
	if clng(status and 256)<>0 then	'����ͳ��
		StatusStatistics=true
	else
		StatusStatistics=false
	end if
end if

IPConStatus=lrs("ipconstatus")	'IP���β��ԣ���4λ����IPv4����4λ����IPv6
IPv4ConStatus=IPConStatus mod 16
if IPv4ConStatus<0 or IPv4ConStatus>2 then IPv4ConStatus=0
IPv6ConStatus=IPConStatus \ 16
if IPv6ConStatus<0 or IPv6ConStatus>2 then IPv6ConStatus=0

HomeLogo=lrs("homelogo")			'��վLogo��ַ
HomeName=lrs("homename")			'��վ����
HomeAddr=lrs("homeaddr")	'��վ��ַ



'========HTMLȨ���趨========
AdminHTMLSupport=false
AdminUBBSupport=false
AdminAllowNewLine=false
AdminViewCode=false

HTMLSupport=false
UBBSupport=false
AllowNewLine=false

dim adminlimit,guestlimit
adminlimit=lrs("adminhtml")
if clng(adminlimit and 1) <>0 then AdminHTMLSupport=true		'����Ա�ظ�������Ĭ���Ƿ�֧��HTML True:�� False:��
if clng(adminlimit and 2) <>0 then AdminUBBSupport=true		'����Ա�ظ�������Ĭ���Ƿ�֧��UBB���
if clng(adminlimit and 4) <>0 then AdminAllowNewLine=true	'����Ա��֧��HTML��UBBʱ��Ĭ���Ƿ�����س�����
if clng(adminlimit and 8) <>0 then AdminViewCode=true		'Ϊ����Ա��ʾʵ��HTML����

guestlimit=lrs("guesthtml")
if clng(guestlimit and 1) <>0 then HTMLSupport=true			'�ÿ������Ƿ�֧��HTML
if clng(guestlimit and 2) <>0 then UBBSupport=true			'�ÿ������Ƿ�֧��UBB���
if clng(guestlimit and 4) <>0 then AllowNewLine=true			'�ÿͲ�֧��HTML��UBBʱ���Ƿ�����س�����


'========��ȫ������========
AdminTimeOut=lrs("admintimeout")		'����Ա��¼��ʱ(��)
ShowIP=lrs("showip")			'������IP��ʾ����4λ����IPv4����4λ����IPv6
ShowIPv4=ShowIP mod 16
ShowIPv6=ShowIP \ 16
AdminShowIP=lrs("adminshowip")	'Ϊ����Ա��ʾIPλ������4λ����IPv4����4λ����IPv6
AdminShowIPv4=AdminShowIP mod 16
AdminShowIPv6=AdminShowIP \ 16
AdminShowOriginalIP=lrs("adminshoworiginalip")	'Ϊ����Ա��ʾԭʼIPλ������4λ����IPv4����4λ����IPv6
AdminShowOriginalIPv4=AdminShowOriginalIP mod 16
AdminShowOriginalIPv6=AdminShowOriginalIP \ 16

if web_ShowIPv4<ShowIPv4 then ShowIPv4=web_ShowIPv4		'WebMaster��������
if web_ShowIPv6<ShowIPv6 then ShowIPv6=web_ShowIPv6		'WebMaster��������
if web_AdminShowIPv4<AdminShowIPv4 then AdminShowIPv4=web_AdminShowIPv4		'WebMaster��������
if web_AdminShowIPv6<AdminShowIPv6 then AdminShowIPv6=web_AdminShowIPv6		'WebMaster��������
if web_AdminShowOriginalIPv4<AdminShowOriginalIPv4 then AdminShowOriginalIPv4=web_AdminShowOriginalIPv4		'WebMaster��������
if web_AdminShowOriginalIPv6<AdminShowOriginalIPv6 then AdminShowOriginalIPv6=web_AdminShowOriginalIPv6		'WebMaster��������

VcodeCount=clng(lrs("vcodecount") and &H0F)		'��¼��֤�볤��
if VcodeCount<web_VcodeCount then VcodeCount=web_VcodeCount		'WebMaster��������

WriteVcodeCount=clng(lrs("vcodecount") and &HF0) \ &H10		'������֤�볤��
if WriteVcodeCount<web_WriteVcodeCount then WriteVcodeCount=web_WriteVcodeCount		'WebMaster��������

'========�ʼ�����========
MailFlag=lrs("mailflag")
if clng(Mailflag and 1)<>0 then MailNewInform=true else MailNewInform=false		'������֪ͨ
if clng(Mailflag and 2)<>0 then MailReplyInform=true else MailReplyInform=false		'�ظ�֪ͨ
MailReceive=lrs("mailreceive")			'������֪ͨ���յ�ַ
MailFrom=lrs("mailfrom")				'�����˵�ַ
MailSmtpServer=lrs("mailsmtpserver")	'������SMTP��������ַ
MailUserId=lrs("mailuserid")			'��¼�û���
MailUserPass=lrs("mailuserpass")		'��¼����
MailLevel=lrs("maillevel")				'�����̶�

'========��������========
CssFontFamily=lrs("cssfontfamily")
CssFontSize=lrs("cssfontsize")
CssLineHeight=lrs("csslineheight")

VisualFlag=lrs("visualflag")
if clng(VisualFlag and 1)<>0 then ReplyInWord=true else ReplyInWord=false					'�ظ���Ƕ������
if clng(VisualFlag and 2)<>0 then ShowUbbTool=true else ShowUbbTool=false					'��ʾUBB������
if clng(VisualFlag and 4)<>0 then ShowTopPageList=true else ShowTopPageList=false			'�Ϸ���ʾ��ҳ
if clng(VisualFlag and 8)<>0 then ShowBottomPageList=true else ShowBottomPageList=false		'�·���ʾ��ҳ
if ShowTopPageList=false and ShowBottomPageList=false then ShowBottomPageList=true
if clng(VisualFlag and 16)<>0 then ShowTopSearchBox=true else ShowTopSearchBox=false		'�Ϸ���ʾ����
if clng(VisualFlag and 32)<>0 then ShowBottomSearchBox=true else ShowBottomSearchBox=false	'�·���ʾ����
if ShowTopSearchBox=false and ShowBottomSearchBox=false then ShowBottomSearchBox=true
if clng(VisualFlag and 64)<>0 then ShowAdvPageList=true else ShowAdvPageList=false			'����ʽ��ҳ
if clng(VisualFlag and 128)<>0 then HideHidden=true else HideHidden=false					'�������ݱ����ص�����
if clng(VisualFlag and 256)<>0 then HideAudit=true else HideAudit=false						'���ش��������
if clng(VisualFlag and 512)<>0 then HideWhisper=true else HideWhisper=false					'���ذ���δ�ظ����Ļ�
if clng(VisualFlag and 1024)<>0 then DisplayMode="forum" else DisplayMode="book"			'Ĭ�ϰ���ģʽ
if clng(VisualFlag and 2048)<>0 then LogoBannerMode=true else LogoBannerMode=false			'Logo��ʾģʽ

AdvPageListCount=lrs("advpagelistcount")	'����ʽ��ҳ����

'========UBB Flags========
UbbFlag=lrs("ubbflag")
if clng(UbbFlag and 1)<>0 then UbbFlag_image=true else UbbFlag_image=false
if clng(UbbFlag and 2)<>0 then UbbFlag_url=true else UbbFlag_url=false
if clng(UbbFlag and 4)<>0 then UbbFlag_autourl=true else UbbFlag_autourl=false
if clng(UbbFlag and 8)<>0 then UbbFlag_player=true else UbbFlag_player=false
if clng(UbbFlag and 16)<>0 then UbbFlag_paragraph=true else UbbFlag_paragraph=false
if clng(UbbFlag and 32)<>0 then UbbFlag_fontstyle=true else UbbFlag_fontstyle=false
if clng(UbbFlag and 64)<>0 then UbbFlag_fontcolor=true else UbbFlag_fontcolor=false
if clng(UbbFlag and 128)<>0 then UbbFlag_alignment=true else UbbFlag_alignment=false
'if clng(UbbFlag and 256)<>0 then UbbFlag_movement=true else UbbFlag_movement=false
'if clng(UbbFlag and 512)<>0 then UbbFlag_cssfilter=true else UbbFlag_cssfilter=false
if clng(UbbFlag and 1024)<>0 then UbbFlag_face=true else UbbFlag_face=false

TableAlign=lrs("tablealign")		'�����뷽ʽ left:����� center:���� right:�Ҷ���
TableWidth=lrs("tablewidth")		'����ȣ����ðٷֱ�
TableLeftWidth=lrs("tableleftwidth")	'����󷽿�ȣ����ðٷֱ�
WindowSpace=lrs("windowspace")			'����������

LeaveContentHeight=lrs("leavecontentheight")		'���������ݡ��ı��߶�
SearchTextWidth=lrs("searchtextwidth")				'�������������ı����
ReplyTextHeight=lrs("replytextheight")				'�ظ�������༭��߶�

ItemsPerPage=lrs("itemsperpage")		'ÿҳ��ʾ��������
TitlesPerPage=lrs("titlesperpage")		'ÿҳ��ʾ�ı�����(��̳ģʽ)
PicturesPerRow=lrs("picturesperrow")		'ÿ����ʾ��ͷ����
FrequentFaceCount=lrs("frequentfacecount")	'���������ͷ����
if FrequentFaceCount>FaceCount then FrequentFaceCount=FaceCount

PageControl=clng(lrs("pagecontrol"))
ShowBorder=false
ShowTitle=false
ShowContext=false
SelectContent=false
CopyContent=false
BeFramed=false
if clng(PageControl and 1)<> 0 then ShowBorder=true			'��ʾ�߿�
if clng(PageControl and 2)<> 0 then ShowTitle=true			'��ʾ������
if clng(PageControl and 4)<> 0 then ShowContext=true		'��ʾ�Ҽ��˵�
if clng(PageControl and 8)<> 0 then SelectContent=true		'����ѡ��
if clng(PageControl and 16)<> 0 then CopyContent=true		'������
if clng(PageControl and 32)<> 0 then BeFramed=true			'����frame
if ShowBorder=true then
	TableBorderWidth=1
else
	TableBorderWidth=0
end if

WordsLimit=abs(lrs("wordslimit"))		'���������

DelTip=false
DelReTip=false
DelSelTip=false
DelAdvTip=false
PassAuditTip=false
PassSelTip=false
PubWhisperTip=false
Bring2TopTip=false
Lock2TopTip=false
ReorderTip=false
DelConfirm=lrs("delconfirm")
if clng(DelConfirm and 1)<>0 then DelTip=true
if clng(DelConfirm and 2)<>0 then DelReTip=true
if clng(DelConfirm and 4)<>0 then DelSelTip=true
if clng(DelConfirm and 8)<>0 then DelAdvTip=true
if clng(DelConfirm and 16)<>0 then PassAuditTip=true
if clng(DelConfirm and 32)<>0 then PassSelTip=true
if clng(DelConfirm and 64)<>0 then PubWhisperTip=true
if clng(DelConfirm and 128)<>0 then Bring2TopTip=true
if clng(DelConfirm and 256)<>0 then Lock2TopTip=true
if clng(DelConfirm and 512)<>0 then ReorderTip=true

dim styleid
styleid=lrs("styleid")
lrs.Close
lrs.Open Replace(sql_loadconfig_style,"{0}",styleid),lcn,0,1,1
if lrs.EOF=true then
	lrs.Close
	lrs.Open sql_loadconfig_top1style,lcn,0,1,1
end if

ThemePath=lrs("themepath")

lrs.Close
lrs.Open Replace(sql_loadconfig_floodconfig,"{0}",adminid),lcn,0,1,1
flood_minwait=abs(lrs.Fields("minwait"))
flood_searchrange=abs(lrs.Fields("searchrange"))
flood_searchflag=lrs.Fields("searchflag")
if clng(flood_searchflag and 1)<>0 then flood_sfnewword=true else flood_sfnewword=false
if clng(flood_searchflag and 2)<>0 then flood_sfnewreply=true else flood_sfnewreply=false
if clng(flood_searchflag and 16)<>0 then flood_include=true else flood_include=false
if clng(flood_searchflag and 32)<>0 then flood_equal=true else flood_equal=false
if clng(flood_searchflag and 256)<>0 then flood_sititle=true else flood_sititle=false
if clng(flood_searchflag and 512)<>0 then flood_sicontent=true else flood_sicontent=false

lrs.Close : lcn.Close : set lrs=nothing : set lcn=nothing
%>