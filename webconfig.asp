<!-- #include file="include/sql/loadconfig.asp" -->
<%
web_BookName="���û����Ա�ϵͳ"

dim lcn,lrs
set lcn=server.CreateObject("ADODB.Connection")
set lrs=server.CreateObject("ADODB.Recordset")

Call CreateConn(lcn)
lrs.Open Replace(sql_loadconfig_config,"{0}",wm_id),lcn,,,1

status=lrs("status")
StatusOpen=CBool(status and 1)	'���Ա�����
StatusWrite=CBool(status and 2)	'����Ȩ�޿���
StatusSearch=CBool(status and 4)	'����Ȩ�޿���
StatusReg=CBool(status and 8)	'����Ȩ�޿���
StatusFindkey=CBool(status and 16)	'�һ�����Ȩ�޿���
StatusLogin=CBool(status and 32)	'�û���¼Ȩ�޿���
StatusUnreg=CBool(status and 64)	'��ɾ�ʺŹ��ܿ���
StatusStatistics=CBool(status and 256)  	'ͳ�ƿ���

web_IPConStatus=lrs("ipconstatus")	'IP���β��ԣ���4λ����IPv4����4λ����IPv6
web_IPv4ConStatus=web_IPConStatus mod 16
if web_IPv4ConStatus<0 or web_IPv4ConStatus>2 then web_IPv4ConStatus=0
web_IPv6ConStatus=web_IPConStatus \ 16
if web_IPv6ConStatus<0 or web_IPv6ConStatus>2 then web_IPv6ConStatus=0

StatusShowHead=true

'========HTMLȨ���趨========
AdminViewCode=false
adminlimit=lrs("adminhtml")
if clng(adminlimit and 8) <>0 then AdminViewCode=true		'Ϊ����Ա��ʾʵ��HTML����


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

VcodeCount=lrs("vcodecount") mod 16		'��¼��֤�볤��
WriteVcodeCount=lrs("vcodecount") \ 16	'������֤�볤��

'========�ʼ�����========
MailFlag=lrs("mailflag")
MailNewInform=CBool(Mailflag and 1)		'������֪ͨ
MailReplyInform=CBool(Mailflag and 2)		'�ظ�֪ͨ
if CBool(Mailflag and 4) then MailComponent="cdo" else MailComponent="jmail"

'========��������========
CssFontFamily=lrs("cssfontfamily")
CssFontSize=lrs("cssfontsize")
CssLineHeight=lrs("csslineheight")

VisualFlag=lrs("visualflag")
ReplyInWord=CBool(VisualFlag and 1)					'�ظ���Ƕ������
ShowUbbTool=CBool(VisualFlag and 2)					'��ʾUBB������
ShowTopPageList=CBool(VisualFlag and 4)			'�Ϸ���ʾ��ҳ
ShowBottomPageList=CBool(VisualFlag and 8)		'�·���ʾ��ҳ
if ShowTopPageList=false and ShowBottomPageList=false then ShowBottomPageList=true
'ShowTopSearchBox=CBool(VisualFlag and 16)		'�Ϸ���ʾ����
'ShowBottomSearchBox=CBool(VisualFlag and 32)	'�·���ʾ����
'if ShowTopSearchBox=false and ShowBottomSearchBox=false then ShowBottomSearchBox=true
ShowAdvPageList=CBool(VisualFlag and 64)			'����ʽ��ҳ
if CBool(VisualFlag and 1024) then DisplayMode="forum" else DisplayMode="book"			'Ĭ�ϰ���ģʽ

AdvPageListCount=lrs("advpagelistcount")	'����ʽ��ҳ����

UbbFlag=lrs("ubbflag")
UbbFlag_image=CBool(UbbFlag and 1)
UbbFlag_url=CBool(UbbFlag and 2)
UbbFlag_autourl=CBool(UbbFlag and 4)
UbbFlag_player=CBool(UbbFlag and 8)
UbbFlag_paragraph=CBool(UbbFlag and 16)
UbbFlag_fontstyle=CBool(UbbFlag and 32)
UbbFlag_fontcolor=CBool(UbbFlag and 64)
UbbFlag_alignment=CBool(UbbFlag and 128)
'UbbFlag_movement=CBool(UbbFlag and 256)
'UbbFlag_cssfilter=CBool(UbbFlag and 512)
UbbFlag_face=CBool(UbbFlag and 1024)
web_UbbFlag_image=UbbFlag_image
web_UbbFlag_url=UbbFlag_url
web_UbbFlag_autourl=UbbFlag_autourl
web_UbbFlag_player=UbbFlag_player
web_UbbFlag_paragraph=UbbFlag_paragraph
web_UbbFlag_fontstyle=UbbFlag_fontstyle
web_UbbFlag_fontcolor=UbbFlag_fontcolor
web_UbbFlag_alignment=UbbFlag_alignment
'web_UbbFlag_movement=UbbFlag_movement
'web_UbbFlag_cssfilter=UbbFlag_cssfilter
web_UbbFlag_face=UbbFlag_face

TableBorderWidth=1
TableAlign=lrs("tablealign")		'�����뷽ʽ left:����� center:���� right:�Ҷ���
TableWidth=lrs("tablewidth")		'����ȣ����ðٷֱ�
WindowSpace=lrs("windowspace")		'����������
TableLeftWidth=lrs("tableleftwidth")	'����󷽿�ȣ����ðٷֱ�

UBBSupport=true
web_UBBSupport=true
ShowUbbTool=true

SearchTextWidth=lrs("searchtextwidth")				'��������
ReplyTextHeight=lrs("replytextheight")				'����༭��߶�

ItemsPerPage=lrs("itemsperpage")		'ÿҳ��ʾ����Ŀ��
TitlesPerPage=lrs("titlesperpage")		'ÿҳ��ʾ�ı�����

PageControl=clng(lrs("pagecontrol"))
ShowBorder=CBool(PageControl and 1)			'��ʾ�߿�

DelConfirm=lrs("delconfirm")
DelTip=CBool(DelConfirm and 1)
DelReTip=CBool(DelConfirm and 2)
DelSelTip=CBool(DelConfirm and 4)
DelAdvTip=CBool(DelConfirm and 8)
DelDecTip=CBool(DelConfirm and 16)
DelSelDecTip=CBool(DelConfirm and 32)

'==========================
dim styleid
styleid=lrs("styleid")
lrs.Close
lrs.Open Replace(sql_loadconfig_style,"{0}",styleid),lcn,0,1,1
if lrs.EOF=true then
	lrs.Close
	lrs.Open sql_loadconfig_top1style,lcn,0,1,1
end if

ThemePath=lrs("themepath")

'==========================
lrs.Close
lrs.Open Replace(sql_loadconfig_floodconfig,"{0}",wm_id),lcn,,,1
flood_minwait=abs(lrs.Fields("minwait"))
flood_searchrange=abs(lrs.Fields("searchrange"))
flood_searchflag=lrs.Fields("searchflag")
flood_sfnewword=CBool(flood_searchflag and 1)
flood_sfnewreply=CBool(flood_searchflag and 2)
flood_include=CBool(flood_searchflag and 16)
flood_equal=CBool(flood_searchflag and 32)
flood_sititle=CBool(flood_searchflag and 256)
flood_sicontent=CBool(flood_searchflag and 512)


lrs.Close : lcn.Close : set lrs=nothing : set lcn=nothing
%>