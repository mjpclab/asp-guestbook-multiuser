<!-- #include file="common.asp" -->
<%
web_BookName="���û����Ա�ϵͳ"

dim lcn,lrs
set lcn=server.CreateObject("ADODB.Connection")
set lrs=server.CreateObject("ADODB.Recordset")

CreateConn lcn,dbtype
lrs.Open Replace(sql_loadconfig_config,"{0}",wm_id),lcn,,,1

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
if clng(status and 8)<>0 then	'����Ȩ�޿���
	StatusReg=true
else
	StatusReg=false
end if
if clng(status and 16)<>0 then	'�һ�����Ȩ�޿���
	StatusFindkey=true
else
	StatusFindkey=false
end if
if clng(status and 32)<>0 then	'�û���¼Ȩ�޿���
	StatusLogin=true
else
	StatusLogin=false
end if
if clng(status and 64)<>0 then	'��ɾ�ʺŹ��ܿ���
	StatusUnreg=true
else
	StatusUnreg=false
end if
if clng(status and 256)<>0 then	'����ͳ��
	StatusStatistics=true
else
	StatusStatistics=false
end if

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

VcodeCount=clng(lrs("vcodecount") and &H0F)		'��¼��֤�볤��
WriteVcodeCount=clng(lrs("vcodecount") and &HF0) \ &H10		'������֤�볤��

'========�ʼ�����========
MailFlag=lrs("mailflag")
if clng(Mailflag and 1)<>0 then MailNewInform=true else MailNewInform=false		'������֪ͨ
if clng(Mailflag and 2)<>0 then MailReplyInform=true else MailReplyInform=false		'�ظ�֪ͨ
if clng(Mailflag and 4)<>0 then MailComponent="cdo" else MailComponent="jmail"

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
'if clng(VisualFlag and 16)<>0 then ShowTopSearchBox=true else ShowTopSearchBox=false		'�Ϸ���ʾ����
'if clng(VisualFlag and 32)<>0 then ShowBottomSearchBox=true else ShowBottomSearchBox=false	'�·���ʾ����
'if ShowTopSearchBox=false and ShowBottomSearchBox=false then ShowBottomSearchBox=true
if clng(VisualFlag and 64)<>0 then ShowAdvPageList=true else ShowAdvPageList=false			'����ʽ��ҳ
if clng(VisualFlag and 1024)<>0 then DisplayMode="forum" else DisplayMode="book"			'Ĭ�ϰ���ģʽ

AdvPageListCount=lrs("advpagelistcount")	'����ʽ��ҳ����

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

UbbToolWidth=lrs("ubbtoolwidth")					'��д���ԡ�UBB���߿�
UbbToolHeight=lrs("ubbtoolheight")					'��д���ԡ�UBB���߸�
SearchTextWidth=lrs("searchtextwidth")				'��������
AdvDelTextWidth=lrs("advdeltextwidth")				'���߼�ɾ�������ı���
SetInfoTextWidth=lrs("setinfotextwidth")			'���޸����ϡ����ı���
FilterTextWidth=lrs("filtertextwidth")				'�����ݹ��ˡ����ı���
ReplyTextWidth=lrs("replytextwidth")				'����༭����
ReplyTextHeight=lrs("replytextheight")				'����༭��߶�

ItemsPerPage=lrs("itemsperpage")		'ÿҳ��ʾ����Ŀ��
TitlesPerPage=lrs("titlesperpage")		'ÿҳ��ʾ�ı�����

DelTip=false
DelReTip=false
DelSelTip=false
DelAdvTip=false
DelDecTip=false
DelSelDecTip=false
DelConfirm=lrs("delconfirm")
if clng(DelConfirm and 1)<>0 then DelTip=true
if clng(DelConfirm and 2)<>0 then DelReTip=true
if clng(DelConfirm and 4)<>0 then DelSelTip=true
if clng(DelConfirm and 8)<>0 then DelAdvTip=true
if clng(DelConfirm and 16)<>0 then DelDecTip=true
if clng(DelConfirm and 32)<>0 then DelSelDecTip=true

'==========================
dim styleid
styleid=lrs("styleid")
lrs.Close
lrs.Open Replace(sql_loadconfig_style,"{0}",styleid),lcn,,,1
if lrs.EOF=true then
	lrs.Close
	lrs.Open sql_loadconfig_top1style,lcn,,,1
end if

PageBackColor=lrs("pagebackcolor")			'��ҳ����ɫ
PageBackImage=lrs("pagebackimage")			'��ҳ����ͼƬ��ʹ�����·��

TableBorderColor=lrs("tablebordercolor")		'���߿���ɫ
TableBorderColorInactive=lrs("tablebordercolorinactive")		'���߿���ɫ���ǻ��
TableBorderWidth=TableBorderWidth*lrs("tableborderwidth")		'�������ϸ
TableBorderPadding=lrs("tableborderpadding")					'������߾�
TableBGC=lrs("tablebgc")						'������屳��ɫ
TablePic=lrs("tablepic")						'������ⱳ��ͼƬ

TitleColor=lrs("titlecolor")		'���Ա�����������ɫ
TitleBGC=lrs("titlebgc")			'���Ա����ⱳ��ɫ
TitlePic=lrs("titlepic")			'���Ա����ⱳ��ͼƬ

TableBulletinTitleColor=lrs("tablebulletintitlecolor")	'��񣭹������������ɫ
TableBulletinTitleBGC=lrs("tablebulletintitlebgc")		'��񣭹������������ɫ
TableBulletinTitlePic=lrs("tablebulletintitlepic")		'��񣭹������������ɫ

TableTitleColor=lrs("tabletitlecolor")		'������Ա���������ɫ
TableTitleBGC=lrs("tabletitlebgc")			'�������/������ⱳ��ɫ
TableTitlePic=lrs("tabletitlepic")			'�������/������ⱳ��ͼƬ

TableContentColor=lrs("tablecontentcolor")		'�����������������ɫ
TableContentBGC=lrs("tablecontentbgc")			'����������ݱ���ɫ

TableReplyColor=lrs("tablereplycolor")		'��񣭰����ظ�����������ɫ
TableReplyBGC=lrs("tablereplybgc")			'��񣭰����ظ����ⱳ��ɫ
TableReplyPic=lrs("tablereplypic")			'��񣭰����ظ����ⱳ��ͼƬ

TableReplyContentColor=lrs("tablereplycontentcolor")	'��񣭰����ظ�����������ɫ

TableGuestInfoColor=lrs("tableguestinfocolor")	'�����������Ϣ��(��)������ɫ
TableGuestInfoBGC=lrs("tableguestinfobgc")		'�����������Ϣ�򱳾�ɫ
TableGuestInfoPic=lrs("tableguestinfopic")		'�����������Ϣ�򱳾�ͼƬ

BlockBGC=lrs("blockbgc")					'��Ƕ���鱳��ɫ

FormColor=lrs("formcolor")
FormBGC=lrs("formbgc")

LinkNormal=lrs("linknormal")		'�����ӣ�������ɫ
LinkHover=lrs("linkhover")			'�����ӣ������ͣ��ɫ
LinkActive=lrs("linkactive")		'�����ӣ����ɫ
LinkVisited=lrs("linkvisited")		'�����ӣ��ѷ�����ɫ

PageNumColor_Curr=lrs("pagenumcolor_curr")			'��ǰҳ����ɫ
PageNumColor_Normal=lrs("pagenumcolor_normal")		'һ��ҳ����ɫ

Additional_Css=lrs("additional_css")	'����CSS

'==========================
lrs.Close
lrs.Open Replace(sql_loadconfig_floodconfig,"{0}",wm_id),lcn,,,1
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