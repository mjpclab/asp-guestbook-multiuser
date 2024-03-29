<!-- #include file="include/sql/loadconfig.asp" -->
<%
dim lcn,lrs
set lcn=server.CreateObject("ADODB.Connection")
set lrs=server.CreateObject("ADODB.Recordset")

Call CreateConn(lcn)
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
StatusOpen=CBool(status and 1)	'留言本开启
StatusWrite=CBool(status and 2)	'留言权限开启
StatusSearch=CBool(status and 4)	'搜索权限开启
StatusLogin=CBool(status and 32)	'用户登录权限开启
StatusStatistics=CBool(status and 256)	'统计开启
StatusNeedAudit=CBool(status and 512)	'全局审核开启

web_StatusOpen=StatusOpen
web_StatusWrite=StatusWrite
web_StatusSearch=StatusSearch
web_StatusStatistics=StatusStatistics
web_StatusNeedAudit=StatusNeedAudit

web_IPConStatus=lrs("ipconstatus")	'IP屏蔽策略，低4位用于IPv4，高4位用于IPv6
web_IPv4ConStatus=web_IPConStatus mod 16
if web_IPv4ConStatus<0 or web_IPv4ConStatus>2 then web_IPv4ConStatus=0
web_IPv6ConStatus=web_IPConStatus \ 16
if web_IPv6ConStatus<0 or web_IPv6ConStatus>2 then web_IPv6ConStatus=0


'========HTML权限设定========
dim web_adminlimit,web_guestlimit
web_adminlimit=lrs("adminhtml")
web_AdminHTMLSupport=CBool(web_adminlimit and 1)		'管理员回复、公告默认是否支持HTML True:是 False:否
web_AdminUBBSupport=CBool(web_adminlimit and 2)		'管理员回复、公告默认是否支持UBB标记
web_AdminAllowNewLine=CBool(web_adminlimit and 4)	'管理员不支持HTML和UBB时，默认是否允许回车换行

web_guestlimit=lrs("guesthtml")
web_HTMLSupport=CBool(web_guestlimit and 1)			'访客留言是否支持HTML
web_UBBSupport=CBool(web_guestlimit and 2)			'访客留言是否支持UBB标记
web_AllowNewLine=CBool(web_guestlimit and 4)			'访客不支持HTML和UBB时，是否允许回车换行


'========安全性设置========
web_ShowIP=lrs("showip")			'留言者IP显示，低4位用于IPv4，高4位用于IPv6
web_ShowIPv4=web_ShowIP mod 16
web_ShowIPv6=web_ShowIP \ 16
web_AdminShowIP=lrs("adminshowip")	'为管理员显示IP位数，低4位用于IPv4，高4位用于IPv6
web_AdminShowIPv4=web_AdminShowIP mod 16
web_AdminShowIPv6=web_AdminShowIP \ 16
web_AdminShowOriginalIP=lrs("adminshoworiginalip")	'为管理员显示原始IP位数，低4位用于IPv4，高4位用于IPv6
web_AdminShowOriginalIPv4=web_AdminShowOriginalIP mod 16
web_AdminShowOriginalIPv6=web_AdminShowOriginalIP \ 16

web_VcodeCount=clng(lrs("vcodecount") and &H0F)		'登录验证码长度
web_WriteVcodeCount=clng(lrs("vcodecount") and &HF0) \ &H10		'留言验证码长度

'========行为========
ServerTimezoneOffset=lrs("servertimezoneoffset")

'========邮件设置========
web_MailFlag=lrs("mailflag")
web_MailNewInform=CBool(web_Mailflag and 1)		'新留言通知
web_MailReplyInform=CBool(web_Mailflag and 2)		'回复通知
if CBool(web_Mailflag and 4) then MailComponent="cdo" else MailComponent="jmail"

'========web UBB Flags========
web_UbbFlag=lrs("ubbflag")
web_UbbFlag_image=CBool(web_UbbFlag and 1)
web_UbbFlag_url=CBool(web_UbbFlag and 2)
web_UbbFlag_autourl=CBool(web_UbbFlag and 4)
web_UbbFlag_player=CBool(web_UbbFlag and 8)
web_UbbFlag_paragraph=CBool(web_UbbFlag and 16)
web_UbbFlag_fontstyle=CBool(web_UbbFlag and 32)
web_UbbFlag_fontcolor=CBool(web_UbbFlag and 64)
web_UbbFlag_alignment=CBool(web_UbbFlag and 128)
'web_UbbFlag_movement=CBool(web_UbbFlag and 256)
'web_UbbFlag_cssfilter=CBool(web_UbbFlag and 512)
web_UbbFlag_face=CBool(web_UbbFlag and 1024)
web_UbbFlag_markdown_paragraph=CBool(web_UbbFlag and 32768)
web_UbbFlag_markdown_fontstyle=CBool(web_UbbFlag and 65536)

'========================
lrs.Close
lrs.Open Replace(sql_loadconfig_floodconfig,"{0}",wm_id),lcn,0,1,1
web_flood_minwait=abs(lrs.Fields("minwait"))
web_flood_searchrange=abs(lrs.Fields("searchrange"))
web_flood_searchflag=lrs.Fields("searchflag")
web_flood_sfnewword=CBool(web_flood_searchflag and 1)
web_flood_sfnewreply=CBool(web_flood_searchflag and 2)
web_flood_include=CBool(web_flood_searchflag and 16)
web_flood_equal=CBool(web_flood_searchflag and 32)
web_flood_sititle=CBool(web_flood_searchflag and 256)
web_flood_sicontent=CBool(web_flood_searchflag and 512)

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
if StatusOpen then StatusOpen=CBool(status and 1)	'留言本开启
if StatusWrite then StatusWrite=CBool(status and 2)	'留言权限开启
if StatusSearch then StatusSearch=CBool(status and 4)	'搜索权限开启
StatusShowHead=CBool(status and 8)	'访客头像显示开关
if Not StatusNeedAudit then StatusNeedAudit=CBool(status and 16)	'留言需要审核
StatusWhisper=CBool(status and 32)	'允许悄悄话
StatusEncryptWhisper=CBool(status and 64)	'允许访客加密悄悄话回复
StatusGuestReply=CBool(status and 128)	'允许访客回复
if StatusStatistics then StatusStatistics=CBool(status and 256)	'统计开启
StatusAccountEnabled=Not CBool(status and 1073741824)	'账号启用状态
StatusAccountLoginEnabled=Not CBool(status and 536870912)	'账号启用状态
StatusAccountLeavewordEnabled=Not CBool(status and 268435456)	'账号启用状态

IPConStatus=lrs("ipconstatus")	'IP屏蔽策略，低4位用于IPv4，高4位用于IPv6
IPv4ConStatus=IPConStatus mod 16
if IPv4ConStatus<0 or IPv4ConStatus>2 then IPv4ConStatus=0
IPv6ConStatus=IPConStatus \ 16
if IPv6ConStatus<0 or IPv6ConStatus>2 then IPv6ConStatus=0

HomeLogo=lrs("homelogo")			'网站Logo地址
HomeName=lrs("homename")			'网站名称
HomeAddr=lrs("homeaddr")	'网站地址



'========HTML权限设定========
dim adminlimit,guestlimit
adminlimit=lrs("adminhtml")
AdminHTMLSupport=CBool(adminlimit and 1)	'管理员回复、公告默认是否支持HTML True:是 False:否
AdminUBBSupport=CBool(adminlimit and 2)		'管理员回复、公告默认是否支持UBB标记
AdminAllowNewLine=CBool(adminlimit and 4)	'管理员不支持HTML和UBB时，默认是否允许回车换行
AdminViewCode=CBool(adminlimit and 8)		'为管理员显示实际HTML代码

guestlimit=lrs("guesthtml")
HTMLSupport=CBool(guestlimit and 1)			'访客留言是否支持HTML
UBBSupport=CBool(guestlimit and 2)			'访客留言是否支持UBB标记
AllowNewLine=CBool(guestlimit and 4)		'访客不支持HTML和UBB时，是否允许回车换行


'========安全性设置========
AdminTimeOut=lrs("admintimeout")		'管理员登录超时(分)
ShowIP=lrs("showip")			'留言者IP显示，低4位用于IPv4，高4位用于IPv6
ShowIPv4=ShowIP mod 16
ShowIPv6=ShowIP \ 16
AdminShowIP=lrs("adminshowip")	'为管理员显示IP位数，低4位用于IPv4，高4位用于IPv6
AdminShowIPv4=AdminShowIP mod 16
AdminShowIPv6=AdminShowIP \ 16
AdminShowOriginalIP=lrs("adminshoworiginalip")	'为管理员显示原始IP位数，低4位用于IPv4，高4位用于IPv6
AdminShowOriginalIPv4=AdminShowOriginalIP mod 16
AdminShowOriginalIPv6=AdminShowOriginalIP \ 16

if web_ShowIPv4<ShowIPv4 then ShowIPv4=web_ShowIPv4		'WebMaster设置优先
if web_ShowIPv6<ShowIPv6 then ShowIPv6=web_ShowIPv6		'WebMaster设置优先
if web_AdminShowIPv4<AdminShowIPv4 then AdminShowIPv4=web_AdminShowIPv4		'WebMaster设置优先
if web_AdminShowIPv6<AdminShowIPv6 then AdminShowIPv6=web_AdminShowIPv6		'WebMaster设置优先
if web_AdminShowOriginalIPv4<AdminShowOriginalIPv4 then AdminShowOriginalIPv4=web_AdminShowOriginalIPv4		'WebMaster设置优先
if web_AdminShowOriginalIPv6<AdminShowOriginalIPv6 then AdminShowOriginalIPv6=web_AdminShowOriginalIPv6		'WebMaster设置优先

VcodeCount=lrs("vcodecount") mod 16		'登录验证码长度
if VcodeCount<web_VcodeCount then VcodeCount=web_VcodeCount		'WebMaster设置优先

WriteVcodeCount=lrs("vcodecount") \ 16	'留言验证码长度
if WriteVcodeCount<web_WriteVcodeCount then WriteVcodeCount=web_WriteVcodeCount		'WebMaster设置优先

'========邮件设置========
MailFlag=lrs("mailflag")
MailNewInform=CBool(Mailflag and 1)		'新留言通知
MailReplyInform=CBool(Mailflag and 2)	'回复通知
MailReceive=lrs("mailreceive")			'新留言通知接收地址
MailFrom=lrs("mailfrom")				'发件人地址
MailSmtpServer=lrs("mailsmtpserver")	'发件人SMTP服务器地址
MailUserId=lrs("mailuserid")			'登录用户名
MailUserPass=lrs("mailuserpass")		'登录密码
MailLevel=lrs("maillevel")				'紧急程度

'========界面设置========
CssFontFamily=lrs("cssfontfamily")
CssFontSize=lrs("cssfontsize")
CssLineHeight=lrs("csslineheight")

DisplayTimezoneOffset=lrs("displaytimezoneoffset")

VisualFlag=lrs("visualflag")
ReplyInWord=CBool(VisualFlag and 1)					'回复内嵌于留言
ShowUbbTool=CBool(VisualFlag and 2)					'显示UBB工具栏
ShowTopPageList=CBool(VisualFlag and 4)			'上方显示分页
ShowBottomPageList=CBool(VisualFlag and 8)		'下方显示分页
if Not ShowTopPageList and Not ShowBottomPageList then ShowBottomPageList=true
ShowTopSearchBox=CBool(VisualFlag and 16)		'上方显示搜索
ShowBottomSearchBox=CBool(VisualFlag and 32)	'下方显示搜索
if Not ShowTopSearchBox and Not ShowBottomSearchBox then ShowBottomSearchBox=true
ShowAdvPageList=CBool(VisualFlag and 64)			'区段式分页
HideHidden=CBool(VisualFlag and 128)					'隐藏内容被隐藏的留言
HideAudit=CBool(VisualFlag and 256)						'隐藏待审核留言
HideWhisper=CBool(VisualFlag and 512)					'隐藏版主未回复悄悄话
if CBool(VisualFlag and 1024) then DisplayMode="forum" else DisplayMode="book"			'默认版面模式
LogoBannerMode=CBool(VisualFlag and 2048)			'Logo显示模式

AdvPageListCount=lrs("advpagelistcount")	'区段式分页项数

'========UBB Flags========
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
UbbFlag_markdown_paragraph=CBool(UbbFlag and 32768)
UbbFlag_markdown_fontstyle=CBool(UbbFlag and 65536)

TableWidth=lrs("tablewidth")		'表格宽度，可用百分比
TableLeftWidth=lrs("tableleftwidth")	'表格左方宽度，可用百分比
WindowSpace=lrs("windowspace")			'窗格区块间距

LeaveContentHeight=lrs("leavecontentheight")		'“留言内容”文本高度
SearchTextWidth=lrs("searchtextwidth")				'“留言搜索”文本宽度
ReplyTextHeight=lrs("replytextheight")				'回复、公告编辑框高度

ItemsPerPage=lrs("itemsperpage")		'每页显示的留言数
TitlesPerPage=lrs("titlesperpage")		'每页显示的标题数(论坛模式)
PicturesPerRow=lrs("picturesperrow")		'每行显示的头像数
FrequentFaceCount=lrs("frequentfacecount")	'少量载入的头像数
if FrequentFaceCount>FaceCount then FrequentFaceCount=FaceCount

PageControl=clng(lrs("pagecontrol"))
'ShowBorder=CBool(PageControl and 1)		'显示边框
ShowTitle=CBool(PageControl and 2)			'显示标题栏
ShowContext=CBool(PageControl and 4)		'显示右键菜单
SelectContent=CBool(PageControl and 8)		'允许选择
CopyContent=CBool(PageControl and 16)		'允许复制
BeFramed=CBool(PageControl and 32)			'允许frame

WordsLimit=abs(lrs("wordslimit"))		'最大留言数

DelConfirm=lrs("delconfirm")
DelTip=CBool(DelConfirm and 1)
DelReTip=CBool(DelConfirm and 2)
DelSelTip=CBool(DelConfirm and 4)
DelAdvTip=CBool(DelConfirm and 8)
PassAuditTip=CBool(DelConfirm and 16)
PassSelTip=CBool(DelConfirm and 32)
PubWhisperTip=CBool(DelConfirm and 64)
Bring2TopTip=CBool(DelConfirm and 128)
Lock2TopTip=CBool(DelConfirm and 256)
ReorderTip=CBool(DelConfirm and 512)

dim styleid
styleid=lrs("styleid")
lrs.Close
lrs.Open Replace(sql_loadconfig_style,"{0}",styleid),lcn,0,1,1
if lrs.EOF then
	lrs.Close
	lrs.Open sql_loadconfig_top1style,lcn,0,1,1
end if

ThemePath=lrs("themepath")

lrs.Close
lrs.Open Replace(sql_loadconfig_floodconfig,"{0}",adminid),lcn,0,1,1
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