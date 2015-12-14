<!-- #include file="include/sql/loadconfig.asp" -->
<%
web_BookName="多用户留言本系统"

dim lcn,lrs
set lcn=server.CreateObject("ADODB.Connection")
set lrs=server.CreateObject("ADODB.Recordset")

Call CreateConn(lcn)
lrs.Open Replace(sql_loadconfig_config,"{0}",wm_id),lcn,,,1

status=lrs("status")
StatusOpen=CBool(status and 1)	'留言本开启
StatusWrite=CBool(status and 2)	'留言权限开启
StatusSearch=CBool(status and 4)	'搜索权限开启
StatusReg=CBool(status and 8)	'申请权限开启
StatusFindkey=CBool(status and 16)	'找回密码权限开启
StatusLogin=CBool(status and 32)	'用户登录权限开启
StatusUnreg=CBool(status and 64)	'自删帐号功能开启
StatusStatistics=CBool(status and 256)  	'统计开启

web_IPConStatus=lrs("ipconstatus")	'IP屏蔽策略，低4位用于IPv4，高4位用于IPv6
web_IPv4ConStatus=web_IPConStatus mod 16
if web_IPv4ConStatus<0 or web_IPv4ConStatus>2 then web_IPv4ConStatus=0
web_IPv6ConStatus=web_IPConStatus \ 16
if web_IPv6ConStatus<0 or web_IPv6ConStatus>2 then web_IPv6ConStatus=0

StatusShowHead=true

'========HTML权限设定========
AdminViewCode=false
adminlimit=lrs("adminhtml")
if clng(adminlimit and 8) <>0 then AdminViewCode=true		'为管理员显示实际HTML代码


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

VcodeCount=lrs("vcodecount") mod 16		'登录验证码长度
WriteVcodeCount=lrs("vcodecount") \ 16	'留言验证码长度

'========邮件设置========
MailFlag=lrs("mailflag")
MailNewInform=CBool(Mailflag and 1)		'新留言通知
MailReplyInform=CBool(Mailflag and 2)		'回复通知
if CBool(Mailflag and 4) then MailComponent="cdo" else MailComponent="jmail"

'========界面设置========
CssFontFamily=lrs("cssfontfamily")
CssFontSize=lrs("cssfontsize")
CssLineHeight=lrs("csslineheight")

VisualFlag=lrs("visualflag")
ReplyInWord=CBool(VisualFlag and 1)					'回复内嵌于留言
ShowUbbTool=CBool(VisualFlag and 2)					'显示UBB工具栏
ShowTopPageList=CBool(VisualFlag and 4)			'上方显示分页
ShowBottomPageList=CBool(VisualFlag and 8)		'下方显示分页
if ShowTopPageList=false and ShowBottomPageList=false then ShowBottomPageList=true
'ShowTopSearchBox=CBool(VisualFlag and 16)		'上方显示搜索
'ShowBottomSearchBox=CBool(VisualFlag and 32)	'下方显示搜索
'if ShowTopSearchBox=false and ShowBottomSearchBox=false then ShowBottomSearchBox=true
ShowAdvPageList=CBool(VisualFlag and 64)			'区段式分页
if CBool(VisualFlag and 1024) then DisplayMode="forum" else DisplayMode="book"			'默认版面模式

AdvPageListCount=lrs("advpagelistcount")	'区段式分页项数

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
TableAlign=lrs("tablealign")		'表格对齐方式 left:左对齐 center:居中 right:右对齐
TableWidth=lrs("tablewidth")		'表格宽度，可用百分比
WindowSpace=lrs("windowspace")		'窗格区块间距
TableLeftWidth=lrs("tableleftwidth")	'表格左方宽度，可用百分比

UBBSupport=true
web_UBBSupport=true
ShowUbbTool=true

SearchTextWidth=lrs("searchtextwidth")				'搜索框宽度
ReplyTextHeight=lrs("replytextheight")				'公告编辑框高度

ItemsPerPage=lrs("itemsperpage")		'每页显示的项目数
TitlesPerPage=lrs("titlesperpage")		'每页显示的标题数

PageControl=clng(lrs("pagecontrol"))
ShowBorder=CBool(PageControl and 1)			'显示边框

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