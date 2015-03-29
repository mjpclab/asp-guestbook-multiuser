<!-- #include file="common.asp" -->
<%
web_BookName="多用户留言本系统"

dim lcn,lrs
set lcn=server.CreateObject("ADODB.Connection")
set lrs=server.CreateObject("ADODB.Recordset")

CreateConn lcn,dbtype
lrs.Open Replace(sql_loadconfig_config,"{0}",wm_id),lcn,,,1

status=lrs("status")
if clng(status and 1)<>0 then	'留言本开启
	StatusOpen=true
else
	StatusOpen=false
end if
if clng(status and 2)<>0 then	'留言权限开启
	StatusWrite=true
else
	StatusWrite=false
end if
if clng(status and 4)<>0 then	'搜索权限开启
	StatusSearch=true
else
	StatusSearch=false
end if
if clng(status and 8)<>0 then	'申请权限开启
	StatusReg=true
else
	StatusReg=false
end if
if clng(status and 16)<>0 then	'找回密码权限开启
	StatusFindkey=true
else
	StatusFindkey=false
end if
if clng(status and 32)<>0 then	'用户登录权限开启
	StatusLogin=true
else
	StatusLogin=false
end if
if clng(status and 64)<>0 then	'自删帐号功能开启
	StatusUnreg=true
else
	StatusUnreg=false
end if
if clng(status and 256)<>0 then	'开启统计
	StatusStatistics=true
else
	StatusStatistics=false
end if

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

VcodeCount=clng(lrs("vcodecount") and &H0F)		'登录验证码长度
WriteVcodeCount=clng(lrs("vcodecount") and &HF0) \ &H10		'留言验证码长度

'========邮件设置========
MailFlag=lrs("mailflag")
if clng(Mailflag and 1)<>0 then MailNewInform=true else MailNewInform=false		'新留言通知
if clng(Mailflag and 2)<>0 then MailReplyInform=true else MailReplyInform=false		'回复通知
if clng(Mailflag and 4)<>0 then MailComponent="cdo" else MailComponent="jmail"

'========界面设置========
CssFontFamily=lrs("cssfontfamily")
CssFontSize=lrs("cssfontsize")
CssLineHeight=lrs("csslineheight")

VisualFlag=lrs("visualflag")
if clng(VisualFlag and 1)<>0 then ReplyInWord=true else ReplyInWord=false					'回复内嵌于留言
if clng(VisualFlag and 2)<>0 then ShowUbbTool=true else ShowUbbTool=false					'显示UBB工具栏
if clng(VisualFlag and 4)<>0 then ShowTopPageList=true else ShowTopPageList=false			'上方显示分页
if clng(VisualFlag and 8)<>0 then ShowBottomPageList=true else ShowBottomPageList=false		'下方显示分页
if ShowTopPageList=false and ShowBottomPageList=false then ShowBottomPageList=true
'if clng(VisualFlag and 16)<>0 then ShowTopSearchBox=true else ShowTopSearchBox=false		'上方显示搜索
'if clng(VisualFlag and 32)<>0 then ShowBottomSearchBox=true else ShowBottomSearchBox=false	'下方显示搜索
'if ShowTopSearchBox=false and ShowBottomSearchBox=false then ShowBottomSearchBox=true
if clng(VisualFlag and 64)<>0 then ShowAdvPageList=true else ShowAdvPageList=false			'区段式分页
if clng(VisualFlag and 1024)<>0 then DisplayMode="forum" else DisplayMode="book"			'默认版面模式

AdvPageListCount=lrs("advpagelistcount")	'区段式分页项数

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
TableAlign=lrs("tablealign")		'表格对齐方式 left:左对齐 center:居中 right:右对齐
TableWidth=lrs("tablewidth")		'表格宽度，可用百分比
WindowSpace=lrs("windowspace")		'窗格区块间距
TableLeftWidth=lrs("tableleftwidth")	'表格左方宽度，可用百分比

UBBSupport=true
web_UBBSupport=true
ShowUbbTool=true

UbbToolWidth=lrs("ubbtoolwidth")					'“写留言”UBB工具宽
UbbToolHeight=lrs("ubbtoolheight")					'“写留言”UBB工具高
SearchTextWidth=lrs("searchtextwidth")				'搜索框宽度
AdvDelTextWidth=lrs("advdeltextwidth")				'“高级删除”中文本宽
SetInfoTextWidth=lrs("setinfotextwidth")			'“修改资料”中文本宽
FilterTextWidth=lrs("filtertextwidth")				'“内容过滤”中文本宽
ReplyTextWidth=lrs("replytextwidth")				'公告编辑框宽度
ReplyTextHeight=lrs("replytextheight")				'公告编辑框高度

ItemsPerPage=lrs("itemsperpage")		'每页显示的项目数
TitlesPerPage=lrs("titlesperpage")		'每页显示的标题数

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

PageBackColor=lrs("pagebackcolor")			'网页背景色
PageBackImage=lrs("pagebackimage")			'网页背景图片，使用相对路径

TableBorderColor=lrs("tablebordercolor")		'表格边框颜色
TableBorderColorInactive=lrs("tablebordercolorinactive")		'表格边框颜色（非活动）
TableBorderWidth=TableBorderWidth*lrs("tableborderwidth")		'表格外框粗细
TableBorderPadding=lrs("tableborderpadding")					'表格外框边距
TableBGC=lrs("tablebgc")						'表格主体背景色
TablePic=lrs("tablepic")						'表格主题背景图片

TitleColor=lrs("titlecolor")		'留言本标题文字颜色
TitleBGC=lrs("titlebgc")			'留言本标题背景色
TitlePic=lrs("titlepic")			'留言本标题背景图片

TableBulletinTitleColor=lrs("tablebulletintitlecolor")	'表格－公告标题文字颜色
TableBulletinTitleBGC=lrs("tablebulletintitlebgc")		'表格－公告标题文字颜色
TableBulletinTitlePic=lrs("tablebulletintitlepic")		'表格－公告标题文字颜色

TableTitleColor=lrs("tabletitlecolor")		'表格－留言标题文字颜色
TableTitleBGC=lrs("tabletitlebgc")			'表格－留言/公告标题背景色
TableTitlePic=lrs("tabletitlepic")			'表格－留言/公告标题背景图片

TableContentColor=lrs("tablecontentcolor")		'表格－文章内容文字颜色
TableContentBGC=lrs("tablecontentbgc")			'表格－文章内容背景色

TableReplyColor=lrs("tablereplycolor")		'表格－版主回复标题文字颜色
TableReplyBGC=lrs("tablereplybgc")			'表格－版主回复标题背景色
TableReplyPic=lrs("tablereplypic")			'表格－版主回复标题背景图片

TableReplyContentColor=lrs("tablereplycontentcolor")	'表格－版主回复内容文字颜色

TableGuestInfoColor=lrs("tableguestinfocolor")	'表格－留言者信息框(左方)文字颜色
TableGuestInfoBGC=lrs("tableguestinfobgc")		'表格－留言者信息框背景色
TableGuestInfoPic=lrs("tableguestinfopic")		'表格－留言者信息框背景图片

BlockBGC=lrs("blockbgc")					'内嵌区块背景色

FormColor=lrs("formcolor")
FormBGC=lrs("formbgc")

LinkNormal=lrs("linknormal")		'超链接－正常颜色
LinkHover=lrs("linkhover")			'超链接－鼠标悬停颜色
LinkActive=lrs("linkactive")		'超链接－活动颜色
LinkVisited=lrs("linkvisited")		'超链接－已访问颜色

PageNumColor_Curr=lrs("pagenumcolor_curr")			'当前页号颜色
PageNumColor_Normal=lrs("pagenumcolor_normal")		'一般页号颜色

Additional_Css=lrs("additional_css")	'附加CSS

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