<!-- #include file="webconfig.asp" -->
<link rel="stylesheet" type="text/css" href="style.css"/>
<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
<link rel="stylesheet" type="text/css" href="web_adminstyle.css"/>
<!-- #include file="style.asp" -->
<!-- #include file="adminstyle.asp" -->
<!-- #include file="web_adminstyle.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%

sub errbox(byval errmsg)
	Call MessagePage(errmsg,"web_config.asp?page=" &Request.Form("page"))
	Response.End
end sub

function isnum(byval pdnum)
if isnumeric(pdnum)=true then
	isnum=true
	exit function
end if

if right(pdnum,1)="%" then
	pdnum=left(pdnum,len(pdnum)-1)
	if isnumeric(pdnum) then
		isnum=true
		exit function
	end if
end if

isnum=false	
end function

Response.Expires=-1

showpage=15
if isnumeric(Request.Form("page")) and len(Request.Form("page"))<=2 and Request.Form("page")<>"" then showpage=clng(Request.Form("page"))



if isnumeric(request.Form("admintimeout"))=false and clng(showpage and 1)<> 0 then
	errbox "“管理员登录超时”必须为数字。"
elseif isnumeric(request.Form("showip"))=false and clng(showpage and 1)<> 0 then
	errbox "“为访客显示IP”必须为数字。"
elseif isnumeric(request.Form("adminshowip"))=false and clng(showpage and 1)<> 0 then
	errbox "“为管理员显示IP”必须为数字。"
elseif isnumeric(request.Form("adminshoworiginalip"))=false and clng(showpage and 1)<> 0 then
	errbox "“为管理员显示原IP”必须为数字。"
elseif isnumeric(request.Form("vcodecount"))=false and clng(showpage and 1)<> 0 then
	errbox "“登录验证码长度”必须为数字。"
elseif isnumeric(request.Form("writevcodecount"))=false and clng(showpage and 1)<> 0 then
	errbox "“留言验证码长度”必须为数字。"
elseif isnum(request.Form("tablewidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“留言本最大宽度”必须为数字或百分比。"
elseif isnumeric(request.Form("windowspace"))=false and clng(showpage and 4)<> 0 then
	errbox "“窗口区块间距”必须为数字。"
elseif isnum(request.Form("tableleftwidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“留言本左窗格宽度”必须为数字或百分比。"
elseif isnumeric(request.Form("advdeltextwidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“‘高级删除’中文本宽”必须为数字。"
elseif isnumeric(request.Form("setinfotextwidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“‘修改密码’中文本宽”必须为数字。"
elseif isnumeric(request.Form("ubbtoolwidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“‘写留言’UBB工具宽”必须为数字。"
elseif isnumeric(request.Form("ubbtoolheight"))=false and clng(showpage and 4)<> 0 then
	errbox "“‘写留言’UBB工具高”必须为数字。"
elseif isnumeric(request.Form("searchtextwidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“搜索框宽度”必须为数字。"
elseif isnumeric(request.Form("filtertextwidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“‘内容过滤’中文本宽”必须为数字。"
elseif isnumeric(request.Form("replytextwidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“公告编辑框宽度”必须为数字。"
elseif isnumeric(request.Form("replytextheight"))=false and clng(showpage and 4)<> 0 then
	errbox "“公告编辑框高度”必须为数字。"
elseif isnumeric(request.Form("itemsperpage"))=false and clng(showpage and 4)<> 0 then
	errbox "“每页显示的项目数”必须为数字。"
elseif isnumeric(request.Form("titlesperpage"))=false and clng(showpage and 4)<> 0 then
	errbox "“每页显示的标题数”必须为数字。"
elseif isnumeric(request.Form("advpagelistcount"))=false and clng(showpage and 8)<> 0 then
	errbox "“区段式分页项数”必须为数字。"
else

	if clng(showpage and 1)<> 0 then
		tstatus1=Request.Form("status1")
		if tstatus1<>"0" and tstatus1<>"1" then tstatus1=1
				
		tstatus2=Request.Form("status2")
		if tstatus2<>"0" and tstatus2<>"2" then tstatus2=2
				
		tstatus3=Request.Form("status3")
		if tstatus3<>"0" and tstatus3<>"4" then tstatus3=4
				
		tstatus4=Request.Form("status4")
		if tstatus4<>"0" and tstatus4<>"8" then tstatus4=8
				
		tstatus5=Request.Form("status5")
		if tstatus5<>"0" and tstatus5<>"16" then tstatus5=16
				
		tstatus6=Request.Form("status6")
		if tstatus6<>"0" and tstatus6<>"32" then tstatus6=32
				
		tstatus7=Request.Form("status7")
		if tstatus7<>"0" and tstatus7<>"64" then tstatus6=64

		tstatus9=Request.Form("status9")
		if tstatus9<>"0" and tstatus9<>"256" then tstatus9=256

		tstatus=clng(tstatus1)+clng(tstatus2)+clng(tstatus3)+clng(tstatus4)+clng(tstatus5)+clng(tstatus6)+clng(tstatus7)+clng(tstatus9)
				
		tadminhtml=0
		if Request.Form("adminhtml")="1" then tadminhtml=tadminhtml+1
		if Request.Form("adminubb")="1" then tadminhtml=tadminhtml+2
		if Request.Form("adminertn")="1" then tadminhtml=tadminhtml+4
		if Request.Form("adminviewcode")="1" then tadminhtml=tadminhtml+8

		tguesthtml=0
		if Request.Form("guesthtml")="1" then tguesthtml=tguesthtml+1
		if Request.Form("guestubb")="1" then tguesthtml=tguesthtml+2
		if Request.Form("guestertn")="1" then tguesthtml=tguesthtml+4

		tadmintimeout=Request.Form("admintimeout")
		if len(cstr(tadmintimeout))>4 or isnumeric(tadmintimeout)=false then tadmintimeout=1440
		if clng(tadmintimeout)>1440 then tadmintimeout=1440
		if clng(tadmintimeout)<1 then tadmintimeout="20"
		
		tshowip=Request.Form("showip")
		if clng(tshowip)<0 or clng(tshowip)>4 then tshowip="0"

		tadminshowip=Request.Form("adminshowip")
		if clng(tadminshowip)<0 or clng(tadminshowip)>4 then tadminshowip="0"

		tadminshoworiginalip=Request.Form("adminshoworiginalip")
		if clng(tadminshoworiginalip)<0 or clng(tadminshoworiginalip)>4 then tadminshoworiginalip="0"

		tvcodecount=Request.Form("vcodecount")
		if len(cstr(tvcodecount))>2 then tvcodecount=4
		if clng(tvcodecount)>10 then tvcodecount=4
		if clng(tvcodecount)<0 then tvcodecount=4
		tvcodecount=clng(tvcodecount)
	
		twritevcodecount=Request.Form("writevcodecount")
		if len(cstr(twritevcodecount))>2 then twritevcodecount=4
		if clng(twritevcodecount)>10 then twritevcodecount=4
		if clng(twritevcodecount)<0 then twritevcodecount=4
		twritevcodecount=clng(twritevcodecount) * &H10

	end if
	
	if clng(showpage and 2)<> 0 then
		tmailflag=0
		if Request.Form("mailnewinform")="1" then tmailflag=tmailflag+1
		if Request.Form("mailreplyinform")="1" then tmailflag=tmailflag+2
		if Request.Form("mailcomponent")="1" then tmailflag=tmailflag+4	
	end if
	
	if clng(showpage and 4)<> 0 then
		tcssfontfamily=Request.Form("cssfontfamily")
		if len(tcssfontfamily)>48 then tcssfontfamily=left(tcssfontfamily,48)
		
		tcssfontsize=Request.Form("cssfontsize")
		if len(tcssfontsize)>8 then tcssfontsize=left(tcssfontsize,8)
		
		tcsslineheight=Request.Form("csslineheight")
		if len(tcsslineheight)>8 then tcsslineheight=left(tcsslineheight,8)
	
		ttablewidth=Request.Form("tablewidth")
		if len(cstr(ttablewidth))>5 then ttablewidth="630"
		if isnumeric(ttablewidth) then
			if clng(ttablewidth)<1 then ttablewidth="630"
			ttablewidth=cstr(round(clng(ttablewidth)))
		else
			if left(ttablewidth,1)="-" then ttablewidth=right(ttablewidth,len(ttablewidth)-1)
		end if
	
		twindowspace=Request.Form("windowspace")
		if len(cstr(twindowspace))>3 then twindowspace=20
		if clng(twindowspace)>255 then twindowspace=20
		if clng(twindowspace)<1 then twindowspace=20

		ttableleftwidth=Request.Form("tableleftwidth")
		if len(cstr(ttableleftwidth))>5 then ttableleftwidth="150"
		if isnumeric(ttableleftwidth) then
			if clng(ttableleftwidth)<1 then ttableleftwidth="150"
			ttableleftwidth=cstr(round(clng(ttableleftwidth)))
		else
			if left(ttableleftwidth,1)="-" then ttableleftwidth=right(ttableleftwidth,len(ttableleftwidth)-1)
		end if
	
		tubbtoolwidth=Request.Form("ubbtoolwidth")
		if len(cstr(tubbtoolwidth))>4 then tubbtoolwidth=320
		if clng(tubbtoolwidth)>9999 then tubbtoolwidth=320
		if clng(tubbtoolwidth)<1 then tubbtoolwidth=320
	
		tubbtoolheight=Request.Form("ubbtoolheight")
		if len(cstr(tubbtoolheight))>4 then tubbtoolheight=48
		if clng(tubbtoolheight)>9999 then tubbtoolheight=48
		if clng(tubbtoolheight)<1 then tubbtoolheight=48
	
		tsearchtextwidth=Request.Form("searchtextwidth")
		if len(cstr(tsearchtextwidth))>3 then tsearchtextwidth=20
		if clng(tsearchtextwidth)>255 then tsearchtextwidth=255
		if clng(tsearchtextwidth)<1 then tsearchtextwidth=20
	
		tadvdeltextwidth=Request.Form("advdeltextwidth")
		if len(cstr(tadvdeltextwidth))>3 then tadvdeltextwidth=20
		if clng(tadvdeltextwidth)>255 then tadvdeltextwidth=255
		if clng(tadvdeltextwidth)<1 then tadvdeltextwidth=20
	
		tsetinfotextwidth=Request.Form("setinfotextwidth")
		if len(cstr(tsetinfotextwidth))>3 then tsetinfotextwidth=20
		if clng(tsetinfotextwidth)>255 then tsetinfotextwidth=255
		if clng(tsetinfotextwidth)<1 then tsetinfotextwidth=20

		tfiltertextwidth=Request.Form("filtertextwidth")
		if len(cstr(tfiltertextwidth))>3 then tfiltertextwidth=62
		if clng(tfiltertextwidth)>255 then tfiltertextwidth=255
		if clng(tfiltertextwidth)<1 then tfiltertextwidth=62
	
		treplytextwidth=Request.Form("replytextwidth")
		if len(cstr(treplytextwidth))>3 then treplytextwidth=62
		if clng(treplytextwidth)>255 then treplytextwidth=255
		if clng(treplytextwidth)<1 then treplytextwidth=62
	
		treplytextheight=Request.Form("replytextheight")
		if len(cstr(treplytextheight))>3 then treplytextheight=62
		if clng(treplytextheight)>255 then treplytextheight=255
		if clng(treplytextheight)<1 then treplytextheight=62

		titemsperpage=Request.Form("itemsperpage")
		if len(cstr(titemsperpage))>5 then titemsperpage=5
		if clng(titemsperpage)>32767 then titemsperpage=32767
		if clng(titemsperpage)<1 then titemsperpage=5

		ttitlesperpage=Request.Form("titlesperpage")
		if len(cstr(ttitlesperpage))>5 then ttitlesperpage=5
		if clng(ttitlesperpage)>32767 then ttitlesperpage=32767
		if clng(ttitlesperpage)<1 then ttitlesperpage=5

	end if
	
	if clng(showpage and 8)<> 0	then
		tvisualflag=0
		if Request.Form("replyinword")="1" then tvisualflag=tvisualflag+1
		select case Request.Form("showpagelist") : case "1","2","3"
			tvisualflag=tvisualflag+clng(Request.Form("showpagelist"))*4				'左移2位后累加
		end select
		if Request.Form("advpagelist")="1" then tvisualflag=tvisualflag+64
		if Request.Form("displaymode")="1" then tvisualflag=tvisualflag+1024

		tadvpagelistcount=Request.Form("advpagelistcount")
		if len(cstr(tadvpagelistcount))>3 then
			tadvpagelistcount=10
		else
			tadvpagelistcount=clng(tadvpagelistcount)
		end if
		if tadvpagelistcount>255 then tadvpagelistcount=255
		if tadvpagelistcount<1 then tadvpagelistcount=1

		tubbflag=0
		if Request.Form("ubbflag_image")="1" then tubbflag=tubbflag+1
		if Request.Form("ubbflag_url")="1" then tubbflag=tubbflag+2
		if Request.Form("ubbflag_autourl")="1" then tubbflag=tubbflag+4
		if Request.Form("ubbflag_player")="1" then tubbflag=tubbflag+8
		if Request.Form("ubbflag_paragraph")="1" then tubbflag=tubbflag+16
		if Request.Form("ubbflag_fontstyle")="1" then tubbflag=tubbflag+32
		if Request.Form("ubbflag_fontcolor")="1" then tubbflag=tubbflag+64
		if Request.Form("ubbflag_alignment")="1" then tubbflag=tubbflag+128
		'if Request.Form("ubbflag_movement")="1" then tubbflag=tubbflag+256
		'if Request.Form("ubbflag_cssfilter")="1" then tubbflag=tubbflag+512
		if Request.Form("ubbflag_face")="1" then tubbflag=tubbflag+1024

		ttablealign=Request.Form("tablealign")
		if ttablealign<>"left" and ttablealign<>"center" and ttablealign<>"right" then ttablealign="left"
		
		tdelconfirm=0
		if Request.Form("deltip")="1" then tdelconfirm=tdelconfirm+1
		if Request.Form("delretip")="1" then tdelconfirm=tdelconfirm+2
		if Request.Form("delseltip")="1" then tdelconfirm=tdelconfirm+4
		if Request.Form("deladvtip")="1" then tdelconfirm=tdelconfirm+8
		if Request.Form("deldectip")="1" then tdelconfirm=tdelconfirm+16
		if Request.Form("delseldectip")="1" then tdelconfirm=tdelconfirm+32

		tstylename=server.HTMLEncode(Request.Form("style"))
		tstylename=replace(tstylename," ","")
		
	end if

	set cn1=server.CreateObject("ADODB.Connection")
	set rs1=server.CreateObject("ADODB.Recordset")

	CreateConn cn1,dbtype

	rs1.open Replace(sql_adminsaveconfig,"{0}",wm_name),cn1,0,3,1

	if clng(showpage and 1)<> 0	then
		rs1("status")=tstatus
		rs1("adminhtml")=tadminhtml
		rs1("guesthtml")=tguesthtml
		rs1("admintimeout")=tadmintimeout
		rs1("showip")=tshowip
		rs1("adminshowip")=tadminshowip
		rs1("adminshoworiginalip")=tadminshoworiginalip
		rs1("vcodecount")=tvcodecount + twritevcodecount
	end if
	if clng(showpage and 2)<> 0	then
		rs1("mailflag")=tmailflag
	end if
	if clng(showpage and 4)<> 0	then
		rs1("cssfontfamily")=tcssfontfamily
		rs1("cssfontsize")=tcssfontsize
		rs1("csslineheight")=tcsslineheight
		rs1("tablewidth")=ttablewidth
		rs1("windowspace")=twindowspace
		rs1("tableleftwidth")=ttableleftwidth
		rs1("ubbtoolwidth")=tubbtoolwidth
		rs1("ubbtoolheight")=tubbtoolheight
		rs1("searchtextwidth")=tsearchtextwidth
		rs1("advdeltextwidth")=tadvdeltextwidth
		rs1("setinfotextwidth")=tsetinfotextwidth		
		rs1("filtertextwidth")=tfiltertextwidth
		rs1("replytextwidth")=treplytextwidth
		rs1("replytextheight")=treplytextheight
		rs1("itemsperpage")=titemsperpage
		rs1("titlesperpage")=ttitlesperpage
	end if
	if clng(showpage and 8)<> 0	then
		rs1("visualflag")=tvisualflag
		rs1("advpagelistcount")=tadvpagelistcount
		rs1("ubbflag")=tubbflag
		rs1("tablealign")=ttablealign
		rs1("delconfirm")=tdelconfirm
		rs1("stylename")=tstylename
	end if
	rs1.Update
	
	rs1.Close
	cn1.close
	set rs1=nothing
	set cn1=nothing
end if

Response.Redirect "web_config.asp?page=" & showpage
%>