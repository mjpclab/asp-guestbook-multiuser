<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/write.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/mail.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="tips.asp" -->
<%
'======================================================
sub wordsbaned
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	if StatusStatistics then call addstat("banned")
	Response.Redirect "err.asp?user=" &ruser& "&number=4"
	Response.End
end sub
'======================================================
sub floodbaned
	rs.Close : rs2.Close : cn.Close : set rs=nothing : set rs2=nothing : set cn=nothing
	if StatusStatistics then call addstat("banned")
	Response.Redirect "err.asp?user=" &ruser& "&number=7"
	Response.End
end sub
'======================================================
sub bancheck(byref re,byref field,byval tfiltermode,byval bitflag)
	if clng(tfiltermode and bitflag)<>0 then
		if re.Test(field)=true then Call wordsbaned
	end if
end sub
'======================================================
sub filtercheck(byref re,byref field,byval tfiltermode,byval bitflag,byref treplacestr,byref filtered)
	if clng(tfiltermode and bitflag)<>0 then
		if re.Test(field)=true then
			if clng(tfiltermode and 4096)<>0 then	'wait to audit
				guestflag=clng(guestflag or 16)
			else
				filtered=true
				field=re.Replace(field,treplacestr)
			end if
		end if
	end if
end sub
'======================================================

Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif checkIsBannedIP then
	Response.Redirect "err.asp?user=" &ruser& "&number=1"
	Response.End
elseif StatusOpen=false then
	Response.Redirect "err.asp?user=" &ruser& "&number=2"
	Response.End
elseif StatusWrite=false then
	Response.Redirect "err.asp?user=" &ruser& "&number=3"
	Response.End
elseif StatusUserBanned then
	Response.Redirect "err.asp?user=" &ruser& "&number=100"
	Response.End
elseif (flood_minwait>0 or web_flood_minwait>0) and isdate(session.Contents("wrote_time")) then
	if datediff("s",session.Contents("wrote_time"),now())<=flood_minwait or datediff("s",session.Contents("wrote_time"),now())<=web_flood_minwait then
		if StatusStatistics then call addstat("banned")
		Response.Redirect "err.asp?user=" &ruser& "&number=6"
		Response.End
	end if
elseif (flood_minwait>0 or web_flood_minwait>0) and isdate(Request.Cookies("wrote_time")) then
	if datediff("s",Request.Cookies("wrote_time"),now())<=flood_minwait or datediff("s",Request.Cookies("wrote_time"),now())<=web_flood_minwait then
		if StatusStatistics then call addstat("banned")
		Response.Redirect "err.asp?user=" &ruser& "&number=6"
		Response.End
	end if
end if
if Request.Form("iname")="" or Request.Form("ititle")="" then Response.Redirect("face.asp")

Session(InstanceName & "_ititle_" & ruser)=Request.Form("ititle")
Session(InstanceName & "_icontent_" & ruser)=Request.Form("icontent")

if WriteVcodeCount>0 and (Request.Form("ivcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	Call TipsPage("��֤�����","leaveword.asp?" & Request.Form("qstr"))
	Response.End
else
	session("vcode")=""
end if
'===================================================================

SetTimelessCookie "iname",Request.form("iname")
SetTimelessCookie "imail",Request.form("imail")
SetTimelessCookie "iqq",Request.form("iqq")
SetTimelessCookie "imsn",Request.form("imsn")
SetTimelessCookie "ihomepage",Request.form("ihomepage")
SetTimelessCookie "ihead",Request.form("ihead")

'===================================================================

Dim cn,rs,sql

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
set rs2=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

dim content1,name1,title1,email1,qqid1,msnid1,homepage1,ipv4addr1,ipv6addr1,originalipv41,originalipv61,head1,guestflag,whisperpwd
dim tmpAddr
'---------------------
name1=Request.form("iname")
title1=request.form("ititle")
email1=request.form("imail")
qqid1=request.form("iqq")
msnid1=request.form("imsn")
homepage1=request.form("ihomepage")

tmpAddr=CStr(request.ServerVariables("REMOTE_ADDR"))
if IsIPv4(tmpAddr) then
	ipv4addr1=Left(tmpAddr,15)
elseif IsIPv6(tmpAddr) then
	ipv6addr1=expandIPv6(tmpAddr,false)
end if
tmpAddr=CStr(request.ServerVariables("HTTP_X_FORWARDED_FOR"))
if IsIPv4(tmpAddr) then
	originalipv41=Left(tmpAddr,15)
elseif IsIPv6(tmpAddr) then
	originalipv61=expandIPv6(tmpAddr,false)
end if
face1=request.form("ihead")

content1=request.form("icontent")
if WordsLimit<>0 and len(content1)>WordsLimit then content1=left(content1,WordsLimit)

guestflag=guestlimit AND web_guestlimit
whisperpwd=""
if StatusNeedAudit=true then guestflag=guestflag OR 16
if StatusWhisper=true and Request.Form("chk_whisper")="1" then guestflag=guestflag OR 32
if StatusEncryptWhisper=true and Request.Form("chk_encryptwhisper")="1" and Request.Form("iwhisperpwd")<>"" then
	guestflag=guestflag OR 64
	whisperpwd=md5(Request.Form("iwhisperpwd"),32)
end if
if Request.Form("imailreplyinform")="1" then guestflag=guestflag OR 128
if Request.Form("hidecontact")="1" then guestflag=guestflag OR 256
'-------------------------
dim tregexp,tfiltermode,treplacestr,re,filtered
set re=new RegExp
filtered=false

for i=1 to 2
	if i=1 then
		rs.Open Replace(sql_write_filter,"{0}",wm_id),cn,0,1,1
	elseif i=2 then
		rs.Open Replace(sql_write_filter,"{0}",adminid),cn,0,1,1
	end if

	while rs.EOF=false
		tregexp=rs("regexp")
		tfiltermode=rs("filtermode")
		treplacestr=rs("replacestr")

		if tregexp<>"" then
			re.Multiline=(clng(tfiltermode and 2048)<>0)
			re.IgnoreCase=(clng(tfiltermode and 8192)=0)
			re.Global=true
			re.Pattern=tregexp

			if clng(tfiltermode and 16384)<>0 then    'baned
				bancheck re,name1,tfiltermode,1
				bancheck re,email1,tfiltermode,2
				bancheck re,qqid1,tfiltermode,4
				bancheck re,msnid1,tfiltermode,8
				bancheck re,homepage1,tfiltermode,16
				bancheck re,title1,tfiltermode,32
				bancheck re,content1,tfiltermode,64
			else
				filtercheck re,name1,tfiltermode,1,treplacestr,filtered
				filtercheck re,email1,tfiltermode,2,treplacestr,filtered
				filtercheck re,qqid1,tfiltermode,4,treplacestr,filtered
				filtercheck re,msnid1,tfiltermode,8,treplacestr,filtered
				filtercheck re,homepage1,tfiltermode,16,treplacestr,filtered
				filtercheck re,title1,tfiltermode,32,treplacestr,filtered
				filtercheck re,content1,tfiltermode,64,treplacestr,filtered
			end if
		end if
		rs.MoveNext
	wend

	rs.Close
next

set re=nothing
if filtered=true and StatusStatistics then addstat("filtered")
'-------------------------
if name1="" or title1="" then Response.Redirect "index.asp?user=" & ruser

name1=server.htmlEncode(name1)
name1=textfilter(name1,false)
name1=left(name1,20)

title1=server.htmlEncode(title1)
title1=textfilter(title1,false)
title1=left(title1,30)

email1=replace(server.htmlEncode(email1)," ","")
email1=textfilter(email1,true)
email1=left(email1,50)

qqid1=replace(server.htmlEncode(qqid1)," ","")
qqid1=textfilter(qqid1,false)
qqid1=left(qqid1,16)

msnid1=replace(server.htmlEncode(msnid1)," ","")
msnid1=textfilter(msnid1,false)
msnid1=left(msnid1,50)

content1=replace(content1,"<%","< %")

logdate1=now()

homepage1=replace(server.HTMLEncode(homepage1)," ","")
if lcase(left(homepage1,7))<>"http://" and lcase(left(homepage1,6))<>"ftp://" and homepage1<>"" then homepage1="http://" & homepage1
homepage1=textfilter(homepage1,true)
homepage1=left(homepage1,127)

if isnumeric(face1)=false or face1="" then
	face1=0
elseif len(cstr(face1))>len(cstr(FaceCount))then
	face1=0
elseif clng(face1)>clng(FaceCount) then
	face1=0
else
	face1=abs(cbyte(face1))
end if


'web_flood check
if web_flood_searchrange>0 and (web_flood_sfnewword or web_flood_sfnewreply) and (web_flood_include or web_flood_equal) and (web_flood_sititle or web_flood_sicontent) then
	if isnumeric(Request.Form("follow")) and Request.Form("follow")<>"" and StatusGuestReply then
		rs2.Open sql_write_flood_ids & Request.Form("follow"),cn,,,1
	else
		rs2.Open sql_write_flood_idnull,cn,,,1
	end if

	if web_flood_sfnewword and rs2.EOF then
		sql=sql_write_flood_head

		if web_flood_sititle then
			if web_flood_include and title1<>"" then sql=sql & Replace(sql_write_flood_titlelike,"{0}",FilterSqlStr(title1))
			if web_flood_equal or title1="" then sql=sql & Replace(sql_write_flood_titleequal,"{0}",FilterSqlStr(title1))
			if web_flood_sicontent then sql=sql & " OR"
		end if

		if web_flood_sicontent then
			if web_flood_include and content1<>"" then sql=sql & Replace(sql_write_flood_articlelike,"{0}",FilterSqlStr(content1))
			if web_flood_equal or content1="" then sql=sql & Replace(sql_write_flood_articleequal,"{0}",FilterSqlStr(content1))
		end if
		
		sql=sql & Replace(Replace(sql_write_flood_wordstail,"{0}",web_flood_searchrange),"{1}",adminid)
		rs.Open sql,cn,,,1
		if not rs.EOF then floodbaned()
		rs.Close
	end if

	if web_flood_sfnewreply and (not rs2.EOF) then
		sql=sql_write_flood_head

		if web_flood_sititle and title1<>"Re:" then
			if web_flood_include and title1<>"" then sql=sql & Replace(sql_write_flood_titlelike,"{0}",FilterSqlStr(title1))
			if web_flood_equal or title1="" then sql=sql & Replace(sql_write_flood_titleequal,"{0}",FilterSqlStr(title1))
			if web_flood_sicontent then sql=sql & " OR"
		end if

		if web_flood_sicontent then
			if web_flood_include and content1<>"" then sql=sql & Replace(sql_write_flood_articlelike,"{0}",FilterSqlStr(content1))
			if web_flood_equal or content1="" then sql=sql & Replace(sql_write_flood_articleequal,"{0}",FilterSqlStr(content1))
		end if
		
		sql=sql & Replace(Replace(Replace(sql_write_flood_replytail,"{0}",web_flood_searchrange),"{1}",rs2.Fields("id")),"{2}",adminid)
		rs.Open sql,cn,,,1
		if not rs.EOF then floodbaned()
		rs.Close
	end if
	rs2.Close
end if
'flood check
if flood_searchrange>0 and (flood_sfnewword or flood_sfnewreply) and (flood_include or flood_equal) and (flood_sititle or flood_sicontent) then
	if isnumeric(Request.Form("follow")) and Request.Form("follow")<>"" and StatusGuestReply then
		rs2.Open sql_write_flood_ids & Request.Form("follow"),cn,,,1
	else
		rs2.Open sql_write_flood_idnull,cn,,,1
	end if

	if flood_sfnewword and rs2.EOF then
		sql=sql_write_flood_head

		if flood_sititle then
			if flood_include and title1<>"" then sql=sql & Replace(sql_write_flood_titlelike,"{0}",FilterSqlStr(title1))
			if flood_equal or title1="" then sql=sql & Replace(sql_write_flood_titleequal,"{0}",FilterSqlStr(title1))
			if flood_sicontent then sql=sql & " OR"
		end if

		if flood_sicontent then
			if flood_include and content1<>"" then sql=sql & Replace(sql_write_flood_articlelike,"{0}",FilterSqlStr(content1))
			if flood_equal or content1="" then sql=sql & Replace(sql_write_flood_articleequal,"{0}",FilterSqlStr(content1))
		end if
		
		sql=sql & Replace(Replace(sql_write_flood_wordstail,"{0}",flood_searchrange),"{1}",adminid)
		rs.Open sql,cn,,,1
		if not rs.EOF then floodbaned()
		rs.Close
	end if

	if flood_sfnewreply and (not rs2.EOF) then
		sql=sql_write_flood_head

		if flood_sititle and title1<>"Re:" then
			if flood_include and title1<>"" then sql=sql & Replace(sql_write_flood_titlelike,"{0}",FilterSqlStr(title1))
			if flood_equal or title1="" then sql=sql & Replace(sql_write_flood_titleequal,"{0}",FilterSqlStr(title1))
			if flood_sicontent then sql=sql & " OR"
		end if

		if flood_sicontent then
			if flood_include and content1<>"" then sql=sql & Replace(sql_write_flood_articlelike,"{0}",FilterSqlStr(content1))
			if flood_equal or content1="" then sql=sql & Replace(sql_write_flood_articleequal,"{0}",FilterSqlStr(content1))
		end if
		
		sql=sql & Replace(Replace(Replace(sql_write_flood_replytail,"{0}",flood_searchrange),"{1}",rs2.Fields("id")),"{2}",adminid)
		rs.Open sql,cn,,,1
		if not rs.EOF then floodbaned()
		rs.Close
	end if
	rs2.Close
end if

'------------------------
rs.Open sql_write_idnull,cn,0,3,1
rs.AddNew
rs("adminid")=adminid
rs("name")=name1
rs("title")=title1
rs("email")=email1
rs("qqid")=qqid1
rs("msnid")=msnid1
rs("homepage")=homepage1
rs("logdate")=logdate1
rs("lastupdated")=logdate1
rs("ipv4addr")=ipv4addr1
rs("ipv6addr")=ipv6addr1
rs("originalipv4")=originalipv41
rs("originalipv6")=originalipv61
rs("faceid")=face1
rs("guestflag")=guestflag
rs("whisperpwd")=whisperpwd
rs("article")=content1
rs("parent_id")=0
if isnumeric(Request.Form("follow")) and Request.Form("follow")<>"" and StatusGuestReply then
	rs2.Open Replace(Replace(sql_write_verify_repliable,"{0}",Request.Form("follow")),"{1}",adminid),cn,0,1,1
	if not rs2.EOF then
		rs("parent_id")=Request.Form("follow")
		cn.Execute Replace(Replace(Replace(sql_write_updatelastupdated,"{0}",now()),"{1}",Request.Form("follow")),"{2}",adminid),,1
		cn.Execute Replace(Replace(sql_write_updateparentflag,"{0}",Request.Form("follow")),"{1}",adminid),,1
	end if
	rs2.Close : set rs2=nothing
end if

if IsAccess then	'update root_id info
	if rs("parent_id")<=0 then
		rs("root_id")=rs("id")
	else
		rs("root_id")=rs("parent_id")
	end if
end if
rs.Update

rs.Close : cn.Close : set rs=nothing : set cn=nothing

if StatusStatistics then call addstat("written")

SetTimelessCookie "wrote_time",now()
Session.Contents("wrote_time")=now()
Session(InstanceName & "_ititle_" & ruser)=""
Session(InstanceName & "_icontent_" & ruser)=""
Session.Contents("guestflag")=guestflag
Session.Contents("guestflag_user")=ruser

if web_MailNewInform=true and MailNewInform=true then newinform()

if Request.Form("return")="showword" and isnumeric(Request.Form("follow")) and Request.Form("follow")<>"" then
	Response.Redirect "showword.asp?user=" & ruser & "&id=" & Request.Form("follow")
else
	Response.Redirect "index.asp?user=" & ruser
end if
%>