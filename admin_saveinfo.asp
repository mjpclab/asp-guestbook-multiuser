<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_saveinfo.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if Not IsEmpty(Request.Form) then
	tname=HtmlEncode(Request.Form("aname"))

	tfaceid=Request.Form("afaceid")
	if len(cstr(tfaceid))>3 or Request.Form("afaceurl")<>"" then
		tfaceid=0
	elseif isnumeric(tfaceid)=false or tfaceid="" then
		tfaceid=0
	elseif clng(tfaceid)<0 or clng(tfaceid)>255 then
		tfaceid=0
	else
		tfaceid=cbyte(tfaceid)
	end if

	tfaceurl=replace(HtmlEncode(Request.Form("afaceurl"))," ","")
	if tfaceurl<>"" then
		if InStr(tfaceurl,"://")=0 then tfaceurl="http://" & tfaceurl
		tfaceurl=Left(tfaceurl,127)
	end if

	temail=replace(HtmlEncode(Request.Form("aemail"))," ","")
	tqqid=replace(HtmlEncode(Request.Form("aqqid"))," ","")
	tmsnid=replace(HtmlEncode(Request.Form("amsnid"))," ","")

	thomepage=replace(HtmlEncode(Request.Form("ahomepage"))," ","")
	if thomepage<>"" then
		if InStr(thomepage,"://")=0 then thomepage="http://" & thomepage
		thomepage=Left(thomepage,127)
	end if

	set cn1=server.CreateObject("ADODB.Connection")
	set rs1=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn1)
	rs1.open Replace(sql_adminsaveinfo,"{0}",adminid),cn1,0,3,1

	rs1("name")=textfilter(tname,false)
	rs1("faceid")=tfaceid
	rs1("faceurl")=textfilter(tfaceurl,true)
	rs1("email")=textfilter(temail,true)
	rs1("qqid")=textfilter(tqqid,false)
	rs1("msnid")=textfilter(tmsnid,false)
	rs1("homepage")=textfilter(thomepage,true)
	rs1.Update

	rs1.Close : cn1.close : set rs1=nothing : set cn1=nothing
end if
Response.Redirect "admin.asp?user=" &ruser
%>