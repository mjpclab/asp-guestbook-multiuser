<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/web_admin_mdeldec.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1
if Request.Form("users")="" then
	Response.Redirect "web_searchdec.asp"
	Response.End
end if

dim users,affected
affected=0

users=FilterSql(Request.Form("users"))
if Len(users)>0 then
	users="'" & Replace(users,",","','") & "'"

	set cn=server.CreateObject("ADODB.Connection")
	Call CreateConn(cn)
	
	cn.execute Replace(sql_webmdeldec,"{0}",users),affected,1
	
	cn.Close
	set cn=nothing
end if

Call TipsPage("已删除" &affected& "条公告。",Replace(Replace(Replace("web_searchdec.asp?adminname={0}&searchtxt={1}&page={2}","{0}",server.URLEncode(Request.Form("adminname"))),"{1}",server.URLEncode(request.Form("searchtxt"))),"{2}",request.Form("page")))
%>