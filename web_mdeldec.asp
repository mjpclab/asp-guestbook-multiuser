<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
if Request.Form("users")="" then
	Response.Redirect "web_searchdec.asp"
	Response.End
end if

dim arr,users,affected,i
affected=0

users=replace(Request.Form("users"),"'","")
users=replace(users,";","")
users=replace(users,"#","")
users=replace(users,"%","")
users=replace(users,"&","")
users=replace(users,"=","")
users=replace(users," ","")
users=replace(users,"[","")
users=replace(users,"]","")
users=replace(users,"\","")

arr=split(users,",")
users=""
for i=0 to ubound(arr)
	if i>0 then
		users=users & ",'" & FilterSql(arr(i)) & "'"
	else
		users="'" & FilterSql(arr(i)) & "'"
	end if
next

if users<>"" then
	set cn=server.CreateObject("ADODB.Connection")
	CreateConn cn,dbtype
	
	cn.execute Replace(sql_webmdeldec,"{0}",users),affected,1
	
	cn.Close
	set cn=nothing
end if

Call MessagePage("已删除" &affected& "条公告。",Replace(Replace(Replace("web_searchdec.asp?adminname={0}&searchtxt={1}&page={2}","{0}",server.URLEncode(Request.Form("adminname"))),"{1}",server.URLEncode(request.Form("searchtxt"))),"{2}",request.Form("page")))
%>