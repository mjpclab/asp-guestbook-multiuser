<!-- #include file="webconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

dim re
set re=new RegExp
re.Pattern="^\w+$"

ruser=Request.QueryString("user")
if ruser="" then
	Response.Write "用户名不能为空。"
elseif re.Test(ruser)=false then
	Response.Write "用户名中只能包含英文字母、数字和下划线。"
else
	dim cn,rs
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype

	'=============================存在性验证
	rs.Open Replace(sql_checkuser,"{0}",ruser),cn,,,1
	if not rs.EOF then
		Response.Write "用户名： " & ruser & " 已存在，请换用其它用户名。"
	else
		Response.Write "用户名： " & ruser & " 可以使用。"
	end if
end if
%>