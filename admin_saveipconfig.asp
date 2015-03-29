<!-- #include file="loadconfig.asp" -->
<link rel="stylesheet" type="text/css" href="css/style.css"/>
<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
<!-- #include file="css/style.asp" -->
<!-- #include file="css/adminstyle.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
checkuser cn,rs,true

'IPConStatus
Dim tipconstatus
tipconstatus=clng(Request.Form("ipv4constatus"))+clng(Request.Form("ipv6constatus"))*16
cn.Execute Replace(Replace(sql_adminsaveipconfig_update,"{0}",tipconstatus),"{1}",adminid),,1


function deleteSaved(requestField,sql)
	Dim inputIds,listid,listids
	inputIds=split(Request.Form(requestField),",")
	listids=""
	for each listid in inputIds
		if isnumeric(listid) then
			listids=listids & "," & clng(listid)
		end if
	next
	if Len(listids)>0 then
		listids=Mid(listids,2)
		cn.Execute Replace(Replace(sql,"{0}",listids),"{1}",adminid),,1
	end if
end function

'IPv4 delete status 1
deleteSaved "savedipv4status1",sql_adminsaveipv4config_delete

'IPv4 delete status 2
deleteSaved "savedipv4status2",sql_adminsaveipv4config_delete

'IPv6 delete status 1
deleteSaved "savedipv6status1",sql_adminsaveipv6config_delete

'IPv6 delete status 2
deleteSaved "savedipv6status2",sql_adminsaveipv6config_delete


function addNewIPv4(requestField,sql)
	Dim entries,iprange,maxIndex,ipfrom,ipto

	entries=split(Request.Form(requestField),chr(13)&chr(10))
	for each iprange in entries
		ips=split(iprange,"-")
		ipfrom=""
		ipto=""
		maxIndex=ubound(ips)

		if maxIndex=0 then
			ipfrom=ips(0)
			ipto=ips(0)
		elseif maxIndex>=1 then
			ipfrom=ips(0)
			ipto=ips(1)
		end if

		if isIPv4Range(ipfrom) and isIPv4Range(ipto) then
			cn.Execute Replace(Replace(Replace(sql,"{0}",ipv4ToHex(ipfrom,false)),"{1}",ipv4ToHex(ipto,true)),"{2}",adminid),,1
		end if
	next
end function
function addNewIPv6(requestField,sql)
	Dim entries,iprange,maxIndex,ipfrom,ipto

	entries=split(Request.Form(requestField),chr(13)&chr(10))
	for each iprange in entries
		ips=split(iprange,"-")
		ipfrom=""
		ipto=""
		maxIndex=ubound(ips)

		if maxIndex=0 then
			ipfrom=ips(0)
			ipto=ips(0)
		elseif maxIndex>=1 then
			ipfrom=ips(0)
			ipto=ips(1)
		end if

		if isIPv6Range(ipfrom) and isIPv6Range(ipto) then
			cn.Execute Replace(Replace(Replace(sql,"{0}",ipv6ToHex(ipfrom,false)),"{1}",ipv6ToHex(ipto,true)),"{2}",adminid),,1
		end if
	next
end function

'IPv4 add status 1
addNewIPv4 "newipv4status1",sql_adminsaveipv4config_insert1

'IPv4 add status 2
addNewIPv4 "newipv4status2",sql_adminsaveipv4config_insert2

'IPv6 add status 1
addNewIPv6 "newipv6status1",sql_adminsaveipv6config_insert1

'IPv6 add status 2
addNewIPv6 "newipv6status2",sql_adminsaveipv6config_insert2

cn.Close : set cn=nothing
Response.Redirect "admin_ipconfig.asp?user=" & ruser & "&tabIndex=" & Request.Form("tabIndex")
%>