<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%

Response.Expires=-1

tfindexp=Request.Form("findexp")
if tfindexp<>"" then
	tfiltermode=0

	arr_findrange=split(Request.Form("findrange"),",")
	for fi=0 to ubound(arr_findrange)
		if len(arr_findrange(fi))<=5 then
			if isnumeric(arr_findrange(fi)) then
				if cint(arr_findrange(fi))<16384 then
					tfiltermode=tfiltermode+cint(arr_findrange(fi))
				end if
			end if
		end if
	next

	if Request.Form("multiline")="2048" then tfiltermode=tfiltermode+2048
	if Request.Form("matchcase")="8192" then tfiltermode=tfiltermode+8192
	if Request.Form("filtermethod")="16384" then
		tfiltermode=tfiltermode+16384
	elseif Request.Form("filtermethod")="4096" then
		tfiltermode=tfiltermode+4096
	end if

	if Request.Form("filtermethod")="0" then
		treplacestr=Request.Form("replacetxt")
	else
		treplacestr=""
	end if

	if len(Request.Form("memo"))>25 then
		tmemo=left(Request.Form("memo"),25)
	else
		tmemo="" & Request.Form("memo") & ""
	end if

	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype

	cn.BeginTrans
	
	rs.Open sql_adminfilter_null,cn,0,3,1
	rs.AddNew
	rs("adminname")=wm_name
	rs("regexp")=tfindexp
	rs("filtermode")=tfiltermode
	rs("replacestr")=treplacestr
	rs("memo")=tmemo
	rs.Update
	rs.Close

	rs.Open Replace(sql_adminfilter_max,"{0}",wm_name),cn,,,1
	pfilterid=rs(0)
	rs.Close
	
	cn.Execute Replace(Replace(sql_adminfilter_update,"{0}",pfilterid),"{1}",wm_name),,1

	cn.CommitTrans
	
	cn.Close
	set rs=nothing
	set cn=nothing
end if

Response.Redirect "web_filter.asp"
%>