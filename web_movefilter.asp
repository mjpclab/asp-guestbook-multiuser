<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%

Response.Expires=-1

if isnumeric(Request.Form("filterid")) and Request.Form("filterid")<>"" and (Request.Form("movedirection")="up" or Request.Form("movedirection")="down" or Request.Form("movedirection")="top" or Request.Form("movedirection")="bottom") then
	filterid1=clng(Request.Form("filterid"))

	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype
	
	rs.Open Replace(Replace(sql_adminmovefilter_sort1,"{0}",filterid1),"{1}",wm_name),cn,,,1
	if rs.EOF=false then
		qid1=rs(0)
		rs.Close
	
		if Request.Form("movedirection")="up" then
			rs.Open Replace(Replace(sql_adminmovefilter_sort2_up,"{0}",qid1),"{1}",wm_name),cn,,,1
		elseif Request.Form("movedirection")="down" then
			rs.Open Replace(Replace(sql_adminmovefilter_sort2_down,"{0}",qid1),"{1}",wm_name),cn,,,1
		elseif Request.Form("movedirection")="top" then
			rs.Open Replace(sql_adminmovefilter_sort2_top,"{0}",wm_name),cn,,,1
		elseif Request.Form("movedirection")="bottom" then
			rs.Open Replace(sql_adminmovefilter_sort2_bottom,"{0}",wm_name),cn,,,1
		end if
		if rs.EOF=false then
			qid2=rs(0)
			rs.Close
			
			if qid2<>"" and qid1<>qid2 then
				rs.Open Replace(Replace(sql_adminmovefilter_id2,"{0}",qid2),"{1}",wm_name),cn,,,1
				filterid2=rs(0)
				rs.Close

				cn.BeginTrans
				
				select case Request.Form("movedirection")
				case "up","down"
					cn.Execute Replace(Replace(sql_adminmovefilter_update_updown1,"{0}",filterid1),"{1}",wm_name),,1
					cn.Execute Replace(Replace(Replace(sql_adminmovefilter_update_updown2,"{0}",qid1),"{1}",qid2),"{2}",wm_name),,1
					cn.Execute Replace(Replace(sql_adminmovefilter_update_updown3,"{0}",qid2),"{1}",wm_name),,1
				case "top"
					cn.Execute Replace(Replace(sql_adminmovefilter_update_top1,"{0}",qid1),"{1}",wm_name),,1
					cn.Execute Replace(Replace(Replace(sql_adminmovefilter_update_top2,"{0}",qid2),"{1}",filterid1),"{2}",wm_name),,1
				case "bottom"
					cn.Execute Replace(Replace(sql_adminmovefilter_update_bottom1,"{0}",qid1),"{1}",wm_name),,1
					cn.Execute Replace(Replace(Replace(sql_adminmovefilter_update_bottom2,"{0}",qid2),"{1}",filterid1),"{2}",wm_name),,1
				end select
									
				cn.CommitTrans
			end if
		else
			rs.Close
		end if
	else
		rs.Close
	end if
	cn.Close
	set rs=nothing
	set cn=nothing
end if

Response.Redirect "web_filter.asp"
%>