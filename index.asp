<!-- #include file="loadconfig.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"

if web_checkIsBannedIP then
	Response.Redirect "web_err.asp?number=4"
	Response.End
elseif checkIsBannedIP then
	Response.Redirect "err.asp?user=" &ruser& "&number=1"
	Response.End
elseif StatusOpen=false then
	Response.Redirect "err.asp?user=" &ruser& "&number=2"
	Response.End
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
checkuser cn,rs,false
if StatusStatistics then call addstat("view")

Dim WordsPerPage
if GuestDisplayMode()="book" then
	WordsPerPage=ItemsPerPage
elseif GuestDisplayMode()="forum" then
	WordsPerPage=TitlesPerPage
end if

Dim ItemsCount,PagesCount,CurrentItemsCount,ipage

dim local_sql_count,local_sql_query
local_sql_count=sql_index_words_count & GetHiddenWordCondition()
local_sql_query=sql_index_words_query & GetHiddenWordCondition()
get_divided_page cn,rs,sql_pk_main,Replace(local_sql_count,"{0}",adminid),Replace(local_sql_query,"{0}",adminid),"parent_id INC,lastupdated DEC,id DEC",Request.QueryString("page"),WordsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա�</title>
	<!-- #include file="inc_stylesheet.asp" -->
	</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">
	<%if ShowTitle=true then show_book_title 2,""%>

	<%RPage="index.asp"%><!-- #include file="func_guest.inc" -->
	
	<%dim sys_bul_flag
	sys_bul_flag=64%>
	<!-- #include file="sysbulletin.inc" -->
	<!-- #include file="topbulletin.inc" -->
	<!-- #include file="hidetip.inc" -->
	<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"index.asp","[���Է�ҳ]",""%>
	<%if ItemsCount>0 and StatusSearch and ShowTopSearchBox then%><!-- #include file="searchbox_guest.inc" --><%end if%>

	<%
	if ItemsCount=0 then
		Response.Write "<br/><br/><div class=""centertext"">Ŀǰ�������ԣ�������ǩд���ԡ���</div><br/><br/>"
	else
		dim pagename
		pagename="index"
		if GuestDisplayMode()="book" then
			%><!-- #include file="listword_guest.inc" --><%
		elseif GuestDisplayMode()="forum" then
			%><!-- #include file="listtitle_guest.inc" --><%
		end if
		rs.Close
	end if
	cn.Close : set rs=nothing : set cn=nothing%>
	
	<!-- #include file="func_guest.inc" -->

	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"index.asp","[���Է�ҳ]",""%>
	<%if ItemsCount>0 and StatusSearch and ShowBottomSearchBox then%><!-- #include file="searchbox_guest.inc" --><%end if%>

</div>

<!-- #include file="bottom.asp" -->
<%if StatusStatistics then%><script type="text/javascript" src="getclientinfo.asp?user=<%=ruser%>" defer="defer" async="async"></script><%end if%>
</body>
</html>