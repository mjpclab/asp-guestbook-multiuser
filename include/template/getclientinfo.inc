<%if StatusStatistics then%>
<%if session("gotclientuser")="" then session("gotclientuser")="|"%>
<%if instr(session("gotclientuser"),"|" &ruser& "|")=0 then%>
<script type="text/javascript" src="asset/js/stats.js"></script>
<script type="text/javascript">
	var url=location.href.substring(0,location.href.lastIndexOf('/')+1);
	url+='saveclientinfo.asp?sys=' +getsysname() +
		'&brow=' + encodeURIComponent(getbrowsername()) +
		'&sw=' + encodeURIComponent(getscreenwidth()) +
		'&sh=' + encodeURIComponent(getscreenheight()) +
		'&src=' + encodeURIComponent(getsourceaddr()) +
		'&fsrc=' + encodeURIComponent(document.referrer) +
		'&user=' + encodeURIComponent('<%=Request.QueryString("user")%>');
	RequestScript(url);
</script>
<%end if%>
<%end if%>