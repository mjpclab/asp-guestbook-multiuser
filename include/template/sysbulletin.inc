<%set rs2=server.CreateObject("ADODB.Recordset")
rs2.open sql_sysbulletin,cn,0,1,1

pub_flag=rs2("bulletinflag")
if clng(pub_flag and sys_bul_flag)<>0 then
	pub="" &rs2("webbulletin")& ""
	convertstr pub,pub_flag,1
else
	pub=""
end if
rs2.close
set rs2=nothing

if pub<>"" then
%>
<div class="region">
	<h3 class="title"><img src="asset/image/icon_bulletin.gif"/>留言本系统公告：</h3>
	<div class="content"><%=pub%></div>
</div>
<%end if%>