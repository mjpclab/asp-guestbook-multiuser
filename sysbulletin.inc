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
<table border="1" cellpadding="2" class="onetopic" style="overflow:hidden;">
	<tr style="height:25px">
		<td class="bulletintitle"><img alt="" src="image/icon_bulletin.gif" class="imgicon" />留言本系统公告：</td>
	</tr>
	<tr>
		<td class="wordscontent"><%=pub%></td>
	</tr>
</table>
<%end if%>