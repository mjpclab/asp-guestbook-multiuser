<%
sql_submitreg_checkuser="SELECT 1 FROM " &table_supervisor& " WHERE adminname='{0}'"
if IsAccess then
	sql_submitreg_init1=	"INSERT INTO " &table_supervisor& "(adminname,adminpass,name,faceid,regdate,lastlogin,question,[key],[declare]) VALUES('{0}','{1}','{2}',{3},'{4}','{5}','{6}','{7}','{8}')"
	sql_submitreg_init2=	"INSERT INTO " &table_config& "(adminid,status,adminhtml,guesthtml,admintimeout,showip,adminshowip,adminshoworiginalip,displaytimezoneoffset,visualflag,ubbflag,cssfontfamily,cssfontsize,csslineheight,styleid) SELECT TOP 1 adminid,{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},'{11}','{12}','{13}',{14} FROM " &table_supervisor& " WHERE adminname='{0}'"
elseif IsSqlServer then
	sql_submitreg_init1=	"INSERT INTO " &table_supervisor& "(adminname,adminpass,name,faceid,regdate,lastlogin,question,[key],[declare]) VALUES('{0}','{1}',N'{2}',{3},'{4}','{5}',N'{6}','{7}','{8}')"
	sql_submitreg_init2=	"INSERT INTO " &table_config& "(adminid,status,adminhtml,guesthtml,admintimeout,showip,adminshowip,adminshoworiginalip,displaytimezoneoffset,visualflag,ubbflag,cssfontfamily,cssfontsize,csslineheight,styleid) SELECT TOP 1 adminid,{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},N'{11}','{12}','{13}',{14} FROM " &table_supervisor& " WHERE adminname='{0}'"
end if
sql_submitreg_init3=	"INSERT INTO " &table_floodconfig& "(adminid) SELECT TOP 1 adminid FROM " &table_supervisor& " WHERE adminname='{0}'"
%>
