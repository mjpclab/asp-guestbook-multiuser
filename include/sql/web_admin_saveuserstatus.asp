<%
if IsAccess then
	sql_web_enableuser="UPDATE " &table_config& " SET status=status-1073741824 WHERE status>=1073741824 AND adminid={0}"
	sql_web_disableuser="UPDATE " &table_config& " SET status=status+1073741824 WHERE status<1073741824 AND adminid={0}"
	sql_web_enableuserlogin="UPDATE " &table_config& " SET status=status-536870912 WHERE (status mod 1073741824)\536870912<>0 AND adminid={0}"
	sql_web_disableuserlogin="UPDATE " &table_config& " SET status=status+536870912 WHERE (status mod 1073741824)\536870912=0 AND adminid={0}"
	sql_web_enableuserleaveword="UPDATE " &table_config& " SET status=status-268435456 WHERE (status mod 536870912)\268435456<>0 AND adminid={0}"
	sql_web_disableuserleaveword="UPDATE " &table_config& " SET status=status+268435456 WHERE (status mod 536870912)\268435456=0 AND adminid={0}"
elseif IsSqlServer then
	sql_web_enableuser="UPDATE " &table_config& " SET status=status-1073741824 WHERE status & 1073741824<>0 AND adminid={0}"
	sql_web_disableuser="UPDATE " &table_config& " SET status=status+1073741824 WHERE status & 1073741824=0 AND adminid={0}"
	sql_web_enableuserlogin="UPDATE " &table_config& " SET status=status-536870912 WHERE status & 536870912<>0 AND adminid={0}"
	sql_web_disableuserlogin="UPDATE " &table_config& " SET status=status+536870912 WHERE status & 536870912=0 AND adminid={0}"
	sql_web_enableuserleaveword="UPDATE " &table_config& " SET status=status-268435456 WHERE status & 268435456<>0 AND adminid={0}"
	sql_web_disableuserleaveword="UPDATE " &table_config& " SET status=status+268435456 WHERE status & 268435456=0 AND adminid={0}"
end if
%>