<%
sql_adminverify="SELECT adminpass FROM " &table_supervisor& " WHERE adminid={0}"
sql_updatelastlogin="UPDATE " &table_supervisor& " SET lastlogin='{0}' WHERE adminid={1}"
%>
