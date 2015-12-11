<%
sql_findkey4_checkuser="SELECT adminname,[key] FROM " &table_supervisor& " WHERE adminname='{0}'"
sql_findkey4_resetpass="UPDATE " &table_supervisor& " SET adminpass='{0}' WHERE adminname='{1}'"
%>
