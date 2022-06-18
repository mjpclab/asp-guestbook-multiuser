<%
if IsAccess then
	sql_websearchdec_words_count="SELECT COUNT(adminname) FROM " &table_supervisor& " WHERE adminname LIKE '%{0}%' AND [declare] LIKE '%{1}%' AND Len([declare])>0"
	sql_websearchdec_words_query="SELECT adminname,declareflag,[declare] FROM " &table_supervisor& " WHERE adminname LIKE '%{0}%' AND [declare] LIKE '%{1}%' AND Len([declare])>0"
elseif IsSqlServer then
	sql_websearchdec_words_count="SELECT COUNT(adminname) FROM " &table_supervisor& " WHERE adminname LIKE N'%{0}%' AND [declare] LIKE N'%{1}%' AND [declare] LIKE N'_%'"
	sql_websearchdec_words_query="SELECT adminname,declareflag,[declare] FROM " &table_supervisor& " WHERE adminname LIKE N'%{0}%' AND [declare] LIKE N'%{1}%' AND [declare] LIKE N'_%'"
end if
%>
