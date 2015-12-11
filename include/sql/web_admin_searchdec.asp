<%
sql_websearchdec_words_count="SELECT COUNT(adminname) FROM " &table_supervisor& " WHERE adminname LIKE '%{0}%' AND [declare] LIKE '%{1}%' AND [declare] LIKE '_%'"
sql_websearchdec_words_query="SELECT adminname,declareflag,[declare] FROM " &table_supervisor& " WHERE adminname LIKE '%{0}%' AND [declare] LIKE '%{1}%' AND [declare] LIKE '_%'"
%>
