<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->

<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=web_BookName%> Webmaster�������� ��������</title>
	<!-- #include file="style.asp" -->
</head>

<body>

<%
function filterstr(byref str,byval delquote)
	if delquote=true then
		filterstr=replace(replace(replace(str,"'",""),"[","[\[]"),chr(34),"")
	else
		filterstr=replace(replace(replace(str,"'","''"),"[","[\[]"),chr(34),"")
	end if
end function


'<ShowArticle>
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
'====================================

	'=============��ҳ����=============
	Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
	ItemsCount=0:PagesCount=0
	
	dim count_sql,sql
	count_sql=Replace(Replace(sql_websearchdec_words_count,"{0}",filterstr(Request("adminname"),true)),"{1}",filterstr(Request("searchtxt"),false))
	sql=Replace(Replace(sql_websearchdec_words_query,"{0}",filterstr(Request("adminname"),true)),"{1}",filterstr(Request("searchtxt"),false))
	
	get_divided_page cn,rs,sql_pk_supervisor,count_sql,sql,"adminname INC",Request.QueryString("page"),ItemsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
	'==============================
%>

<div id="outerborder" class="outerborder">

	<!-- #include file="web_admintitle.inc" -->
	<!-- #include file="web_admintool.inc" -->

	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">�����ö�����</td>
		</tr>
		<tr>
			<td class="wordscontent" style="padding:20px 2px;">
			<form method="post" action="web_searchdec.asp" name="form1">
			������("%"����������ַ���"_"����һ���ַ�)<br/><br/>
			�û�������<input type="text" name="adminname" value="<%=Request("adminname")%>" size="<%=SearchTextWidth%>" /><br/>
			�������ݣ�<input type="text" name="searchtxt" value="<%=Request("searchtxt")%>" size="<%=SearchTextWidth%>" /><br/>
			<br/><input type="submit" value="��������" name="searchsubmit" />
			</form>
			</td>
		</tr>
	</table>

<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"web_searchdec.asp","[���������ҳ]","center","adminname=" &server.URLEncode(Request("adminname"))& "&searchtxt=" & server.URLEncode(Request("searchtxt")) end if%>
<form method="post" action="web_mdeldec.asp" name="form7">

<input type="hidden" name="adminname" value="<%=request("adminname")%>" />
<input type="hidden" name="searchtxt" value="<%=request("searchtxt")%>" />
<input type="hidden" name="page" value="<%if isnumeric(request("page")) and request("page")<>"" then response.write request("page")%>" />
<span class="generalwindow noborder" style="text-align:left; display:block;"><a href="javascript:for(var i=0;i<=form7.elements.length-1;i++)if(form7.elements[i].name=='users' && form7.elements[i].checked){<%if DelSelDecTip=true then Response.Write "if (confirm('ȷʵҪɾ��ѡ��������')==true)"%>form7.submit();break;}else if(i==form7.elements.length-1)alert('����ѡ��Ҫɾ���Ĺ��档');" onmouseover="return true;"><img src="image/icon_mdel.gif" style="border-width:0px;" />ɾ��ѡ������</a></span>

<%
dim pub_flag,pub
if ItemsCount=0 then
	Response.Write "<br/><br/><div class=""centertext"">û���ҵ����������Ĺ��档</div><br/><br/>"
else
for i=1 to ItemsPerPage
if rs.eof=false and Response.IsClientConnected then
%>

<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="0" cellspacing="0" class="onetopic">
  <tr>
	<td style="width:100%; height:25px; text-align:left; vertical-align:middle; font-weight:bold; color:<%=TableTitleColor%>; background-color:<%=TableTitleBGC%>; background-image:url(<%=TableTitlePic%>);">�����û���<a href="web_userinfo.asp?user=<%=rs("adminname")%>" target="_blank"><%=rs("adminname")%></a></td>
  </tr>
  <tr>
	<td class="wordscontent">
		<%
		pub_flag=rs("declareflag")
		pub="" &rs("declare")& ""

		if AdminViewCode=true then		'Only view HTML code
			pub=server.HTMLEncode(pub)
		else
			convertstr pub,pub_flag,2
		end if

		Response.Write pub
		%>
	</td>
  </tr>
  <tr style="height:25px;">
	<td style="width:100%; color:<%=TableContentColor%>; background-color:<%=TableContentBGC%>;">
		<table cellpadding="0" cellspacing="0" style="border-width:0px;">
			<tr>
				<td style="width:<%=TableLeftWidth%>px; border-right-width:0px;">
					<input type="checkbox" name="users" id="c<%=i%>" value="<%=rs("adminname")%>"><label for="c<%=i%>">(ѡ��)</label>
				</td>
				<td style="border-left-width:0px;">
					<a href="web_deldec.asp?user=<%=rs("adminname")%><%if isnumeric(request("page")) and request("page")<>"" then response.write "&page=" & request("page")%>&adminname=<%=server.URLEncode(request("adminname"))%>&searchtxt=<%=server.URLEncode(request("searchtxt"))%>" title="ɾ������"<%if DelDecTip=true then Response.Write " onclick=""return confirm('ȷʵҪɾ��������');"""%>><img border="0" src="image/icon_del.gif" class="imgicon" />[ɾ������]</a> 
				</td>
			</tr>
		</table>
	</td>
  </tr>
</table>

<%
		rs.MoveNext
		else
			exit for
		end if
	next
end if	'��Ӧfor����һ�е�if
%>
<span class="generalwindow noborder" style="text-align:left; display:block;"><a href="javascript:for(var i=0;i<=form7.elements.length-1;i++)if(form7.elements[i].name=='users' && form7.elements[i].checked){<%if DelSelDecTip=true then Response.Write "if (confirm('ȷʵҪɾ��ѡ��������')==true)"%>form7.submit();break;}else if(i==form7.elements.length-1)alert('����ѡ��Ҫɾ���Ĺ��档');" onmouseover="return true;"><img src="image/icon_mdel.gif" style="border-width:0px;">ɾ��ѡ������</a></span>
</form>

<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"web_searchdec.asp","[���������ҳ]","center","adminname=" &server.URLEncode(Request("adminname"))& "&searchtxt=" & server.URLEncode(Request("searchtxt")) end if%>
</div>

<%
	if ItemsCount<>0 then rs.Close
	cn.Close
	set rs=nothing
	set cn=nothing
%>


<!-- #include file="bottom.asp" -->
</body>
</html>