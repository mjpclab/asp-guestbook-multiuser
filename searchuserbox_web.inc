<script type="text/javascript">
//<![CDATA[

function searchsubmitcheck()
{
	if (form1.searchtxt.value=='') {alert('请输入搜索内容！');form1.searchtxt.focus();return false;}
	form1.searchsubmit.disabled=true;
	return true;
}

//]]>
</script>

		<table cellpadding="2" class="generalwindow">
			<tr>
				<td class="centertitle">[搜索用户]</td>
			</tr>
			<tr>
				<td class="wordscontent" style="text-align:center; padding:20px 0px;">
					<form method="post" action="web_searchuser.asp" name="form1" onsubmit="return searchsubmitcheck();">
					搜索：<input type="text" name="searchtxt" size="<%=SearchTextWidth%>" value="<%=request("searchtxt")%>" />
					<input type="submit" value="搜索" name="searchsubmit" />
					<select name="type" size="1" onchange="searchtxt.focus();">
						<option value="adminname"<%=seled(request("type")<>"adminname")%>>按用户名搜索</option>
						<option value="name"<%=seled(request("type")="name")%>>按昵称搜索</option>
					</select>
					<br/>("%"代表任意个字符，"_"代表一个字符)
					</form>
				</td>
			</tr>
		</table>
