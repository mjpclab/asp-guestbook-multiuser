<script type="text/javascript">
function searchsubmitcheck()
{
	if (form1.searchtxt.value.length===0) {alert('请输入搜索内容！');form1.searchtxt.focus();return false;}
	form1.searchsubmit.disabled=true;
	return true;
}
</script>

<div class="region form-region center-region">
	<h3 class="title">[搜索用户]</h3>
	<div class="content">
		<form method="post" action="web_searchuser.asp" name="form1" onsubmit="return searchsubmitcheck();">
			搜索：<input type="text" name="searchtxt" size="<%=SearchTextWidth%>" value="<%=request("searchtxt")%>" />
			<input type="submit" value="搜索" name="searchsubmit" />
			<select name="type" size="1" onchange="searchtxt.focus();">
				<option value="adminname"<%=seled(request("type")<>"adminname")%>>按用户名搜索</option>
				<option value="name"<%=seled(request("type")="name")%>>按昵称搜索</option>
			</select>
			<br/>("%"代表任意个字符，"_"代表一个字符)
		</form>
	</div>
</div>
