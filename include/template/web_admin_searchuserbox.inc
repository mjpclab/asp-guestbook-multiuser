<script type="text/javascript">
function searchsubmitcheck()
{
	if (form1.searchtxt.value.length===0) {alert('�������������ݣ�');form1.searchtxt.focus();return false;}
	form1.searchsubmit.disabled=true;
	return true;
}
</script>

<div class="region form-region center-region">
	<h3 class="title">[�����û�]</h3>
	<div class="content">
		<form method="post" action="web_searchuser.asp" name="form1" onsubmit="return searchsubmitcheck();">
			������<input type="text" name="searchtxt" size="<%=SearchTextWidth%>" value="<%=request("searchtxt")%>" />
			<input type="submit" value="����" name="searchsubmit" />
			<select name="type" size="1" onchange="searchtxt.focus();">
				<option value="adminname"<%=seled(request("type")<>"adminname")%>>���û�������</option>
				<option value="name"<%=seled(request("type")="name")%>>���ǳ�����</option>
			</select>
			<br/>("%"����������ַ���"_"����һ���ַ�)
		</form>
	</div>
</div>