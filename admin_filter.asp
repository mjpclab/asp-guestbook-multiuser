<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_filter.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/user.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="web_error.asp" -->
<%
Response.Expires=-1
if web_checkIsBannedIP() then
	Call WebErrorPage(4)
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 内容过滤策略</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="include/template/admin_mainmenu.inc" -->

	<div class="region form-region region-filter">
		<h3 class="title">内容过滤策略</h3>
		<div class="content">
			<form method="post" name="newfilter" action="admin_appendfilter.asp" onsubmit="if(findexp.value.length===0){alert('请输入查找内容。');findexp.focus();return false;}submit1.disabled=true;">
			<input type="hidden" name="user" value="<%=ruser%>" />
			<h4>添加新过滤策略：</h4>
			<p>查找内容(可用正则表达式,多个过滤词间用“|”分割)<br/>
			<input type="text" name="findexp" /><br/>
			<input type="checkbox" name="matchcase" id="matchcase" value="8192" /><label for="matchcase">区分大小写</label>
			<input type="checkbox" name="multiline" id="multiline" value="2048" /><label for="multiline">正则多行模式</label>
			</p>
			<p>查找范围<br/>
			<input type="checkbox" name="findrange" id="findname" value="1" checked="checked" /><label for="findname">称呼</label>
			<input type="checkbox" name="findrange" id="findmail" value="2" checked="checked" /><label for="findmail">邮件</label>
			<input type="checkbox" name="findrange" id="findqq" value="4" checked="checked" /><label for="findqq">QQ号</label>
			<input type="checkbox" name="findrange" id="findmsn" value="8" checked="checked" /><label for="findmsn">Skype</label>
			<input type="checkbox" name="findrange" id="findhome" value="16" checked="checked" /><label for="findhome">主页</label>
			<input type="checkbox" name="findrange" id="findtitle" value="32" checked="checked" /><label for="findtitle">标题</label>
			<input type="checkbox" name="findrange" id="findcontent" value="64" checked="checked" /><label for="findcontent">内容</label>
			</p>
			<p>处理方式<br/>
			<input type="radio" name="filtermethod" id="filtermethod" value="0" checked="checked" onclick="if(typeof(newfilter.replacetxt.disabled)!=='undefined')newfilter.replacetxt.disabled=false;" /><label for="filtermethod">替换为下面的文本</label>
			<input type="radio" name="filtermethod" id="filtermethod2" value="4096" onclick="if(typeof(newfilter.replacetxt.disabled)!=='undefined')newfilter.replacetxt.disabled=true;" /><label for="filtermethod2">等待审核</label>
			<input type="radio" name="filtermethod" id="filtermethod3" value="16384" onclick="if(typeof(newfilter.replacetxt.disabled)!=='undefined')newfilter.replacetxt.disabled=true;" /><label for="filtermethod3">拒绝留言</label><br/>
			<input type="text" name="replacetxt" />
			</p>
			<p>备注<br/>
			<input type="text" name="memo" maxlength="25" />
			</p>
			<div class="field-command"><input type="submit" value="添加过滤策略" name="submit1" /></div>
			</form>

			<%
			rs.Open Replace(sql_adminfilter,"{0}",adminid),cn,,,1

			while rs.EOF=false%>
			<form method="post" action="admin_updatefilter.asp" class="detail-item">
				<input type="hidden" name="user" value="<%=ruser%>" />
				<%tfilterid=rs("filterid")%>
				<input type="hidden" name="filterid" value="<%=tfilterid%>" />
				<%tfiltermode=clng(rs("filtermode"))%>
				<p>查找内容<br/>
				<input type="text" name="findexp" value="<%=rs("regexp")%>" /><br/>
				<input type="checkbox" name="matchcase" id="matchcase<%=tfilterid%>" value="8192"<%=cked(clng(tfiltermode and 8192)<>0)%> /><label for="matchcase<%=tfilterid%>">区分大小写</label>
				<input type="checkbox" name="multiline" id="multiline<%=tfilterid%>" value="2048"<%=cked(clng(tfiltermode and 2048)<>0)%> /><label for="multiline<%=tfilterid%>">正则多行模式</label>
				</p>
				<p>查找范围<br/>
				<input type="checkbox" name="findrange" id="findname<%=tfilterid%>" value="1"<%=cked(clng(tfiltermode and 1)<>0)%> /><label for="findname<%=tfilterid%>">称呼</label>
				<input type="checkbox" name="findrange" id="findmail<%=tfilterid%>" value="2"<%=cked(clng(tfiltermode and 2)<>0)%> /><label for="findmail<%=tfilterid%>">邮件</label>
				<input type="checkbox" name="findrange" id="findqq<%=tfilterid%>" value="4"<%=cked(clng(tfiltermode and 4)<>0)%> /><label for="findqq<%=tfilterid%>">QQ号</label>
				<input type="checkbox" name="findrange" id="findmsn<%=tfilterid%>" value="8"<%=cked(clng(tfiltermode and 8)<>0)%> /><label for="findmsn<%=tfilterid%>">Skype</label>
				<input type="checkbox" name="findrange" id="findhome<%=tfilterid%>" value="16"<%=cked(clng(tfiltermode and 16)<>0)%> /><label for="findhome<%=tfilterid%>">主页</label>
				<input type="checkbox" name="findrange" id="findtitle<%=tfilterid%>" value="32"<%=cked(clng(tfiltermode and 32)<>0)%> /><label for="findtitle<%=tfilterid%>">标题</label>
				<input type="checkbox" name="findrange" id="findcontent<%=tfilterid%>" value="64"<%=cked(clng(tfiltermode and 64)<>0)%> /><label for="findcontent<%=tfilterid%>">内容</label>
				</p>
				<p>处理方式<br/>
				<input type="radio" name="filtermethod" id="filtermethoda<%=tfilterid%>" value="0"<%=cked(clng(tfiltermode and 16384)=0)%> onclick="if(typeof(this.form.replacetxt.disabled)!='undefined')this.form.replacetxt.disabled=false;" /><label for="filtermethoda<%=tfilterid%>">替换为下面的文本</label>
				<input type="radio" name="filtermethod" id="filtermethodb<%=tfilterid%>" value="4096"<%=cked(clng(tfiltermode and 4096)<>0)%> onclick="if(typeof(this.form.replacetxt.disabled)!='undefined')this.form.replacetxt.disabled=true;" /><label for="filtermethodb<%=tfilterid%>">等待审核</label>
				<input type="radio" name="filtermethod" id="filtermethodc<%=tfilterid%>" value="16384"<%=cked(clng(tfiltermode and 16384)<>0)%> onclick="if(typeof(this.form.replacetxt.disabled)!='undefined')this.form.replacetxt.disabled=true;" /><label for="filtermethodc<%=tfilterid%>">拒绝留言</label><br/>
				<input type="text" name="replacetxt" value="<%=rs("replacestr")%>"<%=dised(clng(tfiltermode and 16384+4096)<>0)%> />
				</p>
				<p>备注<br/>
				<input type="text" name="memo" maxlength="25" value="<%=rs("memo")%>" />
				</p>
				<div class="field-command">
				<input type="submit" value="更新" onclick="if (this.form.findexp.value.length===0) {alert('请输入查找内容。');this.form.findexp.focus();return false;}" />
				<input type="button" value="删除" onclick="if (confirm('确实要删除该条策略吗？')) {this.form.action='admin_delfilter.asp';this.form.submit();}" />
				<input type="button" value="上移" onclick="{this.form.movedirection.value='up';this.form.action='admin_movefilter.asp';this.form.submit();}" />
				<input type="button" value="下移" onclick="{this.form.movedirection.value='down';this.form.action='admin_movefilter.asp';this.form.submit();}" />
				<input type="button" value="移至顶层" onclick="{this.form.movedirection.value='top';this.form.action='admin_movefilter.asp';this.form.submit();}" />
				<input type="button" value="移至底层" onclick="{this.form.movedirection.value='bottom';this.form.action='admin_movefilter.asp';this.form.submit();}" />
				<input type="hidden" name="movedirection" value="" />
				</div>
			</form>
			<%
			rs.MoveNext
			wend

			rs.Close : cn.Close : set rs=nothing : set cn=nothing
			%>
		</div>
	</div>

</div>

<!-- #include file="include/template/footer.inc" -->
</body>
</html>