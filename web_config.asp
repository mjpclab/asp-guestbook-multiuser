<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_config.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="webconfig.asp" -->
<!-- #include file="web_admin_verify.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

Call CreateConn(cn)
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=web_BookName%> Webmaster管理中心 留言本参数</title>
	<!-- #include file="inc_web_admin_stylesheet.asp" -->
</head>

<body>

<div id="outerborder" class="outerborder">

	<%Call WebInitHeaderData("","Webmaster管理中心","","")%><!-- #include file="include/template/header.inc" -->
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/web_admin_mainmenu.inc" -->
	<%rs.Open Replace(sql_adminconfig_config,"{0}",wm_id),cn,,,1%>

	<div class="region region-config admin-tools">
		<h3 class="title">留言本参数设置</h3>
		<div class="content">
			<form method="post" action="web_saveconfig.asp" name="configform" onsubmit="return check();">

			<div id="tabContainer"></div>

			<div id="tab-basic">
				<h4>基本配置（当某功能关闭时，用户设置无效，IP显示优先于用户设置，验证码长度作为用户设置最小值）：</h4>
				<%tstatus=rs("status")%>
				<div class="field">
					<span class="label">留言本创建功能：</span>
					<span class="value"><input type="radio" name="status4" value="1" id="status41"<%=cked(CBool(tstatus AND 8))%> /><label for="status41">开启</label>　　<input type="radio" name="status4" value="0" id="status42"<%=cked(Not CBool(tstatus AND 8))%> /><label for="status42">关闭</label></span>
				</div>
				<div class="field">
					<span class="label">找回密码功能：</span>
					<span class="value"><input type="radio" name="status5" value="1" id="status51"<%=cked(CBool(tstatus AND 16))%> /><label for="status51">开启</label>　　<input type="radio" name="status5" value="0" id="status52"<%=cked(Not CBool(tstatus AND 16))%> /><label for="status52">关闭</label></span>
				</div>
				<div class="field">
					<span class="label">用户登录维护功能：</span>
					<span class="value"><input type="radio" name="status6" value="1" id="status61"<%=cked(CBool(tstatus AND 32))%> /><label for="status61">开启</label>　　<input type="radio" name="status6" value="0" id="status62"<%=cked(Not CBool(tstatus AND 32))%> /><label for="status62">关闭</label></span>
				</div>
				<div class="field">
					<span class="label">用户自删帐号功能：</span>
					<span class="value"><input type="radio" name="status7" value="1" id="status71"<%=cked(CBool(tstatus AND 64))%> /><label for="status71">开启</label>　　<input type="radio" name="status7" value="0" id="status72"<%=cked(Not CBool(tstatus AND 64))%> /><label for="status72">关闭</label></span>
				</div>
				<div class="field">
					<span class="label">留言本状态：</span>
					<span class="value"><input type="radio" name="status1" value="1" id="status11"<%=cked(CBool(tstatus AND 1))%> /><label for="status11">开启</label>　　<input type="radio" name="status1" value="0" id="status12"<%=cked(Not CBool(tstatus AND 1))%> /><label for="status12">关闭</label></span>
				</div>
				<div class="field">
					<span class="label">访客留言权限：</span>
					<span class="value"><input type="radio" name="status2" value="1" id="status21"<%=cked(CBool(tstatus AND 2))%> /><label for="status21">开启</label>　　<input type="radio" name="status2" value="0" id="status22"<%=cked(Not CBool(tstatus AND 2))%> /><label for="status22">关闭</label></span>
				</div>
				<div class="field">
					<span class="label">访客搜索留言权限：</span>
					<span class="value"><input type="radio" name="status3" value="1" id="status31"<%=cked(CBool(tstatus AND 4))%> /><label for="status31">开启</label>　　<input type="radio" name="status3" value="0" id="status32"<%=cked(Not CBool(tstatus AND 4))%> /><label for="status32">关闭</label></span>
				</div>
				<div class="field">
					<span class="label">访问统计：</span>
					<span class="value"><input type="radio" name="status9" value="1" id="status91"<%=cked(CBool(tstatus AND 256))%> /><label for="status91">开启</label>　　<input type="radio" name="status9" value="0" id="status92"<%=cked(Not CBool(tstatus AND 256))%> /><label for="status92">关闭</label></span>
				</div>
				<%adminlimit=rs("adminhtml")%>
				<div class="field">
					<span class="label">管理中心安全性设置：</span>
					<span class="value"><input type="checkbox" value="1" name="adminviewcode" id="adminviewcode"<%=cked(CBool(adminlimit AND 8))%> /><label for="adminviewcode">用户及访客留言显示实际HTML或UBB代码</label></span>
				</div>
				<div class="field">
					<span class="label">用户管理员HTML权限：</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="adminhtml" id="adminhtml"<%=cked(CBool(adminlimit AND 1))%> /><label for="adminhtml">开启HTML权限（不推荐）</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminubb" id="adminubb"<%=cked(CBool(adminlimit AND 2))%> /><label for="adminubb">开启UBB权限</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminertn" id="adminertn"<%=cked(CBool(adminlimit AND 4))%> /><label for="adminertn">不支持HTML和UBB时，允许回车换行</label></span>
					</span>
				</div>
				<%guestlimit=rs("guesthtml")%>
				<div class="field">
					<span class="label">访客HTML权限：</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="guesthtml" id="guesthtml"<%=cked(CBool(guestlimit AND 1))%> /><label for="guesthtml">开启HTML权限（不推荐）</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestubb" id="guestubb"<%=cked(CBool(guestlimit AND 2))%> /><label for="guestubb">开启UBB权限</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestertn" id="guestertn"<%=cked(CBool(guestlimit AND 4))%> /><label for="guestertn">不支持HTML和UBB时，允许回车换行</label></span>
					</span>
				</div>
				<div class="field">
					<span class="label">管理员登录超时：</span>
					<span class="value"><input type="text" size="4" maxlength="4" name="admintimeout" value="<%=rs("admintimeout")%>" />分 (默认=20)</span>
				</div>
				<div class="field">
					<span class="label">为访客显示IPv4前：</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="showipv4" value="<%=ShowIPv4%>" />字节 (可选值：0～4)</span>
				</div>
				<div class="field">
					<span class="label">为访客显示IPv6前：</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="showipv6" value="<%=ShowIPv6%>" />组 (可选值：0～8)</span>
				</div>
				<div class="field">
					<span class="label">为管理员显示IPv4前：</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshowipv4" value="<%=AdminShowIPv4%>" />字节 (可选值：0～4)</span>
				</div>
				<div class="field">
					<span class="label">为管理员显示IPv6前：</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshowipv6" value="<%=AdminShowIPv6%>" />组 (可选值：0～8)</span>
				</div>
				<div class="field">
					<span class="label">为管理员显示原IPv4前：</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshoworiginalipv4" value="<%=AdminShowOriginalIPv4%>" />字节 (可选值：0～4,使用代理服务器时此项显示原始IP)</span>
				</div>
				<div class="field">
					<span class="label">为管理员显示原IPv6前：</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshoworiginalipv6" value="<%=AdminShowOriginalIPv6%>" />组 (可选值：0～8,使用代理服务器时此项显示原始IP)</span>
				</div>
				<div class="field">
					<span class="label">登录验证码长度：</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="vcodecount" value="<%=rs("vcodecount") AND &H0F%>" />位 (可选值：0～10)</span>
				</div>
				<div class="field">
					<span class="label">留言验证码长度：</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="writevcodecount" value="<%=(rs("vcodecount") and &HF0) \ &H10%>" />位 (可选值：0～10)</span>
				</div>
			</div>

			<div id="tab-mail">
				<h4>邮件通知（设置是否对用户开放）：</h4>
				<%MailFlag=rs("mailflag")%>
				<div class="field">
					<span class="label">新留言到达通知版主：</span>
					<span class="value"><input type="checkbox" value="1" name="mailnewinform" id="mailnewinform"<%=cked(CBool(MailFlag AND 1))%> /><label for="mailnewinform">启用</label></span>
				</div>
				<div class="field">
					<span class="label">版主回复通知留言人：</span>
					<span class="value"><input type="checkbox" value="1" name="mailreplyinform" id="mailreplyinform"<%=cked(CBool(MailFlag AND 2))%> /><label for="mailreplyinform">开启</label></span>
				</div>
				<div class="field">
					<span class="label">邮件发送组件：</span>
					<span class="value"><input type="radio" value="0" name="mailcomponent" id="mailcomponent0"<%=cked(Not CBool(MailFlag AND 4))%> /><label for="mailcomponent0">JMail</label>　<input type="radio" value="1" name="mailcomponent" id="mailcomponent1"<%=cked(CBool(MailFlag AND 4))%> /><label for="mailcomponent1">CDO</label></span>
				</div>
			</div>

			<div id="tab-ui-size">
				<h4>界面与尺寸（此项只影响本管理页面）：</h4>
				<div class="field">
					<span class="label">字体列表（","分隔）：</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="cssfontfamily" value="<%=rs("cssfontfamily")%>" /></span>
				</div>
				<div class="field">
					<span class="label">字体大小：</span>
					<span class="value"><input type="text" size="10" maxlength="8" name="cssfontsize" value="<%=rs("cssfontsize")%>" /></span>
				</div>
				<div class="field">
					<span class="label">文字行间距：</span>
					<span class="value"><input type="text" size="10" maxlength="8" name="csslineheight" value="<%=rs("csslineheight")%>" /></span>
				</div>
				<div class="field">
					<span class="label">留言本最大宽度：</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="tablewidth" value="<%=rs("tablewidth")%>" /> (默认=630,可用百分比)</span>
				</div>
				<div class="field">
					<span class="label">窗格区块间距：</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="windowspace" value="<%=rs("windowspace")%>" /> (默认=20,单位:象素)</span>
				</div>
				<div class="field">
					<span class="label">留言本左窗格宽度：</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="tableleftwidth" value="<%=rs("tableleftwidth")%>" /> (默认=150,可用百分比)</span>
				</div>
				<div class="field">
					<span class="label">搜索框宽度：</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="searchtextwidth" value="<%=rs("searchtextwidth")%>" /> (默认=20,单位:字母宽度)</span>
				</div>
				<div class="field">
					<span class="label">公告编辑框高度：</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="replytextheight" value="<%=rs("replytextheight")%>" /> (默认=10,单位:字母高度)</span>
				</div>
				<div class="field">
					<span class="label">每页显示的项目数：</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="itemsperpage" value="<%=rs("itemsperpage")%>" /> (默认=5)</span>
				</div>
				<div class="field">
					<span class="label">每页显示的标题数：</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="titlesperpage" value="<%=rs("titlesperpage")%>" /> (默认=20)</span>
				</div>
				<div class="field">
					<span class="label">留言本配色方案：</span>
					<span class="value">
						<select name="style">
						<%
						styleid=rs("styleid")

						set rs1=server.CreateObject("ADODB.Recordset")
						rs1.Open sql_adminconfig_style,cn,,,1

						dim onestyleid,onestylename
						while rs1.EOF=false
							onestyleid=rs1("styleid")
							onestylename=rs1("stylename")
							Response.Write "<option value=""" & onestyleid & """"
							Response.Write seled(onestyleid=styleid or onestyleid="")
							Response.Write ">" &onestylename& "</option>"
							rs1.MoveNext
						wend
						rs1.Close : set rs1=nothing
						%>
						</select>
					</span>
				</div>
			</div>

			<div id="tab-function">
				<h4>功能设置：</h4>
				<%tvisualflag=rs("visualflag")%>
				<div class="field">
					<span class="label">默认版面模式：</span>
					<span class="value"><input type="radio" name="displaymode" value="1" id="displaymode1"<%=cked(CBool(tvisualflag AND 1024))%> /><label for="displaymode1">标题模式</label>　　<input type="radio" name="displaymode" value="0" id="displaymode2"<%=cked(Not CBool(tvisualflag AND 1024))%> /><label for="displaymode2">完整模式</label></span>
				</div>
				<div class="field">
					<span class="label">回复内容显示位置：</span>
					<span class="value"><input type="radio" name="replyinword" value="1" id="replyinword1"<%=cked(CBool(tvisualflag AND 1))%> /><label for="replyinword1">内嵌于访客留言</label>　　<input type="radio" name="replyinword" value="0" id="replyinword2"<%=cked(Not CBool(tvisualflag AND 1))%> /><label for="replyinword2">显示在访客留言下方</label></span>
				</div>
				<div class="field">
					<span class="label">分页窗口显示位置：</span>
					<span class="value"><input type="radio" name="showpagelist" value="3" id="showpagelist3"<%=cked((tvisualflag and 12)=12)%> /><label for="showpagelist3">上下方</label>　<input type="radio" name="showpagelist" value="1" id="showpagelist1"<%=cked((tvisualflag and 12)=4)%> /><label for="showpagelist1">上方</label>　　<input type="radio" name="showpagelist" value="2" id="showpagelist2"<%=cked((tvisualflag and 12)=8)%> /><label for="showpagelist2">下方</label></span>
				</div>
				<div class="field">
					<span class="label">分页列表模式：</span>
					<span class="value"><input type="radio" name="advpagelist" value="1" id="advpagelist1"<%=cked(CBool(tvisualflag AND 64))%> /><label for="advpagelist1">区段式</label>　<input type="radio" name="advpagelist" value="0" id="advpagelist2"<%=cked(Not CBool(tvisualflag AND 64))%> /><label for="advpagelist2">平面式</label></span>
				</div>
				<div class="field">
					<span class="label">区段式分页项数：</span>
					<span class="value"><input type="text" size="5" maxlength="3" name="advpagelistcount" value="<%=rs("advpagelistcount")%>" /> (默认=10)</span>
				</div>
				<div class="field">
					<span class="label">UBB开关(须启用UBB)：</span>
					<span class="value">
						<span class="row">
							<input type="checkbox" name="ubbflag_image" id="ubbflag_image" value="1"<%=cked(UbbFlag_image)%> /><label for="ubbflag_image">图片</label>
							<input type="checkbox" name="ubbflag_url" id="ubbflag_url" value="1"<%=cked(UbbFlag_url)%> /><label for="ubbflag_url">URL、Email</label>
							<input type="checkbox" name="ubbflag_autourl" id="ubbflag_autourl" value="1"<%=cked(UbbFlag_autourl)%> /><label for="ubbflag_autourl">自动识别网址</label>
						</span>
						<span class="row">
							<input type="checkbox" name="ubbflag_player" id="ubbflag_player" value="1"<%=cked(UbbFlag_player)%> /><label for="ubbflag_player">播放控件</label>
							<input type="checkbox" name="ubbflag_paragraph" id="ubbflag_paragraph" value="1"<%=cked(UbbFlag_paragraph)%> /><label for="ubbflag_paragraph">段落样式</label>
							<input type="checkbox" name="ubbflag_fontstyle" id="ubbflag_fontstyle" value="1"<%=cked(UbbFlag_fontstyle)%> /><label for="ubbflag_fontstyle">字体样式</label>
						</span>
						<span class="row">
							<input type="checkbox" name="ubbflag_fontcolor" id="ubbflag_fontcolor" value="1"<%=cked(UbbFlag_fontcolor)%> /><label for="ubbflag_fontcolor">字体颜色</label>
							<input type="checkbox" name="ubbflag_alignment" id="ubbflag_alignment" value="1"<%=cked(UbbFlag_alignment)%> /><label for="ubbflag_alignment">对齐方式</label>
							<input type="checkbox" name="ubbflag_face" id="ubbflag_face" value="1"<%=cked(UbbFlag_face)%> /><label for="ubbflag_face">表情图标</label>
						</span>
					</span>
				</div>
				<%tpagecontrol=rs("pagecontrol")%>
				<%tdelconfirm=rs("delconfirm")%>
				<div class="field">
					<span class="label">删除留言时提示：</span>
					<span class="value"><input type="radio" name="deltip" value="1" id="deltip1"<%=cked(CBool(tdelconfirm AND 1))%> /><label for="deltip1">提示</label>　　<input type="radio" name="deltip" value="0" id="deltip2"<%=cked(Not CBool(tdelconfirm AND 1))%> /><label for="deltip2">不提示</label></span>
				</div>
				<div class="field">
					<span class="label">删除回复时提示：</span>
					<span class="value"><input type="radio" name="delretip" value="1" id="delretip1"<%=cked(CBool(tdelconfirm AND 2))%> /><label for="delretip1">提示</label>　　<input type="radio" name="delretip" value="0" id="delretip2"<%=cked(Not CBool(tdelconfirm AND 2))%> /><label for="delretip2">不提示</label></span>
				</div>
				<div class="field">
					<span class="label">删除选定留言时提示：</span>
					<span class="value"><input type="radio" name="delseltip" value="1" id="delseltip1"<%=cked(CBool(tdelconfirm AND 4))%> /><label for="delseltip1">提示</label>　　<input type="radio" name="delseltip" value="0" id="delseltip2"<%=cked(Not CBool(tdelconfirm AND 4))%> /><label for="delseltip2">不提示</label></span>
				</div>
				<div class="field">
					<span class="label">执行高级删除时提示：</span>
					<span class="value"><input type="radio" name="deladvtip" value="1" id="deladvtip1"<%=cked(CBool(tdelconfirm AND 8))%> /><label for="deladvtip1">提示</label>　　<input type="radio" name="deladvtip" value="0" id="deladvtip2"<%=cked(Not CBool(tdelconfirm AND 8))%> /><label for="deladvtip2">不提示</label></span>
				</div>
				<div class="field">
					<span class="label">删除置顶公告时提示：</span>
					<span class="value"><input type="radio" name="deldectip" value="1" id="deldectip1"<%=cked(CBool(tdelconfirm AND 16))%> /><label for="deldectip1">提示</label>　　<input type="radio" name="deldectip" value="0" id="deldectip2"<%=cked(Not CBool(tdelconfirm AND 16))%> /><label for="deldectip2">不提示</label></span>
				</div>
				<div class="field">
					<span class="label">删除选定置顶公告时提示：</span>
					<span class="value"><input type="radio" name="delseldectip" value="1" id="delseldectip1"<%=cked(CBool(tdelconfirm AND 32))%> /><label for="delseldectip1">提示</label>　　<input type="radio" name="delseldectip" value="0" id="delseldectip2"<%=cked(Not CBool(tdelconfirm AND 32))%> /><label for="delseldectip2">不提示</label></span>
				</div>
			</div>

			<input type="hidden" name="tabIndex" id="tabIndex" value="<%=Request.QueryString("tabIndex")%>" />
			<div class="command"><input value="更新数据" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>

<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

<script type="text/javascript" defer="defer">
function check()
{
	var tv;

	if (isNaN(tv=Number(document.configform.admintimeout.value)))
		{alert('“管理员登录超时”必须为数字。');document.configform.admintimeout.select();return false;}
	else if (tv<1 || tv>1440)
		{alert('“管理员登录超时”必须在1～1440的范围内。');document.configform.admintimeout.select();return false;}

	if (isNaN(tv=Number(document.configform.showipv4.value)))
		{alert('“为访客显示IPv4”必须为数字。');document.configform.showipv4.select();return false;}
	else if (tv<0 || tv>4 || document.configform.showipv4.value==='')
		{alert('“为访客显示IPv4”必须在0～4的范围内。');document.configform.showipv4.select();return false;}

	if (isNaN(tv=Number(document.configform.showipv6.value)))
		{alert('“为访客显示IPv6”必须为数字。');document.configform.showipv6.select();return false;}
	else if (tv<0 || tv>8 || document.configform.showipv6.value==='')
		{alert('“为访客显示IPv6”必须在0～8的范围内。');document.configform.showipv6.select();return false;}

	if (isNaN(tv=Number(document.configform.adminshowipv4.value)))
		{alert('“为管理员显示IPv4”必须为数字。');document.configform.adminshowipv4.select();return false;}
	else if (tv<0 || tv>4 || document.configform.adminshowipv4.value==='')
		{alert('“为管理员显示IPv4”必须在0～4的范围内。');document.configform.adminshowipv4.select();return false;}

	if (isNaN(tv=Number(document.configform.adminshowipv6.value)))
		{alert('“为管理员显示IPv6”必须为数字。');document.configform.adminshowipv6.select();return false;}
	else if (tv<0 || tv>8 || document.configform.adminshowipv6.value==='')
		{alert('“为管理员显示IPv6”必须在0～8的范围内。');document.configform.adminshowipv6.select();return false;}

	if (isNaN(tv=Number(document.configform.adminshoworiginalipv4.value)))
		{alert('“为管理员显示原IPv4”必须为数字');document.configform.adminshoworiginalipv4.select();return false;}
	else if (tv<0 || tv>4 || document.configform.adminshoworiginalipv4.value==='')
		{alert('“为管理员显示原IPv4”必须在0～4的范围内。');document.configform.adminshoworiginalipv4.select();return false;}

	if (isNaN(tv=Number(document.configform.adminshoworiginalipv6.value)))
		{alert('“为管理员显示原IPv6”必须为数字');document.configform.adminshoworiginalipv6.select();return false;}
	else if (tv<0 || tv>8 || document.configform.adminshoworiginalipv6.value==='')
		{alert('“为管理员显示原IPv6”必须在0～8的范围内。');document.configform.adminshoworiginalipv6.select();return false;}

	if (isNaN(tv=Number(document.configform.vcodecount.value)))
		{alert('“登录验证码长度”必须为数字。');document.configform.vcodecount.select();return false;}
	else if (tv<0 || tv>10)
		{alert('“登录验证码长度”必须在0～10的范围内。');document.configform.vcodecount.select();return false;}

	if (isNaN(tv=Number(document.configform.writevcodecount.value)))
		{alert('“留言验证码长度”必须为数字。');document.configform.writevcodecount.select();return false;}
	else if (tv<0 || tv>10)
		{alert('“留言验证码长度”必须在0～10的范围内。');document.configform.writevcodecount.select();return false;}


	if (isNaN(tv=Number(document.configform.tablewidth.value)))
		{if (/^\d+%$/.test(document.configform.tablewidth.value)==false) {alert('“留言本最大宽度”必须为正数或正百分比。');document.configform.tablewidth.select();return false;}}
	else if (tv<1)
		{alert('“留言本最大宽度”必须大于零。');document.configform.tablewidth.select();return false;}

	if (isNaN(tv=Number(document.configform.windowspace.value)))
		{alert('“窗口区块间距”必须为数字。');document.configform.windowspace.select();return false;}
	else if (tv<1 || tv>255)
		{alert('“窗口区块间距”必须在1～255的范围内。');document.configform.windowspace.select();return false;}

	if (isNaN(tv=Number(document.configform.tableleftwidth.value)))
		{if (/^\d+%$/.test(document.configform.tableleftwidth.value)==false) {alert('“留言本左窗格宽度”必须为正数或正百分比。');document.configform.tableleftwidth.select();return false;}}
	else if (tv<1)
		{alert('“留言本左窗格宽度”必须大于零。');document.configform.tableleftwidth.select();return false;}

	if (isNaN(tv=Number(document.configform.searchtextwidth.value)))
		{alert('“搜索框宽度”必须为数字。');document.configform.searchtextwidth.select();return false;}
	else if (tv<1 || tv>255)
		{alert('“搜索框宽度”必须在1～255的范围内。');document.configform.searchtextwidth.select();return false;}

	if (isNaN(tv=Number(document.configform.replytextheight.value)))
		{alert('“回复、公告编辑框高度”必须为数字。');document.configform.replytextheight.select();return false;}
	else if (tv<1 || tv>255)
		{alert('“回复、公告编辑框高度”必须在1～255的范围内。');document.configform.replytextheight.select();return false;}

	if (isNaN(tv=Number(document.configform.itemsperpage.value)))
		{alert('“每页显示的留言数”必须为数字。');document.configform.itemsperpage.select();return false;}
	else if (tv<1 || tv>32767)
		{alert('“每页显示的留言数”必须在1～32767的范围内。');document.configform.itemsperpage.select();return false;}

	if (isNaN(tv=Number(document.configform.titlesperpage.value)))
		{alert('“每页显示的标题数”必须为数字。');document.configform.titlesperpage.select();return false;}
	else if (tv<1 || tv>32767)
		{alert('“每页显示的标题数”必须在1～32767的范围内。');document.configform.titlesperpage.select();return false;}


	if (isNaN(tv=Number(document.configform.advpagelistcount.value)))
		{alert('“区段式分页项数”必须为数字。');document.configform.advpagelistcount.select();return false;}
	else if (tv<1 || tv>255 || document.configform.advpagelistcount.value=='')
		{alert('“区段式分页项数”必须在1～255的范围内。');document.configform.advpagelistcount.select();return false;}

	document.configform.submit1.disabled=true;
	return true;
}
</script>
<script type="text/javascript" src="asset/js/tabcontrol.js"></script>
<script type="text/javascript">
var tab=new TabControl('tabContainer');

tab.addPage('tab-basic','基本配置');
tab.addPage('tab-mail','邮件通知');
tab.addPage('tab-ui-size','界面与尺寸');
tab.addPage('tab-function','功能设置');

tab.restoreFromField('tabIndex');
</script>
</body>
</html>