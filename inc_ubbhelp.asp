<table id="ubblist" cellpadding="2" class="grid" style="width:auto; border:solid 1px <%=TableBorderColor%>; border-collapse:collapse; margin:0px auto;">
	<tr style="font-weight:bold;"><th>功能</th><th>格式</th><th>状态</th><th>分组</th></tr>
	<tr><td>图片</td>					<td>[img]图片地址[/img]<br/>[img=宽,高]图片地址[/img]</td>			<td><%=getstatus(web_UbbFlag_image and UbbFlag_image)%></td>			<td>图片</td></tr>
	<tr><td>超链接</td>					<td>[url]资源地址[/url]<br/>[url=资源地址]链接文字[/url]</td>		<td><%=getstatus(web_UbbFlag_url and UbbFlag_url)%></td>				<td>URL、Email</td></tr>
	<tr><td>电子邮件链接</td>			<td>[email]邮件地址[/email]</td>										<td><%=getstatus(web_UbbFlag_url and UbbFlag_url)%></td>				<td>URL、Email</td></tr>
	<tr><td>自动识别网址</td>			<td>输入网址后自动变为超链接</td>									<td><%=getstatus(web_UbbFlag_autourl and UbbFlag_autourl)%></td>		<td>自动识别网址</td></tr>
	<tr><td>Flash</td>					<td>[flash]Flash地址[/flash]<br/>[flash=宽,高]Flash地址[/flash]</td>	<td><%=getstatus(web_UbbFlag_player and UbbFlag_player)%></td>			<td>播放控件</td></tr>
	<tr><td>Windows Media Player</td>	<td>[mp]媒体文件地址[/mp]<br/>[mp=宽,高]媒体文件地址[/mp]</td>		<td><%=getstatus(web_UbbFlag_player and UbbFlag_player)%></td>			<td>播放控件</td></tr>
	<tr><td>Real Player</td>			<td>[rm]媒体文件地址[/rm]<br/>[rm=宽,高]媒体文件地址[/rm]</td>		<td><%=getstatus(web_UbbFlag_player and UbbFlag_player)%></td>			<td>播放控件</td></tr>
	<tr><td>引用（缩进）</td>			<td>[quote]引用文字[/quote]</td>										<td><%=getstatus(web_UbbFlag_paragraph and UbbFlag_paragraph)%></td>	<td>段落样式</td></tr>
	<tr><td>无序列表</td>				<td>[li]列表项[/li]</td>												<td><%=getstatus(web_UbbFlag_paragraph and UbbFlag_paragraph)%></td>	<td>段落样式</td></tr>
	<tr><td>指定字体</td>				<td>[font=字体名]文字[/font]</td>										<td><%=getstatus(web_UbbFlag_fontstyle and UbbFlag_fontstyle)%></td>	<td>字体样式</td></tr>
	<tr><td>粗体字</td>					<td>[b]粗体文字[/b]</td>												<td><%=getstatus(web_UbbFlag_fontstyle and UbbFlag_fontstyle)%></td>	<td>字体样式</td></tr>
	<tr><td>斜体字</td>					<td>[i]斜体文字[/i]</td>												<td><%=getstatus(web_UbbFlag_fontstyle and UbbFlag_fontstyle)%></td>	<td>字体样式</td></tr>
	<tr><td>下划线</td>					<td>[u]加下划线的文字[/u]</td>										<td><%=getstatus(web_UbbFlag_fontstyle and UbbFlag_fontstyle)%></td>	<td>字体样式</td></tr>
	<tr><td>删除线</td>					<td>[strike]加删除线的文字[/strike]</td>								<td><%=getstatus(web_UbbFlag_fontstyle and UbbFlag_fontstyle)%></td>	<td>字体样式</td></tr>
	<tr><td>文字颜色</td>				<td>[color=颜色]应用颜色的文字[/color]</td>							<td><%=getstatus(web_UbbFlag_fontcolor and UbbFlag_fontcolor)%></td>	<td>字体颜色</td></tr>
	<tr><td>文字背景色</td>			<td>[bgcolor=颜色]含背景色的文字[/bgcolor]</td>						<td><%=getstatus(web_UbbFlag_fontcolor and UbbFlag_fontcolor)%></td>	<td>字体颜色</td></tr>
	<tr><td>指定对齐方式</td>			<td>[align=方式(left,center,right)]应用于对齐的文字[/align]</td>		<td><%=getstatus(web_UbbFlag_alignment and UbbFlag_alignment)%></td>	<td>对齐方式</td></tr>
	<tr><td>居中对齐</td>				<td>[center]应用于居中对齐的文字[/center]</td>						<td><%=getstatus(web_UbbFlag_alignment and UbbFlag_alignment)%></td>	<td>对齐方式</td></tr>
	<tr><td>来回滚动的文字</td>		<td>[fly]滚动文字[/fly]</td>											<td><%=getstatus(web_UbbFlag_movement and UbbFlag_movement)%></td>		<td>移动效果</td></tr>
	<tr><td>向左滚动的文字</td>		<td>[move]滚动文字[/move]</td>											<td><%=getstatus(web_UbbFlag_movement and UbbFlag_movement)%></td>		<td>移动效果</td></tr>
	<tr><td>背景发光效果</td>			<td>[glow=背景发光宽度,颜色,强度]文字[/glow]</td>					<td><%=getstatus(web_UbbFlag_cssfilter and UbbFlag_cssfilter)%></td>	<td>滤镜效果</td></tr>
	<tr><td>背景阴影效果</td>			<td>[shadow=背景阴影宽度,颜色,强度]文字[/shadow]</td>				<td><%=getstatus(web_UbbFlag_cssfilter and UbbFlag_cssfilter)%></td>	<td>滤镜效果</td></tr>
	<tr><td>表情图标</td>				<td>[faceX] （X为表情编号）</td>										<td><%=getstatus(web_UbbFlag_face and UbbFlag_face)%></td>				<td>表情图标</td></tr>
</table>
<!-- #include file="rowlight.inc" -->