<table id="ubblist" class="topic-list">
<thead>
	<tr><th style="width:8em;">功能</th><th>格式</th><th style="width:3em;">状态</th><th style="width:8em;">分组</th></tr>
</thead>
<tbody>
	<%status_image=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_image And UbbFlag_image)%>
	<tr><td>图片</td>					<td>[img]图片地址[/img]<br/>[img=宽,高]图片地址[/img]</td>		<td><%=status_image%></td>		<td>图片</td></tr>
	<%status_url=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_url And UbbFlag_url)%>
	<tr><td>超链接</td>					<td>[url]资源地址[/url]<br/>[url=资源地址]链接文字[/url]</td>	<td><%=status_url%></td>		<td>URL、Email</td></tr>
	<tr><td>电子邮件链接</td>			    <td>[email]邮件地址[/email]</td>								<td><%=status_url%></td>		<td>URL、Email</td></tr>
	<%status_autourl=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_autourl And UbbFlag_autourl)%>
	<tr><td>自动识别网址</td>			    <td>输入网址后自动变为超链接</td>								<td><%=status_autourl%></td>	<td>自动识别网址</td></tr>
	<%status_player=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_player And UbbFlag_player)%>
	<tr><td>Flash</td>					<td>[flash]Flash地址[/flash]<br/>[flash=宽,高]Flash地址[/flash]</td>	<td><%=status_player%></td>	<td>播放控件</td></tr>
	<tr><td>Windows Media Player</td>	<td>[mp]媒体文件地址[/mp]<br/>[mp=宽,高]媒体文件地址[/mp]</td>	<td><%=status_player%></td>		<td>播放控件</td></tr>
	<tr><td>HTML5 视频</td>				<td>[video]视频文件地址[/video]<br/>[video=宽,高]视频文件地址[/video]</td>	<td><%=status_player%></td>	<td>播放控件</td></tr>
	<tr><td>HTML5 音频</td>				<td>[audio]音频文件地址[/audio]<br/>[audio=宽,高]音频文件地址[/audio]</td>	<td><%=status_player%></td>	<td>播放控件</td></tr>
	<%status_paragraph=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_paragraph And UbbFlag_paragraph)%>
	<tr><td>引用（缩进）</td>			    <td>[quote]引用文字[/quote]</td>								<td><%=status_paragraph%></td>	<td>段落样式</td></tr>
	<tr><td>无序列表</td>				<td>[li]列表项[/li]</td>										<td><%=status_paragraph%></td>	<td>段落样式</td></tr>
	<%status_markdown_paragraph=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_markdown_paragraph And UbbFlag_markdown_paragraph)%>
	<tr><td>无序列表</td>				<td>-列表项</td>												<td><%=status_markdown_paragraph%></td>	<td>Markdown段落样式</td></tr>
	<%status_fontstyle=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_fontstyle And UbbFlag_fontstyle)%>
	<tr><td>指定字体</td>				<td>[font=字体名]文字[/font]</td>								<td><%=status_fontstyle%></td>	<td>字体样式</td></tr>
	<tr><td>粗体字</td>					<td>[b]粗体文字[/b]</td>										<td><%=status_fontstyle%></td>	<td>字体样式</td></tr>
	<tr><td>斜体字</td>					<td>[i]斜体文字[/i]</td>										<td><%=status_fontstyle%></td>	<td>字体样式</td></tr>
	<tr><td>下划线</td>					<td>[u]加下划线的文字[/u]</td>								<td><%=status_fontstyle%></td>	<td>字体样式</td></tr>
	<tr><td>删除线</td>					<td>[strike]加删除线的文字[/strike]</td>						<td><%=status_fontstyle%></td>	<td>字体样式</td></tr>
	<%status_markdown_fontstyle=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_markdown_fontstyle And UbbFlag_markdown_fontstyle)%>
	<tr><td>粗体字</td>					<td>**粗体文字**</td>										<td><%=status_markdown_fontstyle%></td>	<td>Markdown字体样式</td></tr>
	<tr><td>斜体字</td>					<td>__斜体文字__</td>										<td><%=status_markdown_fontstyle%></td>	<td>Markdown字体样式</td></tr>
	<%status_fontcolor=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_fontcolor And UbbFlag_fontcolor)%>
	<tr><td>文字颜色</td>				<td>[color=颜色]应用颜色的文字[/color]</td>					<td><%=status_fontcolor%></td>	<td>字体颜色</td></tr>
	<tr><td>文字背景色</td>				<td>[bgcolor=颜色]含背景色的文字[/bgcolor]</td>				<td><%=status_fontcolor%></td>	<td>字体颜色</td></tr>
	<%status_alignment=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_alignment And UbbFlag_alignment)%>
	<tr><td>指定对齐方式</td>			    <td>[align=left|center|right]对齐的文字[/align]</td>			<td><%=status_alignment%></td>	<td>对齐方式</td></tr>
	<tr><td>居中对齐</td>				<td>[center]应用于居中对齐的文字[/center]</td>					<td><%=status_alignment%></td>	<td>对齐方式</td></tr>
	<%status_face=getstatus(web_UBBSupport And UBBSupport AND web_UbbFlag_face And UbbFlag_face)%>
	<tr><td>表情图标</td>				<td>[faceX] （X为表情编号）</td>								<td><%=status_face%></td>		<td>表情图标</td></tr>
</tbody>
</table>
