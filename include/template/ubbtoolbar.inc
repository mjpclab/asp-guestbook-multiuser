<%Sub ShowUbbToolBar(byval allflags)%>
<%if ShowUbbTool then%>
<div class="ubbtoolbar">
	<div class="row">
		<%if ubblimit=1 or (ubblimit=2 and web_UbbFlag_fontstyle) or (web_UbbFlag_fontstyle and UbbFlag_fontstyle) then%><a name="ubbbutton" class="ubbbutton ubb_bold" onclick="sendUbbCode &&sendUbbCode('b')"><img alt="粗体" src="asset/image/icon_ubb_bold.gif"/></a><%end if%>
		<%if ubblimit=1 or (ubblimit=2 and web_UbbFlag_fontstyle) or (web_UbbFlag_fontstyle and UbbFlag_fontstyle) then%><a name="ubbbutton" class="ubbbutton ubb_italic" onclick="sendUbbCode && sendUbbCode('i')"><img alt="斜体" src="asset/image/icon_ubb_italic.gif"/></a><%end if%>
		<%if ubblimit=1 or (ubblimit=2 and web_UbbFlag_fontcolor) or (web_UbbFlag_fontcolor and UbbFlag_fontcolor) then%><a name="ubbbutton" class="ubbbutton ubb_color" onclick="toggleColorPalette && toggleColorPalette(this,arguments[0]||window.event);"><img alt="文字颜色" src="asset/image/icon_ubb_color.gif"/></a><%end if%>
		<%if ubblimit=1 or (ubblimit=2 and web_UbbFlag_image) or (web_UbbFlag_image and UbbFlag_image) then%><a name="ubbbutton" class="ubbbutton ubb_image" onclick="sendUbbCode && sendUbbCode('img')"><img alt="图片" src="asset/image/icon_ubb_img.gif"/></a><%end if%>
		<%if ubblimit=1 or (ubblimit=2 and web_UbbFlag_player) or (web_UbbFlag_player and UbbFlag_player) then%><a name="ubbbutton" class="ubbbutton ubb_flash" onclick="sendUbbCode && sendUbbCode('flash')"><img alt="Flash" src="asset/image/icon_ubb_flash.gif"/></a><%end if%>
		<%if ubblimit=1 or (ubblimit=2 and web_UbbFlag_player) or (web_UbbFlag_player and UbbFlag_player) then%><a name="ubbbutton" class="ubbbutton ubb_mplayer" onclick="sendUbbCode && sendUbbCode('mp')"><img alt="Windows Media Player" src="asset/image/icon_ubb_mplayer.gif"/></a><%end if%>
		<%if ubblimit=1 or (ubblimit=2 and web_UbbFlag_player) or (web_UbbFlag_player and UbbFlag_player) then%><a name="ubbbutton" class="ubbbutton ubb_html5_video" onclick="sendUbbCode && sendUbbCode('video')"><img alt="HTML5视频" src="asset/image/icon_ubb_html5_video.gif"/></a><%end if%>
		<%if ubblimit=1 or (ubblimit=2 and web_UbbFlag_player) or (web_UbbFlag_player and UbbFlag_player) then%><a name="ubbbutton" class="ubbbutton ubb_html5_audio" onclick="sendUbbCode && sendUbbCode('audio')"><img alt="HTML5音频" src="asset/image/icon_ubb_html5_audio.gif"/></a><%end if%>
	</div>
	<%if UbbFlag_face Or allflags then%>
	<div class="row">
		<%for i=1 to SmallFaceCount%><a name="ubbface" class="ubbface face<%=i%>" title="[face<%=i%>]"><img alt="[face<%=i%>]" src="asset/smallface/<%=i%>.gif"/></a><%next%>
	</div>
	<%end if%>
</div>

<div class="colorsContainer" id="colorsContainer">
	<div class="colorsInfo" id="colorsInfo">
		<div class="colorsPreview" id="colorsPreview"></div>
		<div class="colorsNumber" id="colorsNumber">#??????</div>
		<div class="colorsCancel" id="colorsCancel" onclick="this.parentNode.parentNode.style.display='none';">关闭</div>
	</div>
	<div class="colorsPalette" id="colorsPalette"></div>
</div>
<script type="text/javascript" src="asset/js/ubb-hook.js" defer="defer" async="async"></script>
<%end if%>
<%End Sub%>