<%
const FaceCount=51			'头像总数
const SmallFaceCount=77		'表情总数
const InstanceName=""		'留言本实例名，部署多个留言本时避免Session冲突
const ServerTimezoneOffset=480  '服务器时区偏移，单位分钟
const DisplayTimezoneOffset=480  '显示时区偏移，单位分钟

'此部分代码供内部处理之用,请勿修改
wm_name="@admin"
wm_id=-1
Dim adminid
Dim ruser
%>