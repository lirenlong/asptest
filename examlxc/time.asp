<%'---------------------------------------------------------------
' AspStudio_Codepage="936"
' 上面这行是软件使用的代码页标记，请不要删除。详情请参考帮助文件。
'
' 档案名称：time.asp
' 原创作者：番茄花园
' 作者邮件：
' 创建日期：星期六，2008年05月24日 22:56:37
' 版权所有(C)番茄花园
'--------------------------------------------------------------%>

<HTML>
<HEAD>
	<Title>time.asp</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<META name="Generator" content="Asp Studio 1.0">
</HEAD>

<BODY>

<!-- 请在这里输入您的HTML代码 -->

<%
	'请在这里输入您的ASP代码
		
  '保存开始时间
  dim starttime
	'session("starttime")=hour(now())*60 + minute(now())
	'starttime=session("starttime")
	starttime = hour(now())*60*60 + minute(now())*60 + second(now())
	dim lefttime
  lefttime=1
  do while lefttime>0
    'lefttime=Request.Cookies("testtime")-(hour(now())*60+minute(now())+ Request.Cookies("starttime"))
	  'lefttime = session("testtime") - (hour(now()))*60- minute(now()) +session("starttime")
	  lefttime = session("testtime") * 60 + starttime - (hour(now()))*60*60- minute(now())*60 -second(now())
  loop
  if lefttime=0 or lefttime<0 then
%>
<script language=vbscript>
  msgbox "时间到了！请交卷" 
</script>
<%
 end if 
%>
	
</BODY>

</HTML>
