<%'---------------------------------------------------------------
' AspStudio_Codepage="936"
' �������������ʹ�õĴ���ҳ��ǣ��벻Ҫɾ����������ο������ļ���
'
' �������ƣ�time.asp
' ԭ�����ߣ����ѻ�԰
' �����ʼ���
' �������ڣ���������2008��05��24�� 22:56:37
' ��Ȩ����(C)���ѻ�԰
'--------------------------------------------------------------%>

<HTML>
<HEAD>
	<Title>time.asp</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<META name="Generator" content="Asp Studio 1.0">
</HEAD>

<BODY>

<!-- ����������������HTML���� -->

<%
	'����������������ASP����
		
  '���濪ʼʱ��
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
  msgbox "ʱ�䵽�ˣ��뽻��" 
</script>
<%
 end if 
%>
	
</BODY>

</HTML>
