<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<%
if session("studentname")="" then
  Response.Redirect "default.asp"
end if

if Request.Form("submit")="ȷ��" then  '���ѡ���˿��Կ�Ŀ������뿼�Խ���
  if Request.Form("selectsubject")="" then
	  response.write " <center>��û��ѡ���Կ�Ŀ����ѡ���Կ�Ŀ��</center>"
  else
    dim rs,sql    
	  session("selectsubjectname")=Request.Form("selectsubject")
		set rs = server.createobject("adodb.recordset")
		sql="select * from subject where subjectname='"&session("selectsubjectname")&"'"
		rs.open sql,conn,1,1	
	  '���浥ѡ��������
	  'Response.Cookies("singlenumber") = rs("singlenumber")
	  session("singlenumber")=rs("singlenumber")
	  '�����ѡ��������
	  'Response.Cookies("multinumber") = rs("multinumber")
	  session("multinumber")=rs("multinumber")
	  '���浥ѡ�����ֵ
	  'Response.Cookies("singleper") = rs("singleper")
	  session("singleper")=rs("singleper")
	  '�����ѡ�����ֵ
	  'Response.Cookies("multiper") = rs("multiper")
	  session("multiper")=rs("multiper")
	  '���濼��ʱ��
	  'Response.Cookies("testime") = rs("testtime")
  	session("testtime")=rs("testtime")	
	  '���濼�Կ�Ŀ����
	  'Response.Cookies("selectsubjectname") = request.Form("selectsubject")
	  session("selectsubjectname")=request.form("selectsubject")
	  rs.close
		set rs=nothing
		
	 '���뿼�Խ���
	  Response.Redirect "test.asp"
  end if  
end if  

%>
<html>
<head>
<title>���Կ�Ŀѡ��-----���߿���ϵͳ</title>
</head>
<body bgcolor="#66CCCC">
<table border="0" cellspacing="0" cellpadding="0" width="500" height="156" align="center" border=1 bordercolor=lightgreen>
<tr> 
<td><FONT size=4 color=red face=����>
<%Response.Write session("studentname")%></FONT><font face ="�����п�" size="5" color=blue>,��ӭ���������߿���ϵͳ</font>
</td></tr>
<tr>
<td>
<br>
<form action="selectsubject.asp" method="post" id="form" name="form">
<p align=left><FONT color=green face=���� size=4>��������ѡ��Ҫ���ԵĿ�Ŀ</FONT> 
<br>
<%
set rs = server.createobject("adodb.recordset")
sql="select * from subject"
rs.open sql,conn,1,1
if err.number<>0 then 
	response.write "���ݿ����ʧ�ܣ�"&err.description
elseif rs.bof and rs.eof then
	response.write "<center>�Բ���,��ʱû���κο��Կ�Ŀ��</center>"
	rs.close		    
else                          
	do while not rs.eof
		Response.Write( "<input name=selectsubject type=radio value=" & rs("subjectname") & ">" & rs("subjectname") & "," & rs("testtime") & "����<br>")
		rs.movenext         
	loop   
end if
rs.close
set rs=nothing
call endConnection()
%>
<p align=center ><input  name="submit" type="submit" value="ȷ��"></p>
</form>
</td>
</tr>
</table>
</body>
</html>




