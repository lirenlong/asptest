<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim id,studentname,studentpassword'�������
dim sql,rs,rsc
if request("submit")="ע��" then   '������û�
	if trim(request("studentname"))="" or trim(request("studentpassword"))="" then
		Response.Redirect "register.asp?errMessage=����!�û��������벻��Ϊ��!"
	  response.end
	end if
	
  if trim(request("studentpassword")) <> trim(request("studentpassword2")) then
    Response.Redirect "register.asp?errMessage=����!������������벻һ��!"
	  response.end
	end if
	
	set rs=server.createobject("adodb.recordset")   '���ѧ���Ƿ�����
	rs.open "select * from student where studentname='" & cstr(trim(request("studentname"))) & "'",conn,1,1
	if err.number <> 0 then
	  Response.Redirect "register.asp?errMessage=���ݿ����!"
	  response.end
	else  if not rs.bof and not rs.eof then
	  Response.Redirect "register.asp?errMessage=����!��ѧ���Ѿ�����!!"
		rs.close
		response.end
	end if
	rs.close
  set rs=nothing
  
  sql="insert into student(studentname,studentpassword) values('" & cstr(trim(request("studentname"))) & "','" & cstr(trim(request("studentpassword"))) & "')"
	conn.execute sql
	if err.number <> 0 then
	  Response.Redirect "register.asp?errMessage=" & "���ݿ��������:" & err.description
		Response.End
	else 
	  session("studentname")=request("studentname") 'ͨ��session����studentname��־һ��ѧ����½��ϵͳ
	  response.write "<script language=javascript>window.alert('ע��ɹ�!')</script>"
		session("studentname")=request("studentname")
		Response.Redirect "selectsubject.asp"
  end if
end if
end if
%>

<html>
<head>
<title>���û�ע��----���߿���ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#66CCCC" background="images/backgrand.jpg" onresize=hero(); scroll="no"> 
<%
  if Request("errMessage") <> null or Request("errMessage") <> "" then
    Response.Write("<center><font color=red>" & Request("errMessage") & "</font></center>")
	end if
%>
<form action="register.asp" method="post">
<table border=0 cellpadding=0 cellspacing=0 bordercolor=lightgreen align="center" width=350>
<tr><td colspan=2 align="center"><font color="green">���û�ע��</font></td></tr>
<tr><td>�û�����:</td><td><input type="text" name="studentname" class=input maxlength=14 size="16"></td></tr>
<tr><td>�û�����:</td><td><input type="password" name="studentpassword"  class=input maxlength=12 size="16"></td></tr>
<tr><td>����ȷ��:</td><td><input type="password" name="studentpassword2"  class=input maxlength=12 size="16"></td></tr>
<tr><td colspan=2 align="center">
 <input type=submit name="submit" value="ע��" class=button>
</td></tr>
<tr><td colspan=2 align="center"><a href="lo.asp"><font color=black size=+0>�������߿���ϵͳ��¼����</font></a></td></tr>
</table>
</form>
</body>
</html>