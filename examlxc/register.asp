<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim id,studentname,studentpassword'定义变量
dim sql,rs,rsc
if request("submit")="注册" then   '添加新用户
	if trim(request("studentname"))="" or trim(request("studentpassword"))="" then
		Response.Redirect "register.asp?errMessage=错误!用户名或密码不能为空!"
	  response.end
	end if
	
  if trim(request("studentpassword")) <> trim(request("studentpassword2")) then
    Response.Redirect "register.asp?errMessage=错误!两次输入的密码不一致!"
	  response.end
	end if
	
	set rs=server.createobject("adodb.recordset")   '检查学生是否重名
	rs.open "select * from student where studentname='" & cstr(trim(request("studentname"))) & "'",conn,1,1
	if err.number <> 0 then
	  Response.Redirect "register.asp?errMessage=数据库出错!"
	  response.end
	else  if not rs.bof and not rs.eof then
	  Response.Redirect "register.asp?errMessage=错误!该学生已经存在!!"
		rs.close
		response.end
	end if
	rs.close
  set rs=nothing
  
  sql="insert into student(studentname,studentpassword) values('" & cstr(trim(request("studentname"))) & "','" & cstr(trim(request("studentpassword"))) & "')"
	conn.execute sql
	if err.number <> 0 then
	  Response.Redirect "register.asp?errMessage=" & "数据库操作出错:" & err.description
		Response.End
	else 
	  session("studentname")=request("studentname") '通过session变量studentname标志一个学生登陆了系统
	  response.write "<script language=javascript>window.alert('注册成功!')</script>"
		session("studentname")=request("studentname")
		Response.Redirect "selectsubject.asp"
  end if
end if
end if
%>

<html>
<head>
<title>新用户注册----在线考试系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#66CCCC" background="images/backgrand.jpg" onresize=hero(); scroll="no"> 
<%
  if Request("errMessage") <> null or Request("errMessage") <> "" then
    Response.Write("<center><font color=red>" & Request("errMessage") & "</font></center>")
	end if
%>
<form action="register.asp" method="post">
<table border=0 cellpadding=0 cellspacing=0 bordercolor=lightgreen align="center" width=350>
<tr><td colspan=2 align="center"><font color="green">新用户注册</font></td></tr>
<tr><td>用户名称:</td><td><input type="text" name="studentname" class=input maxlength=14 size="16"></td></tr>
<tr><td>用户密码:</td><td><input type="password" name="studentpassword"  class=input maxlength=12 size="16"></td></tr>
<tr><td>密码确认:</td><td><input type="password" name="studentpassword2"  class=input maxlength=12 size="16"></td></tr>
<tr><td colspan=2 align="center">
 <input type=submit name="submit" value="注册" class=button>
</td></tr>
<tr><td colspan=2 align="center"><a href="lo.asp"><font color=black size=+0>返回在线考试系统登录界面</font></a></td></tr>
</table>
</form>
</body>
</html>