<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
	if Request.Form("submit")="登    录" then
		'学生登录的处理
	  dim rs,sql	
	  set rs = server.createobject("adodb.recordset")
		sql="select * from student where studentname='" & Request.Form("studentname") & "' and studentpassword='" & Request.Form("studentpassword") & "'"
		rs.open sql,conn,1,1
		if err.number<>0 then 
		  response.write "数据库操作失败："&err.description
		else if rs.bof and rs.eof then
			response.write "<center>对不起，请输入正确的用户名和密码。</center>"
		  rs.close		    
		else		
			rs.close
			session("studentname")=request.form("studentname")
			set rs=nothing
			call endConnection()
			Response.Redirect "selectsubject.asp"
		end if
  end if	
		
		'用户注册
  elseif Request.Form("submit")="注    册" then       
		Response.Redirect "register.asp"
		'管理员进行管理
	
  end if
%>

<html>
<head>
<title>在线考试系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" background="images/backgrand.jpg" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<FORM action="lo.asp" method=post id=form name=form>
<table id="__01" width="1005" height="564" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
	<tr>
    
		<td rowspan="6">&nbsp;</td>
		<td background="images/index_03.png" width="284" height="28" align="right"><INPUT id=1 type="text" name=studentname>
	  </td>
		<td rowspan="6">&nbsp;</td>
	</tr>
	<tr>
		<td background="images/index_05.jpg" width="284" height="28" align="right">
			<INPUT id=2 type="password" name=studentpassword ></td>
	</tr>
	<tr>
		<td height="56" align="center">
			<INPUT id=submit1 name=submit type=submit value="登    录">&nbsp;&nbsp;&nbsp;
  <INPUT id=submit2 name=submit type=submit value="注    册">
  </td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td background="images/index_08.png" width="284" height="28" align="center">
			<font size="-1"><a href="admin/login.asp">管理登录</a></font></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
</FORM>
</body>
</html>