<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="file:///D|/Program Files/Tencent/QQ/Users/510890210/FileRecv/examlxc/admin/conn.asp"-->
<% 
	if Request.Form("submit")="登录" then
		'管理员登录的处理
		session("name")=request.form("name")
		session("password")=request.form("password")
		dim rs,sql	
		set rs = server.createobject("adodb.recordset")
		sql="select * from admin where name='" & Request.Form("name") & "' and password='" & Request.Form("password") & "'"
		rs.open sql,conn,1,1
		if err.number<>0 then 
		  response.write "数据库操作失败："&err.description
		elseif rs.bof and rs.eof then
			response.write "<center>对不起，请输入正确的用户名和密码。如果您不是管理员，请退出！</center>"
			rs.close		    
		else		
			rs.close
			session("name")=request.form("name")
			set rs=nothing
			call endConnection()
			Response.Redirect "primarypage.asp"
		end if
	elseif Request.Form("submit")="退出" then
	  Response.Redirect "../index.asp"
  end if	
%>
<HTML>
<HEAD>
<title>管理员登陆----在线考试系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></HEAD>
<BODY bgcolor="#66CCCC" background="file:///C|/Documents and Settings/Administrator/桌面/examlxc/images/backgrand.jpg">
<FORM action="file:///D|/Program Files/Tencent/QQ/Users/510890210/FileRecv/examlxc/admin/login.asp" method=post id=form name=form>
<table align="center" width=314 border=1  cellpadding=0 cellspacing=0>
<tr>
  <td colspan=2 align="center"><p><FONT color=green size=6>教师登录</FONT></p>
    <p>&nbsp;</p>
    <p>&nbsp;</p>
    <p>&nbsp;</p>
    <p>&nbsp;</p></td>
</tr>
<tr><td width="40">姓名:</td>
<td width="268"><input id=1 type=text name=name></td></tr>
<tr><td>密码:</td><td><input id=2 type=password name=password></td></tr>
<tr><td height="144" colspan=2 align="center"><INPUT id=submit1 name=submit type=submit value="登录">
<INPUT id=submit2 name=submit type=submit value="退出"></td></tr>
</table>
</FORM>
</BODY>
</HTML>
