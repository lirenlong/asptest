<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim isedit '是否在编辑状态
dim color  '表格颜色
dim id,studentname'定义变量，用户的id
dim sql,rs,rsc
color=1
isedit=false
if request("action")="edit" then
    isedit=true
end if
if request("action")="edit" then   '修改用户
    if trim(request("studentpassword"))="" then
	    response.write "错误!密码不能为空! <a href=mgstudent.asp>返回</a>"
        response.end
    end if
	sql="update student set studentname='" & cstr(trim(request("studentname"))) & "',studentpassword='" & cstr(trim(request("studentpassword")))
	conn.execute sql
	if err.number <> 0 then
	    response.write "数据库操作出错:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "操作成功!用户 <%=trim(request("studentname"))%> 的信息已经更新!"
		</script>
  <%end if
end if
if request("action")="add" then   '添加新用户
    if trim(request("studentname"))="" or trim(request("studentpassword"))="" then
	    response.write "错误!用户名或密码不能为空! <a href=# onclick='javascript:window.history.go(-1)'>返回</a>"
        response.end
    end if
	set rs=server.createobject("adodb.recordset")   '检查学生是否重名
    rs.open "select * from student where studentname='" & cstr(trim(request("studentname"))) & "'",conn,1,1
    if err.number <> 0 then
	          response.write "数据库出错"
    else  if not rs.bof and not rs.eof then
	          response.write "错误!该用学生存在! <a href=# onclick='javascript:window.history.go(-1)'>返回</a>"
              response.end
          end if
    end if
	rs.close
	set rs=nothing
	sql="insert into student(studentname,studentpassword) values('" & cstr(trim(request("studentname"))) & "','" & cstr(trim(request("studentpassword"))) & "')"
	conn.execute sql
	if err.number <> 0 then
	    response.write "数据库操作出错:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "操作成功!新用户 <%=trim(request("studentname"))%> 的信息添加成功!"
		</script>
  <%end if
end if
if request("action")="del" then   '删除用户
	sql="delete from student where id=" &request("id")
	conn.execute sql
	if err.number <> 0 then
		response.write "数据库操作错误：" + err.description
		err.clear
	else %>
        <script language=vbscript>
		msgbox "操作成功!用户 <%=trim(request("studentname"))%> 的信息已删除!"
		</script>
<%  end if
end if
%>
<html>
<head>
<title>管理学生----在线考试系统</title>
<script language=javascript>
function SureDel(id)
{
    if ( confirm("您确定要删除该用户吗？"))
        {
            window.location.href = "mgstudent.asp?action=del&id=" +id
        }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background=file:///C|/Documents and Settings/Administrator/桌面/examlxc/images/88.jpg > 
<center><font color="lime" size=+3>学生管理</font></center>
<table width=528 border="1" cellspacing="0" cellpadding="0" align="center" bordercolor=blue>
  <tr> 
    <td width="25%"> 
      <div align="center">学生姓名</div>
    </td>
    <td width="20%"> 
      <div align="center">密码</div>
    </td>
    <td width="20%"> 
      <div align="center">操作</div>
    </td>
  </tr>
  <%
  set rs=server.createobject("adodb.recordset")
  rs.open "select * from student ",conn,1,1
  if err.number <> 0 then
	           response.write "数据库出错"
           else
	           if rs.bof and rs.eof then
		           rs.close
		           response.write "目前没有学生"
	           else
			       do while not rs.eof %>
  <tr> 
    <td width="25%" height="78" > 
      <div align="center"><%=rs("studentname")%></div>
    </td>
    <td width="20%" height="78" > 
      <div align="center"><%=rs("studentpassword")%></div>
    </td>
    <td width="20%" height="78" >
      <div align="center">
        <% 					 
	      response.write "<a href='javascript:SureDel(" & cstr(rs("id")) & ")'>删除</a>"		     
	 %>
      </div>
    </td>
  </tr>
  <% rs.movenext
	color=color+1				   
	loop			   
	end if	       
       end if 
	'rs.close
	'set rs=nothing %>
</table>
 <p align="center"> 
    <%  response.write "<font size=3>添 加 新 的 学 生</font><br>" %> 
<form action="mgstudent.asp" method="post">
	    <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else  %>add<% End If %>'>
		<%If isedit then%>
              <input type="Hidden" name="studentname" value='<%=cstr(request("studentname"))%>'>
        <%End If%>
	    用户名称:<input type="text" name="studentname" class=input maxlength=14 size="16"><br>
	    用户密码:<input type="password" name="studentpassword"  class=input maxlength=12 size="16"><br>
	            <input type="submit" name="submit" value="确 定" class=button>
</form>
     <p align=center><a href="primarypage.asp"><font color=red size=+0 face=楷体>返回管理界面</font></a></p>
 </p>
</body>
</html>