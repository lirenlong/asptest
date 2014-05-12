<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim isedit '是否在编辑状态
dim color  '表格颜色
dim id,name'定义变量，用户的id
dim sql,rs,rsc
color=1
isedit=false
if request("action")="edit" then
    isedit=true
end if
if request("action")="edit" then   '修改管理员
    if trim(request("password"))="" then
	    response.write "错误!密码不能为空! <a href=mgadmin.asp>返回</a>"
        response.end
    end if
	sql="update admin set name='" & cstr(trim(request("name"))) & "',password='" & cstr(trim(request("password")))
	conn.execute sql
	if err.number <> 0 then
	    response.write "数据库操作出错:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "操作成功!管理员<%=trim(request("name"))%>的信息已经更新!"
		</script>
  <%end if
end if
if request("action")="add" then   '添加新管理员
    if trim(request("name"))="" or trim(request("password"))="" then
	    response.write "错误!用户名或密码不能为空! <a href=# onclick='javascript:window.history.go(-1)'>返回</a>"
        response.end
    end if
	set rs=server.createobject("adodb.recordset")   '检查学生是否重名
    rs.open "select * from admin where name='" & cstr(trim(request("name"))) & "'",conn,1,1
    if err.number <> 0 then
	          response.write "数据库出错"
    else  if not rs.bof and not rs.eof then
	          response.write "错误!该管理员存在! <a href=# onclick='javascript:window.history.go(-1)'>返回</a>"
              response.end
          end if
    end if
	rs.close
	set rs=nothing
	sql="insert into admin(name,password) values('" & cstr(trim(request("name"))) & "','" & cstr(trim(request("password"))) & "')"
	conn.execute sql
	if err.number <> 0 then
	    response.write "数据库操作出错:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "操作成功!新管理员<%=trim(request("name"))%>的信息添加成功!"
		</script>
  <%end if
end if
if request("action")="del" then   '删除管理员
	sql="delete from admin where id=" &request("id")
	conn.execute sql
	if err.number <> 0 then
		response.write "数据库操作错误：" + err.description
		err.clear
	else %>
        <script language=vbscript>
		msgbox "操作成功!管理员<%=trim(request("name"))%> 的信息已删除!"
		</script>
<%  end if
end if
%>
<html>
<head>
<title>管理管理员----在线考试系统</title>
<script language=javascript>
function SureDel(id)
{
    if ( confirm("您确定要删除该管理员吗？"))
        {
            window.location.href = "mgadmin.asp?action=del&id=" +id
        }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background=file:///C|/Documents and Settings/Administrator/桌面/examlxc/images/88.jpg > 
<center>
  <font color="lime" size=+3>教师管理</font>
</center>
<table width=442 height="86" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor=lightgreen>
  <tr> 
    <td width="25%"> 
      <div align="center">管理员姓名</div>
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
  rs.open "select * from admin ",conn,1,1
  if err.number <> 0 then
	           response.write "数据库出错"
           else
	           if rs.bof and rs.eof then
		           rs.close
		           response.write "目前没有管理员"
	           else
			       do while not rs.eof %>
  <tr> 
    <td width="25%" height="21" > 
      <div align="center"><%=rs("name")%></div>
    </td>
    <td width="20%" height="21" > 
      <div align="center"><%=rs("password")%></div>
    </td>
    <td width="20%" height="21" >
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
    <%  response.write "<font size=3>添加新的管理员</font><br>" %> 
<form action="mgadmin.asp" method="post">
	    <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else  %>add<% End If %>'>
		<%If isedit then%>
              <input type="Hidden" name="name" value='<%=cstr(request("name"))%>'>
        <%End If%>
	    用户名称:<input type="text" name="name" class=input maxlength=14 size="16"><br>
	    用户密码:
	    <input type="password" name="password"  class=input maxlength=12 size="16"><br>
	            <input type=submit value="确 定" class=button>
                <p align=center><a href="primarypage.asp"><font color=red size=+0 face=楷体></font></a></p>
                <div align="center"><a href="primarypage.asp"><font color=red size=+0 face=楷体>返回管理界面</font></a>
                </div>
</form>
     </p>
</body>
</html>