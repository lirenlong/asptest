<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim isedit '是否在编辑状态
dim id,subjectname'定义变量，科目的id
dim sql,rs,rsc
isedit=false
if request("action")="edit" then
    isedit=true
end if
if request("action")="modify" then   '修改用户
    if trim(request("subjectname"))="" or trim(request("testtime"))="" or trim(request("multinumber"))=""or trim(request("multiper"))=""or trim(request("singlenumber"))=""or trim(request("singleper"))=""then
	    response.write "错误!请正确填写各项，且不能为空! <a href=mgsubject.asp>返回</a>"
        response.end
    end if
	sql="update subject set subjectname='" & cstr(trim(request("subjectname"))) & "',testtime=" & cstr(trim(request("testtime")))&","&cstr(trim(request("singlenumber"))) & "," & cstr(trim(request("singleper"))) & "," & cstr(trim(request("multinumber"))) & "," & cstr(trim(request("multiper")))
	conn.execute sql
	if err.number <> 0 then
	    response.write "数据库操作出错:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "操作成功!<%=trim(request("subjectname"))%>科目的信息已经更新!"
		</script>
  <%end if
end if
if request("action")="add" then   '添加新用户
    if trim(request("subjectname"))="" or trim(request("testtime"))="" or trim(request("multinumber"))=""or trim(request("multiper"))=""or trim(request("singlenumber"))=""or trim(request("singleper"))=""then
	    response.write "错误!科目名或密码以及其余各项不能为空! <a href=# onclick='javascript:window.history.go(-1)'>返回</a>"
        response.end
    end if
	set rs=server.createobject("adodb.recordset")   '检查科目名是否重名
    rs.open "select * from subject where subjectname='" & cstr(trim(request("subjectname"))) & "'",conn,1,1
    if err.number <> 0 then
	          response.write "数据库出错"
    else  if not rs.bof and not rs.eof then
	          response.write "错误!该科目已经存在! <a href=# onclick='javascript:window.history.go(-1)'>返回</a>"
              response.end
          end if
    end if
	rs.close
	set rs=nothing
	sql="insert into subject(subjectname,testtime,singlenumber,singleper,multinumber,multiper) values('" & cstr(trim(request("subjectname"))) & "'," & cstr(trim(request("testtime"))) & "," & cstr(trim(request("singlenumber"))) & "," & cstr(trim(request("singleper"))) & "," & cstr(trim(request("multinumber"))) & "," & cstr(trim(request("multiper"))) & ")"
	conn.execute sql
	if err.number <> 0 then
	    response.write "数据库操作出错:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "操作成功!新科目<%=trim(request("subjectname"))%>的信息添加成功!"
		</script>
  <%end if
end if
if request("action")="del" then   '删除用户
	sql="delete from subject where id=" &request("id")
	conn.execute sql
	if err.number <> 0 then
		response.write "数据库操作错误：" + err.description
		err.clear
	else %>
        <script language=vbscript>
		msgbox "操作成功!科目<%=trim(request("subjectname"))%>的信息已删除!"
		</script>
<%  end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理科目----在线考试系统</title>
<script language=javascript>
function SureDel(id)
{
    if ( confirm("您确定要删除该科目吗？"))
        {
            window.location.href = "mgsubject.asp?action=del&id=" +id
        }
}
</script>
</head>
<body background=file:///C|/Documents and Settings/Administrator/桌面/examlxc/images/88.jpg > 
<center><font color="lime" size=+3>科目管理</font></center>
<table width=621 border="1" cellspacing="0" cellpadding="0" align="center" bordercolor=lightgreen>
  <tr> 
    <td width="20%"> 
      <div align="center">科目名称</div>
    </td>
    <td width="20%"> 
      <div align="center">考试时间(分钟)</div>
    </td>
    <td width="12%"> 
      <div align="center">单选题量</div>
    </td>
    <td width="12%"> 
      <div align="center">单选分值</div>
    </td>
    <td width="12%"> 
      <div align="center">多选题量</div>
    </td>
    <td width="12%"> 
      <div align="center">多选分值</div>
    </td>
    <td width="20%"> 
      <div align="center">操作</div>
    </td>
  </tr>
  <%
  set rs=server.createobject("adodb.recordset")
  rs.open "select * from subject ",conn,1,1
  if err.number <> 0 then
	           response.write "数据库出错"
           else
	           if rs.bof and rs.eof then
		           rs.close
		           response.write "目前没有科目"
	           else
			       do while not rs.eof %>
  <tr> 
    <td width="20%" height="88" > 
      <div align="center"><%=rs("subjectname")%></div>
    </td>
    <td width="12%" height="88" > 
      <div align="center"><%=rs("testtime")%></div>
    </td>
    <td width="12%" height="88" > 
      <div align="center"><%=rs("singlenumber")%></div>
    </td>
    <td width="12%" height="88" > 
      <div align="center"><%=rs("singleper")%></div>
    </td>
    <td width="12%" height="88" > 
      <div align="center"><%=rs("multinumber")%></div>
    </td>
    <td width="12%" height="88" > 
      <div align="center"><%=rs("multiper")%></div>
    </td>
    <td width="20%" height="88" >
      <div align="center">
    <a href='javascript:SureDel(<%=rs("id") %>)'>删除</a></div></td>
      </div>
    </td>
  </tr>
  <% rs.movenext				   
	loop			   
	end if	       
       end if 
	'rs.close
	'set rs=nothing %>
</table>
 <p align="center"> 
    <%  response.write "<font size=3>添 加 新 的 科 目</font><br>" %> 
<form action="mgsubject.asp" method="post">
	    <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else  %>add<% End If %>'>
		<%If isedit then%>
              <input type="Hidden" name="subjectname" value='<%=cstr(request("subjectname"))%>'>
        <%End If%>
	    科目名称:<input type="text" name="subjectname"  value='<% if isedit then response.write trim(rs("subjectname")) end if %>'><br>
	    考试时间:<input type="text" name="testtime"  value='<% if isedit then response.write trim(rs("testtime")) end if %>'><br>
	    单选题量:<input type="text" name="singlenumber" value='<% if isedit then response.write trim(rs("singlenumber")) end if %>'><br>
	    单选分值:<input type="text" name="singleper"  value='<% if isedit then response.write trim(rs("singleper")) end if %>'><br>
	    多选题量:<input type="text" name="multinumber" value='<% if isedit then response.write trim(rs("multiumber")) end if %>'><br>
	    多选分值:<input type="text" name="multiper" value='<% if isedit then response.write trim(rs("multiper")) end if %>'><br>
	            <input type=submit value="确 定">
</form>
     <p align=center><a href="primarypage.asp"><font color=red size=+0 face=楷体>返回管理界面</font></a></p>
 </p>
</body>
</html>