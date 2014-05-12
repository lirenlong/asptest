<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim id'定义变量，用户的id
dim sql,rs,rsc
    if request("action")="del" then   '删除纪录
	sql="delete from score where id=" &request("id")
	conn.execute sql
	if err.number <> 0 then
		response.write "数据库操作错误：" + err.description
		err.clear
	else %>
        <script language=vbscript>
		msgbox "操作成功!号码为<%=trim(request("id"))%>的考试纪录已删除!"
		</script>
<%  end if
end if
%>
<html>
<head>
<title>管理考试成绩----在线考试系统</title>
<script language=javascript>
function SureDel(id)
{
    if ( confirm("您确定要删除该考试纪录吗？"))
        {
            window.location.href = "mgscore.asp?action=del&id=" +id
        }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background=file:///C|/Documents and Settings/Administrator/桌面/examlxc/images/88.jpg > 
<center><font color="lime" size=+3>考试成绩管理</font></center>
<table width=624 border="1" cellspacing="0" cellpadding="0" align="center" bordercolor=lightgreen>
  <tr> 
    <td width="20%"> 
      <div align="center">学生姓名</div>
    </td>
    <td width="20%"> 
      <div align="center">考试科目</div>
    </td>
    <td width="20%"> 
      <div align="center">考试时间</div>
    </td>
    <td width="20%"> 
      <div align="center">考试分数</div>
    </td>
    <td width="20%"> 
      <div align="center">操作</div>
    </td>
  </tr>
  <%
  set rs=server.createobject("adodb.recordset")
  rs.open "select * from score order by studentname ",conn,1,1
  if err.number <> 0 then
	           response.write "数据库出错"
           else
	           if rs.bof and rs.eof then
		           rs.close
		           response.write "目前没有考试纪录"
	           else
			       do while not rs.eof %>
  <tr> 
    <td width="20%" height="126" > 
      <div align="center"><%=rs("studentname")%></div>
    </td>
    <td width="20%" height="126" > 
      <div align="center"><%=rs("subjectname")%></div>
    </td>
    <td width="20%" height="126" > 
      <div align="center"><%=rs("endtime")%></div>
    </td>
    <td width="20%" height="126" > 
      <div align="center"><%=rs("score")%></div>
    </td>
    <td width="20%" height="126" >
      <div align="center">
        <% 					 
	      response.write "<a href='javascript:SureDel(" & cstr(rs("id")) & ")'>删除</a>"		     
	 %>
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
 <p align=center><a href="primarypage.asp"><font color=red size=+0 face=楷体>返回管理界面</font></a></p>
</body>
</html>