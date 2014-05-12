<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
  dim n
  dim rs,sql
  n=2
%>
<html>
<head>
<title>在线考试系统管理界面</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background="file:///C|/Documents and Settings/Administrator/桌面/examlxc/images/backgrand.jpg">
<center>
<table border="0" cellspacing="0" cellpadding="0" width="309" height="156">
<td> 
        <p>
          <br>
	    <p ><font size=+2 color=green face=宋体><center>管理页面菜单</center></font></p>
		  <script language="JavaScript1.2" src="menu.js"></script>
        </p>
        <div id="KB1Parent" class="parent"><font size=+1 face=宋体><a href="#" onClick="expandIt('KB1'); return false" onMouseOver="window.status='管理学生和科目';return true;" onMouseOut="window.status='';return true;"><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0>管理员学生和科目</a></font></div>
        <div id="KB1Child" class="child">
        <a href="mgstudent.asp"  onMouseOver="window.status='管理学生';return true;" onMouseOut="window.status='';return true;"><img src="../images/bag.gif" width=20 height=11 border=0 alt=""><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><font size=+0 color=green face=宋体>管理学生</font></a><br>
        <a href="mgadmin.asp"  onMouseOver="window.status='管理管理员';return true;" onMouseOut="window.status='';return true;"><img src="../images/bag.gif" width=20 height=11 border=0 alt=""><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><font size=+0 color=green face=宋体>管理管理员</font></a><br>
        <a href="mgsubject.asp"  onMouseOver="window.status='管理考试科目';return true;" onMouseOut="window.status='';return true;"><img src="../images/bag.gif" width=20 height=11 border=0 alt=""><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><font size=+0 color=green face=宋体>管理考试科目</font></a><br> 
        <a href="mgscore.asp"  onMouseOver="window.status='查看及管理学生考分';return true;" onMouseOut="window.status='';return true;"><img src="../images/bag.gif" width=20 height=11 border=0 alt=""><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><font size=+0 color=green face=宋体>查看及管理学生考分</font></a><br></div>      
    <% set rs = server.createobject("adodb.recordset")
	   rs.open "select * from subject",conn,1,1
       if err.number <> 0 then
	       response.write "数据库出错"
       else
	       if rs.bof and rs.eof then
		       rs.close
		       response.write "没有科目"
			   response.end
	       else
			   do while not rs.eof
		   %>
		   <div id="KB<%=n%>Child" class="parent"><font size=+1><a href="mgquestion.asp?subjectname=<%=rs("subjectname")%>" onMouseOver="window.status='<%=rs("subjectname") %>';return true;" onMouseOut="window.status='';return true;"><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><%=rs("subjectname") %></a></font><br></div>
		          <% rs.movenext
			 n=n+1
			 loop
		rs.close
		set rs=nothing
	  end if
  end if %>
                  <!--刷新和退出-->
<div id="KB<%=n%>Parent" class="parent"><a href="primarypage.asp" target="_top" onMouseOver="window.status='刷新页面';return true;" onMouseOut="window.status='';return true;"><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><font size=+1 face=宋体>刷新页面</font></a></div>
<div id="KB<%=n+1%>Parent" class="parent"><font size=+1 face=宋体><a href="login.asp"  onMouseOver="window.status='退  出';return true;" onMouseOut="window.status='';return true;"><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0>退  出</a></font></div>

<script language="JavaScript">
if (NS4) {
        firstEl = "KB1Parent";
        firstInd = getIndex(firstEl);
        arrange();
}
</script>
</td>
</table>
</center>
</body>
</html>