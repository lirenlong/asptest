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
<title>���߿���ϵͳ�������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background="file:///C|/Documents and Settings/Administrator/����/examlxc/images/backgrand.jpg">
<center>
<table border="0" cellspacing="0" cellpadding="0" width="309" height="156">
<td> 
        <p>
          <br>
	    <p ><font size=+2 color=green face=����><center>����ҳ��˵�</center></font></p>
		  <script language="JavaScript1.2" src="menu.js"></script>
        </p>
        <div id="KB1Parent" class="parent"><font size=+1 face=����><a href="#" onClick="expandIt('KB1'); return false" onMouseOver="window.status='����ѧ���Ϳ�Ŀ';return true;" onMouseOut="window.status='';return true;"><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0>����Աѧ���Ϳ�Ŀ</a></font></div>
        <div id="KB1Child" class="child">
        <a href="mgstudent.asp"  onMouseOver="window.status='����ѧ��';return true;" onMouseOut="window.status='';return true;"><img src="../images/bag.gif" width=20 height=11 border=0 alt=""><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><font size=+0 color=green face=����>����ѧ��</font></a><br>
        <a href="mgadmin.asp"  onMouseOver="window.status='�������Ա';return true;" onMouseOut="window.status='';return true;"><img src="../images/bag.gif" width=20 height=11 border=0 alt=""><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><font size=+0 color=green face=����>�������Ա</font></a><br>
        <a href="mgsubject.asp"  onMouseOver="window.status='�����Կ�Ŀ';return true;" onMouseOut="window.status='';return true;"><img src="../images/bag.gif" width=20 height=11 border=0 alt=""><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><font size=+0 color=green face=����>�����Կ�Ŀ</font></a><br> 
        <a href="mgscore.asp"  onMouseOver="window.status='�鿴������ѧ������';return true;" onMouseOut="window.status='';return true;"><img src="../images/bag.gif" width=20 height=11 border=0 alt=""><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><font size=+0 color=green face=����>�鿴������ѧ������</font></a><br></div>      
    <% set rs = server.createobject("adodb.recordset")
	   rs.open "select * from subject",conn,1,1
       if err.number <> 0 then
	       response.write "���ݿ����"
       else
	       if rs.bof and rs.eof then
		       rs.close
		       response.write "û�п�Ŀ"
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
                  <!--ˢ�º��˳�-->
<div id="KB<%=n%>Parent" class="parent"><a href="primarypage.asp" target="_top" onMouseOver="window.status='ˢ��ҳ��';return true;" onMouseOut="window.status='';return true;"><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0><font size=+1 face=����>ˢ��ҳ��</font></a></div>
<div id="KB<%=n+1%>Parent" class="parent"><font size=+1 face=����><a href="login.asp"  onMouseOver="window.status='��  ��';return true;" onMouseOut="window.status='';return true;"><IMG SRC="../images/pagepic.gif" WIDTH=18 HEIGHT=16 BORDER=0>��  ��</a></font></div>

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