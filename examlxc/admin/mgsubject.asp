<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim isedit '�Ƿ��ڱ༭״̬
dim id,subjectname'�����������Ŀ��id
dim sql,rs,rsc
isedit=false
if request("action")="edit" then
    isedit=true
end if
if request("action")="modify" then   '�޸��û�
    if trim(request("subjectname"))="" or trim(request("testtime"))="" or trim(request("multinumber"))=""or trim(request("multiper"))=""or trim(request("singlenumber"))=""or trim(request("singleper"))=""then
	    response.write "����!����ȷ��д����Ҳ���Ϊ��! <a href=mgsubject.asp>����</a>"
        response.end
    end if
	sql="update subject set subjectname='" & cstr(trim(request("subjectname"))) & "',testtime=" & cstr(trim(request("testtime")))&","&cstr(trim(request("singlenumber"))) & "," & cstr(trim(request("singleper"))) & "," & cstr(trim(request("multinumber"))) & "," & cstr(trim(request("multiper")))
	conn.execute sql
	if err.number <> 0 then
	    response.write "���ݿ��������:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "�����ɹ�!<%=trim(request("subjectname"))%>��Ŀ����Ϣ�Ѿ�����!"
		</script>
  <%end if
end if
if request("action")="add" then   '������û�
    if trim(request("subjectname"))="" or trim(request("testtime"))="" or trim(request("multinumber"))=""or trim(request("multiper"))=""or trim(request("singlenumber"))=""or trim(request("singleper"))=""then
	    response.write "����!��Ŀ���������Լ���������Ϊ��! <a href=# onclick='javascript:window.history.go(-1)'>����</a>"
        response.end
    end if
	set rs=server.createobject("adodb.recordset")   '����Ŀ���Ƿ�����
    rs.open "select * from subject where subjectname='" & cstr(trim(request("subjectname"))) & "'",conn,1,1
    if err.number <> 0 then
	          response.write "���ݿ����"
    else  if not rs.bof and not rs.eof then
	          response.write "����!�ÿ�Ŀ�Ѿ�����! <a href=# onclick='javascript:window.history.go(-1)'>����</a>"
              response.end
          end if
    end if
	rs.close
	set rs=nothing
	sql="insert into subject(subjectname,testtime,singlenumber,singleper,multinumber,multiper) values('" & cstr(trim(request("subjectname"))) & "'," & cstr(trim(request("testtime"))) & "," & cstr(trim(request("singlenumber"))) & "," & cstr(trim(request("singleper"))) & "," & cstr(trim(request("multinumber"))) & "," & cstr(trim(request("multiper"))) & ")"
	conn.execute sql
	if err.number <> 0 then
	    response.write "���ݿ��������:" + err.description
	else %>
	    <script language=vbscript>
			msgbox "�����ɹ�!�¿�Ŀ<%=trim(request("subjectname"))%>����Ϣ��ӳɹ�!"
		</script>
  <%end if
end if
if request("action")="del" then   'ɾ���û�
	sql="delete from subject where id=" &request("id")
	conn.execute sql
	if err.number <> 0 then
		response.write "���ݿ��������" + err.description
		err.clear
	else %>
        <script language=vbscript>
		msgbox "�����ɹ�!��Ŀ<%=trim(request("subjectname"))%>����Ϣ��ɾ��!"
		</script>
<%  end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����Ŀ----���߿���ϵͳ</title>
<script language=javascript>
function SureDel(id)
{
    if ( confirm("��ȷ��Ҫɾ���ÿ�Ŀ��"))
        {
            window.location.href = "mgsubject.asp?action=del&id=" +id
        }
}
</script>
</head>
<body background=file:///C|/Documents and Settings/Administrator/����/examlxc/images/88.jpg > 
<center><font color="lime" size=+3>��Ŀ����</font></center>
<table width=621 border="1" cellspacing="0" cellpadding="0" align="center" bordercolor=lightgreen>
  <tr> 
    <td width="20%"> 
      <div align="center">��Ŀ����</div>
    </td>
    <td width="20%"> 
      <div align="center">����ʱ��(����)</div>
    </td>
    <td width="12%"> 
      <div align="center">��ѡ����</div>
    </td>
    <td width="12%"> 
      <div align="center">��ѡ��ֵ</div>
    </td>
    <td width="12%"> 
      <div align="center">��ѡ����</div>
    </td>
    <td width="12%"> 
      <div align="center">��ѡ��ֵ</div>
    </td>
    <td width="20%"> 
      <div align="center">����</div>
    </td>
  </tr>
  <%
  set rs=server.createobject("adodb.recordset")
  rs.open "select * from subject ",conn,1,1
  if err.number <> 0 then
	           response.write "���ݿ����"
           else
	           if rs.bof and rs.eof then
		           rs.close
		           response.write "Ŀǰû�п�Ŀ"
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
    <a href='javascript:SureDel(<%=rs("id") %>)'>ɾ��</a></div></td>
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
    <%  response.write "<font size=3>�� �� �� �� �� Ŀ</font><br>" %> 
<form action="mgsubject.asp" method="post">
	    <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else  %>add<% End If %>'>
		<%If isedit then%>
              <input type="Hidden" name="subjectname" value='<%=cstr(request("subjectname"))%>'>
        <%End If%>
	    ��Ŀ����:<input type="text" name="subjectname"  value='<% if isedit then response.write trim(rs("subjectname")) end if %>'><br>
	    ����ʱ��:<input type="text" name="testtime"  value='<% if isedit then response.write trim(rs("testtime")) end if %>'><br>
	    ��ѡ����:<input type="text" name="singlenumber" value='<% if isedit then response.write trim(rs("singlenumber")) end if %>'><br>
	    ��ѡ��ֵ:<input type="text" name="singleper"  value='<% if isedit then response.write trim(rs("singleper")) end if %>'><br>
	    ��ѡ����:<input type="text" name="multinumber" value='<% if isedit then response.write trim(rs("multiumber")) end if %>'><br>
	    ��ѡ��ֵ:<input type="text" name="multiper" value='<% if isedit then response.write trim(rs("multiper")) end if %>'><br>
	            <input type=submit value="ȷ ��">
</form>
     <p align=center><a href="primarypage.asp"><font color=red size=+0 face=����>���ع������</font></a></p>
 </p>
</body>
</html>