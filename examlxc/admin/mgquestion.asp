<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
  dim isedit  '�Ƿ����޸�״̬
  dim color  '�����ɫ
  dim sql,rs
  dim subjectname
  dim number  'ÿҳ��ʾ��������Ŀ
  dim curpage, i,page  
  subjectname=trim(request("subjectname"))
  color=1
  function invert(str) 
    invert=replace(replace(replace(replace(str,"&lt;","<"),"&gt;",">"),"<br>",chr(13)),"&nbsp;"," ")
  end function
  number=5  '��ʾ������Ĭ��ֵ
  isedit=false
  if request("action")="edit" then
      isedit=true
  end if
  if request("action")="del" then  'ɾ��
     sql="delete from question where id=" &request("id")
	 conn.execute sql
     %>
     <script language=vbscript>
	     msgbox "�����ɹ�!!��������ɾ��!" 
     </script>
<% end if
%>
<html>
<head>
<title>��������----���߿���ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background=file:///C|/Documents and Settings/Administrator/����/examlxc/images/88.jpg > 
<center><font color="lime" size=+3>�������</font></center>
<script language=javascript>
function SureDel(id,subjectname)
{
    if ( confirm("���Ƿ����Ҫɾ�������⣿"))   
        {
            window.location.href = "mgquestion.asp?action=del&id="+id+"&subjectname="+subjectname
        }
}
</script>
 <%
  set rs=server.createobject("adodb.recordset")
  rs.open "select * from question where subjectname='" & cstr(trim(request("subjectname"))) & "' order by id desc ",conn,1,1
  if err.number <> 0 then
	           response.write "���ݿ����"
           else
	           if rs.bof and rs.eof then
		           rs.close
		           response.write "�ÿ�Ŀû������"
	           else
			        %>
                   <table width=800 border="1" cellspacing="0" cellpadding="0" align="center" height="44" bordercolor=lightgreen>
                       <tr>
                          <td width="20%"><div align="center">����</div></td>
                          <td width="10%"> <div align="center">ѡ��A</div></td>
                          <td width="10%"> <div align="center">ѡ��B</div></td>	  
			  <td width="10%"> <div align="center">ѡ��C</div></td>
			  <td width="10%"> <div align="center">ѡ��D</div></td>
			  <td width="10%"> <div align="center">��</div></td>
			  <td width="10%"> <div align="center">����</div></td>
			  <td width="10%"> <div align="center">��Ŀ</div></td>
			  <td width="10%"> <div align="center">����</div></td>
                       </tr> <%
               if request("page")="" then
  	               curpage = 1
               else
	               curpage = cint(request("page"))
               end if

               rs.pagesize=cint(number)
               rs.absolutepage = curpage
               for i = 1 to rs.pagesize
                   %><tr>
                   <td width="20%" height="23" >
                   <div align="center"><%=rs("question")%></div></td>
                   <td width="10%" height="23" > 
                       <div align="center"><%=rs("A")%></div></td>  
                   <td width="10%" height="23" > 
                       <div align="center"><%=rs("B")%></div></td>
                   <td width="10%" height="23" > 
                       <div align="center"><%=rs("C")%></div></td> 
                   <td width="10%" height="23" > 
                       <div align="center"><%=rs("D")%></div></td> 
                   <td width="10%" height="23" > 
                       <div align="center"><%=rs("answer")%></div></td> 
                   <td width="10%" height="23" > 
                       <div align="center"><%=rs("type")%></div></td> 
                   <td width="10%" height="23" > 
                       <div align="center"><%=rs("subjectname")%></div></td>
                   <td width=10%" height="23" > 
                       <div align="center">		  
		       <a href='mgquestion.asp?type=<%=trim(rs("type"))%>&subjectname=<%=trim(rs("subjectname"))%>&action=edit&id=<%= trim(rs("id"))%>&page=<%=request("page")%>'>�༭</a>&nbsp<a href='javascript:SureDel(<%=rs("id") %>)'>ɾ��</a></div></td>
                 </tr>
 <% rs.movenext
    color=color+1
      if rs.eof then
	      i = i + 1
	      exit for
     end if
               next %>
</table>
 
	      <% response.write "<hr size=0 width='100%'><div align=center>"
		response.write "��<font color=red>" + cstr(curpage) + "</font>ҳ/��<font color=red>" + cstr(rs.pagecount) + "</font>ҳ "
		response.write "��ҳ<font color=red>" + cstr(i-1) + "</font>��/��<font color=red>" + cstr(rs.recordcount) + "</font>�� "
		if curpage = 1 then 
			
		else
			response.write "<a href='mgquestion.asp?type=" & cstr(request("type")) & "&subjectname=" & cstr(request("subjectname")) & "&page=1'>��ҳ</a> <a href='mgquestion.asp?type=" & cstr(request("type"))  & "&subjectname=" & cstr(request("subjectname"))& "&page=" & cstr(curpage-1) & "'>ǰҳ</a> "
		end if
		if  curpage = rs.pagecount then
			
		else
			response.write "<a href='mgquestion.asp?type=" & cstr(request("type"))& "&subjectname=" & cstr(request("subjectname")) & "&page=" + cstr(curpage+1) + "'>��ҳ</a> <a href='mgquestion.asp?subjectname=" & cstr(request("subjectname")) + "&page=" + cstr(rs.pagecount) + "'>ĩҳ</a>"
		end if
	end If
end if
 'rs.close
set rs=nothing
 %>
<hr size=0 width=100%></div> 
 <%  if isedit then
	   set rs=server.createobject("adodb.recordset")
	   rs.open "select * from question where id=" & cstr(request("id")),conn,1,1
	   response.write "<p align='center'><font size=3>�� �� �� ��</font></p>"
   else
	   response.write "<p align='center'><font size=3>�� �� �� ��</font></p>"
   end if %>
 <form action="addquestion.asp"  method="post">
	<input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'>
        <input type="Hidden" name="id" value='<%=request("id")%>'>
	<input type="Hidden" name="page" value='<%=request("page")%>'>

     <p align=center><font color=red size=+0>��*�����������ȫ��д</font></p>
	 
  <p align="left"><font color=red>*</font>����:
     <input type="text" name="question" class=input maxlength=100 size="50" value='<% if isedit then response.write trim(rs("question")) end if %>'>
  </p> 
  <p align="left"><font color=red>*</font>��Ŀ:
     <input type="text" name="subjectname" class=input maxlength=30 size="10" value='<% if isedit then  response.write trim(rs("subjectname")) else response.write trim(request("subjectname")) end if %>'>
  </p> 
  <p align="left">
     ѡ��A:<input type="text" name="A" class=input maxlength=15 size="10" value='<% if isedit then response.write trim(rs("A")) end if %>'>
     ѡ��B<input type="text" name="B" class=input maxlength=15 size="10" value='<% if isedit then response.write trim(rs("B")) end if %>'>
     ѡ��C<input type="text" name="C" class=input maxlength=15 size="10" value='<% if isedit then response.write trim(rs("C")) end if %>'>
     ѡ��D<input type="text" name="D" class=input maxlength=15 size="10" value='<% if isedit then response.write trim(rs("D")) end if %>'>
  </p> 
  <p align="left"><font color=red>*</font>�𰸣�����дѡ�����ĸ��:
     <input type="text" name="answer" class=input maxlength=15 size="10" value='<% if isedit then response.write trim(rs("answer")) end if %>'>
  </p>
  <p align="left"><font color=red>*</font>���ͣ���ѡ�⻹�Ƕ�ѡ�⣩:<input type="radio" name="leixing" value="��ѡ��" 
  <% if isedit then
     if  rs("type")="��ѡ��" then
	       response.write "checked"
	   end if
	end if  %>>��ѡ��&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
	 <input type="radio" name="leixing" value="��ѡ��" 
	<% if isedit then
	      if rs("type")="��ѡ��" then
	         response.write "checked"
	    end if
	  end if
         %>>��ѡ��</p>  
	 <p align="center"><input type=submit value="  ȷ  ��  " class=button></p>
 </form>
 <p align=center><a href="primarypage.asp"><font color=red size=+0 face=����>���ع������</font></a></p>
</body>
</html>