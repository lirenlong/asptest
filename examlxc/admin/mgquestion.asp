<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
  dim isedit  '是否在修改状态
  dim color  '表格颜色
  dim sql,rs
  dim subjectname
  dim number  '每页显示的文章数目
  dim curpage, i,page  
  subjectname=trim(request("subjectname"))
  color=1
  function invert(str) 
    invert=replace(replace(replace(replace(str,"&lt;","<"),"&gt;",">"),"<br>",chr(13)),"&nbsp;"," ")
  end function
  number=5  '显示试题数默认值
  isedit=false
  if request("action")="edit" then
      isedit=true
  end if
  if request("action")="del" then  '删除
     sql="delete from question where id=" &request("id")
	 conn.execute sql
     %>
     <script language=vbscript>
	     msgbox "操作成功!!该试题已删除!" 
     </script>
<% end if
%>
<html>
<head>
<title>管理试题----在线考试系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background=file:///C|/Documents and Settings/Administrator/桌面/examlxc/images/88.jpg > 
<center><font color="lime" size=+3>试题管理</font></center>
<script language=javascript>
function SureDel(id,subjectname)
{
    if ( confirm("你是否真的要删除该试题？"))   
        {
            window.location.href = "mgquestion.asp?action=del&id="+id+"&subjectname="+subjectname
        }
}
</script>
 <%
  set rs=server.createobject("adodb.recordset")
  rs.open "select * from question where subjectname='" & cstr(trim(request("subjectname"))) & "' order by id desc ",conn,1,1
  if err.number <> 0 then
	           response.write "数据库出错"
           else
	           if rs.bof and rs.eof then
		           rs.close
		           response.write "该科目没有试题"
	           else
			        %>
                   <table width=800 border="1" cellspacing="0" cellpadding="0" align="center" height="44" bordercolor=lightgreen>
                       <tr>
                          <td width="20%"><div align="center">问题</div></td>
                          <td width="10%"> <div align="center">选项A</div></td>
                          <td width="10%"> <div align="center">选项B</div></td>	  
			  <td width="10%"> <div align="center">选项C</div></td>
			  <td width="10%"> <div align="center">选项D</div></td>
			  <td width="10%"> <div align="center">答案</div></td>
			  <td width="10%"> <div align="center">题型</div></td>
			  <td width="10%"> <div align="center">科目</div></td>
			  <td width="10%"> <div align="center">操作</div></td>
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
		       <a href='mgquestion.asp?type=<%=trim(rs("type"))%>&subjectname=<%=trim(rs("subjectname"))%>&action=edit&id=<%= trim(rs("id"))%>&page=<%=request("page")%>'>编辑</a>&nbsp<a href='javascript:SureDel(<%=rs("id") %>)'>删除</a></div></td>
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
		response.write "第<font color=red>" + cstr(curpage) + "</font>页/共<font color=red>" + cstr(rs.pagecount) + "</font>页 "
		response.write "本页<font color=red>" + cstr(i-1) + "</font>条/共<font color=red>" + cstr(rs.recordcount) + "</font>条 "
		if curpage = 1 then 
			
		else
			response.write "<a href='mgquestion.asp?type=" & cstr(request("type")) & "&subjectname=" & cstr(request("subjectname")) & "&page=1'>首页</a> <a href='mgquestion.asp?type=" & cstr(request("type"))  & "&subjectname=" & cstr(request("subjectname"))& "&page=" & cstr(curpage-1) & "'>前页</a> "
		end if
		if  curpage = rs.pagecount then
			
		else
			response.write "<a href='mgquestion.asp?type=" & cstr(request("type"))& "&subjectname=" & cstr(request("subjectname")) & "&page=" + cstr(curpage+1) + "'>后页</a> <a href='mgquestion.asp?subjectname=" & cstr(request("subjectname")) + "&page=" + cstr(rs.pagecount) + "'>末页</a>"
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
	   response.write "<p align='center'><font size=3>编 辑 试 题</font></p>"
   else
	   response.write "<p align='center'><font size=3>添 加 试 题</font></p>"
   end if %>
 <form action="addquestion.asp"  method="post">
	<input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'>
        <input type="Hidden" name="id" value='<%=request("id")%>'>
	<input type="Hidden" name="page" value='<%=request("page")%>'>

     <p align=center><font color=red size=+0>带*各项均必须完全填写</font></p>
	 
  <p align="left"><font color=red>*</font>问题:
     <input type="text" name="question" class=input maxlength=100 size="50" value='<% if isedit then response.write trim(rs("question")) end if %>'>
  </p> 
  <p align="left"><font color=red>*</font>科目:
     <input type="text" name="subjectname" class=input maxlength=30 size="10" value='<% if isedit then  response.write trim(rs("subjectname")) else response.write trim(request("subjectname")) end if %>'>
  </p> 
  <p align="left">
     选项A:<input type="text" name="A" class=input maxlength=15 size="10" value='<% if isedit then response.write trim(rs("A")) end if %>'>
     选项B<input type="text" name="B" class=input maxlength=15 size="10" value='<% if isedit then response.write trim(rs("B")) end if %>'>
     选项C<input type="text" name="C" class=input maxlength=15 size="10" value='<% if isedit then response.write trim(rs("C")) end if %>'>
     选项D<input type="text" name="D" class=input maxlength=15 size="10" value='<% if isedit then response.write trim(rs("D")) end if %>'>
  </p> 
  <p align="left"><font color=red>*</font>答案（请填写选项的字母）:
     <input type="text" name="answer" class=input maxlength=15 size="10" value='<% if isedit then response.write trim(rs("answer")) end if %>'>
  </p>
  <p align="left"><font color=red>*</font>类型（单选题还是多选题）:<input type="radio" name="leixing" value="单选题" 
  <% if isedit then
     if  rs("type")="单选题" then
	       response.write "checked"
	   end if
	end if  %>>单选题&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
	 <input type="radio" name="leixing" value="多选题" 
	<% if isedit then
	      if rs("type")="多选题" then
	         response.write "checked"
	    end if
	  end if
         %>>多选题</p>  
	 <p align="center"><input type=submit value="  确  认  " class=button></p>
 </form>
 <p align=center><a href="primarypage.asp"><font color=red size=+0 face=宋体>返回管理界面</font></a></p>
</body>
</html>