<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="conn.asp"-->
<% 
dim question,subjectname,A,B,C,D,answer,leixing,page,action,rs,id
 function invert(str) 
    invert=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
 end function
     id=trim(request.form("id"))
     action=trim(request.form("action"))
     question=trim(Request.form("question"))
     subjectname=trim(Request.form("subjectname"))
     A=trim(Request.form("A"))
     B=trim(Request.form("B"))
     C=trim(Request.form("C"))
     D=trim(Request.form("D"))
     answer=trim(Request.form("answer")) 
     leixing=trim(Request.form("leixing"))
     'response.write(question)
     page=trim(request("page"))
  if question="" or subjectname="" or answer="" or leixing=""  then
      response.write "错误!!带<font color=red>*</font>号的为必填项!  <a href='javascript:history.go(-1)'>返回</a>"
	  response.end
	  else
  end if
      if action="modify" then '修改问题  
	     set rs=server.createobject("ADODB.recordset") 
         rs.Open "SELECT * from question Where id=" & id,conn,1,3 
	     rs("question")=question
             rs("subjectname")=subjectname
             rs("A")=A
             rs("B")=B
             rs("C")=C
             rs("D")=D
             rs("answer")=answer
             rs("type")=leixing  
             rs("haveselect")=0
                		
			 rs.update
			 rs.close
			 set rs=nothing
	     response.redirect "mgquestion.asp?id=" & id & "&subjectname=" & subjectname & "&page=" & page
     end if 
     if action="add" then '添加新问题
	   set rs=server.createobject("ADODB.recordset") 
           rs.Open "SELECT * from question",conn,1,3 
           rs.addnew     
             rs("question")=question
             rs("subjectname")=subjectname
             rs("A")=A
             rs("B")=B
             rs("C")=C
             rs("D")=D
             rs("answer")=answer
             rs("type")=leixing 
             rs("haveselect")=0
             
          rs.update
          rs.close
	  set rs=nothing
	response.redirect "mgquestion.asp?id=" & id & "&subjectname=" & subjectname & "&page=" & page
     end if  
 %>