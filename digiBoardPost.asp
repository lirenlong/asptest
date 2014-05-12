<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/llxxcc.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_llxxcc_STRING
  MM_editTable = "board"
  MM_editRedirectUrl = "1.asp"
  MM_fieldsStr  = "digiB_subject|value|digiB_name|value|digiB_face|value|digiB_email|value|digiB_Web|value|digiB_content|value"
  MM_columnsStr = "digiB_subject|',none,''|digiB_name|',none,''|digiB_face|',none,''|digiB_email|',none,''|digiB_web|',none,''|digiB_content|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")

  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_llxxcc_STRING
Recordset1.Source = "SELECT * FROM board"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<html>
<head>
<title>数字留言板</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<style type="text/css">
<!--
.lineBox {
	border: 1px dotted #CCCCCC;
}
.Box {
	border: 1px solid #666666;
}
a:hover.ImgButton {
	left: 2px;
	top: 2px;
	right: 0px;
	bottom: 0px;
	position: relative;
}
form {
	margin: 0px;
}
-->
</style>
</head>
<body bgcolor="#ffffff" onLoad="MM_preloadImages('images/digiBoard_r2_c2_f2.gif','images/digiBoard_r2_c4_f2.gif','images/digiBoard_r2_c3_f2.gif')">
<table width="700" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <!-- fwtable fwsrc="easyBoard.png" fwbase="digiBoard.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="1" -->
  <tr>
    <td height="91"><img name="digiBoard_r1_c1" src="images/digiBoard_r1_c1.gif" width="700" height="91" border="0" alt=""></td>
  </tr>
  <tr>
    <td height="23"> <table border="0" cellpadding="0" cellspacing="0" width="700">
        <tr>
          <td><img name="digiBoard_r2_c1" src="images/digiBoard_r2_c1.gif" width="367" height="23" border="0" alt=""></td>
          <td><a href="digiBoardPost.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('digiBoard_r2_c2','','images/digiBoard_r2_c2_f2.gif',1);"><img name="digiBoard_r2_c2" src="images/digiBoard_r2_c2.gif" width="108" height="23" border="0" alt=""></a></td>
          <td><a href="1.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('digiBoard_r2_c3','','images/digiBoard_r2_c3_f2.gif',1);"><img name="digiBoard_r2_c3" src="images/digiBoard_r2_c3.gif" width="108" height="23" border="0" alt=""></a></td>
          <td><a href="digiBoardLogin.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('digiBoard_r2_c4','','images/digiBoard_r2_c4_f2.gif',1);"><img name="digiBoard_r2_c4" src="images/digiBoard_r2_c4.gif" width="108" height="23" border="0" alt=""></a></td>
          <td><img name="digiBoard_r2_c5" src="images/digiBoard_r2_c5.gif" width="9" height="23" border="0" alt=""></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="36"><img name="digiBoard_r3_c1" src="images/digiBoard_r3_c1.gif" width="700" height="36" border="0" alt=""></td>
  </tr>
  <tr>
    <td valign="top" background="images/digiBoard_r4_c1.gif"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="14">　</td>
          <td><table width="96%" border="0" align="center" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" class="Box">
              <tr valign="top">
                <td width="100" align="center" bgcolor="#000066"><font color="#FFFFFF" size="-1">&nbsp;<img src="images/post.jpg" width="100" height="250"></font></td>
                <td><form action="<%=MM_editAction%>" method="POST" name="form1" onSubmit="MM_validateForm('digiB_subject','','R','digiB_name','','R','digiB_email','','NisEmail','digiB_content','','R');return document.MM_returnValue">
                    <table width="100%" border="0" cellpadding="2" cellspacing="1">
                      <tr valign="middle">
                        <td width="80" align="right" bgcolor="#EAEAF4"><font size="-1">标题：</font></td>
                        <td><font size="-1">
                          <input name="digiB_subject" type="text" id="digiB_subject" size="20">
                          <font color="#FF0000"><strong>*必填</strong></font> </font></td>
                      </tr>
                      <tr valign="middle">
                        <td width="80" align="right" bgcolor="#EAEAF4"><font size="-1">姓名：</font></td>
                        <td><font size="-1">
                          <input name="digiB_name" type="text" id="digiB_name" size="20">
                          <font color="#FF0000"><strong>*必填</strong></font></font></td>
                      </tr>
                      <tr valign="top">
                        <td width="80" align="right" bgcolor="#EAEAF4"><font size="-1">发言图标：</font></td>
                        <td><font size="-1">
                          <input name="digiB_face" type="radio" value="m01.jpg" checked>
                          <img src="images/face/m01.jpg" width="60" height="60">
                          <input type="radio" name="digiB_face" value="m02.jpg">
                          <img src="images/face/m02.jpg" width="60" height="60">
                          <input type="radio" name="digiB_face" value="m03.jpg">
                          <img src="images/face/m03.jpg" width="60" height="60">
                          <input type="radio" name="digiB_face" value="m04.jpg">
                          <img src="images/face/m04.jpg" width="60" height="60">
                          <input type="radio" name="digiB_face" value="m05.jpg">
                          <img src="images/face/m05.jpg" width="60" height="60">
                          <br>
                          <input type="radio" name="digiB_face" value="w01.jpg">
                          <img src="images/face/w01.jpg" width="60" height="60">
                          <input type="radio" name="digiB_face" value="w02.jpg">
                          <img src="images/face/w02.jpg" width="60" height="60">
                          <input type="radio" name="digiB_face" value="w03.jpg">
                          <img src="images/face/w03.jpg" width="60" height="60">
                          <input type="radio" name="digiB_face" value="w04.jpg">
                          <img src="images/face/w04.jpg" width="60" height="60">
                          <input type="radio" name="digiB_face" value="w05.jpg">
                          <img src="images/face/w05.jpg" width="60" height="60">
                          </font></td>
                      </tr>
                      <tr valign="middle">
                        <td width="80" align="right" bgcolor="#EAEAF4"><font size="-1">电子邮件：</font></td>
                        <td><font size="-1">
                          <input name="digiB_email" type="text" id="digiB_email" size="20">
                          </font></td>
                      </tr>
                      <tr valign="middle">
                        <td width="80" align="right" bgcolor="#EAEAF4"><font size="-1">个人网页：</font></td>
                        <td><font size="-1">
                          <input name="digiB_Web" type="text" id="digiB_Web" size="20">
                          (请加上http://)</font></td>
                      </tr>
                      <tr valign="middle">
                        <td width="80" align="right" valign="top" bgcolor="#EAEAF4"><font size="-1">留言内容：</font></td>
                        <td><font size="-1">
                          <textarea name="digiB_content" cols="40" rows="5" id="digiB_content"></textarea>
                          <font color="#FF0000"><strong>*必填</strong></font></font></td>
                      </tr>
                      <tr valign="top">
                        <td width="80" align="right" bgcolor="#EAEAF4">　</td>
                        <td> <input type="submit" name="Submit" value="送出">                          <input type="reset" name="Submit2" value="重设"></td>
                      </tr>
                    </table>
                    <input type="hidden" name="MM_insert" value="form1">
                  </form></td>
              </tr>
            </table></td>
          <td width="10">　</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="21"><img name="digiBoard_r5_c1" src="images/digiBoard_r5_c1.gif" width="700" height="21" border="0" alt=""></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
