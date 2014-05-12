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
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_llxxcc_STRING
  MM_editTable = "board"
  MM_editColumn = "digiB_id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "digiBoardAdmin.asp"

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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
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
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("digiB_id") <> "") Then
  Recordset1__MMColParam = Request.QueryString("digiB_id")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_llxxcc_STRING
Recordset1.Source = "SELECT * FROM board WHERE digiB_id = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function CleanHtmlBr(str)												
// 织梦平台 http://www.e-dreamer.idv.tw												
	CleanHtmlBr =  Replace(Replace(Replace(Replace(str,"<","&lt;"),">","&gt;"),vbCrlf, "<br>"), chr(32)&chr(32),"&nbsp;&nbsp;")			
End Function														
</SCRIPT>									
<html>
<head>
<title>数字留言板</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!-- Fireworks MX Dreamweaver MX target.  Created Tue Jan 28 14:10:55 GMT+0800 (￥x￥_?D·CRE?!) 2003-->
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

//-->
</script>
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
-->
</style>
<style type="text/css">
<!--
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
    <td height="23">
      <table border="0" cellpadding="0" cellspacing="0" width="700">
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
    <td valign="top" background="images/digiBoard_r4_c1.gif">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="14">　</td>
          <td> <form name="form1" method="POST" action="<%=MM_editAction%>">

              <table width="96%" height="30" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E5E5E5">
                <tr>
                  <td><strong><font color="#990000">** 您确定要删除下面这一则留言吗？</font></strong></td>
                </tr>
              </table>
              <hr width="96%" size="1" noshade class="lineBox">
              <table width="96%" border="0" align="center" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" class="Box">

                <tr valign="top">
                  <td width="120" align="center" class="backIMG"><strong><font color="#000066" size="-1">&nbsp;<img src="images/face/<%=(Recordset1.Fields.Item("digiB_face").Value)%>">
                    <br>
                    <%=(Recordset1.Fields.Item("digiB_name").Value)%> </font></strong></td>
                  <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><font size="-1">[<%=(Recordset1.Fields.Item("digiB_id").Value)%>] 
                          <strong><font color="#FF6600" size="3"><%=(Recordset1.Fields.Item("digiB_subject").Value)%></font></strong> </font></td>
                        <td align="right"><strong><font color="#999999" size="-1"><%=(Recordset1.Fields.Item("digiB_potime").Value)%></font></strong></td>
                      </tr>
                    </table>
                    <hr size="1" noshade class="lineBox"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td valign="top" style="word-break:break-all"><font color="#666666" size="-1"><%= CleanHtmlBr(Recordset1.Fields.Item("digiB_content").Value)%><wbr></font></td>
                      </tr>
                    </table>
                    <hr size="1" noshade> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td align="right"><font size="-1">
                          <% if Recordset1.Fields.Item("digiB_email").Value <> "" Then 'start db_sc script %>
                          <a href="mailto:<%=(Recordset1.Fields.Item("digiB_email").Value)%>" class="ImgButton"><img src="images/icon_email.gif" width="60" height="19" border="0"></a>
                          <% end if 'end db_sc script %>
                          <% if Recordset1.Fields.Item("digiB_web").Value <> "" Then 'start db_sc script %>
                          <a href="<%=(Recordset1.Fields.Item("digiB_web").Value)%>" target="_blank" class="ImgButton"><img src="images/icon_www.gif" width="60" height="19" border="0"></a>
                          <% end if 'end db_sc script %>
                          </font></td>
                      </tr>
                    </table></td>
                </tr>
              </table>

              <hr width="96%" size="1" noshade class="lineBox">
              <div align="center">
                <input type="hidden" name="MM_delete" value="form1">
                <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("digiB_id").Value %>">
                <input type="submit" name="Submit" value="确定删除">
                <input type="button" name="Submit2" value="回上一页" onClick="window.history.back();">
              </div>
            </form>

          </td>
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