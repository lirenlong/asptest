<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/llxxcc.asp" -->
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_llxxcc_STRING
Recordset1.Source = "SELECT * FROM denglu"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername=CStr(Request.Form("denglu_name_f"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="doexercises.html"
  MM_redirectLoginFailed="personal.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_llxxcc_STRING
  MM_rsUser.Source = "SELECT denglu_name, denglu_mina"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM denglu WHERE denglu_name='" & Replace(MM_valUsername,"'","''") &"' AND denglu_mina='" & Replace(Request.Form("denglu_mina_f"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>教学资源</title>
<style type="text/css">
<!--
.STYLE2 {color: #CCCCCC}
body {
	background-image: url(images/backgrand.jpg);
	background-repeat: repeat;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.STYLE7 {
	font-family: "宋体";
	font-weight: bold;
}
-->
</style>
</head>

<body>
<table width="780%" border="0">
  <tr>
    <td width="335" height="1729" valign="top"><p><img src="images/biaoti.jpg" width="300" height="80" />
    </p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <h1 align="center">&nbsp;</h1>
      <h1 align="center"><span class="STYLE7"><a href="jxdg.doc">教学大纲</a></span></h1>
      <h1 align="center" class="STYLE7"><a href="jxrl.doc">教学日历</a></h1>
      <h1 align="center"><a href="ja.zip">教案</a></h1>
      <h1 align="center"><strong><a href="syzds.doc">实验指导书</a></strong></h1>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p align="center"><a href="default.asp">返回首页</a> </p>
      <p>&nbsp;</p>
    <td width="537" height="1729" valign="top"><h1 align="center"><a name="_Toc176143865" id="_Toc176143865"><strong>微波技术与天线课程教学大纲</strong></a> <br />
      </h1>
      <p>课程编码：11211361 学时：32 学分：2<br />
        适用专业：通信工程 <br />
      </p>
      <p>一、课程的性质和任务 <br />
        微波技术与天线课程是通信工程专业的专业方向必修课程，考查课。 <br />
        由于微波的频率很高，波长很短，这就使得微波器件和电路具有它固有的特性。通过本课程学习，使学生初步掌握微波技术中各种微波导波结构、器件的工作原理和相应的分析手段。使学生在掌握传输线理论，微波网络理论，微波器件的工作原理和相应的分析手段的同时，对天线理论也有一定的了解。 <br />
        二、课程教学内容的基本要求、重点和难点及学时分配 <br />
        1.理论教学 24 学时 <br />
        1.1传输线理论 （10学时） <br />
        1.1.1传输线理论引言 <br />
        1.1.2无耗传输线方程及其解 <br />
        1.1.3无耗传输线的基本特性 <br />
        1.1.4均匀无耗传输线工作状态的分析 <br />
        1.1.5有耗传输线 <br />
        1.1.6阻抗圆图及其应用 <br />
        1.1.7传输线阻抗匹配 <br />
        必须掌握：微波的范围和特点；长线的概念；传输线的分布参数及传输线分布参数电路；列出传输线方程；传输线方程的解；传输线的特性阻抗、传播参数；传输线的输入阻抗；传输线反射系数；传输线的阻抗匹配；归一化阻抗的概念；圆图的构成及工作原理；如何使用圆图。 <br />
        掌握：传输线的概念；均匀无耗传输线的工作状态分析中的行波状态条件，行波状态下的电压、电流的沿线分布，行波状态下的沿线阻抗分布；均匀无耗传输线的驻波状态条件，驻波状态下的沿线电压、电流分布及沿线阻抗分布。 <br />
        重点：掌握无耗传输线的基本特点和工作状态；掌握使用阻抗圆图。 <br />
        难点：输入阻抗、反射系数和驻波比。 <br />
        1.2规则波导和空腔谐振器 （4学时） <br />
        必须掌握：谐振腔中的电磁能量关系；品质因素的概念；矩形谐振腔的谐振波长 <br />
        重点：品质因素的概念；矩形谐振腔的谐振波长 <br />
        1.3微波网络基础 （4学时） <br />
        1.3.1微波网络基础引言 <br />
        1.3.2波导传输线与双线传输线的等效 <br />
        1.3.3微波元件等效为微波网络的原理 <br />
        1.3.4二端口微波网络 <br />
        必须掌握： 场、路两种方法分析；波导等效为双线的原则；波导的等效特性阻抗；归一化阻抗、归一化电压和归一化电流；两端口微波网络参量的阻抗参量；两端口微波网络参量的散射参量；两端口微波网络参量的散射参量与阻抗参量之间的变换。 <br />
        掌握：两端口微波网络参量的导纳参量；两端口微波网络参量的转移参量；匹配元件和连接元件；分路元件。 <br />
        重点：两端口微波网络参量的散射参量。 <br />
        1.4天线 (6学时)<br />
        必须掌握：基本辐射单元；发射天线的主要参数。 <br />
        掌握：接收天线；对称振子天线。 <br />
        重点：对称振子的结构及辐射场；天线的特性参量。 <br />
        2.实验 8学时 <br />
        2.1微波测试系统的调试（2学时） <br />
        2.2波导波长和驻波比的测量（2学时） <br />
        2.3阻抗测量（2学时） <br />
        2.4阻抗匹配（2学时） <br />
        三、参考教材和主要参考文献 <br />
        参考教材：机械工业出版社《微波技术与天线》编著：傅文斌 <br />
        参考文献： </p>
      <ol>
        <ol>
          <ol>
            <li>北京邮电大学出版社《微波技术基础与应用》编著：陈振国； </li>
            <li>电子工业出版社《电磁场与微波工程基础》编著：毛均杰； </li>
            <li>西安电子科技大学出版社《电磁场微波技术与天线》编著：盛振华 。 </li>
          </ol>
        </ol>
      </ol>
      <p>四、考核形式和考核要求 <br />
        开卷笔试，卷面满分为100分 ，占期末总成绩的70－90%。其余为平时成绩，平时成绩根据测验、实验、出勤情况。考核覆盖率在85%以上。 <br />
        考核要求： </p>
      <ol>
        <li>掌握传输线的特性阻抗、传播参数； </li>
        <li>掌握无耗传输线的基本特点，工作状态及输入阻抗反射系数等传播参数的求解； </li>
        <li>掌握传输线的阻抗匹配的概念；无反射匹配的概念及应用；四分之一波长匹配器； </li>
        <li>掌握导行电磁波的分类；导行电磁波的传输特征； </li>
        <li>掌握矩形波导和圆波导的场结构、主模； </li>
        <li>掌握两端口微波网络参量的阻抗参量和散射参量； </li>
        <li>掌握基本辐射单元；天线的特性参量。 </li>
      </ol>
      <p>五、有关说明 <br />
        执笔人：李路 <br />
        审核人：周昕 <br />
        批准人：范立南 <br />
        六、本门课程主要概念的中英文对照表 <br />
        微波 Microwave<br />
        电磁场 Electromagnetic field<br />
        横电磁波 Transverse  electro-magnetic wave<br />
        天线 Antenna<br />
        电压 Voltage<br />
        电流 Electric current<br />
        频率 Frequency<br />
        电路 Circuit<br />
        断路 Open circuit<br />
        短路 Short circuit<br />
        衰减器 Attenuator<br />
        振幅 Amplitude<br />
        散度 Divergence<br />
        旋度 Rotation or curl<br />
        梯度 Gradient<br />
        群速 Group velocity<br />
        波导 Waveguide<br />
        阻抗 Impedance<br />
        导纳 Admittance<br />
        归一化 Normalization<br />
        谐振腔 Resonant cavity<br />
        驻波比 Standing wave ratio<br />
        辐射 Radiation<br />
        垂直 Vertical separation<br />
        传播特性 Propagation characteristic<br />
        主模 Dominant mode<br />
        波长 Wave length<br />
        纵向分量 Longitudinal component<br />
        横向分量 Transverse component<br />
        相位滞后 Lagging phase<br />
        边界条件 Boundary condition<br />
        匹配 Matching<br />
        调制 Modulate<br />
        全反射 Total reflection<br />
        反射系数 Reflectance<br />
        解调 Demodulation<br />
        传输线 Transmission line<br />
        放大器 Amplifier<br />
        混频器 Mixer<br />
        微波集成电路 Microwave Integrated Circuit (MIC)<br />
        振荡器 Oscillator<br />
        耦合 Coupling<br />
        品质因素 Goodness<br />
        频分复用 Frequency-division multiplexing<br />
        信噪比 Signal-to-noise ratio</p>
      <p>&nbsp;</p>
      <p><a href="default.asp">返回首页</a></p>
    <br clear="all" /></td>
    <td width="5073"><span class="STYLE2"></span></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
