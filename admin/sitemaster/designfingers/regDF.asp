<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
'##############################################
' History: 2008.03.12 modify - 2008 ������ �߰� ��� ����
' Description: ������ �ΰŽ�
'##############################################
Dim clsDF
Dim iDFSeq,sMode
Dim iDFType,sTitle,tContents, dPrizeDate,blnDisplay,blnOtherMall,sUserid,dRegDate,sComment,blnMainDisplay
Dim arrImg,arrItemid, tmpItemid, sProdName, sProdSize, sProdColor, sProdJe, sProdGu, sProdSpe, iTotImgCnt, vIsMovie, vOpenDate, vTag
Dim intLoop, iCurrpage, edid, emktid
Dim strImgUrl, strLink, strImgDesc
Dim i3dvPCode : i3dvPCode = 8

iDFSeq  = requestCheckVar(request("iDFS"),10)
iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
blnDisplay = False
sMode = "I"	'// insert
vIsMovie = "N"

IF iDFSeq <> "" THEN
	sMode = "U" '// update
	set clsDF = new  CDesignFingers
		clsDF.FDFSeq = iDFSeq
		clsDF.fnGetDFCont
		iDFType 	 	= clsDF.FDFType
		sTitle  	 	= clsDF.FTitle
		tContents 	 	= clsDF.FContents
		dPrizeDate   	= clsDF.FPrizeDate
		sComment		= clsDF.FComment
		blnDisplay 	 	= clsDF.FIsDisplay
		blnOtherMall 	= clsDF.FIsOtherMall
		sUserid	 	 	= clsDF.FUserid
		dRegDate	 	= clsDF.FRegDate
		arrItemid    	= clsDF.FItemid
		arrImg		 	= clsDF.FImg
		sProdName		= clsDF.FProdName
		sProdSize		= clsDF.FProdSize
		sProdColor		= clsDF.FProdColor
		sProdJe			= clsDF.FProdJe
		sProdGu			= clsDF.FProdGu
		sProdSpe		= clsDF.FProdSpe
		iTotImgCnt		= clsDF.FTotImgCnt
		vIsMovie		= clsDF.FIsMovie
		vOpenDate		= clsDF.FOpenDate
		vTag			= clsDF.FTag
		edid			= clsDF.FEDId
		If isNull(edid) Then
			edid = ""
		End If
		emktid			= clsDF.FEMKTId
		If isNull(emktid) Then
			emktid = ""
		End If

		blnMainDisplay	= clsDF.FIsMainDisplay
	set clsDF = nothing

	IF isArray(arrItemid) THEN
		For intLoop =0 To UBound(arrItemid,2)
		 	IF intLoop = 0 THEN
		 		tmpItemid =  arrItemid(0,intLoop)
		 	ELSE
				tmpItemid = tmpItemid&","& arrItemid(0,intLoop)
			END IF
		Next
		arrItemid = tmpItemid
	END IF
END IF

	'//icon �̹��� url ����
function fnSetIconImage(ByVal sImgURL)
	Dim tmpImg
	tmpImg = split(sImgURL,"/")

	fnSetIconImage = replace(sImgURL,tmpImg(Ubound(tmpImg)),"icon/icon_"&tmpImg(Ubound(tmpImg)))
End Function
%>
<script language="javascript">
<!--

//-- jsPopCal : �޷� �˾� --//
 function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
 }


 //�̹���÷��
 function jsPopAddImg(sFolder,sImgID){
 document.domain ="10x10.co.kr";
 	var chkIcon = 0;
 	var winImg;
 	var sImgURL;

 		if(sFolder =="3dv"){chkIcon = 1;} //3dview �̹����ϋ��� ������ ����
 		if (sImgID == 0){
 			sImgURL = eval("document.frmReg.img"+sFolder).value;
 		}else{
 			sImgURL = eval("document.frmReg.img"+sFolder+"["+(sImgID-1)+"]").value;
 		}
 		winImg = window.open('popAddImage.asp?sF='+sFolder+'&sID='+sImgID+'&chkI='+chkIcon+'&sIU='+sImgURL,'popImg','width=380,height=150');
 		winImg.focus();
 }

 //�̹��� ����
 function jsDelImg(sValue,sID){
 document.domain ="10x10.co.kr";
 	if ( sID == 0 ){
 	eval("document.all.div"+sValue).innerHTML = "";
 	eval("document.all.img"+sValue).value = "";
 	}else{
 	eval("document.all.div"+sValue+"["+(sID-1)+"]").innerHTML = "";
 	eval("document.all.img"+sValue+"["+(sID-1)+"]").value = "";
 	}
 }

 //���� ���
 function jsDFSubmit(){
 var frm = document.frmReg;
 	 if(!frm.selDFT.value){
 	 	alert("������ ���� ���ּ���");
	  	frm.selDFT.focus();
	  	return false;
 	 }
 	 if(!frm.sT.value){
	  	alert("������ �Է����ּ���");
	  	frm.sT.focus();
	  	return false;
	  }
 	 if(!frm.opendate.value){
	  	alert("�������� �Է����ּ���");
	  	return false;
	  }
	 if(frm.arrI.value != "")
	 {
	 	var tlng = frm.arrI.value.length;
		var chkchar = frm.arrI.value.substring(tlng,tlng-1);
	 	if(chkchar == ",")
	 	{
	 		alert("��ǰID ������ �޸��� �����Դϴ�.");
	 		return false;
	 	}
	 }
 }

 //-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/lib/showimage.asp?img='+sImgUrl,'pImg','width=50,height=50');
	 wImgView.focus();
	}


function AutoInsert() {
	var f = document.all;

	var rowLen = f.imgIn.rows.length;
	var i = rowLen;
	var r  = f.imgIn.insertRow(rowLen++);
	var c0 = r.insertCell(0);
	var Html;

	c0.innerHTML = "&nbsp;";
	var inHtml = "<tr>"
				+ "	<td style='padding:5 5 5 5'><table width='100%' height='100%' cellpadding='0' cellspacing='0' border='0' bgcolor='#FFFFFF'><tr><td>"
				+ "		<input type='button' id='A"+(i-1)+"' value='�̹���÷��' onClick=jsPopAddImg('add',"+(i-1)+"); class='button'>"
				+ "		<input type='hidden' name='imgadd' value=''>"
				+ "		<span id='divadd'></span><br><br><font color='blue'>+ map name='add"+(i-1)+"Map'</font>"
				+ "		<textarea name='tA"+(i-1)+"' rows='10' cols='75'></textarea>"
				+ "	</td></tr></table></td>"
				+ "</tr>"
 c0.innerHTML = inHtml;
}

function findProd()
{
	window.open('pop_additemlist.asp','findProd','width=900,height=600,scrollbars=yes')
}
function workerlist()
{
	var openWorker = null;
	var worker = frmReg.selMKTId.value;
	openWorker = window.open('PopWorkerList.asp?worker='+worker+'&team=11','openWorker','width=570,height=570,scrollbars=yes');
	openWorker.focus();
}
 //-->
</script>

<table width="800" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
<td colspan="2">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
  <form name="frmReg" method="post" action="procDF.asp" onSubmit="return jsDFSubmit();">
  	<input type="hidden" name="iC" value="<%= iCurrpage%>">
  	<input type="hidden" name="menupos" value="<%= menupos%>">
    <input type="hidden" name="sM" value="<%= sMode%>">
    <input type="hidden" name="iDFS" value="<%=iDFSeq %>">
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">ID</td>
	  <td>
	  	<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a"><tr>
	  	<td><%=iDFSeq%></td>
	  	<td align="right"><a href="<%=wwwUrl%>/designfingers/designfingers.asp?fingerid=<%=iDFSeq%>" target="_blank">[Front��������]</a></td>
		</tr></table>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">����</td>
	  <td>
		<select name="selDFT">
			<option value="">����</option>
			<%sbOptCommCode 10, iDFType%>
		</select>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">����</td>
	  <td><input type="text" name="sT" value="<%= sTitle %>" size="50" maxlength="32"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">��ǰID</td>
	  <td>
	  	<input type="text" name="arrI" value="<%= arrItemid %>" size="50" maxlength="256">
	  	<input type="button" class="button" value="��ǰã��" onClick="findProd()">
	  	<br>�޸�(<font color="red">,</font>)�� �������ּ���. ������ �޸��� ����
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">���ÿ���</td>
	  <td>
	  	<input type="radio" name="rdoD" value="1" <% if blnDisplay then response.write "checked" %>>������
	  	<input type="radio" name="rdoD" value="0" <% if Not blnDisplay then response.write "checked" %>>���þ���
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">�ܺθ���뿩��</td>
	  <td>
	  	<input type="radio" name="rdoOM" value="1" <% if blnOtherMall then response.write "checked" %>>�����
	  	<input type="radio" name="rdoOM" value="0" <% if Not blnOtherMall then response.write "checked" %>>������
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">Main Top<br>��뿩��</td>
	  <td>
	  	<input type="radio" name="rdoMD" value="1" <% if blnMainDisplay then response.write "checked" %>>������
	  	<input type="radio" name="rdoMD" value="0" <% if Not blnMainDisplay then response.write "checked" %>>���þ���
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">��÷��ǥ��</td>
	  <td><input type="text" name="dPD" value="<%=dPrizeDate%>" size="10" maxlength="10" onClick="jsPopCal('dPD');"  style="cursor:hand;" class="input_b"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">������</td>
	  <td><input type="text" name="opendate" value="<%=vOpenDate%>" size="10" maxlength="10" onClick="jsPopCal('opendate');"  style="cursor:hand;" class="input_b"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">��ǰ</td>
	  <td><input type="text" name="sPdN" value="<%=sProdName%>" size="50" maxlength="50" > ����)���� ���ν�Ÿ�� �ð��н�����</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">ũ��</td>
	  <td><input type="text" name="sPdS" value="<%=sProdSize%>" size="50" maxlength="50" > ����)small : 9 x 14, pocket : 9 x 14, large 9 x 14 (cm)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">����</td>
	  <td><input type="text" name="sPdC" value="<%=sProdColor%>" size="50" maxlength="50" > ����)black/black, blue/navy, pink/rose, red</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">���</td>
	  <td><input type="text" name="sPdJ" value="<%=sProdJe%>" size="50" maxlength="50" > ����)����</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">����</td>
	  <td><input type="text" name="sPdG" value="<%=sProdGu%>" size="50" maxlength="50" > ����)80pages</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">Ư¡</td>
	  <td><input type="text" name="sPdP" value="<%=sProdSpe%>" size="50" maxlength="50" > ����)���� ���� ���� ����</td>
	</tr>
   	<tr>
   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�������̳�</td>
   		<td bgcolor="#FFFFFF">
   			<%sbGetDesignerid "selDId",edid,""%>
   		</td>
   	</tr>
   	<tr>
   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">����ȹ��</td>
   		<td bgcolor="#FFFFFF">
   			<%sbGetwork "selMKTId",emktid,""%>
   		</td>
   	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">Comment</td>
	  <td><input type="text" name="sCom" value="<%=sComment%>" size="80" maxlength="100" ></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">��ǰ����</td>
	  <td><textarea name="txtC" rows="10" cols="75"><%=tContents%></textarea></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">Tag</td>
	  <td><input type="text" name="sTag" value="<%=vTag%>" size="80" maxlength="100" ></td>
	</tr>
	</table>
</td>

</tr>
<tr>
<td colspan="2">
	<font color="red">+ �̹����� 300kb ���Ϸ� ���� �ּ��� (�ִ� 500kb)</font><br>
	<font color="blue">+ �ΰŽ� ���� ������ �÷��ô� [ON]����Ʈ����>>[����]���������� �� ������ġ�� ���� 2010�ΰŽ������÷��ö�� �ֽ��ϴ�.</font><br>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<!--- event Left --------->
<%
IF iDFSeq <> "" THEN
	If isNull(arrImg(22,1,3)) = True Then
		strImgUrl = ""
	Else
		strImgUrl = arrImg(22,1,3)
	End If
End If
%>
<!--
	<tr bgcolor="#FFFFFF" height="35">
	  <td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">Event ��÷ �̹���<br>Left(361*178)<br>Right(569*128)</td>
	  <td>
	  	  <input type="button" id="eventLeft" value="Left �̹���÷��" onClick="jsPopAddImg('eventLeft',0);" class="button">
	  	  <input type="hidden" name="imgeventLeft" value="<%=strImgUrl%>">
	  	  <span id="diveventLeft">
	  	  <%IF strImgUrl <> "" THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=strImgUrl%>" width="200" height="100" border="0"></a> <a href="javascript:jsDelImg('eventLeft',0);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
//-->
<!--- /event Left --------->
<!--- event Right --------->
<%
IF iDFSeq <> "" THEN
	If isNull(arrImg(23,1,3)) = True Then
		strImgUrl = ""
	Else
		strImgUrl = arrImg(23,1,3)
	End If
End If
%>
	<tr bgcolor="#FFFFFF" height="35">
		<td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">�̺�Ʈ �̹���</td>
	  <td>
	  	  <input type="button" id="eventRight" value="�̹���÷��" onClick="jsPopAddImg('eventRight',0);" class="button">
	  	  <input type="hidden" name="imgeventRight" value="<%=strImgUrl%>">
	  	  <span id="diveventRight">
	  	   <%IF strImgUrl <> "" THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=strImgUrl%>" width="200" height="100" border="0"></a> <a href="javascript:jsDelImg('eventRight',0);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<!--- /event Right --------->
<!--- main top --------->
<%
IF iDFSeq <> "" THEN
	If isNull(arrImg(21,1,3)) = True Then
		strImgUrl = ""
	Else
		strImgUrl = arrImg(21,1,3)
	End If
End If
%>
<!--
	<tr bgcolor="#FFFFFF">
	  <td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">Main Top �̹���(739*297)</td>
	  <td>
	  	  <input type="button" id="main_top" value="�̹���÷��" onClick="jsPopAddImg('main_top',0);" class="button">
	  	  <input type="hidden" name="imgmain_top" value="<%=strImgUrl%>">
	  	  <span id="divmain_top">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=strImgUrl%>" width="200" height="100" border="0"></a> <a href="javascript:jsDelImg('main_top',0);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
//-->
<!--- /main top --------->
<!--- top --------->
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(2,1,3)
%>
<!--
	<tr bgcolor="#FFFFFF">
	  <td width="100" rowspan="3"  bgcolor="<%= adminColor("gray") %>" align="center">Top �̹���</td>
	  <td>
	  	  <input type="button" id="top1" value="�̹���÷��" onClick="jsPopAddImg('top',1);" class="button">
	  	  <input type="hidden" name="imgtop" value="<%=strImgUrl%>">
	  	  <span id="divtop">
	  	  <%IF strImgUrl <> "" THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=strImgUrl%>" width="200" height="100" border="0"></a> <a href="javascript:jsDelImg('top',1);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
//-->
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(2,2,3)
%>
<!--
	<tr bgcolor="#FFFFFF">
	  <td>
	  	  <input type="button" id="top2" value="�̹���÷��" onClick="jsPopAddImg('top',2);" class="button">
	  	  <input type="hidden" name="imgtop" value="<%=strImgUrl%>">
	  	  <span id="divtop">
	  	   <%IF strImgUrl <> "" THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=strImgUrl%>" width="200" height="100" border="0"></a> <a href="javascript:jsDelImg('top',2);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
//-->
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(2,3,3)
%>
<!--
	<tr bgcolor="#FFFFFF">
	  <td>
	  	  <input type="button" id="top3" value="�̹���÷��" onClick="jsPopAddImg('top',3);" class="button">
	  	   <input type="hidden" name="imgtop" value="<%=strImgUrl%>">
	  	  <span id="divtop">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=strImgUrl%>" width="200" height="100" border="0"></a> <a href="javascript:jsDelImg('top',3);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
//-->
<!--- /top --------->
<!--- play --------->
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(8,1,3)
%>
	<tr bgcolor="#FFFFFF">
	  <td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">play �̹���</td>
	  <td>
	  	  <input type="button" id="play" value="�̹���÷��" onClick="jsPopAddImg('play',0);" class="button">
	  	  <input type="hidden" name="imgplay" value="<%=strImgUrl%>">
	  	  <span id="divplay">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<img src="<%=strImgUrl%>" > <a href="javascript:jsDelImg('play',0);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<!--- /play --------->
<!--- small --------->
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(3,1,3)
%>
	<tr bgcolor="#FFFFFF">
	  <td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">Small �̹���(50*50)</td>
	  <td>
	  	  <input type="button" id="small" value="�̹���÷��" onClick="jsPopAddImg('small',0);" class="button">
	  	  <input type="hidden" name="imgsmall" value="<%=strImgUrl%>">
	  	  <span id="divsmall">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<img src="<%=strImgUrl%>" > <a href="javascript:jsDelImg('small',0);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<!--- /small --------->
<!--- list --------->
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(4,1,3)
%>
	<tr bgcolor="#FFFFFF">
	  <td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">List �̹���(150*150)</td>
	  <td>
	  	 <input type="button" id="list" value="�̹���÷��" onClick="jsPopAddImg('list',0);" class="button">
	  	  <input type="hidden" name="imglist" value="<%=strImgUrl%>">
	  	  <span id="divlist">
	  	  	  <%IF strImgUrl <> ""  THEN%>
	  	  	<img src="<%=strImgUrl%>"> <a href="javascript:jsDelImg('list',0);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<!--- /list --------->
<!--- 3dview --------->
<%
IF iDFSeq <> "" THEN  strImgUrl = arrImg(7,1,3) :  strImgDesc = arrImg(7,1,5)
%>
	<tr bgcolor="#FFFFFF">
	  <td width="100" rowspan="10"  bgcolor="<%= adminColor("gray") %>" align="center">3d View</td>
	  <td>
	  	 <select name="sel3dv">
	  	 <%sbOptCommCode i3dvPCode,strImgDesc%>
	  	 </select>
	  	  <input type="button" id="3dv1" value="�̹���÷��" onClick="jsPopAddImg('3dv',1);" class="button">
	  	  <input type="hidden" name="img3dv" value="<%=strImgUrl%>">
	  	  <span id="div3dv">
	  	   <%IF  strImgUrl <> ""  THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=fnSetIconImage(strImgUrl)%>" border="0"></a> <a href="javascript:jsDelImg('3dv',1);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(7,2,3):  strImgDesc = arrImg(7,2,5)
%>
	<tr bgcolor="#FFFFFF">
	  <td>
	  	 <select name="sel3dv">
	  	 <%sbOptCommCode i3dvPCode,strImgDesc%>
	  	 </select>
	  	  <input type="button" id="3dv2" value="�̹���÷��" onClick="jsPopAddImg('3dv',2);" class="button">
	  	  <input type="hidden" name="img3dv" value="<%=strImgUrl%>">
	  	  <span id="div3dv">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=fnSetIconImage(strImgUrl)%>" border="0"></a> <a href="javascript:jsDelImg('3dv',2);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(7,3,3):  strImgDesc = arrImg(7,3,5)
%>
	<tr bgcolor="#FFFFFF">
	  <td>
	  	 <select name="sel3dv">
	  	 <%sbOptCommCode i3dvPCode,strImgDesc%>
	  	 </select>
	  	  <input type="button" id="3dv3" value="�̹���÷��" onClick="jsPopAddImg('3dv',3);" class="button">
	  	  <input type="hidden" name="img3dv" value="<%=strImgUrl%>">
	  	  <span id="div3dv">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=fnSetIconImage(strImgUrl)%>" border="0"></a> <a href="javascript:jsDelImg('3dv',3);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(7,4,3):  strImgDesc = arrImg(7,4,5)
%>
	<tr bgcolor="#FFFFFF">
	  <td>
	  	 <select name="sel3dv">
	  	 <%sbOptCommCode i3dvPCode,strImgDesc%>
	  	 </select>
	  	  <input type="button" id="3dv4" value="�̹���÷��" onClick="jsPopAddImg('3dv',4);" class="button">
	  	  <input type="hidden" name="img3dv" value="<%=strImgUrl%>">
	  	  <span id="div3dv">
	  	   <%IF strImgUrl <> "" THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=fnSetIconImage(strImgUrl)%>" border="0"></a> <a href="javascript:jsDelImg('3dv',4);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(7,5,3):  strImgDesc = arrImg(7,5,5)
%>
	<tr bgcolor="#FFFFFF">
	  <td>
	  	 <select name="sel3dv">
	  	 <%sbOptCommCode i3dvPCode,strImgDesc%>
	  	 </select>
	  	  <input type="button" id="3dv5" value="�̹���÷��" onClick="jsPopAddImg('3dv',5);" class="button">
	  	  <input type="hidden" name="img3dv" value="<%=strImgUrl%>">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=fnSetIconImage(strImgUrl)%>" border="0"></a> <img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  <span id="div3dv">
	  	  </span>
	  </td>
	</tr>
<%
IF iDFSeq <> "" THEN strImgUrl =arrImg(7,6,3):  strImgDesc = arrImg(7,6,5)
%>
	<tr bgcolor="#FFFFFF">
	  <td>
	  	 <select name="sel3dv">
	  	 <%sbOptCommCode i3dvPCode,strImgDesc%>
	  	 </select>
	  	  <input type="button" id="3dv6" value="�̹���÷��" onClick="jsPopAddImg('3dv',6);" class="button">
	  	  <input type="hidden" name="img3dv" value="<%=strImgUrl%>">
	  	  <span id="div3dv">
	  	   <%IF strImgUrl <> "" THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=fnSetIconImage(strImgUrl)%>" border="0"></a>  <a href="javascript:jsDelImg('3dv',6);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(7,7,3):  strImgDesc = arrImg(7,7,5)
%>
	<tr bgcolor="#FFFFFF">
	  <td>
	  	 <select name="sel3dv">
	  	 <%sbOptCommCode i3dvPCode,strImgDesc%>
	  	 </select>
	  	  <input type="button" id="3dv7" value="�̹���÷��" onClick="jsPopAddImg('3dv',7);" class="button">
	  	  <input type="hidden" name="img3dv" value="<%=strImgUrl%>">
	  	  <span id="div3dv">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=fnSetIconImage(strImgUrl)%>" border="0"></a>  <a href="javascript:jsDelImg('3dv',7);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<%
IF iDFSeq <> "" THEN strImgUrl = arrImg(7,8,3):  strImgDesc = arrImg(7,8,5)
%>
	<tr bgcolor="#FFFFFF">
	  <td>
	  	 <select name="sel3dv">
	  	 <%sbOptCommCode i3dvPCode,strImgDesc%>
	  	 </select>
	  	  <input type="button" id="3dv8" value="�̹���÷��" onClick="jsPopAddImg('3dv',8);" class="button">
	  	  <input type="hidden" name="img3dv" value="<%=strImgUrl%>">
	  	  <span id="div3dv">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=fnSetIconImage(strImgUrl)%>" border="0"></a>  <a href="javascript:jsDelImg('3dv',8);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<%
IF iDFSeq <> "" THEN strImgUrl =  arrImg(7,9,3):  strImgDesc = arrImg(7,9,5)
%>
	<tr bgcolor="#FFFFFF">
	  <td>
	  	 <select name="sel3dv">
	  	 <%sbOptCommCode i3dvPCode,strImgDesc%>
	  	 </select>
	  	  <input type="button" id="3dv9" value="�̹���÷��" onClick="jsPopAddImg('3dv',9);" class="button">
	  	  <input type="hidden" name="img3dv" value="<%=strImgUrl%>">
	  	  <span id="div3dv">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=fnSetIconImage(strImgUrl)%>" border="0"></a>  <a href="javascript:jsDelImg('3dv',9);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<%
IF iDFSeq <> "" THEN strImgUrl =  arrImg(7,10,3):  strImgDesc = arrImg(7,10,5)
%>
	<tr bgcolor="#FFFFFF">
	  <td>
	  	 <select name="sel3dv">
	  	 <%sbOptCommCode i3dvPCode,strImgDesc%>
	  	 </select>
	  	  <input type="button" id="3dv10" value="�̹���÷��" onClick="jsPopAddImg('3dv',10);" class="button">
	  	  <input type="hidden" name="img3dv" value="<%=strImgUrl%>">
	  	  <span id="div3dv">
	  	   <%IF strImgUrl <> ""  THEN%>
	  	  	<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=fnSetIconImage(strImgUrl)%>" border="0"></a>  <a href="javascript:jsDelImg('3dv',10);"><img src='/images/i_delete.gif' border='0'></a>
	  	  <%END IF%>
	  	  </span>
	  </td>
	</tr>
<!--- /3dview --------->
<tr bgcolor="#FFFFFF">
	<td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">����������</td>
	</td>
	<td>
		<input type="radio" name="ismovie" value="N" <% If vIsMovie = "N" Then %>checked<% End If %>>����&nbsp;&nbsp;&nbsp;
		<input type="radio" name="ismovie" value="Y" <% If vIsMovie = "Y" Then %>checked<% End If %>>����
	</td>
</tr>
</table>
<br>
<table width="100%" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
<tr height="50">
	<td  bgcolor="<%= adminColor("gray") %>" align="center">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="10%">
			<input type="button" value="�ҽ��۰����" onclick="window.open('/admin/sitemaster/designfingers/regDF_SourceImage.asp?iDFS=<%=iDFSeq%>','source','width=700,height=527,scrollbars=yes,resizable=yes')" class="button">
			</td>
			<td width="80%" align="center">Add �̹��� (���� 700����)<br>�� ������ �Ҷ��� �̹����� ������ textarea �� �ִ� ������ ��� ����</td>
			<td width="10%" align="right">
			<% If iDFSeq <> "" Then %>
			<input type="button" value="����Ͽ�" onclick="window.open('/admin/sitemaster/designfingers/regDF_MImage.asp?iDFS=<%=iDFSeq%>','mobile','width=400,height=527,scrollbars=yes,resizable=yes')" class="button">
			<% Else %>
			����Ͽ��� �ΰŽ����� ����ϼž� �մϴ�.
			<% End If %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:5 5 5 5">
		<font color="red"><b>�� �ΰŽ� ���� ��Ͻ� ������.</b></font><br>
		1. �������� �� ������ �ݵ�� �� �κп� �����ִ� <b>������������ üũ</b>�ؾ���.<br>
		2. �̹����� �ʼҽ� �κ� �� �� �ϳ��� �Է��� �Ǿ ������ �Ǵ� �����Ͽ� �Է�.<br>
		(����:�ʼҽ� ���� �κп� <b>ĭ��� �̳� ���Ͱ� �� �����δ� ��������� ������ ĭ���, ���Ͱ��� �ϳ��� �����̹Ƿ� ���ǿ��.</b>)<br>
		3. ������ �ҽ��� ����Ҷ��� �̹��� ����� �������� <b>�ʼҽ��� ������ �ҽ��� ���</b>�ؾ���.<br>
		�ƿ﷯ �ҽ� ���뿣 <b>'(��������ǥ) ��� �Ұ�.</b> ����ϰԵǸ� �ҽ����簡 �ȵ�.<br>
		4. mp4�� �ø����ߴ� �������� <strong><FONT style="BACKGROUND-COLOR: #efff81">[ON]����Ʈ����>>������ ���� �� FLV�� ��ȯ</FONT></strong>�Ͽ� �ø�. �ø��� �� �� ����Ʈ���� �ҽ����縦 Ŭ���ؼ� ���.<br>
		5. �Ϲ� �÷��� ������ ����� ���<strong><FONT style="BACKGROUND-COLOR: #efff81">(.swf)�� webimage ������ /video/designfingers/�ش��ΰŽ���ȣ/ �� ���</FONT></strong>�� �����Ͽ� ���.<br>
		<!--
		4. �������� �̸� webimage ������ �÷�����. movie���� �ؿ� designfingers���� �ؿ� �ش� �ΰŽ� ��ȣ ���� �Ʒ��� ������ ��.<br>
		(����:<b>/movie/designfingers/123/test.mp4</b> -> URL�� <b>http://movie.10x10.co.kr/designfingers/123/test.mp4</b> �� ��.)<br>
		5. �������� <b>�׻� mp4�� ��ȯ</b>�ؼ� ����.<br>
		�����̿� mp4�� ��ȯ�ϴ°� ����. [<a href="http://tvpot.daum.net/encoder/PotEncoderSpec.do" target="_blank">http://tvpot.daum.net/encoder/PotEncoderSpec.do</a>]<br>
		��ȯ�� ������������ ��ȯ�ϸ� ��.<br>
		//-->
		<br>
		�� ������ �ҽ� ����<br>
		<font color="blue">* ����</font><br>
		&#60embed src="http://fiximage.10x10.co.kr/flash/flvplayer.swf" width="448" height="324" bgcolor="FFFFFF" allowScriptAccess="always" allowfullscreen="true" flashvars="file=http://movie.10x10.co.kr/designfingers/563/MVI_0630.mp4"&#62<br>
		<font color="blue">* ����Ͽ�</font><br>
		&#60video poster="" src="http://movie.10x10.co.kr/designfingers/544/aaa.mp4" controls="true"&#62&#60/video&#62<br>
		* ùȭ�� �̹��� ���̱� &image=http://webimage.10x10.co.kr/video/vid34.JPEG �Ǵ� poster="http://webimage.10x10.co.kr/image/icon1/16/S1000166075.jpg" �߰��ϸ� ��.<br>
		&#60embed src="http://fiximage.10x10.co.kr/flash/flvplayer.swf" width="448" height="324" bgcolor="FFFFFF" allowScriptAccess="always" allowfullscreen="true" flashvars="file=http://movie.10x10.co.kr/designfingers/563/MVI_0630.mp4&image=http://movie.10x10.co.kr/designfingers/563/vid34.JPEG"&#62<br>
		&#60div style="height:160px; padding:10px 0 0 10px;"&#62&#60video style="position:fixed;" src="http://movie.10x10.co.kr/designfingers/728/otamatone.mp4" controls="true"&#62&#60/video&#62&#60/div&#62<br>
	</td>
</tr>
<!--- Add --------->
<%
Dim i, is_Empty
	If iTotImgCnt = 0 Then
		iTotImgCnt = 5
		is_Empty = "x"
	End If
	For i = 1 To iTotImgCnt
		If is_Empty <> "x" Then
			IF iDFSeq <> "" THEN strImgUrl =arrImg(5,i,3) : strLink =arrImg(5,i,4)
		End If
%>
		<tr bgcolor="#FFFFFF">
			<td style="padding:5 5 5 5">
				<input type="button" id="A<%=i%>" value="�̹���÷��" onClick="jsPopAddImg('add',<%=i%>);" class="button">
				<input type="hidden" name="imgadd" value="<%=strImgUrl%>">
				<span id="divadd">
				<%IF strImgUrl <> "" THEN%>
					<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=strImgUrl%>" border="0" width="200" height="100"></a> <a href="javascript:jsDelImg('add',<%=i%>);"><img src='/images/i_delete.gif' border='0'></a>
				<%END IF%>
				</span>
				<br><br>
				<font color="blue">+ map name="add<%=i%>Map"</font>
				<textarea name="tA<%=i%>" rows="10" cols="75"><%=strLink%></textarea>
			</td>
		</tr>
<%
	Next
%>
<!--- /Add --------->
</table>
</td>
</tr>
<tr>
	<td><input type="button" value="�̹���÷�� �߰�" onClick="Javascript:AutoInsert();" class="button"></td>
	<td align="right"><input type="image" src="/images/icon_save.gif">
		<a href="listDF.asp?menupos=<%=menupos%>&iC=<%=iCurrpage%>"><img src="/images/icon_cancel.gif" border="0"></a></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
