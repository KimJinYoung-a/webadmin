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
' History: 2008.03.12 modify - 2008 리뉴얼 추가 기능 수정
' Description: 디자인 핑거스
'##############################################
Dim clsDF
Dim iDFSeq,sMode
Dim iDFType,sTitle,tContents, dPrizeDate,blnDisplay,blnOtherMall,sUserid,dRegDate,sComment,blnMainDisplay
Dim arrImg,arrItemid, tmpItemid, sProdName, sProdSize, sProdColor, sProdJe, sProdGu, sProdSpe, iTotImgCnt, vIsMovie, vOpenDate, vTag
Dim intLoop, iCurrpage, edid, emktid
Dim strImgUrl, strLink, strImgDesc
Dim i3dvPCode : i3dvPCode = 8

iDFSeq  = requestCheckVar(request("iDFS"),10)
iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
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

	'//icon 이미지 url 생성
function fnSetIconImage(ByVal sImgURL)
	Dim tmpImg
	tmpImg = split(sImgURL,"/")

	fnSetIconImage = replace(sImgURL,tmpImg(Ubound(tmpImg)),"icon/icon_"&tmpImg(Ubound(tmpImg)))
End Function
%>
<script language="javascript">
<!--

//-- jsPopCal : 달력 팝업 --//
 function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
 }


 //이미지첨부
 function jsPopAddImg(sFolder,sImgID){
 document.domain ="10x10.co.kr";
 	var chkIcon = 0;
 	var winImg;
 	var sImgURL;

 		if(sFolder =="3dv"){chkIcon = 1;} //3dview 이미지일떄만 아이콘 생성
 		if (sImgID == 0){
 			sImgURL = eval("document.frmReg.img"+sFolder).value;
 		}else{
 			sImgURL = eval("document.frmReg.img"+sFolder+"["+(sImgID-1)+"]").value;
 		}
 		winImg = window.open('popAddImage.asp?sF='+sFolder+'&sID='+sImgID+'&chkI='+chkIcon+'&sIU='+sImgURL,'popImg','width=380,height=150');
 		winImg.focus();
 }

 //이미지 삭제
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

 //내용 등록
 function jsDFSubmit(){
 var frm = document.frmReg;
 	 if(!frm.selDFT.value){
 	 	alert("구분을 선택 해주세요");
	  	frm.selDFT.focus();
	  	return false;
 	 }
 	 if(!frm.sT.value){
	  	alert("제목을 입력해주세요");
	  	frm.sT.focus();
	  	return false;
	  }
 	 if(!frm.opendate.value){
	  	alert("오픈일을 입력해주세요");
	  	return false;
	  }
	 if(frm.arrI.value != "")
	 {
	 	var tlng = frm.arrI.value.length;
		var chkchar = frm.arrI.value.substring(tlng,tlng-1);
	 	if(chkchar == ",")
	 	{
	 		alert("상품ID 마지막 콤마는 생략입니다.");
	 		return false;
	 	}
	 }
 }

 //-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
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
				+ "		<input type='button' id='A"+(i-1)+"' value='이미지첨부' onClick=jsPopAddImg('add',"+(i-1)+"); class='button'>"
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
	  	<td align="right"><a href="<%=wwwUrl%>/designfingers/designfingers.asp?fingerid=<%=iDFSeq%>" target="_blank">[Front에서보기]</a></td>
		</tr></table>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">구분</td>
	  <td>
		<select name="selDFT">
			<option value="">선택</option>
			<%sbOptCommCode 10, iDFType%>
		</select>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">제목</td>
	  <td><input type="text" name="sT" value="<%= sTitle %>" size="50" maxlength="32"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">상품ID</td>
	  <td>
	  	<input type="text" name="arrI" value="<%= arrItemid %>" size="50" maxlength="256">
	  	<input type="button" class="button" value="상품찾기" onClick="findProd()">
	  	<br>콤마(<font color="red">,</font>)로 구분해주세요. 마지막 콤마는 생략
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">전시여부</td>
	  <td>
	  	<input type="radio" name="rdoD" value="1" <% if blnDisplay then response.write "checked" %>>전시함
	  	<input type="radio" name="rdoD" value="0" <% if Not blnDisplay then response.write "checked" %>>전시안함
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">외부몰사용여부</td>
	  <td>
	  	<input type="radio" name="rdoOM" value="1" <% if blnOtherMall then response.write "checked" %>>사용함
	  	<input type="radio" name="rdoOM" value="0" <% if Not blnOtherMall then response.write "checked" %>>사용안함
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">Main Top<br>사용여부</td>
	  <td>
	  	<input type="radio" name="rdoMD" value="1" <% if blnMainDisplay then response.write "checked" %>>전시함
	  	<input type="radio" name="rdoMD" value="0" <% if Not blnMainDisplay then response.write "checked" %>>전시안함
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">당첨발표일</td>
	  <td><input type="text" name="dPD" value="<%=dPrizeDate%>" size="10" maxlength="10" onClick="jsPopCal('dPD');"  style="cursor:hand;" class="input_b"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">오픈일</td>
	  <td><input type="text" name="opendate" value="<%=vOpenDate%>" size="10" maxlength="10" onClick="jsPopCal('opendate');"  style="cursor:hand;" class="input_b"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">상품</td>
	  <td><input type="text" name="sPdN" value="<%=sProdName%>" size="50" maxlength="50" > 예시)원목 아인슈타인 시계학습보드</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">크기</td>
	  <td><input type="text" name="sPdS" value="<%=sProdSize%>" size="50" maxlength="50" > 예시)small : 9 x 14, pocket : 9 x 14, large 9 x 14 (cm)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">색상</td>
	  <td><input type="text" name="sPdC" value="<%=sProdColor%>" size="50" maxlength="50" > 예시)black/black, blue/navy, pink/rose, red</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">재료</td>
	  <td><input type="text" name="sPdJ" value="<%=sProdJe%>" size="50" maxlength="50" > 예시)종이</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">구성</td>
	  <td><input type="text" name="sPdG" value="<%=sProdGu%>" size="50" maxlength="50" > 예시)80pages</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="pink" align="center">특징</td>
	  <td><input type="text" name="sPdP" value="<%=sProdSpe%>" size="50" maxlength="50" > 예시)고무줄 밴드와 포켓 없음</td>
	</tr>
   	<tr>
   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">담당디자이너</td>
   		<td bgcolor="#FFFFFF">
   			<%sbGetDesignerid "selDId",edid,""%>
   		</td>
   	</tr>
   	<tr>
   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">담당기획자</td>
   		<td bgcolor="#FFFFFF">
   			<%sbGetwork "selMKTId",emktid,""%>
   		</td>
   	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">Comment</td>
	  <td><input type="text" name="sCom" value="<%=sComment%>" size="80" maxlength="100" ></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" bgcolor="<%= adminColor("gray") %>" align="center">상품설명</td>
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
	<font color="red">+ 이미지당 300kb 이하로 맞춰 주세요 (최대 500kb)</font><br>
	<font color="blue">+ 핑거스 메인 페이지 플래시는 [ON]사이트관리>>[메인]페이지관리 에 적용위치에 보면 2010핑거스메인플래시라고 있습니다.</font><br>
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
	  <td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">Event 당첨 이미지<br>Left(361*178)<br>Right(569*128)</td>
	  <td>
	  	  <input type="button" id="eventLeft" value="Left 이미지첨부" onClick="jsPopAddImg('eventLeft',0);" class="button">
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
		<td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">이벤트 이미지</td>
	  <td>
	  	  <input type="button" id="eventRight" value="이미지첨부" onClick="jsPopAddImg('eventRight',0);" class="button">
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
	  <td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">Main Top 이미지(739*297)</td>
	  <td>
	  	  <input type="button" id="main_top" value="이미지첨부" onClick="jsPopAddImg('main_top',0);" class="button">
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
	  <td width="100" rowspan="3"  bgcolor="<%= adminColor("gray") %>" align="center">Top 이미지</td>
	  <td>
	  	  <input type="button" id="top1" value="이미지첨부" onClick="jsPopAddImg('top',1);" class="button">
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
	  	  <input type="button" id="top2" value="이미지첨부" onClick="jsPopAddImg('top',2);" class="button">
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
	  	  <input type="button" id="top3" value="이미지첨부" onClick="jsPopAddImg('top',3);" class="button">
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
	  <td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">play 이미지</td>
	  <td>
	  	  <input type="button" id="play" value="이미지첨부" onClick="jsPopAddImg('play',0);" class="button">
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
	  <td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">Small 이미지(50*50)</td>
	  <td>
	  	  <input type="button" id="small" value="이미지첨부" onClick="jsPopAddImg('small',0);" class="button">
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
	  <td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">List 이미지(150*150)</td>
	  <td>
	  	 <input type="button" id="list" value="이미지첨부" onClick="jsPopAddImg('list',0);" class="button">
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
	  	  <input type="button" id="3dv1" value="이미지첨부" onClick="jsPopAddImg('3dv',1);" class="button">
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
	  	  <input type="button" id="3dv2" value="이미지첨부" onClick="jsPopAddImg('3dv',2);" class="button">
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
	  	  <input type="button" id="3dv3" value="이미지첨부" onClick="jsPopAddImg('3dv',3);" class="button">
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
	  	  <input type="button" id="3dv4" value="이미지첨부" onClick="jsPopAddImg('3dv',4);" class="button">
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
	  	  <input type="button" id="3dv5" value="이미지첨부" onClick="jsPopAddImg('3dv',5);" class="button">
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
	  	  <input type="button" id="3dv6" value="이미지첨부" onClick="jsPopAddImg('3dv',6);" class="button">
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
	  	  <input type="button" id="3dv7" value="이미지첨부" onClick="jsPopAddImg('3dv',7);" class="button">
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
	  	  <input type="button" id="3dv8" value="이미지첨부" onClick="jsPopAddImg('3dv',8);" class="button">
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
	  	  <input type="button" id="3dv9" value="이미지첨부" onClick="jsPopAddImg('3dv',9);" class="button">
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
	  	  <input type="button" id="3dv10" value="이미지첨부" onClick="jsPopAddImg('3dv',10);" class="button">
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
	<td width="100"  bgcolor="<%= adminColor("gray") %>" align="center">동영상유무</td>
	</td>
	<td>
		<input type="radio" name="ismovie" value="N" <% If vIsMovie = "N" Then %>checked<% End If %>>없음&nbsp;&nbsp;&nbsp;
		<input type="radio" name="ismovie" value="Y" <% If vIsMovie = "Y" Then %>checked<% End If %>>있음
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
			<input type="button" value="소스퍼가기용" onclick="window.open('/admin/sitemaster/designfingers/regDF_SourceImage.asp?iDFS=<%=iDFSeq%>','source','width=700,height=527,scrollbars=yes,resizable=yes')" class="button">
			</td>
			<td width="80%" align="center">Add 이미지 (가로 700이하)<br>※ 삭제를 할때는 이미지를 없앤후 textarea 에 있는 내용을 모두 삭제</td>
			<td width="10%" align="right">
			<% If iDFSeq <> "" Then %>
			<input type="button" value="모바일용" onclick="window.open('/admin/sitemaster/designfingers/regDF_MImage.asp?iDFS=<%=iDFSeq%>','mobile','width=400,height=527,scrollbars=yes,resizable=yes')" class="button">
			<% Else %>
			모바일용은 핑거스부터 등록하셔야 합니다.
			<% End If %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:5 5 5 5">
		<font color="red"><b>※ 핑거스 내용 등록시 주의점.</b></font><br>
		1. 동영상이 들어간 내용은 반드시 윗 부분에 나와있는 <b>동영상유무를 체크</b>해야함.<br>
		2. 이미지나 맵소스 부분 둘 중 하나만 입력이 되어도 저장이 되니 주의하여 입력.<br>
		(예시:맵소스 적는 부분에 <b>칸띄움 이나 엔터값 등 눈으로는 비어있으나 엄연히 칸띄움, 엔터값이 하나의 내용이므로 주의요망.</b>)<br>
		3. 동영상 소스를 등록할때는 이미지 등록은 하지말고 <b>맵소스에 동영상 소스만 등록</b>해야함.<br>
		아울러 소스 내용엔 <b>'(작은따옴표) 사용 불가.</b> 사용하게되면 소스복사가 안됨.<br>
		4. mp4로 올리려했던 동영상은 <strong><FONT style="BACKGROUND-COLOR: #efff81">[ON]사이트관리>>동영상 관리 에 FLV로 변환</FONT></strong>하여 올림. 올리고 난 후 리스트에서 소스복사를 클릭해서 사용.<br>
		5. 일반 플래시 파일을 사용할 경우<strong><FONT style="BACKGROUND-COLOR: #efff81">(.swf)는 webimage 서버에 /video/designfingers/해당핑거스번호/ 의 경로</FONT></strong>에 저장하여 사용.<br>
		<!--
		4. 동영상은 미리 webimage 서버에 올려놓음. movie폴더 밑에 designfingers폴더 밑에 해당 핑거스 번호 폴더 아래에 넣으면 됨.<br>
		(예시:<b>/movie/designfingers/123/test.mp4</b> -> URL은 <b>http://movie.10x10.co.kr/designfingers/123/test.mp4</b> 가 됨.)<br>
		5. 동영상은 <b>항상 mp4로 변환</b>해서 저장.<br>
		다음팟에 mp4로 변환하는게 있음. [<a href="http://tvpot.daum.net/encoder/PotEncoderSpec.do" target="_blank">http://tvpot.daum.net/encoder/PotEncoderSpec.do</a>]<br>
		변환시 아이폰용으로 변환하면 됨.<br>
		//-->
		<br>
		※ 동영상 소스 예시<br>
		<font color="blue">* 웹용</font><br>
		&#60embed src="http://fiximage.10x10.co.kr/flash/flvplayer.swf" width="448" height="324" bgcolor="FFFFFF" allowScriptAccess="always" allowfullscreen="true" flashvars="file=http://movie.10x10.co.kr/designfingers/563/MVI_0630.mp4"&#62<br>
		<font color="blue">* 모바일용</font><br>
		&#60video poster="" src="http://movie.10x10.co.kr/designfingers/544/aaa.mp4" controls="true"&#62&#60/video&#62<br>
		* 첫화면 이미지 보이기 &image=http://webimage.10x10.co.kr/video/vid34.JPEG 또는 poster="http://webimage.10x10.co.kr/image/icon1/16/S1000166075.jpg" 추가하면 됨.<br>
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
				<input type="button" id="A<%=i%>" value="이미지첨부" onClick="jsPopAddImg('add',<%=i%>);" class="button">
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
	<td><input type="button" value="이미지첨부 추가" onClick="Javascript:AutoInsert();" class="button"></td>
	<td align="right"><input type="image" src="/images/icon_save.gif">
		<a href="listDF.asp?menupos=<%=menupos%>&iC=<%=iCurrpage%>"><img src="/images/icon_cancel.gif" border="0"></a></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
