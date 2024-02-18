<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
'##############################################
' History: 2008.03.12 modify - 2008 리뉴얼 추가 기능 수정
' Description: 디자인 핑거스
'##############################################
Dim clsDF
Dim iDFSeq,sMode
Dim iDFType,sTitle,tContents, dPrizeDate,blnDisplay,blnOtherMall,sUserid,dRegDate,sComment,blnMainDisplay
Dim arrImg,arrItemid, tmpItemid, sProdName, sProdSize, sProdColor, sProdJe, sProdGu, sProdSpe, iTotImgCnt
Dim intLoop, iCurrpage
Dim strImgUrl, strLink, strImgDesc
Dim i3dvPCode : i3dvPCode = 8
	  	
iDFSeq  = requestCheckVar(request("iDFS"),10)
iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
blnDisplay = False
sMode = "I"	'// insert

IF iDFSeq <> "" THEN
	sMode = "U" '// update
	set clsDF = new  CDesignFingers
		clsDF.FDFSeq = iDFSeq
		clsDF.fnGetDFContSourceImage

		arrImg		 	= clsDF.FImg
		iTotImgCnt		= clsDF.FTotImgCnt
	set clsDF = nothing	
	
END IF	

%>
<script language="javascript">
<!--

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
  
 	var inputs = frm.elements.tags("textarea");
 	
	var count = 0;
	for(var i=0;i<inputs.length;i++)
	{
		count = count + 1;
	}
	
	frm.tempcount.value = count;
	
	//alert(count);
	//return false;
	return true;
 
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
				+ "		<input type='button' id='A"+i+"' value='이미지첨부' onClick=jsPopAddImg('source',"+i+"); class='button'>"
				+ "		<input type='hidden' name='imgsource' value=''>"
				+ "		<span id='divsource'></span><br><br><font color='blue'>+ map name='add"+i+"Map'</font>"
				+ "		<textarea name='tA"+i+"' rows='10' cols='75'></textarea>"
				+ "	</td></tr></table></td>"
				+ "</tr>"
 c0.innerHTML = inHtml;
 
//+ "		<span id='divsource'></span><br><br><font color='blue'>+ map name='add"+i+"Map'</font>"
//+ "		<textarea name='tA"+i+"' rows='10' cols='75'></textarea>"
}

function findProd()
{
	window.open('pop_additemlist.asp','findProd','width=900,height=600,scrollbars=yes')
}
 //-->
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<form name="frmReg" method="post" action="procDF_SourceImage.asp" onSubmit="return jsDFSubmit();">  	
<input type="hidden" name="sM" value="<%= sMode%>">
<input type="hidden" name="iDFS" value="<%=iDFSeq %>">
<tr>
	<td colspan="2">
		<table width="100%" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
		<tr height="50">
			<td  bgcolor="<%= adminColor("gray") %>" align="center">Add 이미지 (가로 700이하)<br>※ 삭제를 할때는 이미지를 없앤후 textarea 에 있는 내용을 모두 삭제</td>
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
					IF iDFSeq <> "" THEN strImgUrl =arrImg(25,i,3) : strLink =arrImg(25,i,4)
				End If
		%>
				<tr bgcolor="#FFFFFF">
					<td style="padding:5 5 5 5">
						<input type="button" id="A<%=i%>" value="이미지첨부" onClick="jsPopAddImg('source',<%=i%>);" class="button">
						<input type="hidden" name="imgsource" value="<%=strImgUrl%>">
						<span id="divsource">
						<%IF strImgUrl <> "" THEN%>  	  
							<a href="javascript:jsImgView('<%=strImgUrl%>')"><img src="<%=strImgUrl%>" border="0" width="200" height="100"></a> <a href="javascript:jsDelImg('source',<%=i%>);"><img src='/images/i_delete.gif' border='0'></a>
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
		<img src="/images/icon_cancel.gif" border="0" style="cursor:pointer" onClick="window.close();"></td>
</tr>
<input type="hidden" name="tempcount" value="">
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
