<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 스타일픽 관리
' Hieditor : 2011.04.07 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->

<%
dim menupos ,omain ,catename ,mainimagelink ,contentsyn , omaincontents ,i , comment
dim mainidx,cd1,mainimage,state,startdate,enddate,isusing,regdate,lastadminid,opendate,closedate,partMDid,partWDid
	mainidx = request("mainidx")
	menupos = request("menupos")
	
'//상세정보
set omain = new cstylepick
	omain.frectmainidx = mainidx
	
	if mainidx <> "" then
		omain.fnGetmain_item()
		
		if omain.ftotalcount > 0 then
			mainimagelink = omain.foneitem.fmainimagelink			
			mainidx = omain.foneitem.fmainidx
			cd1 = omain.foneitem.fcd1
			mainimage = omain.foneitem.fmainimage
			state = omain.foneitem.fstate
			startdate = left(omain.foneitem.fstartdate,10)
			enddate = left(omain.foneitem.fenddate,10)
			isusing = omain.foneitem.fisusing
			regdate = omain.foneitem.fregdate
			lastadminid = omain.foneitem.flastadminid
			opendate = omain.foneitem.fopendate
			closedate = omain.foneitem.fclosedate
			partMDid = omain.foneitem.fpartMDid
			partWDid = omain.foneitem.fpartWDid
			contentsyn = omain.foneitem.fcontentsyn
			comment = omain.foneitem.fcomment
		end if	
	end if
set omain = nothing
	
if isusing = "" then isusing = "Y"
if mainimagelink = "" then mainimagelink = "<map name='Mapmainimage'></map>"	
%>

<script language="javascript">

	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//이미지 확대화면 새창으로 보여주기
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	function jsSetImg(sImg, sName, sSpan){	
		document.domain ="10x10.co.kr";
		
		var winImg;
		winImg = window.open('pop_event_mainimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsNowDate(){
	var mydate=new Date() 
		var year=mydate.getYear() 
		    if (year < 1000) 
		        year+=1900 
		
		var day=mydate.getDay() 
		var month=mydate.getMonth()+1 
		    if (month<10) 
		        month="0"+month 
		
		var daym=mydate.getDate() 
		    if (daym<10) 
		        daym="0"+daym
		        
		return year+"-"+month+"-"+ daym      
	}

	//작업내용 화면 보기
	function jsChkDisp(){
		var tmp = '<%=mainidx%>';
		
		if (tmp==''){
			alert('작업내용 추가등록은 신규등록에서 저장 하실수 없습니다.\n저장후 수정에서 추가등록하세요')
			document.frm.contentsyn.checked = false
			return;
		}
		
		if(document.frm.contentsyn.checked){
			eDetail.style.display = "";
		}else{
			eDetail.style.display = "none";
		}
	}
		
	//저장
	function jsEvtSubmit(){

		if(!frm.cd1.value){
			alert("스타일 카테고리를 선택해주세요");
			frm.cd1.focus();
			return;
		}
	
		if(!frm.state.value){
			alert("상태를 선택해주세요");
			frm.state.focus();
			return;
		}
	
		if(!frm.startdate.value || !frm.enddate.value ){
			alert("이벤트 기간을 입력해주세요");
			return;
		}
	
		if(frm.startdate.value > frm.enddate.value){
			alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			frm.enddate.focus();
			return;
		}
	
		if(!frm.partmdid.value){
			alert("담당 MD를 선택하세요.");
			frm.partmdid.focus();
			return;
		}

		if(!frm.partwdid.value){
			alert("담당 WD를 선택하세요.");
			frm.partwdid.focus();
			return;
		}

		if(!frm.isusing.value){
			alert("사용여부를 선택하세요.");
			frm.isusing.focus();
			return;
		}
					
		var nowDate = jsNowDate();
	
		<%
		'//수정일경우
		if mainidx <> "" then
		%>	
			if(<%=state%>==7 || <%=state%> ==9){
				if(frm.opendate.value != ""){
					nowDate = '<%IF opendate <> "" THEN%><%=FormatDate(opendate,"0000-00-00")%><%END IF%>';
				}
			}
	
			//if(<%=state%>==7 || <%=state%> ==9){
			//	if(frm.startdate.value > nowDate){
			//		alert("시작일이 오픈일보다  늦으면 안됩니다. 시작일을 다시 선택해주세요");
			//	  	frm.startdate.focus();
			//	  	return;
			//	}
			//}
	
			//if(frm.enddate.value < jsNowDate()){
			//	alert("종료일이 현재날짜보다 빠르면 안됩니다. 종료된 이벤트는 수정되지 않습니다");
			//	frm.enddate.focus();
			//	return;
			//}
			
			if (frm.contentsyn.checked){			
				var gubun;	var gubunvaluetmp;
				gubun = document.getElementsByName("gubun")
				gubunvaluetmp = document.getElementsByName("gubunvalue")
	
				for (var i=0; i < gubun.length; i++){			
					if ( gubun[i].value == ''){
						alert("구분을 선택하세요");
						gubun[i].focus();
						return;
					}
				}
											
				for (var i=0; i < gubunvaluetmp.length; i++){			
					if ( gubunvaluetmp[i].value == ''){
						alert("기획전코드나 상품코드를 입력해주세요");
						gubunvaluetmp[i].focus();
						return;
					}
				}
			}
				
		<%
		'//신규등록
		else
		%>
	
			//if(frm.startdate.value < nowDate){
			//	alert("시작일이 현재일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
			//	frm.enddate.focus();
			//	return false;
			//}
	
		<% end if %>
	
		frm.submit();
	}
	
	//tr추가
	function AutoInsert() {
		var f = document.all;
	
		var rowLen = f.div1.rows.length;
		var r  = f.div1.insertRow(rowLen++);
		var c0 = r.insertCell(0);
		
		var Html;
	
		c0.innerHTML = "&nbsp;";
		var inHtml = "&nbsp;&nbsp;&nbsp;<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'><table width='100%' align='center' cellpadding=0 cellspacing=0 border=0 class='a'><tr align='center'><td rowspan=3 valign='top'>디자인소스"+rowLen+"</td><td>구분</td><td align='left'><select name='gubun' onchange='searchcode(this.value,"+rowLen+");'><option value=''>선택하세요</option><option value='1'>스타일픽 기획전</option><option value='2'>상품</option></select><div id='divsub"+rowLen+"'>기획전코드 & 상품코드 : <input type='text' name='gubunvalue' size=10 maxlength=10></div></td></tr><tr align='center'><td>카피</td><td align='left'><input type='text' name='copy' size=90 maxlength=50></td></tr><tr align='center'><td>링크값</td><td align='left'><input type='text' name='link' size=90 maxlength=50></td></tr></table>";
		c0.innerHTML = inHtml;
	}
	
	//tr삭제
	function clearRow(tdObj) {
		if(confirm("저장을 누르셔야 삭제가 완료 됩니다\n 선택하신 파일을 삭제하시겠습니까?") == true) {
			var tblObj = tdObj.parentNode.parentNode.parentNode;
			var trIdx = tdObj.parentNode.parentNode.rowIndex;
		
			tblObj.deleteRow(trIdx);
		} else {
			return false;
		}
	}

	//스타일 & 상품 검색
	function searchcode(gubun,num){
		
		if (gubun.value==''){
			alert('구분을 선택하세요');			
			return;
		}
				
		if (frm.cd1.value==''){
			alert('카테고리를 선택하세요');
			frm.cd1.focus();
			return;
		}
			
		//스타일픽 기획전 
		if (gubun=='1'){
			var searchcode = window.open('/admin/stylepick/stylepick_main_search_event.asp?num='+num+'&cd1=<%=cd1%>','searchcode','width=1024,height=768,scrollbars=yes,resizable=yes');
		
		//상품
		}else if (gubun=='2'){
			var searchcode = window.open('/admin/stylepick/stylepick_main_search_item.asp?num='+num+'&cd1=<%=cd1%>','searchcode','width=1024,height=768,scrollbars=yes,resizable=yes');
		}				
	}
	
</script>

<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="/admin/stylepick/stylepick_main_process.asp">
<input type="hidden" name="mode" value="mainedit">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mainimage" value="<%=mainimage%>">
<input type="hidden" name="opendate" value="<%=opendate%>">
<input type="hidden" name="closedate" value="<%=closedate%>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">번호</td>
	<td bgcolor="#FFFFFF"><%= mainidx %><input type="hidden" name="mainidx" value="<%=mainidx%>"></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">스타일</td>
	<td bgcolor="#FFFFFF"><% Drawcategory "cd1",cd1,"","CD1" %></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">상태</td>
	<td bgcolor="#FFFFFF">
		<% Draweventstate "state" , state ,"" %>		
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">기간</td>
	<td bgcolor="#FFFFFF">
   		<%IF state = "9" THEN%>
   			시작일 : <%=startdate%><input type="hidden" name="startdate" size=10 maxlength=10 value="<%=startdate%>">
   			~ 종료일 : <%=enddate%> <input type="hidden" name="enddate" size=10 maxlength=10 value="<%=enddate%>">
   		<%ELSE%>
   			시작일 : <input type="text" name="startdate" size=10 maxlength=10 value="<%=startdate%>" onClick="jsPopCal('startdate');"  style="cursor:hand;">
   			~ 종료일 : <input type="text" name="enddate" value="<%=enddate%>" size=10 maxlength=10 onClick="jsPopCal('enddate');" style="cursor:hand;">
   		<%END IF%>
   		<%
		if opendate <> "1900-01-01" and opendate <> "" then response.write " 오픈처리일 : " & opendate
		if closedate <> "1900-01-01" and closedate <> "" then response.write " 종료처리일 : " & closedate
		%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">담당MD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partmdid",partmdid,"","11" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">담당WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">사용여부</td>
	<td bgcolor="#FFFFFF">
		<% drawSelectBoxUsingYN "isusing", isusing %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">메인이미지</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnBan2011" value="배너이미지등록" onClick="jsSetImg('<%=mainimage%>','mainimage','mainimagediv')" class="button">
		<div id="mainimagediv" style="padding: 5 5 5 5">
			<%IF mainimage <> "" THEN %>			
				<img src="<%=mainimage%>" border="0" width=100 height=100 onclick="jsImgView('<%=mainimage%>');" alt="누르시면 확대 됩니다">
				<a href="javascript:jsDelImg('mainimage','mainimagediv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">메인이미지맵</td>
	<td bgcolor="#FFFFFF">
		※ 맵 이름 변경 하지 마세요<br>
		<textarea name="mainimagelink" cols="80" rows="6"><%=mainimagelink%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">작업전달사항</td>
	<td bgcolor="#FFFFFF">
		<textarea cols="80" rows="6" name="comment"><%=nl2br(comment)%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" colspan=2>		 
		작업내용 추가등록 <input type="checkbox" name="contentsyn" onClick="jsChkDisp();" <%IF contentsyn= "Y" THEN%>checked<%END IF%>>
		<br><br>※ 프론트 링크 참고
		<br>- 메인
		<br>&nbsp; &nbsp; &nbsp; (카테고리코드링크) : &nbsp; /stylepick/index.asp?cd1=스타일카테고리코드
		<br>&nbsp; &nbsp; &nbsp; (메인번호링크) : &nbsp; /stylepick/index.asp?mainidx=메인번호
		<br>- 기획전
		<br>&nbsp; &nbsp; &nbsp; (카테고리코드링크) : &nbsp; /stylepick/stylepick_collect.asp?cd1=스타일카테고리코드
		<br>&nbsp; &nbsp; &nbsp; (기획전번호링크) : &nbsp; /stylepick/stylepick_collect.asp?evtidx=기획전번호
		<br>- 상품리스트
		<br>&nbsp; &nbsp; &nbsp;  : &nbsp; /stylepick/stylepick_list.asp?cd1=스타일카테고리코드&cd2=분류카테고리코드
		<br>- 상품페이지
		<br>&nbsp; &nbsp; &nbsp;  : &nbsp; /shopping/category_prd_stylepick.asp?itemid=상품번호
	</td>
</tr>
<%
'/수정모드에서만 수정가능
if mainidx <> "" then
%>
<tr align="center"  bgcolor="#FFFFFF" id="eDetail" style="display:<%IF contentsyn="N" THEN%>none;<%END IF%>">
	<td colspan=2>
		<table width="100%" cellpadding="0" cellspacing="0" border=0 class="a" id="div1">
		<%
		set omaincontents = new cstylepick
			omaincontents.frectisusing = "Y"
			omaincontents.frectmainidx = mainidx	
			omaincontents.fnGetmainctList()
		
		'/기존내역 있음
		if omaincontents.fresultcount > 0 then
		
		for i = 0 to omaincontents.fresultcount - 1
		%>
		<tr>
			<td><br>&nbsp;&nbsp;&nbsp;<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
				<table width='100%' align='center' cellpadding=0 cellspacing=0 border=0 class='a'>
				<tr align='center'>
					<td rowspan=3 valign='top'>디자인소스<%=i+1%></td>
					<td>구분</td>
					<td align='left'>
						<select name='gubun' onchange='searchcode(this.value,<%=i+1%>);'>
							<option value='' <% if omaincontents.FItemList(i).fgubun="" then response.write " selected" %>>선택하세요</option>
							<option value='1' <% if omaincontents.FItemList(i).fgubun="1" then response.write " selected" %>>스타일픽 기획전</option>
							<option value='2' <% if omaincontents.FItemList(i).fgubun="2" then response.write " selected" %>>상품</option>
						</select>						
						<div id='divsub<%=i+1%>'>기획전코드 & 상품코드 : <input type='text' name='gubunvalue' value='<%=omaincontents.FItemList(i).fgubunvalue%>' size=10 maxlength=10></div>
					</td>
				</tr>
				<tr align='center'>
					<td>카피</td>
					<td align='left'><input type='text' name='copy' value='<%=trim(omaincontents.FItemList(i).fcopy)%>' size=90 maxlength=50></td>
				</tr>
				<tr align='center'>
					<td>링크값</td>
					<td align='left'><input type='text' name='link' value='<%=trim(omaincontents.FItemList(i).flink)%>' size=90 maxlength=50></td>
				</tr>
				</table>
			</td>
		</tr>
		<%
		next
		
		'/신규
		else
		%>
		<tr>
			<td><br>
				<table width='100%' align='center' cellpadding=0 cellspacing=0 border=0 class='a'>
				<tr align='center'>
					<td rowspan=3 valign='top'>디자인소스1</td>
					<td>구분</td>
					<td align='left'>
						<select name='gubun' onchange='searchcode(this.value,1);'><option value=''>선택하세요</option><option value='1'>스타일픽 기획전</option><option value='2'>상품</option></select>						
						<div id='divsub1'>기획전코드 & 상품코드 : <input type='text' name='gubunvalue' size=10 maxlength=10></div>
					</td>
				</tr>
				<tr align='center'>
					<td>카피</td>
					<td align='left'><input type='text' name='copy' size=90 maxlength=50></td>
				</tr>
				<tr align='center'>
					<td>링크값</td>
					<td align='left'><input type='text' name='link' size=90 maxlength=50></td>
				</tr>
				</table>
			</td>
		</tr>
		<% end if %>			
		</table><br>
		<table width="100%" cellpadding="0" cellspacing="0" border=0 class="a">
		<tr>	
			<td bgcolor="#FFFFFF" colspan=2>	
				<input type="button" value="디자인소스 1개 추가" onClick="AutoInsert()" class="button">
			</td>
		</tr>
		</table>		
	</td>
</tr>
<% end if %>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="button" onclick="jsEvtSubmit();" class="button" value="저장"></td>
</tr>	
</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->