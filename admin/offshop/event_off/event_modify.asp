<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 이벤트
' History : 2010.03.09 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<%
dim evt_code , chkdisp , evt_using , evt_kind , evt_name , evt_startdate ,evt_enddate
dim evt_state , evt_prizedate , opendate ,closedate , brand , partMDid ,evt_forward ,issale
dim evt_comment , regdate , shopid , isgift ,israck ,isprize , isracknum ,racknum  ,img_basic
	evt_code = requestCheckVar(Request("evt_code"),10)	'이벤트코드
	chkdisp	= True
	if evt_using = "" then evt_using = "Y"

	dim cEvtCont, cEvtAddedShop
	set cEvtCont = new cevent_list
		cEvtCont.frectevt_code = evt_code	'이벤트 코드
    set cEvtAddedShop = new cevent_list
		cEvtAddedShop.frectevt_code = evt_code	'이벤트 코드
	'//수정일경우에만 쿼리
	if evt_code <> "" then

		'이벤트 내용 가져오기
		cEvtCont.fnGetEventCont_off
		evt_kind = cEvtCont.FOneItem.fevt_kind
		evt_name = cEvtCont.FOneItem.fevt_name
		evt_startdate = cEvtCont.FOneItem.Fevt_startdate
		evt_enddate = cEvtCont.FOneItem.Fevt_enddate
		evt_prizedate =	cEvtCont.FOneItem.Fevt_prizedate
		evt_state =	cEvtCont.FOneItem.Fevt_state
		IF datediff("d",now,evt_enddate) <0 THEN evt_state = 9 '기간 초과시 종료표기
		regdate	= cEvtCont.FOneItem.fevt_regdate
		evt_using = cEvtCont.FOneItem.Fevt_using
		shopid = cEvtCont.FOneItem.fshopid
		opendate = cEvtCont.FOneItem.fopendate
		closedate = cEvtCont.FOneItem.fclosedate

		'이벤트 화면설정 내용 가져오기
		cEvtCont.fnGetEventDisplay_off
		chkdisp = cEvtCont.FOneItem.FChkDisp
		tmp_cdl = cEvtCont.FOneItem.Fevt_Category
		tmp_cdm	= cEvtCont.FOneItem.fevt_cateMid
		issale = cEvtCont.FOneItem.fissale
		isgift = cEvtCont.FOneItem.fisgift
		israck = cEvtCont.FOneItem.fisrack
		isprize = cEvtCont.FOneItem.fisprize
		isracknum = cEvtCont.FOneItem.fisracknum
		partMDid = cEvtCont.FOneItem.FpartMDid
		evt_forward	= db2html(cEvtCont.FOneItem.Fevt_forward)
		brand = cEvtCont.FOneItem.Fbrand
		evt_comment = cEvtCont.FOneItem.fevt_comment
	 	chkdisp	= cEvtCont.FOneItem.fchkdisp
		img_basic = cEvtCont.FOneItem.fimg_basic


		cEvtAddedShop.getAddedShopList
    end if

Dim i
%>

<script language="javascript">

	//이미지 확대화면 새창으로 보여주기
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function jsSetImg(sImg, sName, sSpan){

		document.domain = '10x10.co.kr';

		var winImg;
		winImg = window.open('pop_event_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	//상세내용 추가등록
	function jsChkDisp(){
	 if(document.frmEvt.chkdisp.checked){
	  	eDetail.style.display = "";
	  }else{
	  	eDetail.style.display = "none";
	  }
	}

	//저장
	function jsEvtSubmit(frm){
		if(!frm.evt_name.value){
			alert("이벤트명을 입력해주세요");
			return;
		}

		if(!frm.shopid.value){
			alert("매장을 선택해주세요");
			return;
		}

		if(!frm.evt_startdate.value || !frm.evt_enddate.value ){
			alert("이벤트 기간을 입력해주세요");
			return;
		}

		if(frm.evt_startdate.value > frm.evt_enddate.value){
			alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			frm.evt_enddate.focus();
			return;
		}

		if(!frm.evt_state.value){
			alert("상태를 선택하세요.");
			return;
		}

		var nowDate = jsNowDate();

		<%
		'//수정일경우
		if evt_code <> "" then
		%>

			if(<%=evt_state%>==7 || <%=evt_state%> ==9){
				if(frm.opendate.value != ""){
					nowDate = '<%IF opendate <> "" THEN%><%=FormatDate(opendate,"0000-00-00")%><%END IF%>';
				}
			}

			//if(<%=evt_state%>==7 || <%=evt_state%> ==9){
			//	if(frm.evt_startdate.value > nowDate){
			//		alert("시작일이 오픈일보다  늦으면 안됩니다. 시작일을 다시 선택해주세요");
			//	  	frm.evt_startdate.focus();
			//	  	return;
			// 	}
			// }

			//if(frm.evt_enddate.value < jsNowDate()){
			//	alert("종료일이 현재날짜보다 빠르면 안됩니다. 종료된 이벤트는 수정되지 않습니다");
			//	return;
			//}

		<%
		'//신규등록
		else
		%>

		  	//if(frm.evt_startdate.value < nowDate){
		  	//	alert("시작일이 현재일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
		  	//	frm.evt_enddate.focus();
		  	//	return false;
		  	//}

		<% end if %>

		if(!frm.evt_comment.value){
			if(GetByteLength(frm.evt_comment.value) > 200){
				alert("comment title은 200자 이내로 작성해주세요");
				frm.evt_comment.focus();
				return;
			}
		}
		frm.submit();
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

	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){

		var winCal;
		var blnSale, blnGift, blnCoupon;
		blnGift= "<%=isprize%>";

		if (sName!="sPD" && blnGift=="isprize"){
			if(confirm("기간을 변경시 해당 이벤트 모두에 적용이 됩니다. 기간을 변경하시겠습니까?")){
				winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
				winCal.focus();
			}
		}else{
				winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
				winCal.focus();
		}
	}

	function jsChType(iVal){
		var frm = document.all;
		if(iVal == "isprize"){
			if (frmEvt.isprize.checked==true){
				frm.div1.style.display = "inline";
			}else{
				frm.div1.style.display = "none";
			}
		}
		if(iVal == "israck"){
			if (frmEvt.israck.checked==true){
				frm.div2.style.display = "inline";
			}else{
				frm.div2.style.display = "none";
			}
		}
		//if(iVal == "issale"){
		//	if(!frmEvt.issale.checked){
		//		if(confirm("할인 설정을 해제할 경우 할인 종료처리됩니다. 설정을 해제하시겠습니까?")){
		//			return;
		//		}else{
		//			frm.checked = true;
		//		}
		//	}
		//}
	}

    // 매장 선택 팝업
	function popShopSelect(){
		var popwin = window.open("/admin/offshop/pop_shopSelect.asp", "popShopSelect","width=460,height=400,scrollbars=yes,resizable=yes");
		popwin.focus();
	}

	// 팝업에서 선택 매장 추가
	function addSelectedShop(shopid,shopname)
	{
	    if (document.frmEvt.shopid.value==shopid){
	        alert("이미 기본 매장에 지정된 매장입니다.");
			return;
	    }

		var lenRow = tbl_addshop.rows.length;

		// 기존에 값에 중복 파트 여부 검사
		if(lenRow>1)	{
			for(l=0;l<document.all.addshopid.length;l++)	{
				if(document.all.addshopid[l].value==shopid) {
					alert("이미 지정된 매장입니다.");
					return;
				}
			}
		}
		else {
			if(lenRow>0) {
				if(document.all.addshopid.value==shopid) {
					alert("이미 지정된 매장입니다.");
					return;
				}
			}
		}

		// 행추가
		var oRow = tbl_addshop.insertRow(lenRow);
		oRow.onmouseover=function(){tbl_addshop.clickedRowIndex=this.rowIndex};

		// 셀추가 (부서,등급,삭제버튼)
		var oCell1 = oRow.insertCell(0);
		var oCell3 = oRow.insertCell(1);

		oCell1.innerHTML = shopid + "/" + shopname + "<input type='hidden' name='addshopid' value='" + shopid + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdShop()' align=absmiddle>";
	}

	// 선택매장 삭제
	function delSelectdShop(){

		if(confirm("선택한 매장을 삭제하시겠습니까?"))
			tbl_addshop.deleteRow(tbl_addshop.clickedRowIndex);
	}

</script>

<form name="frmEvt" method="post" action="event_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="event_edit">
<input type="hidden" name="img_basic" value="<%=img_basic%>">

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<tr>
	<td>  <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이벤트 개요 등록  </font></td>
</tr>
<tr>
	<td>
		<table width="800" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트코드</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0" >
		   			<tr>
		   				<td><%=evt_code%><input type="hidden" name="evt_code" value="<%=evt_code%>"></td>
		   			</tr>
		   			</table>
		   		</td>
		   	</tr>
		    <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>사용유무</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="radio" name="evt_using" value="Y" <%IF evt_using="Y" THEN%>checked<%END IF%>>Yes
		   			<input type="radio" name="evt_using" value="N" <%IF evt_using="N" THEN%>checked<%END IF%>>No
		   		</td>
		   	</tr>
			<tr>
				<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>매장</B></td>
				<td bgcolor="#FFFFFF">
					<% drawSelectBoxOffShopdiv_off "shopid",shopid , "1,3,7" ,"" ,"" %> <!-- shopALL 없앰 -->
				</td>
			</tr>
			<tr>
				<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">추가매장</td>
				<td bgcolor="#FFFFFF">
					<table border="0" cellspacing="0" class="a">
					<tr>
    			        <td >
    			        (사은품 한정 수량이 있을 경우 매장별로 등록 하시기 바랍니다.)
    			        </td>
    			    </tr>
    			    <tr>
    			        <td >
            			    <table name='tbl_addshop' id='tbl_addshop' class=a>
            			    <% if (cEvtAddedShop.FResultCount<1) then %>
            				    <tr onMouseOver='tbl_addshop.clickedRowIndex=this.rowIndex'>
    						    <td><input type='hidden' name='addshopid' value=''></td>
    						    <td></td>
    					        </tr>
    					    <% else %>
    					        <% for i=0 to cEvtAddedShop.FResultCount-1 %>
    					        <tr onMouseOver='tbl_addshop.clickedRowIndex=this.rowIndex'>
    						    <td>
    						        <%= cEvtAddedShop.FItemList(i).FShopid %>/<%= cEvtAddedShop.FItemList(i).FShopname %>
    						        <input type='hidden' name='addshopid' value='<%= cEvtAddedShop.FItemList(i).FShopid %>'></td>
    						    <td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdShop()' align=absmiddle></td>
    					        </tr>
    					        <% next %>
    					    <% end if %>
            			    </table>
    			        </td>
            			<td valign="bottom"><input type="button" class='button' value="추가" onClick="popShopSelect()"></td>
            		</tr>
            		</table>
				</td>
			</tr>
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>종류</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue_off "evt_kind",evt_kind,False,""%>
		   		</td>
		   	</tr>
		   	<tr id="evt_nameTr_A" style="display:<% if evt_kind="16" then Response.Write "none" %>;">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트명</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="evt_name" size="60" maxlength="60" value="<%=evt_name%>">
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>기간</B></td>
		   		<td bgcolor="#FFFFFF">
		   		<%
		   		'// 종료 상태
		   		'IF evt_state = 9 THEN
		   		%>
		   			<!--시작일 : <%'=evt_startdate%><input type="hidden" name="evt_startdate" size="10" value="<%'=evt_startdate%>">
		   			~ 종료일 : <%'=evt_enddate%> <input type="hidden" name="evt_enddate" value="<%'=evt_enddate%>" size="10" >-->
		   		<%
		   		'ELSE
		   		%>
		   			시작일 : <input type="text" name="evt_startdate" size="10" value="<%=evt_startdate%>" onClick="jsPopCal('evt_startdate');"  style="cursor:hand;">
		   			~ 종료일 : <input type="text" name="evt_enddate" value="<%=evt_enddate%>" size="10" onClick="jsPopCal('evt_enddate');" style="cursor:hand;">
		   		<%
		   		'END IF
		   		%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>상태</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptStatusCodeValue_off "evt_state",evt_state,true,""%>
		   			<input type="hidden" name="opendate" value="<%=opendate%>">
		   			<input type="hidden" name="closedate" value="<%=closedate%>">
		   			<%IF opendate <> "" THEN%><span style="padding-left:10px;">  오픈처리일 : <%=opendate%>  </span><%END IF%>
		   			<%IF closedate <> "" THEN%>/ <span style="padding-left:10px;">  종료처리일 : <%=closedate%>  </span><%END IF%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>내용</b></td>
		   		<td bgcolor="#FFFFFF">
		   			상세내용 추가등록 <input type="checkbox" name="chkdisp" onClick="jsChkDisp();" <%IF chkdisp= 1 THEN%>checked<%END IF%>>
		   		</td>
		   	</tr>
		</table>
	</td>

</tr>
<tr>
	<td>
	 <div id="eDetail" style="display:<%IF chkdisp<> 1 THEN%>none;<%END IF%>">
		<table width="800" border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<table width="800" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
					   	<tr>
					   		<td width="100"  align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 카테고리</td>
					   		<td bgcolor="#FFFFFF">
					   			<!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">브랜드</td>
					   		<td bgcolor="#FFFFFF">
					   			<% drawSelectBoxDesignerwithName "brand", brand %>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 타입</td>
					   		<td bgcolor="#FFFFFF">
						    	할인<input type="checkbox" name="issale" value="Y" onclick="jsChType('issale');" <% if issale = "Y" then response.write " checked"%> disabled>
						    	사은품<input type="checkbox" name="isgift" value="Y" <% if isgift = "Y" then response.write " checked"%>>
						    	매대<input type="checkbox" name="israck" value="Y" onclick="jsChType('israck');" <% if israck = "Y" then response.write " checked"%>>
						    	당첨<input type="checkbox" name="isprize" value="Y" onclick="jsChType('isprize');" <% if isprize = "Y" then response.write " checked"%>>
					   			<Br>
								<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
								<tr id="div1" style="display:<% if isprize <> "Y" then response.write "none" %>;">
									<td align="left" bgcolor="FFFFFF">
										당첨 발표일 :
										<input type="text" name="evt_prizedate" value="<%=evt_prizedate%>" size="10" onClick="jsPopCal('evt_prizedate');" style="cursor:hand;">
									</td>
								</tr>
								<tr id="div2" style="display:<% if israck <> "Y" then response.write "none" %>;">
									<td align="left" bgcolor="FFFFFF">
										매대번호:<% getracknum "isracknum" ,isracknum  %>
									</td>
								</tr>
								</table>
					   		</td>
						</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">담당MD</td>
					   		<td bgcolor="#FFFFFF">
					   			<% gettenbytenuser "partMDid", partMDid, "" ,"18" ,"" %>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">작업전달사항</td>
					   		<td bgcolor="#FFFFFF">
					   			<textarea name="evt_forward" rows="15" cols="90"><%=evt_forward%></textarea>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">Comment Title</td>
					   		<td bgcolor="#FFFFFF">
					   			(200자 이내)		   			<Br>
					   			<textarea name="evt_comment" cols="90" rows="2"><%=evt_comment%></textarea>
					   		</td>
					   	</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td style="padding: 10 0 5 0"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 화면이미지 등록</td></tr>
			<tr>
				<td>
					<table width="800" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
					<tr>
				   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">기본이미지</td>
				   		<td bgcolor="#FFFFFF">
				   		<input type="button" name="btnBan2010" value="기본이미지 등록" onClick="jsSetImg('<%=img_basic%>','img_basic','img_basicdiv')" class="button">
				   			<div id="img_basicdiv" style="padding: 5 5 5 5">
				   				<%IF img_basic <> "" THEN %>
				   				※이미지 다운로드 방법: 이미지에서 마우스오른쪽버튼 클릭후	"다른이름으로사진저장" 누르시면 됩니다.
				   				<img src="<%=img_basic%>" border="0" width=400 height=400 onclick="jsImgView('<%=img_basic%>');" alt="누르시면 확대 됩니다">
				   				<a href="javascript:jsDelImg('img_basic','img_basicdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
				   				<%END IF%>
				   			</div>
				   		</td>
				   	</tr>
					</table>
				</td>
			</tr>
		</table>
	</div>
	</td>
</tr>
<tr>
	<td width="800" height="40" align="right">
		<input type="button" onclick="jsEvtSubmit(frmEvt);" value="저장" class="button">
		<input type="button" onclick="self.close();" value="취소" class="button">
	</td>
</tr>
</table>

</form>

<%
set cEvtCont = nothing
set cEvtAddedShop = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
