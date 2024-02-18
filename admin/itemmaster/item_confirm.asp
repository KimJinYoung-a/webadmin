<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  승인대기상품 상세리스트
' History : 2014.01.06 정윤정 생성
'			2019.05.28 한용민 수정
' currstate: 0-승인반려,1-승인대기,2-승인보류,5-승인대기(재요청),7-승인완료,9-업체취소
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"-->
<%
Dim sListType, sCurrState, sSort, sMode
Dim dispCate, makerid, itemname, itemcount, itemid
Dim iCurrpage,ipagesize,iTotCnt,iTotalPage
Dim clsWait,arrList, intLoop
Dim clsPartner,idefaultmargine
dim onlyNotSet
dim cdl, cdm, cds, ctrState, upload
	upload =  requestCheckVar(request("upload"),32)
	sListType =  requestCheckVar(request("sLT"),1)
	sCurrstate =  requestCheckVar(request("sCS"),1)
	sSort =  requestCheckVar(request("sS"),2)
	dispCate = requestCheckvar(request("disp"),16)
	makerid	= requestCheckvar(Request("makerid"),32)
	itemname	= requestCheckvar(Request("itemname"),64)
	itemid= requestCheckvar(Request("itemid"),255)
	onlyNotSet =  requestCheckVar(request("onlyNotSet"),1)

	cdl = requestCheckvar(request("cdl"),10)
	cdm = requestCheckvar(request("cdm"),10)
	cds = requestCheckvar(request("cds"),10)

	ctrState= requestCheckvar(request("selCtr"),1)

 	if sCurrState = "" THEN sCurrState = "1"
 	if sSort = "" THEN sSort = "ID"

  if dispcate <> "" and makerid <> "" then
  	iPageSize = 25
  else
 		iPageSize = 50
	end if

	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1

	if itemid<>"" then
		dim iA ,arrTemp,arrItemid
		itemid = replace(itemid,chr(13),"")
		arrTemp = Split(itemid,chr(10))

		iA = 0
		do while iA <= ubound(arrTemp)

			if trim(arrTemp(iA))<>"" then
				'상품코드 유효성 검사(2008.08.04;허진원)
				if Not(isNumeric(trim(arrTemp(iA)))) then
					Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
					dbget.close()	:	response.End
				else
					arrItemid = arrItemid & trim(arrTemp(iA)) & ","
				end if
			end if
			iA = iA + 1
		loop

		if arrItemid<>"" and not(isnull(arrItemid)) then
			itemid = left(arrItemid,len(arrItemid)-1)
		end if
	end if

	if (onlyNotSet = "Y") then
		dispCate = ""
	end if

set clsWait = new CWaitItemlist2014

	if (onlyNotSet = "Y") then
		clsWait.Fcatecode 	= "n"
	else
		clsWait.Fcatecode 	= dispCate
	end if

	clsWait.FRectCate_Large   = cdl
	clsWait.FRectCate_Mid     = cdm
	clsWait.FRectCate_Small   = cds

	clsWait.Fmakerid		= makerid
	clsWait.Fitemname		= itemname
	clsWait.Fcurrstate		= sCurrstate
	clsWait.FSort			= sSort
	clsWait.FPageSize		= iPageSize
	clsWait.FCurrPage		= iCurrPage
	clsWait.Fitemid			= itemid
	clsWait.FRectctrState	= ctrState
	arrList = clsWait.fnGetWaitItemList
	iTotCnt	= clsWait.FTotCnt
 set clsWait = nothing
'  if dispCate ="n" then dispCate = ""
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
dim dctrState
	'검색조건에 브랜드 있을 경우 대표상품계약상태 확인(일괄승인시 체크위함)
	dctrState = 7 '값이 없는 경우 대표계약상태는 계약완료로..
	if makerid <>"" then
		if isArray(arrList) then
		dctrState = arrList(16,0)
		end if
	end if
%>
<style>
	#dialog {display:none; position:absolute;left:100;top:100; z-index:9100;background:#efefef; padding:10px;width:650;}
	#mask {display:none; position:absolute; left:0; top:0; z-index:9000; background:url(http://webadmin.10x10.co.kr/images/mask_bg.png) left top repeat;}
	#boxes .window {position:fixed; left:0; top:0; display:none; z-index:99999;}
</style>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript">

//검색
function jsSearch(sValue){
	if(sValue!=""){
		document.frm.sCS.value = sValue;
	}
	document.frm.submit();
}

function jsSearchBrand(makerid) {
	if(makerid != "") {
		document.frm.makerid.value = makerid;
	}
	document.frm.submit();
}

//리스트 정렬
function jsSort(sValue,i){
	 	document.frm.sS.value= sValue;

		   if (-1 < eval("document.frmList.img"+i).src.indexOf("_alpha")){
	        document.frm.sS.value= sValue+"D";
	    }else if (-1 < eval("document.frmList.img"+i).src.indexOf("_bot")){
	     		document.frm.sS.value= sValue+"A";
	    }else{
	       document.frm.sS.value= sValue+"D";
	    }
		 document.frm.submit();
	}

// 페이지 이동
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}

//-----------------------------------------------
//상태변경
	function jsUniWaitState(itemid){
	var ret = confirm('승인대기로 변경하시겠습니까?');

	if (ret){
		 document.frmList.hidM.value="U";
		 document.frmList.itemid.value = itemid;
		 document.frmList.sCS.value =5;
		 document.frmList.action ="doitemregboru.asp";
		 document.frmList.submit();
	}
}

var chkCnt = 0 ;
 //다중 선택상품 상태변경
function jsMultiWaitState(currstate){
	var mwdiv=''; var deliverytype=''; var deliverfixday=''; var deliverarea='';
	var frm = document.frmList;
	 var itemcount = 0;
	 var count2 = 0;
	if(typeof(frm.chkitem) !="undefined"){
	 	if(!frm.chkitem.length){
	 		if(!frm.chkitem.checked){
	 			alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
	 			return;
	 		}
	 		 frm.itemidarr.value = frm.chkitem.value;
	 		 itemcount = 1;
	 		 if (frm.hidC2.value==2){
	 		 count2 = 1;
	 		}
	  }else{
	  	for(i=0;i<frm.chkitem.length;i++){
	  		if(frm.chkitem[i].checked) {
	  			mwdiv = frm.chkitem[i].getAttribute("mwdiv");
	  			deliverytype = frm.chkitem[i].getAttribute("deliverytype");
	  			deliverfixday = frm.chkitem[i].getAttribute("deliverfixday");
	  			deliverarea = frm.chkitem[i].getAttribute("deliverarea");

				// 배송방법 해외직구 체크
				if (deliverfixday == 'G'){
					if (mwdiv != 'U'){
						alert('해외직구는 업체배송만 선택 가능 합니다.');
						frm.chkitem[i].focus();
						return;
					}
					if ( !(deliverytype=='2' || deliverytype=='9') ){
						alert('해외직구는 업체무료배송과 업체조건배송만 선택 가능 합니다.');
						frm.chkitem[i].focus();
						return;
					}
					if (deliverarea!=''){
						alert('해외직구는 전국배송만 선택 가능 합니다.');
						frm.chkitem[i].focus();
						return;
					}
				}

	  			if (frm.itemidarr.value==""){
	  			 frm.itemidarr.value =  frm.chkitem[i].value;
	  			}else{
	  			 frm.itemidarr.value = frm.itemidarr.value + "," +frm.chkitem[i].value;
	  			}
	  			 itemcount = itemcount+ 1;
	  			 if (frm.hidC2[i].value==2){
					 		 count2 = 1;
					 		}
	  		}

	  	}

	  	if (frm.itemidarr.value == ""){
	  		alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
	 			return;
	  	}
	  }
	}else{
		alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
		return;
	}

 if (currstate ==1){
 		  if ( chkCnt > 0 ){
 	  	alert("이전 데이터를 처리중입니다.잠시 후 다시 승인처리 해주세요");
 	  	return;
 	  }

 	  var dCtrState = "<%=dctrState%>";
 	   if (dCtrState!="7"){
 	  	alert("계약미완료된 브랜드는 승인이 불가능합니다.\n계약확인 후 처리해주세요");
 	  	return;
 	  }

 		if(confirm("선택하신 상품을 승인하시겠습니까?\n업체배송 상품의 경우 프론트에 바로 적용되며, \n텐바이텐배송상품은 입고 완료 후 상품이 오픈됩니다.")){
				  document.itemArrreg.itemid.value = frm.itemidarr.value ;
				  chkCnt ++;
				  $("#btnSubmit").prop("disabled", true);
				   document.itemArrreg.submit();
		}
	}else{
		if(currstate==5){
			if(confirm("선택하신 상품을 승인대기(재요청) 상태로 변경 하시겠습니까?")){
				jsReConfirm(5);
			}
		}
		else{
			if(count2>0&&currstate==2){
				alert("선택하신 상품 중에 3차 보류건이 있습니다.\n해당하는 상품은 승인보류(재등록요청)를 하여도 승인반려(재등록불가) 처리되므로 참고 부탁 드립니다.");
			}

			frm.itemcount.value = itemcount;
				//	var popWin = window.open("item_confirm_pop.asp?sCS="+currstate+"&itemcount="+itemcount,"popW","width=600,height=500");
			$("#dv2").hide();
			$("#dv0").hide();
			$('html, body').animate({scrollTop:0});

			var maskHeight = $(document).height();
			var maskWidth = $(document).width();
			$('#mask').css({'width':maskWidth,'height':maskHeight});
			$('#boxes').show();
			$('#mask').show();
		//	var winH = $(window).height();
		//	var winW = $(document).width();
		//	$("#dialog").css('top', winH/2-$("#dialog").height()/2);
		//	$("#dialog").css('left', winW/2-$("#dialog").width()/2);
			$("#dialog").show();
			$("#dv"+currstate).show();
		}
	}
 }


	$('#mask').click(function () {
		$('#boxes').hide();
		$('.window').hide();
		$('#dialog').hide();
	});


  function jsCancel(){
	document.frmList.hidM.value= "";
	document.frmList.sMsgcd.value= "";
	document.frmList.sMsg.value = "";
	document.frmList.itemcount.value="";
	document.frmList.itemidarr.value="";
	document.frmList.makerid.value="";
	document.frmList.disp.value="";
	document.frmList.itemname.value="";
	document.frmList.itemidarr.value="";
	document.frmList.cdl.value="";
	document.frmList.cdm.value="";
	document.frmList.cds.value="";
	document.frmList.sellCS.value="";
	document.frmList.onlyNotSet.value="";
	document.frmList.selCtr.value="";
	$( "#dialog" ).hide();
  	$('#mask').hide();
  	$('#boxes').hide();
  }


 //승인거부처리
 function jsConfirm(currstate){
 	var chkCount = 0;
 	var iMsgcd = "";
 	var sMsg = "";
 	var iNo = "";
 	for(i=0;i<eval("document.all.chkV"+currstate).length;i++){
 		if(eval("document.all.chkV"+currstate)[i].checked){
 		chkCount = chkCount + 1;
 		iNo = eval("document.all.chkV"+currstate)[i].value;
 		if (iMsgcd==""){
 			iMsgcd = eval("document.all.chkV"+currstate)[i].value;
 			if (eval("document.all.chkV"+currstate)[i].value==999){
 				sMsg = eval("document.all.sM"+currstate).value;
 			}else{
 				sMsg = $("#sp"+currstate+iNo).text();
 			}
 		}else{
 		    iMsgcd = iMsgcd +"^"+ eval("document.all.chkV"+currstate)[i].value;
 			if (eval("document.all.chkV"+currstate)[i].value==999){
 					sMsg = sMsg +"^"+eval("document.all.sM"+currstate).value;
 			}else{
 				sMsg = sMsg +"^"+ $("#sp"+currstate+iNo).text();
 			}
 		}
 	}
 	}
 	if(chkCount == 0){
 		alert("승인 거부 사유를 한개 이상 선택해주세요");
 		return;
 	}

 	document.frmList.sMsgcd.value= iMsgcd;
 	document.frmList.sMsg.value = sMsg;
 	if(document.frmList.hidM.value!="B") {
		document.frmList.hidM.value = "M";
	}
 	document.frmList.sCS.value = currstate;
  document.frmList.submit();
}


//승인처리
function jsApproval(itemid,makerid,ctrstate, mwdiv, deliverytype, deliverfixday, deliverarea){
	if (ctrstate!="7"){
		alert("계약미완료된 브랜드는 승인이 불가능합니다.\n계약확인 후 처리해주세요");
		return;
	}

	// 배송방법 해외직구 체크
	if (deliverfixday == 'G'){
		if (mwdiv != 'U'){
			alert('해외직구는 업체배송만 선택 가능 합니다.');
			return;
		}
		if ( !(deliverytype=='2' || deliverytype=='9') ){
			alert('해외직구는 업체무료배송과 업체조건배송만 선택 가능 합니다.');
			return;
		}
		if (deliverarea!=''){
			alert('해외직구는 전국배송만 선택 가능 합니다.');
			return;
		}
	}

	if ( chkCnt > 0 ){
		alert("이전 데이터를 처리중입니다.잠시 후 다시 승인처리 해주세요");
		return;
	}

	if(confirm("선택하신 상품을 승인하시겠습니까?\n업체배송 상품의 경우 프론트에 바로 적용되며, \n텐바이텐배송상품은 입고 완료 후 상품이 오픈됩니다.")){
		document.itemreg.itemid.value = itemid;
		document.itemreg.makerid.value = makerid;
		chkCnt ++;
		document.itemreg.submit();
	}
}


//상세내용 수정
	function popItemModify(itemid,designer){
	var popwin = window.open('wait_item_modify.asp?itemid=' + itemid + '&designer=' + designer,'waititemmodify','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//전체 선택
function jsChkAll(){
var frm;
frm = document.frmList;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
	   	   	if(frm.chkitem.disabled==false){
		   	 	frm.chkitem.checked = true;
		   	}
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					 	if(frm.chkitem[i].disabled==false){
					frm.chkitem[i].checked = true;
				}
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}
		}
	  }

	}

}

//진행일자 레이어표시
$(document).ready(function(){
 $("div.dlog").click(function(){
 	var divindex =$("div.dlog").index(this);
 	var itemid =$(this).attr("id") ;
 	var url="item_confirm_ajaxLog.asp";
		 var params = "itemid="+itemid;
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){
		 		$("div.dsub").empty().hide();
		 		$("div.dsub").eq(divindex).show();
		 		$("div.dsub").eq(divindex).html(args);
		 	},
		 	error:function(e){
		 		alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
		 //	 alert(e.responseText);
		 	}
	})
	})

	$("div.dlog").mouseleave(function(){
		$("div.dsub").empty().hide();
		})
});

function ViewItemDetail(itemno){
	window.open('<%=replace(manageUrl,"https:","http:")%>/admin/itemmaster/wait_item_preview.asp?itemid='+itemno ,'window1','');
}

function jsPopOption(itemno){
 var winOpt = window.open("/common/pop_upchewaititemoptionedit.asp?itemid="+itemno,"editItemOption","width=800,height=400,scrollbars=yes,resizable=yes");
 winOpt.focus();
}

//승인대기(재요청)처리
 function jsReConfirm(currstate){
 	var chkCount = 0;
 	var iMsgcd = "";
 	var sMsg = "";
 	var iNo = "";
	document.frmList.sMsgcd.value= iMsgcd;
 	document.frmList.sMsg.value = sMsg;
 	document.frmList.hidM.value = "M";
 	document.frmList.sCS.value = currstate;
	document.frmList.submit();
}

// 검색조건 한번에 반려/보류 처리
function jsBatchWaitState(currstate) {
	//검색조건 선택 여부 확인
	var frm = document.frm;
	var cMsg="";

	if(frm.makerid.value=="" && frm.disp1.value=="" && frm.itemname.value=="" && frm.itemid.value=="" && frm.cdl.value=="") {
		alert("검색 조건을 한가지 이상 지정해주세요.");
		return false;
	}

	if(currstate==2){
		cMsg = "일괄처리되는 상품 중에 3차 보류건이 있으면 해당 상품은 승인반려(재등록불가) 처리되므로 참고 부탁드립니다.\n\n";
	}
	cMsg += "검색된 <%=iTotCnt%>건의 상품을 일괄처리하시겠습니까?"
	
	if(!confirm(cMsg)) {
		return false;
	}


	$("#dv2").hide();
	$("#dv0").hide();
	$('html, body').animate({scrollTop:0});

	var maskHeight = $(document).height();
	var maskWidth = $(document).width();
	$('#mask').css({'width':maskWidth,'height':maskHeight});
	$('#boxes').show();
	$('#mask').show();
	$("#dialog").show();
	$("#dv"+currstate).show();

	document.frmList.hidM.value = "B";	//일괄처리
	//검색 데이터 이관
	document.frmList.makerid.value = frm.makerid.value;
	document.frmList.disp.value = frm.disp.value;
	document.frmList.itemname.value = frm.itemname.value;
	document.frmList.itemidarr.value = frm.itemid.value;
	document.frmList.cdl.value = frm.cdl.value;
	document.frmList.cdm.value = frm.cdm.value;
	document.frmList.cds.value = frm.cds.value;
	document.frmList.sellCS.value = frm.sCS.value;
	document.frmList.onlyNotSet.value = (frm.onlyNotSet.checked)?"Y":"N";
	document.frmList.selCtr.value = frm.selCtr.value;
}
 </script>
<form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/doWaitItemToMultiReg_byadmin.asp" onsubmit="return false;" enctype="multipart/form-data">
	<input type="hidden" name="itemid" value="">
	<input type="hidden" name="makerid" value="">
	<input type="hidden" name="sCS" value="<%=sCurrstate%>">
</form>
<form name="itemArrreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/doWaitItemToMultiOneReg_byadmin.asp" onsubmit="return false;" enctype="multipart/form-data">
	<input type="hidden" name="itemid" value="">
</form>
<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
	<tr>
		<td><!-- 검색---------------------------------->
			<form name="frm" method="get" action="">
			<input type="hidden" name="iCP" value="1">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>">
			<input type="hidden" name="sS" value="<%=sSort%>"><!--정렬-->
			<input type="hidden" name="sLT" value="<%=sListType%>"><!--리스트타입(b:브랜드, c:카테고리)-->
			<input type="hidden" name="sCS" value="<%=sCurrstate%>">
				<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" class="a">
					<tr align="center" bgcolor="#FFFFFF">
						<td rowspan="2" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
						<td  bgcolor="#FFFFFF" align="left">
							<table border="0" cellpadding="3" cellspacing="0" class="a">
								<tr>
									<td>브랜드: <%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
									<td> 상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20"></td>
									<td>임시코드:</td>
									<td rowspan="2">
							 			<textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
									</td>
								</tr>
								<tr>
									<td colspan="3">
										전시 카테고리:  <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
										&nbsp;
										관리<!-- #include virtual="/common/module/categoryselectbox.asp"-->
									</td>
								</tr>
								<tr>
									<td colspan="3">
										<input type="checkbox" name="onlyNotSet" value="Y" <% if (onlyNotSet = "Y") then %>checked<% end if %> > 전시카테고리 미지정 상품만
										&nbsp;&nbsp;
										계약상태:
										<select name="selCtr" class="select">
											<option value="">전체</option>
											<option value="Y" <%IF ctrState ="Y" then%>selected<%END IF%>>계약완료</option>
											<option value="N" <%IF ctrState ="N" then%>selected<%END IF%>>계약미완료</option>
										</select>
									</td>
								</tr>
							</table>
						</td>
						<td rowspan="2"  width="50" bgcolor="#EEEEEE">
							<input type="button" class="button_s" value="검색" onClick="jsSearch('');">
						</td>
					</tr>
				</table>
			</form>
		</td><!-- //검색---------------------------------->
	</tr>
	<tr>
		<td>
			<div style="padding:5px"></div>
		</td>
	</tr>
	<tr>
		<td><!-- action ---------------------------------->
				<table width="100%" border="0" cellpadding="5" cellspacing="1"  class="a">
					<tr>
						<td> + 브랜드, 전시카테고리 선택시 [일괄승인] 버튼이 활성화 됩니다. 많은 양의 상품 처리시 속도가 느려질 수 있으니 기다려주세요.
		 <!--input type="button" class="button" value="오픈예약"--></td>
						<td align="right">
						<% if iTotCnt>0 then %>
							<input type="button" class="button" value="승인보류(재등록요청)" onClick="jsMultiWaitState(2);">
							<input type="button" class="button" value="승인반려(재등록불가)" onClick="jsMultiWaitState(0);">
							<% if sCurrstate="2" then %>
							<input type="button" class="button" value="승인대기(재등록)" onClick="jsMultiWaitState(5);">
							<%end if%>
							<%if dispcate <> "" and makerid <> "" and (sCurrstate="1" or sCurrstate="5") then%>
							&nbsp;/&nbsp;
							<input type="button" class="button" value=" 일괄보류 " onClick="jsBatchWaitState(2);" style="background-color:#E8D6E1;">
							<%end if%>
							<%if dispcate <> "" and makerid <> "" and (sCurrstate="1" or sCurrstate="5" or sCurrstate="2") then%>
							<input type="button" class="button" value=" 일괄반려 " onClick="jsBatchWaitState(0);" style="background-color:#F2D6D1;">
							<%end if%>
							<%if dispcate <> "" and makerid <> "" and (sCurrstate="1" or sCurrstate="5" or sCurrstate="A") then%>
							&nbsp;/&nbsp;
							<input type="button" class="button" id="btnSubmit" value="  일괄승인  " onClick="jsMultiWaitState(1);" style="background-color:#D2E6D1;">
							<%end if%>
						<%end if%>
						</td>
					</tr>
				</table>
		</td><!-- //action ---------------------------------->
	</tr>
	<tr>
		<td><!-- List ---------------------------------->
			<form name="frmList" method="post" action="doitemregboru.asp">
			<input type="hidden" name="hidM" value="">
			<input type="hidden" name="itemidarr" value="">
			<input type="hidden" name="itemid" value="">
			<input type="hidden" name="itemcount" value="">
			<input type="hidden" name="sCS" value="">
			<input type="hidden" name="sMsgcd" value="">
			<input type="hidden" name="sMsg" value="">
			<input type="hidden" name="sS" value="<%=sSort%>">
			<input type="hidden" name="sRU" value="item_confirm.asp?sLT=<%=sListType%>&makerid=<%=makerid%>&disp=<%=dispCate%>&sCS=<%=sCurrstate%>&sS=<%=sSort%>">
			<input type="hidden" name="sellCS" value="">
			<input type="hidden" name="makerid" value="">
			<input type="hidden" name="disp" value="">
			<input type="hidden" name="itemname" value="">
			<input type="hidden" name="cdl" value="">
			<input type="hidden" name="cdm" value="">
			<input type="hidden" name="cds" value="">
			<input type="hidden" name="onlyNotSet" value="">
			<input type="hidden" name="selCtr" value="">
			<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
				<tr bgcolor="#FFFFFF">
					<td colspan="15" height="25" align="left">검색결과: <b><%=iTotCnt%></b> &nbsp; 페이지: <b><%=iCurrpage%>/<%=iTotalPage%></b></td>
				</tr>
				<tr class="a" height="25" bgcolor="#DDDDFF" align="center">
					<td><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
					<td width="80" onClick="javascript:jsSort('I','7');" style="cursor:hand;">임시코드 <img src="/images/list_lineup<%IF sSort="ID" THEN%>_bot<%ELSEIF sSort="IA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img7"></td>
					<td>이미지</td>
					<td>컬러칩</td>
					<td width="90" onClick="javascript:jsSort('B','1');" style="cursor:hand;">브랜드ID <img src="/images/list_lineup<%IF sSort="BD" THEN%>_bot<%ELSEIF sSort="BA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
					<td onClick="javascript:jsSort('N','2');" style="cursor:hand;">상품명 <img src="/images/list_lineup<%IF sSort="ND" THEN%>_bot<%ELSEIF sSort="NA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
					<td width="60" onClick="javascript:jsSort('S','3');" style="cursor:hand;">판매가 <img src="/images/list_lineup<%IF sSort="SD" THEN%>_bot<%ELSEIF sSort="SA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img3"></td>
					<td width="60" onClick="javascript:jsSort('A','4');" style="cursor:hand;">매입가 <img src="/images/list_lineup<%IF sSort="AD" THEN%>_bot<%ELSEIF sSort="AA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img4"></td>
					<td>옵션 추가가격</td>
					<td>거래구분</td>
					<td width="40" onClick="javascript:jsSort('M','5');" style="cursor:hand;">마진 <img src="/images/list_lineup<%IF sSort="MD" THEN%>_bot<%ELSEIF sSort="MA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img5"></td>
					<td>한정</td>
					<td>전시카테고리 <font color="blue">(+추가카테고리)</font></td>
					<td width="160" onClick="javascript:jsSort('L','6');" style="cursor:hand;">진행일자 <img src="/images/list_lineup<%IF sSort="LD" THEN%>_bot<%ELSEIF sSort="LA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img6"></td>
					<td width="100">
						<select name="selCS" class="select" onChange="jsSearch(this.value);">
							<%sbOptItemWaitStatus sCurrState%>
						</select>
					</td>
				</tr>
				<%IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)
					%>
				<tr bgcolor="<%if arrList(16,intLoop) <> 7 THEN%>#DDDDFF<%else%>#FFFFFF<%end if%>" align="center">
				 	<td><input type="checkbox" name="chkitem" value="<%= arrList(0,intLoop) %>" mwdiv="<%= arrList(12,intLoop) %>" deliverytype="<%= arrList(21,intLoop) %>" deliverfixday="<%= arrList(22,intLoop) %>" deliverarea="<%= trim(arrList(23,intLoop)) %>" <%IF arrList(7,intLoop) <> 1 and arrList(7,intLoop) <> 2 and arrList(7,intLoop) <> 5  THEN%>disabled<%END IF%>>
				 		<input type="hidden" name="hidC2" value="<%=arrList(9,intLoop)%>">
				 		</td>
					<td><a href="javascript:popItemModify('<% =arrList(0,intLoop) %>','<%=arrList(2,intLoop) %>')"><%=arrList(0,intLoop)%></a>
						<br/>
						<%if arrList(16,intLoop) <> 7 THEN%>
					 	[계약미완료]
						<%end if%>
						</td>
					<td><%IF arrList(1,intLoop) <> "" THEN
						dim imgsubdir
						imgsubdir = GetImageSubFolderByItemid(arrList(0,intLoop))
						%>
						<img src="<%=partnerUrl%>/waitimage/basic/<%=imgsubdir%>/<%= arrList(1,intLoop)%>" width="50" height="50">
						<%END IF%>
					</td>
					<td><table border="0" cellpadding="0" cellspacing="1" bgcolor="#dddddd"><tr><td bgcolor="#FFFFFF"><img src="<%=webImgUrl & "/color/colorchip/" & arrList(24,intLoop)%>" width="12" height="12" hspace="2" vspace="2"></td></tr></table></td>
					<td><a href="javascript:jsSearchBrand('<%=arrList(2,intLoop)%>')"><%=arrList(2,intLoop)%></a></td>
					<td>
						<%=arrList(3,intLoop)%>
						<a href="javascript:ViewItemDetail('<%=arrList(0,intLoop)%>');"><font color="blue">[미리보기]</font></a>
						<%
							Dim keyword, chk
							keyword = arrList(3,intLoop)
							If InStr(keyword, "_") > 0 Then
								chk = InStr(keyword, "_") - 1
								keyword = Mid(keyword, 1, chk)
							End If
							keyword = URLEncodeUTF8(keyword)

							Response.Write "<a href='http://shopping.naver.com/search/all.nhn?query="& keyword &"&pagingIndex=1&pagingSize=40&viewType=list&sort=rel' target='"& arrList(0,intLoop) &"'><font color='blue'>[최저가 확인하기]</font></a>"
						%><br>
						<font color="grey"><%=DDotFormat(arrList(25,intLoop),40)%></font>
					</td>
					<td width="60" align="right"><%=formatnumber(arrList(5,intLoop),0)%>&nbsp;</td>
					<td width="60" align="right"><%=formatnumber(arrList(4,intLoop),0)%>&nbsp;</td>
					<td><a href="javascript:jsPopOption('<%= arrList(0,intLoop) %>');"><%if arrList(20,intLoop) >0 then%><font color=red>Y</font><%else%>N<%end if%></a></td>
					<td><%IF arrList(11,intLoop) <> arrList(12,intLoop) THEN%><font color="red"><%end if%><%=mwdivName(arrList(12,intLoop))%></td>
					<td width="40" align="right"><%IF arrList(6,intLoop) <> arrList(10,intLoop)  THEN%><font color=red><%END IF%><%=arrList(6,intLoop)%>%&nbsp;</td>
					<td><% if arrList(15,intLoop)="Y" then %>
						<font color=red>한정</font><%=arrList(13,intLoop)-arrList(14,intLoop) %>
						<% end if %>
					</td>
					<td align="left"><a href="javascript:popItemModify('<% =arrList(0,intLoop) %>','<%=arrList(2,intLoop) %>')">
						<% if Not isNull(arrList(18,intLoop)) then Response.write replace(arrList(18,intLoop),"^^",">") %> &nbsp;<%if arrList(19,intLoop)  > 0 then%><font color="blue"><%end if%>(+<%=arrList(19,intLoop)%>)</a></td>
					<td width="160"><div id="<%= arrList(0,intLoop) %>" class="dlog" style="cursor:hand;" ><%=arrList(8,intLoop)%></div>
						<div style="position:relative;background-color:#eeeeee">
						 <div id="dLogSub" class="dsub" style="position:absolute;left:-80px;top:0px;z-index:100;background-color:white;"></div>
					 </div>
						</td>
					<td><font color="<%=GetCurrStateColor(arrList(7,intLoop))%>"><%=GetCurrStateName(arrList(7,intLoop))%></font>
							<% if (arrList(7,intLoop)="2") or (arrList(7,intLoop)="0") then %>
							<span style="line-height:23px;"><a href="javascript:jsUniWaitState('<%=arrList(0,intLoop) %>')"><br><font color="#000000">[승인대기변경]</font></a></span>
							<% elseif  (arrList(7,intLoop)="1") or (arrList(7,intLoop)="5") then%>
						 	&nbsp;<input id="btnApp" name="btnApp" type="button" class="button" value="▶승인" style="color:blue;" onclick="jsApproval('<%=arrList(0,intLoop)%>','<%=arrList(2,intLoop)%>','<%=arrList(16,intLoop)%>','<%= arrList(12,intLoop) %>','<%= arrList(21,intLoop) %>','<%= arrList(22,intLoop) %>','<%= trim(arrList(23,intLoop)) %>')">
							<% end if %>
					</td>
				</tr>
				<%	Next
					ELSE
				%>
				<tr bgcolor="#ffffff">
					<td align="center" colspan="14">등록된 내용이 없습니다.</td>
				</tr>
				<%
				END IF%>
			</table>
</form>
		</td><!-- //List ---------------------------------->
	</tr>
	<!-- 페이지 시작 -->
<% 	Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10

		iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1

		If (iCurrPage mod iPerCnt) = 0 Then
			iEndPage = iCurrPage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
					    <tr valign="bottom" height="25">
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if CLng(ix) = CLng(iCurrPage) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if CLng(iTotalPage) > CLng(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>
					    </tr>
					</table>
				</td>
			</tr>
			</table>
	</td>
	</tr>
</table>
<div id="boxes">
<div id="mask"></div>
<div id="dialog">
<!-- #include virtual="/admin/itemmaster/item_confirm_inc.asp"-->
</div>
</div>

<script type="text/javascript">

<% IF not(isArray(arrList)) THEN %>
	<%
	' 이미지 서버 다녀옴
	if upload="on" then
	%>
		<% if sListType="C" then %>
			frm.makerid.value="";
			frm.submit();
		<% end if %>
	<% end if %>
<% end if %>

</script>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
