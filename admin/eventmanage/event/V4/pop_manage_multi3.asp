<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  멀티3번 이벤트 설정
' History : 2018.11.05 최종원 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V3.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/Multi3Cls.asp" -->
<%
'공통
dim evt_code, mode, i, y
dim encUsrId, tmpTx, tmpRn, userid, unitIdxAddPram
dim Omulti3
userid = session("ssBctId")

'콘텐츠
dim idx
dim main_copy
dim sub_copy
dim main_color
dim main_content
dim regdate
dim background_img
dim reg_name
dim content_order
dim moddate
dim mod_name

unitIdxAddPram = request("unitIdxAddPram")

'파라미터값
evt_code = request("evt_code")

'content
set Omulti3 = new Multi3
Omulti3.FRectEvtCode = evt_code
Omulti3.GetOneContent

if Omulti3.FResultCount > 0 then 
idx				= Omulti3.FOneContent.C_idx
evt_code		= Omulti3.FOneContent.C_evt_code
main_copy		= Omulti3.FOneContent.C_main_copy	
sub_copy		= Omulti3.FOneContent.C_sub_copy
main_color		= Omulti3.FOneContent.C_main_color
main_content	= Omulti3.FOneContent.C_main_content	
regdate			= Omulti3.FOneContent.C_regdate	
background_img	= Omulti3.FOneContent.C_background_img	
reg_name		= Omulti3.FOneContent.C_reg_name	
content_order	= Omulti3.FOneContent.C_content_order	
moddate			= Omulti3.FOneContent.C_moddate
mod_name		= Omulti3.FOneContent.C_mod_name
end if

if Omulti3.FResultCount > 0 then  '결과값 없을 떄 콘텐츠등록
	mode = "mod"
else
	mode = "contentadd"	
end if

'unit
if idx <> "" then
	Omulti3.FRectContentId = idx
	Omulti3.getContentsUnitList
end if

Randomize()
tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
tmpRn = tmpTx(int(Rnd*26))
tmpRn = tmpRn & tmpTx(int(Rnd*26))
encUsrId = tenEnc(tmpRn & userid)	

dim unitModParam
unitModParam = request("unitModParam")
%>
<script language="javascript" src="colorbox.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script>
$(function(){
    // 창 리사이즈시 testarea 높이 조정
    $(window).resize(function() { 
        $('#tGMap').css('height', $(window).height()-340); 
    }); 
	$(".unitDisp").css("display", "none"); 
	<% if unitIdxAddPram <> "" then %>	
	$(".dspCtr<%=unitIdxAddPram%>").css("display", ""); 	
	<% end if %>
});
function jsCheckUpload() {
	var gubun = document.frmUpload.imgtype.value;
	var mainfrm = document.contentfrm
	var test = $("input[id="+gubun+"]").val();
	// console.log(gubun);	
	// console.log(test);
	// return false;
	if($("#fileupload").val()!="") {
		$("#fileupmode").val("upload");

		$('#ajaxform').ajaxSubmit({
			//보내기전 validation check가 필요할경우
			beforeSubmit: function (data, frm, opt) {
				if(!(/\.(jpg|jpeg|png)$/i).test(frm[0].upfile.value)) {
					alert("JPG,PNG 이미지파일만 업로드 하실 수 있습니다.");
					$("#fileupload").val("");
					return false;
				}
				$("#lyrPrgs").show();
			},
			//submit이후의 처리
			success: function(responseText, statusText){
				var resultObj = JSON.parse(responseText)

				if(resultObj.response=="fail") {
					alert(resultObj.faildesc);
				} else if(resultObj.response=="ok") {					
					$("#filepre").val(resultObj.fileurl);
					$("img[id="+gubun+"src]").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
					$("input[id="+gubun+"]").val(resultObj.fileurl);															
				} else {
					alert("처리중 오류가 발생했습니다.\n" + responseText);
				}
				$("#fileupload").val("");
				$("#lyrPrgs").hide();
			},
			//ajax error
			error: function(err){
				alert("ERR: " + err.responseText);
				$("#fileupload").val("");
				$("#lyrPrgs").hide();
			}
		});
	}
}
function setImgType(type){	
	document.frmUpload.imgtype.value = type;
	return false;
}
function validChk(frm, vMode){
	console.log("frmName : " + frm.name);
	console.log("mode : " + vMode);
	if(vMode == "contentadd" || vMode == "contentmodify"){
		if(frm.background_img.value == ""){
			alert("이미지를 넣어주세요.");
			frm.background_img.focus()
			return false;
		}else if(frm.main_copy.value == ""){
			alert("메인카피를 넣어주세요.");
			frm.main_copy.focus();
			return false;
		}else if(frm.main_color.value == ""){
			alert("메인 컬러코드를 넣어주세요");
			frm.main_color.focus();
			return false;
		}		
	}else if(vMode == "unitadd" || vMode == "unitmodify"){
		if(frm.unit_class.value == ""){
			alert("분류를 넣어주세요.");
			frm.unit_class.focus()
			return false;
		}else if(frm.unit_main_copy.value == ""){
			alert("메인카피를 넣어주세요.");
			frm.unit_main_copy.focus();
			return false;
		}else if(frm.unit_main_content.value == ""){
			alert("내용을 넣어주세요.");
			frm.unit_main_content.focus();
			return false;
		}else if(frm.tag.value == ""){
			alert("태그를 넣어주세요.");
			frm.tag.focus();
			return false;
		}
	}else if(vMode == "itemadd" || vMode == "itemmodify"){
		if(frm.itemid.value == ""){
			alert("상품을 추가해주세요.");
			frm.itemAddBtn.focus()
			return false;
		}else if(frm.item_name.value == ""){
			alert("상품명을 넣어주세요.");
			frm.item_name.focus();
			return false;
		}else if(frm.item_order.value == ""){
			alert("상품 노출 순서를 넣어주세요.");
			frm.item_order.focus();
			return false;
		}
	}
	return true;
}	
function submitForm(mode, idx){
	var frm = document.contentfrm;
	var link = "multi3_process.asp"
	frm.action = link;
	if(!chkValidation(frm)) return false;
	if(mode == "unitdelete"){
		frm.mode.value="unitdelete";		
		frm.unitDeleteIdx.value = idx;
		if(!confirm("삭제하시겠습니까?")){
			return false;
		}
	}else if(mode == "itemdelete"){
		frm.mode.value="itemdelete";		
		frm.itemDeleteIdx.value = idx;
		if(!confirm("삭제하시겠습니까?")){
			return false;
		}
	}
	frm.submit();
}
function chkValidation(frm){
	if(frm.background_img.value==""){
		alert("배경 이미지를 넣어주세요");
		frm.background_img.focus();
		return false;
	}else if(frm.main_copy.value==""){		
		alert("메인 카피를 입력해주세요. ");
		frm.main_copy.focus();
		return false;
	}else if(frm.main_color.value==""){
		alert("메인 컬러를 넣어주세요.");
		frm.main_color.focus();
		return false;
	}
	return true;
} 
function addnewItem(unitIdx){		
		var popwin; 		
		popwin = window.open("multi3_unitItemaddPopup.asp?evt_code=<%=evt_code%>&unitIdx="+unitIdx, "popup_item", "width=576,height=423,scrollbars=yes,resizable=yes");
		popwin.focus();
}
function displayCtr(vClass, Elbtn){
	var targetObj = $("."+vClass)
	
	if(targetObj.css("display")=="none"){
		targetObj.css("display", "")
		Elbtn.innerHTML = "접기"
	}else{
		targetObj.css("display", "none")
		Elbtn.innerHTML = "펼치기"
	}		
}
function jsOpen(sPURL,sTG){ 
	if (sTG =="M" ){ 
		var winView = window.open(sPURL,"popView","width=400, height=600,scrollbars=yes,resizable=yes");
	}
}
function contentUnitAddPopup(){
	var popwin; 			
	popwin = window.open("multi3_contentUnitAddPopup.asp?evtcode=<%=evt_code%>&contentIdx=<%=idx%>", "contentUnit_popup", "width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();	
}
function findItem(itemIdx){		
	var popwin; 		
	popwin = window.open("multi3_eventitem_regist.asp?itemIdx="+itemIdx, "popup_item_search", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}
function chkColor(e){		
	$("#colorBoxDisp").css("background-color",e.value)
}
</script>
<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
<input style="display:none" type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
<input type="hidden" name="mode" id="fileupmode" value="upload">
<input type="hidden" name="div" value="TQ">
<input type="hidden" name="upPath" value="/appmanage/multi3img/">
<input type="hidden" name="tuid" value="<%=encUsrId%>">
<input type="hidden" name="prefile" id="filepre" >	
<input type="hidden" name="imgtype">
</form>		
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> <%=chkIIF(Omulti3.FResultCount > 0, "멀티3번 이벤트 수정", "멀티3번 이벤트 등록" )%></div>
<div>
	<div align="right">이벤트코드 : <%=evt_code%><button type="button" onclick="jsOpen('<%=vmobileUrl%>/event/eventmain.asp?eventid=<%=evt_code%>','M');">미리보기</button>	</div>	
	<h3>상단 고정영역 구성</h2>
</div>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="contentfrm" method="post">		
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="evt_code" value="<%=evt_code %>">
<input type="hidden" name="content_idx" value="<%=idx%>">	
<input type="hidden" name="unitDeleteIdx" value="">	
<input type="hidden" name="itemDeleteIdx" value="">	
<tr>
	<td>
		<table width="100%" style="margin-bottom:30px" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">					
		<!--
		<input type="hidden" name="mode" value="<%=mode%>">
		<input type="hidden" name="evt_code" value="<%=evt_code %>">
		<input type="hidden" name="content_idx" value="<%=idx%>">	
		-->						
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">배경이미지<b style="color:red">*</b></td>
				<td bgcolor="#FFFFFF">
					<div class="inTbSet">												
						<div>	
							<p class="registImg">
								<input type="hidden" id="background_img" name="background_img" value="<%=background_img%>" />
								<img id="background_imgsrc" src="<%=chkIIF(background_img="" or isNull(background_img),"/images/admin_login_logo2.png",background_img)%>" style="height:138px; border:1px solid #EEE;"/>																
							</p>																		
							<button type="button"><div onclick="setImgType('background_img')" ><label for="fileupload" style="cursor:pointer;"><%=chkIIF(background_img="","이미지 업로드","이미지 수정")%></label></div></button>							
						</div>	
					</div>	
				</td>
			</tr>			
			</tr>
			<tr> 
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">메인카피<b style="color:red">*</b></td>
				<td bgcolor="#FFFFFF">  
					<textarea name="main_copy" style="width:90%; height:40px;" ><%=main_copy%></textarea>					
				</td>
			</tr>			
			<tr> 
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">서브카피</td>
				<td bgcolor="#FFFFFF">  
					<textarea name="sub_copy" style="width:90%; height:40px;" ><%=sub_copy%></textarea>					
				</td>
			</tr>		
			<tr>
				<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">메인컬러<b style="color:red">*</b></td>
				<td bgcolor="#FFFFFF">									
					<input type="text" name="prvColor" id="colorBoxDisp"  readonly style="background-color:<%=main_color%>;width:21px;height:21px;border:1px solid #606060;cursor:pointer;" onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop)">					
					<input type="text" onkeyup="chkColor(this);" class='text_ro' name="main_color" size="7" maxlength="7" value="<%=main_color%>">					
					<div id='ColorBox' style='position:absolute;visibility:hidden;right:200;top:100;'></div>									
				</td>
			</tr>							
			<tr> 
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">내용</td>
				<td bgcolor="#FFFFFF">  
					<textarea name="main_content" style="width:90%; height:40px;" ><%=main_content%></textarea>					
				</td>
			</tr>	
		</table>		
<!-- ==============================컨텐츠 유닛 구성============================== -->							
<% if Omulti3.FResultCount > 0 then  '결과값 없을 떄 콘텐츠등록 %>
		<div>
			<h3>컨텐츠 유닛 구성</h3>
		</div>		
		<% if Omulti3.FUnitTotalCount > 0 then %>
			<% for i = 0 to Omulti3.FUnitTotalCount - 1 
				Omulti3.FRectUnitIdx = Omulti3.FUnitList(i).U_idx
				Omulti3.getUnitItemsList								
			%>
				<table width="100%" border="0" align="left" style="margin-top:10px;border:2px solid black" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">										
				<input type="hidden" name="unitIdx" value="<%=Omulti3.FUnitList(i).U_idx%>">				
					<tr>
						<td width="80" align="center" bgcolor="<%= adminColor("tabletop") %>">분류<b style="color:red">*</b></td>
						<td bgcolor="#FFFFFF" style="width:300px">
							#<input type="text" name="unit_class" size="15" value="<%=Omulti3.FUnitList(i).U_unit_class%>" maxlength="32">					
							<button type="button" onclick="displayCtr('dspCtr<%=Omulti3.FUnitList(i).U_idx%>', this);">펼치기</button>					
							<input type="button" value="삭제" onclick="submitForm('unitdelete','<%=Omulti3.FUnitList(i).U_idx%>');">					
						</td>
						<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">유닛순서</td>
						<td bgcolor="#FFFFFF">
							<input type="number" style="width:50px" name="unit_order" value="<%=Omulti3.FUnitList(i).U_unit_order%>" maxlength="4">					
						</td>						
					</tr>				
					<tr class="unitDisp dspCtr<%=Omulti3.FUnitList(i).U_idx%>"> 
						<td align="center" bgcolor="<%= adminColor("tabletop") %>">메인카피</td>
						<td bgcolor="#FFFFFF" colspan=3>  
							<textarea name="unit_main_copy" style="width:90%; height:40px;"><%=Omulti3.FUnitList(i).U_unit_main_copy%></textarea>					
						</td>
					</tr>	
					<tr class="unitDisp dspCtr<%=Omulti3.FUnitList(i).U_idx%>"> 
						<td align="center" bgcolor="<%= adminColor("tabletop") %>">내용</td>
						<td bgcolor="#FFFFFF" colspan=3>  
							<textarea name="unit_main_content" style="width:90%; height:40px;"><%=Omulti3.FUnitList(i).U_unit_main_content%></textarea>					
						</td>
					</tr>		
					<tr class="unitDisp dspCtr<%=Omulti3.FUnitList(i).U_idx%>">
						<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">태그</td>
						<td bgcolor="#FFFFFF" colspan=3><input type="text" name="tag" value="<%=Omulti3.FUnitList(i).U_tag%>" maxlength="100"></td>
					</tr>										
					<% if Omulti3.FItemTotalCount > 0 then %>		
					<tr class="unitDisp dspCtr<%=Omulti3.FUnitList(i).U_idx%>">
						<td width="100" align="center" colspan=4 bgcolor="<%= adminColor("tabletop") %>">상품 리스트</td>
					</tr>									
					<tr class="unitDisp dspCtr<%=Omulti3.FUnitList(i).U_idx%>">						
						<td bgcolor="#FFFFFF" colspan=4>
							<ul style="list-style:none;">
							<% for y=0 to Omulti3.FItemTotalCount - 1 %>
								<li>										
									<table style="border:solid 1px black;margin-top:10px;width:550px;" id="itemContainer">						
									<input type="hidden" name="itemIdx" value="<%=Omulti3.FItemList(y).I_idx%>">
										<tr>
											<td rowspan=3>
												<div class="inTbSet" align="center">												
													<div>	
														<p class="registImg">
															<input type="hidden" id="item_img<%=Omulti3.FItemList(y).I_idx%>" name="item_img" value="<%=Omulti3.FItemList(y).I_item_img%>" />
															<img name="item_imgsrc" id="item_img<%=Omulti3.FItemList(y).I_idx%>src" src="<%=chkIIF(Omulti3.FItemList(y).I_item_img="" or isNull(Omulti3.FItemList(y).I_item_img),"/images/admin_login_logo2.png",Omulti3.FItemList(y).I_item_img)%>" style="height:138px; border:1px solid #EEE;"/>																
														</p>																															
														<button type="button">
															<div onclick="setImgType('item_img<%=Omulti3.FItemList(y).I_idx%>')" >
																<label for="fileupload" style="cursor:pointer;"><%=chkIIF(Omulti3.FItemList(y).I_item_img="","이미지 업로드","이미지 수정")%>
																</label>
															</div>
														</button>														
													</div>	
												</div>					
											</td>
											<td style="border-bottom: 1px solid">상품코드</td>
											<td style="border-bottom: 1px solid">										
												<input type="text" name="itemid" readonly value="<%=Omulti3.FItemList(y).I_itemid%>">										
												<input type="button" onclick="findItem('<%=y%>')" value="상품찾기">
												<input type="button" onclick="submitForm('itemdelete','<%=Omulti3.FItemList(y).I_idx%>');" value="삭제">
											</td>						
										</tr>							
										<tr>
											<td style="border-bottom: 1px solid">상품순서</td>
											<td style="border-bottom: 1px solid">
												<input style="width:50px" type="number" name="item_order" value="<%=Omulti3.FItemList(y).I_item_order%>">
											</td>						
										</tr>			
										<tr>
											<td>상품명</td>
											<td>
												<input type="text" name="item_name" value="<%=Omulti3.FItemList(y).I_item_name%>">
											</td>												
										</tr>										
									</table>										
								</li>
							<% next %>									
							</ul>	
						</td>					
					</tr>
					<% end if %>
					<tr class="unitDisp dspCtr<%=Omulti3.FUnitList(i).U_idx%>">
						<td colspan=4>
							<button style="width:100%" type="button" onclick="addnewItem('<%=Omulti3.FUnitList(i).U_idx%>');">+ 상품추가</button>
						</td>						
					</tr>																																										
				</table>
				<br/>											
			<% next %>		
		<% end if %>		
<% end if %>		

<!-- ==============================컨텐츠 유닛 구성============================== -->								
	</td>		
</tr>
</form>
</table>
<button style="width:100%" type="button" onclick="contentUnitAddPopup();">+ 컨텐츠 유닛 추가</button>
<div style="margin-top: 20px">
	<button style="display:block;margin:0 auto;" type="button" onclick="submitForm();">저장</button>	
</div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
