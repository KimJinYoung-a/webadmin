<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v3.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V3.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim eC, page, dispCate, itemname, i, stype, maxDepth, searchdiv, catesort
dispCate = requestCheckvar(request("disp"),10)
stype = requestCheckvar(request("stype"),2)
searchdiv = requestCheckvar(request("searchdiv"),10)
maxDepth = 2 '전시카테고리 2depth까지 보여준다

if eC="" Then eC = requestCheckvar(request("eC"),9)
page = requestCheckvar(request("page"),6)
itemname = requestCheckvar(request("itemname"),100)
if (page="") then page=1

if itemname <> "" then
	if checkNotValidHTML(itemname) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

dim oitem, oitembanner

set oitem = new CItem

oitem.FPageSize         = 5
oitem.FCurrPage         = page
oitem.FRectDispCate   = dispCate
If searchdiv="itemid" Then
oitem.FRectItemid = itemname
Else
oitem.FRectItemName = itemname
End If
oitem.GetItemList

set oitembanner = new CEventBanner
oitembanner.FRectEvtCode = eC
oitembanner.FRectSiteDiv=stype
oitembanner.GetBannerItemList

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
.popWinV17 {overflow:hidden; position:absolute; left:0; top:0; right:0; bottom:0; width:100%; height:100%; font-family:"malgun Gothic","맑은고딕", Dotum, "돋움", sans-serif;}
.popWinV17 h1 {height:40px; padding:12px 15px 0; color:#fff; font-size:17px; background:#c80a0a; border-bottom:1px solid #d80a0a}
.popWinV17 h2 {position:relative; padding:12px 15px; color:#333; font-size:12px; font-weight: bold; background-color:#444; border-top:1px solid #666; font-family:"malgun Gothic","맑은고딕", Dotum, "돋움", sans-serif; z-index:55; color:#fff;}
.popContainerV17 {position:absolute; left:0; top:40px; right:0; bottom:90px; width:100%; border-bottom:1px solid #ddd;}
.contL {position:relative; width:65%; height:100%; border-right:1px solid #ddd; z-index:10; overflow-y:auto;}
.contR {position:absolute; right:0; top:0; bottom:0; width:30%; height:100%; border-left:1px solid #ddd;}
.tbListWrap {position:relative; width:100%; height:100%;}
.tbDataList, .thDataList {display:table; width:100%;}
.tbDataList li, .thDataList li {display:table; width:100%; margin-top:-1px; border-top:1px solid #ddd; border-bottom:1px solid #ddd; }
.thDataList li {height:33px; background-color:#eaeaea; border-top:2px solid #ccc; font-weight:bold;}
.tbDataList li {background-color:#fff; z-index:100;}
.tbDataList li p, .thDataList li p {display:table-cell; padding:7px; text-align:center; vertical-align:middle; line-height:1.4;}
.thDataList li p {white-space:nowrap;}
.handling {background-color:rgba(42,42,57,0.2) !important; height:30px; border:none;}
#sortable li {cursor:move;}
.popBtnWrap {position:absolute; left:0; bottom:0; width:100%; height:60px; text-align:center;}
.textOverflow {width:100%; display:block; text-overflow:ellipsis; overflow:hidden; white-space:nowrap;}
.btnMove {position:absolute; left:67.5%; top:50%; width:40px; height:70px; margin-top:-35px; margin-left:-20px; padding:0; border:none; background:transparent url(/images/btn_move_arrow.png) no-repeat 50% 50%; z-index:1000; cursor:pointer;}
</style>
<script type="text/javascript">
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function selectCategory(frm){
	if ((frm.cdl.selectedIndex<0)||(frm.cdm.selectedIndex<0)||(frm.cds.selectedIndex<0)){
		alert('카테고리를 세단계 모두 선택해주세요.');
		return;
	}

	var cd1 = frm.cdl[frm.cdl.selectedIndex].value;
	var cd2 = frm.cdm[frm.cdm.selectedIndex].value;
	var cd3 = frm.cds[frm.cds.selectedIndex].value;

	var cd1name = frm.cdl[frm.cdl.selectedIndex].text;
	var cd2name = frm.cdm[frm.cdm.selectedIndex].text;
	var cd3name = frm.cds[frm.cds.selectedIndex].text;

	if ((cd1=="")||(cd2=="")||(cd3=="")){
		alert('카테고리를 세단계 모두 선택해주세요.');
		return;
	}

	opener.setCategory(cd1,cd2,cd3,cd1name,cd2name,cd3name);
	window.close();
}

//검색
function jsSearch(){   
	//상품코드 숫자&엔터만 입력가능하도록 체크-----------------------------
	var itemid = document.frm.itemid.value;  
	 itemid =  itemid.replace(",","\r");    //콤마는 줄바꿈처리 
		 for(i=0;i<itemid.length;i++){ 
			if ( itemid.charCodeAt(i) != "13" && itemid.charCodeAt(i) != "10" && "0123456789".indexOf(itemid.charAt(i)) < 0){ 
					alert("상품코드는 숫자만 입력가능합니다.");
					return;
			}
		}  
	//---------------------------------------------------------------------
	if(document.frm.itemname.value==""){
		alert("검색어를 입력하세요.");
		return false;
	}else{
		document.frm.submit();
	}
}

function tnCheckAll(bool, comp){
    var frm = comp.form;
    if (!comp.length){
        comp.checked = bool;
    }else{
        for (var i=0;i<comp.length;i++){
            comp[i].checked = bool;
        }
    }
}

function tnCheckOne(itemid, comp){
	if(comp.value==itemid){
		if(comp.checked){
			comp.checked = false;
		}else{
			comp.checked = true;
		}
	}else{
		for (var i=0;i<comp.length;i++){
			if(comp[i].value==itemid){
				if(comp[i].checked){
					comp[i].checked = false;
				}else{
					comp[i].checked = true;
				}

			}
		}
	}
}

function TnSelectItemReg(rfrm){
	//alert("ok");
	if($("#sortable input:checkbox[name='cksel']:checked").length<1){
		alert("선택된 상품이 없습니다.");
	}else{
		$("#sortable input:checkbox[name='cksel']").attr('checked', true);
		rfrm.target="FrameCKP";
		rfrm.submit();
	}
}

function TnSelectItemDel(rfrm){
	//alert("ok");
	if($("#sortable input:checkbox[name='cksel']:checked").length<1){
		alert("선택된 상품이 없습니다.");
	}else{
		rfrm.target="FrameCKP";
		rfrm.submit();
	}
}

function TnSaveThemeItemBanner(){
	opener.document.frmEvt.target='FrameCKP';
	opener.document.frmEvt.upback.value='Y';
	opener.document.frmEvt.submit();
	//self.close();
}

function TnDelThemeItemBanner(){
	opener.document.frmEvt.target='FrameCKP';
	opener.document.frmEvt.upback.value='Y';
	opener.document.frmEvt.submit();
	$("#sortable input:checkbox[name='cksel']").attr('checked', false);
	location.reload();
}
</script>
<!-- 팝업 사이즈 : 최소 1100*750 -->
<div class="popWinV17">
	<h1>Unit 검색</h1>
	
	<div class="popContainerV17">
		<div class="contL">
			<h2>Unit 선택</h2>
			<div class="tab" style="margin:-1px 0 0 -1px;">
				<ul>
					<li class="col11 selected"><a href="#unitType01">상품</a></li>
					<!-- <li class="col11 "><a href="#unitType02">이벤트</a></li>
					<li class="col11 "><a href="#unitType03">컨텐츠</a></li> -->
				</ul>
			</div>
			<!-- 상품 Tab -->
			<form name="frm" method="get">
			<input type="hidden" name="page" >
			<input type="hidden" name="itemid" >
			<input type='hidden' name='eC' id='eC'value="<%=eC%>">
			<input type='hidden' name='stype' id='stype'value="<%=stype%>">
			<div id="unitType01" class="unitPannel">
				<div class="searchWrap" style="border-top:none;">
					<div class="search">
						<ul>
							<li>
								<label class="formTit">카테고리 :</label>
								<!-- #include virtual="/common/module/dispCateSelectBoxDepth2.asp"-->
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit" for="schWord">검색어 :</label>
								<select name="searchdiv">
									<option value="itemname" selected>상품명</option>
									<option value="itemid">상품코드</option>
								</select>
								<input type="text" class="formTxt" name="itemname" id="schWord" style="width:400px" value="<%=itemname%>" placeholder="검색어를 입력하세요." />
							</li>
						</ul>
					</div>
					<input type="button" class="schBtn" value="검색" onClick="jsSearch();" />
				</div>
				<div class="tbListWrap tMar15">
					<!-- <div class="ftLt lPad10">
						<select class="formSlt" id="deal" title="옵션 선택">
							<option>신상품순</option>
							<option>인기순</option>
						</select>
					</div> -->
					<div class="ftRt pad10">
						<span>검색결과 : <strong><%= FormatNumber(oitem.FTotalCount,0) %></strong></span> <span class="lMar10">페이지 : <strong><%= FormatNumber(page,0) %> / <%= FormatNumber(oitem.FTotalPage,0) %></strong></span>
					</div>
					<ul class="thDataList">
						<li>
							<p class="cell05"><input type="checkbox" onClick="tnCheckAll(this.checked,frm.cksel);" /></p>
							<p class="cell10">상품 ID</p>
							<p class="cell10">이미지</p>
							<p>상품명</p>
							<p class="cell10">가격</p>
							<p class="cell10">업체 ID</p>
							<p class="cell10">판매여부</p>
						</li>
					</ul>
					<ul class="tbDataList" id="items">
						<% for i=0 to oitem.FresultCount-1 %>
						<input type="hidden" id="itemname<%=i%>" value="<%= oitem.FItemList(i).FItemName %>">
						<li>
							<p class="cell05"><input type="checkbox" name="cksel" id="cksel" value="<%= oitem.FItemList(i).FItemID %>" /></p>
							<span onClick="tnCheckOne(<%= oitem.FItemList(i).FItemID %>,frm.cksel);">
							<p class="cell10"><%= oitem.FItemList(i).FItemID %></p>
							<p class="cell10"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0" /></p>
							<p class="lt"><%= oitem.FItemList(i).FItemName %></p>
							<p class="cell10"><%= FormatNumber(oitem.FItemList(i).getRealPrice,0) %></p>
							<p class="cell10"><%= oitem.FItemList(i).FMakerID %></p>
							<p class="cell10"><% if oitem.FItemList(i).Fisusing="R" then %>오픈예약<% ElseIf oitem.FItemList(i).Fisusing="S" then %>일시품절<% ElseIf oitem.FItemList(i).Fisusing="N" then %>품절<% Else %>판매<% End If %></p>
							</span>
						</li>
						<% next %>
					</ul>
					<div class="ct tPad20 cBk1">
						<% if oitem.HasPreScroll then %>
						<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>');">[prev]</a>
						<% else %>
						[pre]
						<% end if %>
						<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
							<% if i>oitem.FTotalpage then Exit for %>
							<% if CStr(page)=CStr(i) then %>
							<span class="cRd1">[<%= i %>]</span>
							<% else %>
							<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
							<% end if %>
						<% next %>
						<% if oitem.HasNextScroll then %>
						<a href="javascript:NextPage('<%= i %>');">[next]</a>
						<% else %>
						[next]
						<% end if %>
					</div>
				</div>
			</div>
			</form>
			<!--// 상품 Tab -->
		</div>
		<input type="button" class="btnMove" id="additem" title="선택해서 담기" />
		<form name="frmArrupdate" method="post" action="doitembannerreg.asp">
		<input type='hidden' name='item_temp'>
		<input type='hidden' name='mode'>
		<input type="hidden" name="check" id="check" value="<% If oitembanner.FresultCount > 0 Then Response.write oitembanner.FresultCount Else Response.write 0 %>">
		<input type='hidden' name='checkcnt' id='checkcnt'>
		<input type='hidden' name='eC' id='eC'value="<%=eC%>">
		<input type='hidden' name='stype' id='stype'value="<%=stype%>">
		<input type='hidden' name='upback' value="N">
		<div class="contR">
			<h2 style="margin-left:-1px;">선택 정보</h2>
			<div class="tbListWrap">
				<ul class="thDataList">
					<li>
						<p class="cell10"><input type="checkbox" onClick="tnCheckAll(this.checked,frmArrupdate.cksel);" /></p>
						<p class="cell25">상품코드</p>
						<p>상품명</p>
					</li>
				</ul>
				<div id="sitem">
				<% If oitembanner.FresultCount > 0 Then %>
				<ul id="sortable" class="tbDataList">
				<% For i=0 To oitembanner.FresultCount-1 %>
					<li id='del<%= oitembanner.FItemList(i).Fitemid %>'>
						<p class='cell10'><input type='checkbox' name='cksel' id="cksel<%=i%>" value='<%= oitembanner.FItemList(i).Fitemid %>' /><input type='hidden' name='sitemname' id='sitemname' value='<%= oitembanner.FItemList(i).Fitemname %>'></p>
						<p class='cell25'><%= oitembanner.FItemList(i).Fitemid %></p>
						<p class='lt'><span class='textOverflow'><%= oitembanner.FItemList(i).Fitemname %></span></p>
					</li>
				<% Next %>
				</ul>
				<% Else %>
				<ul id="sortable" class="tbDataList"></ul>
				<% End If %>
				</div>
				<div class="pad10 rt">
					<input type="button" class="btn" value="선택삭제" id="delitem" />
				</div>
			</div>
		</div>
		</form>
	</div>
	<div class="popBtnWrap">
		<!-- <input type="button" value="선택완료" onclick="TnSelectItemReg(frmArrupdate);" class="cRd1" style="width:100px; height:30px;" /> -->
		<input type="button" value="닫기" onclick="self.close();" style="width:100px; height:30px;" />
	</div>
	
</div>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script>
$(function(){
	$("#additem").click(function(){
		
		var CheckOverlap=false;
		$("#items input:checkbox[name='cksel']").each(function(i){
			if($("#items input:checkbox[name='cksel']:eq(" + i + ")").is(":checked")){
				$("#sortable input:checkbox[name='cksel']").each(function(j){
					//alert($("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + " / " + $("#sortable input:checkbox[name='cksel']:eq(" + j + ")").val());
					if($("#items input:checkbox[name='cksel']:eq(" + i + ")").val() == $("#sortable input:checkbox[name='cksel']:eq(" + j + ")").val()){
						CheckOverlap=true;
						$("#items input:checkbox[name='cksel']:eq(" + i + ")").attr('checked', false);
						return false;
					}
					else{
						CheckOverlap=false;
						return;
					}
				});
			}
		});
		//alert((Number($("#check").val())+Number($("#items input:checkbox[name='cksel']:checked").length)));
		if($("#items input:checkbox[name='cksel']:checked").length<1){
			alert("선택된 상품이 없습니다.");
			return false;
		}else if(CheckOverlap){
			alert("같은 상품이 있습니다.");
			return false;
		}else if((Number($("#check").val())+Number($("#items input:checkbox[name='cksel']:checked").length))>5){
			alert("상품 배너 추가 갯수는 5개 입니다.");
			return false;
		}else{
			// 행추가
			var oRow;
			$("#items input:checkbox[name='cksel']").each(function(i){
				if($("#items input:checkbox[name='cksel']:eq(" + i + ")").is(":checked")){
					oRow = "					<li id='del" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "'>"
					oRow += "						<p class='cell10'><input type='checkbox' name='cksel' id='cksel'" + i + " value='" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "' /><input type='hidden' name='sitemname' id='sitemname'  value='" + $("#itemname"+i).val() + "'></p>"
					oRow += "						<p class='cell25'>" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "</p>"
					oRow += "						<p class='lt'><span class='textOverflow'>" + string_cut($("#itemname"+i).val(), 20, "...") + "</span></p>"
					oRow += "					</li>"
					//alert(oRow);
					$("#sitem ul").append(oRow);
					$("#check").val(Number($("#check").val())+1);//옵션 추가 수량 카운트
				}
			});
		}
		tnCheckAll(true,frmArrupdate.cksel);
		document.frmArrupdate.upback.value="Y";
		TnSelectItemReg(frmArrupdate);
	});

	$("#delitem").click(function(){
		if($("#sortable input:checkbox[name='cksel']:checked").length<1){
			alert("선택");
		}else{
			document.frmArrupdate.mode.value="del";
			document.frmArrupdate.upback.value="Y";
			TnSelectItemDel(frmArrupdate);
		}
	});
});
function string_cut(str, len, tail) {
	return str.substr(0, len)+tail;
}
</script>
<script>
$(function() {
	$("#sortable").sortable({
		placeholder:"handling"
	}).disableSelection();

	$(".tab li").click(function() {
		$(".tab li").removeClass('selected');
		$(this).addClass('selected');
		$('.unitPannel').hide();
		var activeTab = $(this).find("a").attr("href");
		$(activeTab).show();
		return false;
	});
});
chgDispCate('<%=dispCate%>','<%=maxDepth%>');
//document.onload = getOnload();
</script>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->