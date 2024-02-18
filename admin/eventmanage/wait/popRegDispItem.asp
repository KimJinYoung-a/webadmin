<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 상품등록
' History : 
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 

<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<!-- #include virtual="/lib/classes/event/eventPartnerWaitCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
dim evtCode
evtCode =    requestCheckVar(Request("eC"),10)

if evtCode = "" then
		Call Alert_close ("유입경로에 문제가 생겼습니다.  ")
end if

dim evtGCode
evtGCode =   requestCheckVar(Request("eGC"),10)
dim menupos
dim itemid, makerid, itemname, waititemid
dim sellyn, isusing, danjongyn, limityn, mwdiv,sailyn
dim page, cdl, cdm, cds, dispCate
dim infodivYn, itemdiv,overseaYN
dim sSort,iPageSize


itemid  	= requestCheckvar(request("itemid"),255)
makerid 	= RequestCheckVar(request("makerid"),32)
itemname 	= RequestCheckVar(request("itemname"),32)

sellyn		= RequestCheckVar(request("sellyn"),10)
isusing 	= RequestCheckVar(request("isusing"),10)
danjongyn = RequestCheckVar(request("danjongyn"),10)
limityn 	= RequestCheckVar(request("limityn"),10)
mwdiv 		= RequestCheckVar(request("mwdiv"),10)
sailyn 		= RequestCheckVar(request("sailyn"),10)
page 			= RequestCheckVar(request("page"),10)
  

dispCate 	= requestCheckvar(request("disp"),16)
infodivYn = requestCheckvar(request("infodivYn"),10) 
itemdiv 	= requestCheckvar(request("itemdiv"),2)
overseaYN	= requestCheckvar(request("overseaYN"),1)
sSort			= requestCheckvar(request("sSort"),10)
if (sellyn="") then sellyn="A"

if (page="") then page=1
if sSort = "" then sSort = "ID"

isusing="Y"

'상품코드 유효성 검사(2008.08.01;허진원)
if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

'==============================================================================
dim oitem
iPageSize = 5

dim i

dim ClsEvt, arrList , intLoop,iTotCnt  

set ClsEvt = new CEvent
ClsEvt.FevtCode = evtCode
makerid =  ClsEvt.fnGetMakerid 
  
ClsEvt.FMakerid = makerid
ClsEvt.FevtGCode = evtGCode
arrList = ClsEvt.fnGetEventGroupItem
iTotCnt = ClsEvt.FTotCnt
set ClsEvt = nothing


set oitem = new CItem

oitem.FRectMakerId =  makerid
oitem.FRectItemid = itemid
oitem.FRectItemName = itemname
oitem.FPageSize = iPageSize
oitem.FCurrPage = page
oitem.FRectDispCate		= dispCate
oitem.FRectSailYn       = sailyn
oitem.FRectSort      =  sSort
if (sellyn <> "A") then
    oitem.FRectSellYN = sellyn
end if

if (isusing <> "A") then
    oitem.FRectIsUsing = isusing
end if


oitem.GetProductList


%>
<html>
	<head>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />

<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript">
	function SubmitSearch(){
	document.frm.submit();
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
		$("#sortable input:checkbox[name='cksel']").attr('checked', true);
		rfrm.target="FrameCKP";
		rfrm.submit();	 
}

function TnSelectItemDel(rfrm){ 
		rfrm.target="FrameCKP";
		rfrm.submit();
	 
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

function jsSetSort(rfrm){
	$("#sortable input:checkbox[name='cksel']").attr('checked', true);
		rfrm.hidM.value="IS";
	//	rfrm.target="FrameCKP";
		rfrm.submit();
}
</script>
</head>
<body>	
<div class="popWinV17">
	<h1>상품 진열</h1>
	<div class="popContainerV17 noScrl">
		<div class="contL">  
			<h2>상품 선택</h2>
			<div id="unitType01" class="unitPannel">
				<form name="frm" method=get>						
						<input type="hidden" name="page" >
						<input type='hidden' name='eC' id='eC' value="<%=evtCode%>">
						<input type='hidden' name='eGC' id='eGC' value="<%=evtGCode%>">
						<div class="searchWrap" style="border-top:none;">
							<div class="search rowSum1">								
								<ul>
									<li>
										<label class="formTit" for="pdtName">상품명 :</label>
										<input type="text" class="formTxt" id="pdtName" style="width:170px" placeholder="상품명 입력" name="itemname" value="<%= itemname %>" size="20" />
									</li>
								</ul>
								<ul>
									<li>
										<label class="formTit" for="ctgy1">카테고리 :</label>
										<!-- #include virtual="/common/module/dispCateSelectBox_upche.asp"-->
									</li>
								</ul>
								<div class="floating1">
									<label class="formTit" for="pdtCode">상품코드 :</label>
									<textarea class="formTxtA" rows="3" id="pdtCode" style="width:120px" placeholder="상품코드 입력" name="itemid" ><%=replace(itemid,",",chr(10))%></textarea>
								</div>
							</div> 
							<dfn class="line"></dfn>
							<div class="search">
								<ul>
									<li>
										<label class="formTit" for="sell">판매 :</label>
										<% drawSelectBoxSellYN "sellyn", sellyn %>
									</li>
									<li>
							   		<label class="formTit" for="sale">할인 :</label><% drawSelectBoxSailYN "sailyn", sailyn %>
							   	</li>
							  </ul>
							</div>		
							<input type="button" class="schBtn" value="검색" onClick="javascript:SubmitSearch();"/> 
						</div>					
						<div class="tbListWrap tMar15">
							<div class="ftRt pad10">
								<span>검색결과 : <strong><%=formatnumber(oitem.FTotalCount,0) %> </strong></span> <span class="lMar10">페이지 : <strong><%=formatnumber(page,0)%> / <%=formatnumber(oitem.FTotalpage,0)%></strong></span>
							</div>
							<ul class="thDataList">
								<li>
									<p class="cell05"><input type="checkbox" onClick="tnCheckAll(this.checked,frm.cksel);" /></p>
									<p class="cell10">상품 ID</p>
									<p class="cell10">이미지</p>
									<p>상품명</p>
									<p class="cell15">판매가</p>
									<p class="cell15">공급가</p>
									<p class="cell10">판매여부</p>
								</li>
							</ul>											    
							<ul class="tbDataList"  id="items">
						<% if oitem.FresultCount > 0 then %>
						    <% for i=0 to oitem.FresultCount-1 %>			
								<li>
									<p class="cell05"><input type="checkbox" name="cksel" id="cksel" value="<%= oitem.FItemList(i).FItemID %>" /></p>
									<p class="cell10"><a href="<%=wwwUrl%>/<%= oitem.FItemList(i).Fitemid %>" target="_blank"><%= oitem.FItemList(i).Fitemid %></a></p>
									<p class="cell10"><img src="<%= oitem.FItemList(i).FImgSmall %>" width="50" height="50" border="0" alt=""></p>
									<p class="lt"><% =oitem.FItemList(i).Fitemname %><input type="hidden" id="itemname<%=i%>" value="<%= oitem.FItemList(i).FItemName %>"></p>
									<p class="cell15">  
										<%= FormatNumber(oitem.FItemList(i).Forgprice,0) %>
										<%
										'할인가
												if oitem.FItemList(i).Fsailyn="Y" then
													Response.Write "<br><font color=#F08050>("&CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsailprice)/oitem.FItemList(i).Forgprice*100) & "%할)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
												end if
												'쿠폰가
												if oitem.FItemList(i).FitemCouponYn="Y" then
													Select Case oitem.FItemList(i).FitemCouponType
														Case "1"
															Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
														Case "2"
															Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
													end Select
												end if
									    %>
									   </p>
									  <p class="cell15" >
									  	<%= FormatNumber(oitem.FItemList(i).Forgsuplycash,0) %>
											<%
											'할인가
												if oitem.FItemList(i).Fsailyn="Y" then
													Response.Write "<br><font color=#F08050>" & FormatNumber(oitem.FItemList(i).Fsailsuplycash,0) & "</font>"
												end if
												'쿠폰가
												if oitem.FItemList(i).FitemCouponYn="Y" then
													if oitem.FItemList(i).FitemCouponType="1" or oitem.FItemList(i).FitemCouponType="2" then
														if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
															Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Forgsuplycash,0) & "</font>"
														else
															Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Fcouponbuyprice,0) & "</font>"
														end if
													end if
												end if
											%>
									  </p>
									<p class="cell10">
										<%if oitem.FItemList(i).Fsellyn="Y" then%>
										<span class="cBl1">판매함</span>
										<%elseif  oitem.FItemList(i).Fsellyn="S" then%>
										<span class="cRd1">일시품절</span>
										<%else%>
										<span class="cRd1">판매안함</span>
										<%end if%>
									</p>
								</li>
							<%next%>	
						<%else%>								
								<li>
									<p>등록된 상품이 없습니다.</p>									 
								</li>
						<%end if%>		
							</ul>
							</form>
							<div class="ct tPad20 cBk1">
								<center>
								<%sbDisplayPaging "page", page, oitem.FTotalCount, iPageSize, 10,menupos %> 
								</center>  
							</div>	
						</div>
					</div>
				</div>

		<input type="button" class="btnMove"  id="additem" title="선택해서 담기" />
				<form name="frmArrupdate" method="post" action="procEvent.asp">
					<input type='hidden' name='item_temp'>
					<input type='hidden' name='hidM' value="IA">
					<input type="hidden" name="check" id="check" value="<% If iTotCnt > 0 Then Response.write iTotCnt Else Response.write 0 %>">
					<input type='hidden' name='checkcnt' id='checkcnt'>
					<input type='hidden' name='eC' id='eC'value="<%=evtCode%>">
					<input type='hidden' name='eGC' id='eGC'value="<%=evtGCode%>">
					<input type='hidden' name='upback' value="N">
					<input type="hidden" name="delid" value="">
					<input type="hidden" name="page" value="<%=page%>">
					<input type="hidden" name="itemid" value="<%=itemid%>">
					<input type="hidden" name="disp" value="<%=dispcate%>">
					<input type="hidden" name="itemname" value="<%=itemname%>">
					<input type="hidden" name="sellyn" value="<%=sellyn%>">
					<input type="hidden" name="sailyn" value="<%=sailyn%>">
		<div class="contR">
			<h2 style="margin-left:-1px;">PC 진열 정보(<%=iTotCnt%>개)</h2>
			<div class="pad10 rt">
				<% dim eImgSize
				if isArray(arrList) Then
					eImgSize = arrList(3,0)
				Else
					eImgSize =153
				end if
					%> 
				<select class="formSlt" id="eImgSize" name="eImgSize" title="정렬 선택">
					<option value="153" <%if eImgSize ="153" then%>selected<%end if%>>4개씩 상품 정렬</option>
					<option value="150" <%if eImgSize ="150" then%>selected<%end if%>>5개씩 상품 정렬</option>
				</select>
			</div>
			<div class="tbListWrap">
				<ul class="thDataList">
					<li>
						<p class="cell10"><input type="checkbox" onClick="tnCheckAll(this.checked,frmArrupdate.cksel);" /></p>
						<p class="cell15">상품 ID</p>
						<p>상품명</p>
						<p class="cell20">판매가격</p>
						<p class="cell12">정렬순서</p>
					</li>
				</ul>
				<div id="sitem">
				<ul id="sortable" class="tbDataList">
					<%if isArray(arrList) Then
							for intLoop = 0 To ubound(arrList,2)
						%>
					<li>
						<p class="cell10"><input type='checkbox' name='cksel' id="cksel<%=intLoop%>" value='<%= arrList(0,intLoop) %>' /></p>
						<p class="cell15"><%=arrList(0,intLoop)%></p>
						<p class="lt"><span><%=arrList(1,intLoop)%></span></p>
						<p class="cell20"><%=FormatNumber(arrList(7,intLoop),0)%>
									<%
										'할인가
												if arrList(4,intLoop)="Y" then
													Response.Write "<br><font color=#F08050>("&CLng((arrList(7,intLoop)-arrList(9,intLoop))/arrList(7,intLoop)*100) & "%할)" & FormatNumber(arrList(9,intLoop),0) & "</font>"
												end if
												'쿠폰가
												if arrList(11,intLoop)="Y" then
													Select Case arrList(12,intLoop)
														Case "1"
															Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(5,intLoop)-(CLng(arrList(13,intLoop)*arrList(5,intLoop)/100)),0) & "</font>"
														Case "2"
															Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(5,intLoop)-arrList(13,intLoop),0) & "</font>"
													end Select
												end if
									%>
								</p>
						<p class="cell12"><input type="text" id="iSort" name="iSort" class="formTxt ct" maxlength="2" style="width:40px;" value="<%=arrList(2,intLoop)%>" /></p>
					</li>
				<%	next
						end if
				%>	
				 
				</ul>
			</div>
			</form>
				<div class="pad10 rt"> 
					<input type="button" class="btn" value="선택삭제" id="delitem" />
				</div>
			</div>
		</div>
	</div>
	<div class="popBtnWrap">
		 <input type="button" value="저장"  class="cRd1" style="width:100px; height:30px;" onClick="jsSetSort(frmArrupdate);"/>  
	</div>
</div>
 
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
		
		if(CheckOverlap){
			alert("선택된 상품과 같은 상품이 있습니다.");
			return false;
		}else if($("#items input:checkbox[name='cksel']:checked").length<1){
			alert("선택된 상품이 없습니다.");
			return false;	
	}else if((Number($("#check").val())+Number($("#items input:checkbox[name='cksel']:checked").length))>105){
			alert("상품은 최대105개까지 선택할 수 있습니다..");
			return false;	
		}else{
			// 행추가
			var oRow;
			$("#items input:checkbox[name='cksel']").each(function(i){
				if($("#items input:checkbox[name='cksel']:eq(" + i + ")").is(":checked")){
					oRow = "					<li id='del" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "'>"
					oRow += "						<p class='cell10'><input type='checkbox' name='cksel' id='cksel'" + i + " value='" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "' /><input type='hidden' name='sitemname' id='sitemname'  value='" + $("#itemname"+i).val() + "'></p>"
					oRow += "						<p class='cell15'>" + $("#items input:checkbox[name='cksel']:eq(" + i + ")").val() + "</p>"
					oRow += "						<p class='lt'><span class='textOverflow'>" + string_cut($("#itemname"+i).val(), 20, "...") + "</span></p>"
					oRow += "					</li>"
					//alert(oRow);
					$("#sitem ul").append(oRow);
					$("#check").val(Number($("#check").val())+1);//옵션 추가 수량 카운트
				}
			});
		}
		tnCheckAll(true,frmArrupdate.cksel);
		document.frmArrupdate.hidM.value="IA";
		document.frmArrupdate.upback.value="Y";
		TnSelectItemReg(frmArrupdate);
	});

var delid="" ;
	$("#delitem").click(function(){
		if($("#sortable input:checkbox[name='cksel']:checked").length<1){
			alert("선택된 상품이 없습니다.");
		}else{
				$("#sortable input:checkbox[name='cksel']").each(function(i){
					if($("#sortable input:checkbox[name='cksel']:eq(" + i + ")").is(":checked")){
					//	$("#del"+ $("#sortable input:checkbox[name='cksel']:eq(" + i + ")").val()).empty();
						if(delid==""){
							delid = $("#sortable input:checkbox[name='cksel']:eq(" + i + ")").val();
						}else{
							delid = delid +"," + $("#sortable input:checkbox[name='cksel']:eq(" + i + ")").val();
						}
						$("#check").val(Number($("#check").val())-1);//옵션 추가 수량 카운트
					}
				
			document.frmArrupdate.hidM.value="IU";
			document.frmArrupdate.upback.value="Y";
			document.frmArrupdate.delid.value = delid;
			TnSelectItemDel(frmArrupdate);
			});
		}
	});
});
function string_cut(str, len, tail) {
	return str.substr(0, len)+tail;
}
$(function() {
	$("#sortable").sortable({
		placeholder:"handling"	,
		stop: function(){			
			$("#sortable li ").each(function(idx){							
				var i = parseInt(idx)+1; 
				$("input[name^='iSort']:eq("+idx+")").val(i);	  
			}); 	
		}
	}).disableSelection();
	 
});
</script>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe> 
</body>
</html>

   $( "#subList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).parent().find("input[name^='eSArr']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).parent().find("input[name^='eSArr']").each(function(){
				$(this).val(i);
				i++;
			}); 
		}
	});