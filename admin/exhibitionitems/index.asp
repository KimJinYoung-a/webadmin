<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기획전 메인 (크리스마스 , 뭐뭐뭐... 뭐뭐뭐... 등등등)
' History : 2018-11-05 이종화
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionCls.asp"-->
<%
dim limited , itemid , poscode
dim isusing , mdpick
dim mastercode , detailcode
dim menu : menu = "exhibition"

dim sBrand,arrItemid
isusing = request("isusingbox")
sBrand 	= request("ebrand")
arrItemid = request("aitem")
mdpick 	= request("mdpick")
poscode = request("menupos")
itemid	= requestCheckvar(request("itemid"),255)

mastercode = requestCheckvar(request("mastercode"),10)
detailcode = requestCheckvar(request("detailcode"),10)

if mastercode = "" then mastercode = 0
if detailcode = "" then detailcode = 0

dim page , i
	page = requestCheckVar(request("page"),5)
	if page = "" then page = 1

if itemid<>"" then
	dim iA ,arrTemp',arrItemid
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

dim oExhibition
set oExhibition = new ExhibitionCls
	oExhibition.FPageSize = 50
	oExhibition.FCurrPage = page
	oExhibition.FrectMasterCode = mastercode
	oExhibition.FrectDetailCode = detailcode
	oExhibition.FrectMakerid = sBrand
	oExhibition.FRectArrItemid = arrItemid
	oExhibition.Frectpick = mdpick
	oExhibition.getItemsList

%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript" src="/admin/common/lib/js/front.js"></script>
<script type="text/javascript">
// 신규 등록 팝업
function popRegNew(mastercode , detailcode){
	var popRegNew = window.open('/admin/exhibitionitems/pop_reg_item.asp?mastercode='+ mastercode +'&detailcode='+ detailcode,'popRegNew','width=600,height=350,status=yes')
	popRegNew.focus();
}

// 신규 등록 팝업 여러 상품
function popRegItems(mastercode , detailcode){
	var popRegItems = window.open('/admin/exhibitionitems/pop_reg_items.asp?mastercode='+ mastercode +'&detailcode='+ detailcode,'popRegItems','width=1000,height=750,scrollbars=yes,resizable=yes')
	popRegItems.focus();
}

// 수정 팝업
function popRegModi(idx){
	var popRegModi = window.open('/admin/exhibitionitems/pop_reg_item.asp?mode=edit&idx='+ idx,'popRegModi','width=600,height=350')
	popRegModi.focus();
}

// 이벤트 관리 팝업
function popEventManage(m){
	var popEventManage = window.open('/admin/exhibitionitems/pop_list_event.asp?mastercode='+m,'popEventManage','width=1124,height=768,resizable=yes,scrollbars=yes')
	popEventManage.focus();
}

// 브랜드 관리 팝업
function popBrandManage(m){
	var popBrandManage = window.open('/admin/exhibitionitems/pop_list_brand.asp?mastercode='+m,'popBrandManage','width=1124,height=768,resizable=yes,scrollbars=yes')
	popBrandManage.focus();
}

// 이벤트 링크 관리 팝업
function popEventLinkManage(m){
	var popEventLinkManage = window.open('/admin/exhibitionitems/pop_eventLink_list.asp?mastercode='+m,'popEventLinkManage','width=1124,height=768,resizable=yes,scrollbars=yes')
	popEventLinkManage.focus();
}

// 그룹 관리 팝업
function popGroupManage() {
	var popGroupManage = window.open('/admin/exhibitionitems/pop_exhibition_manage.asp','popRegNew','width=750,height=750,status=yes')
	popGroupManage.focus();
}

// mdpick 순서 변경 팝업
function popMdpickSort(m,d){
	var popMdpickSort = window.open('/admin/exhibitionitems/pop_pickitems.asp?mastercode='+ m +'&detailcode='+ d,'popMdpickSort','width=1024,height=768,resizable=yes,scrollbars=yes')
	popMdpickSort.focus();
}

// 상품 삭제
function fnDelItem(idx) {
	if (confirm("상품을 삭제 하시겠습니까?") == true){ 
		var frm = document.itemdel
		frm.eidx.value = idx;
		frm.submit();
	}
}

// pick 전체선택
var ichk;
ichk = 1;
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.name=="chkI")){
			if ((e.type=="checkbox")) {
				e.checked = blnChk ;
			}
		}
	}
}

// formcheck + datainsert
var fnCheckedItem = function() {
	var inputArgument1 = document.forms["fitem"].elements[arguments[0]];
	var inputArgument2 = document.forms["fitem"].elements[arguments[1]];
	var tempOutArgument = document.forms["frmreg"].elements[arguments[2]];
	var sValue = ""; //idx

	if (inputArgument1.length > 1){
		for (var i=0;i<inputArgument1.length;i++){
			if (inputArgument1[i].checked){
				if (sValue==""){
					sValue = inputArgument2[i].value;
				}else{
					sValue =sValue+","+inputArgument2[i].value;
				}
			}
		}
	}else{
		if(inputArgument1.checked){
			sValue = inputArgument2.value;
		}
	}
	return tempOutArgument.value = sValue;
}

// pick 일괄 저장
function jsChangedAction() {
	if (confirm("pick사용여부를 저장 하시겠습니까?") == true){    //확인
		fnCheckedItem('chkI','chkI','eid');
		document.frmreg.mode.value = "mdpick"
		document.frmreg.submit();
	}else{
	    return;
	}
}

// pick 일괄 삭제
function fnDeleteItems() {
	if (confirm("상품을 일괄삭제 하시겠습니까?") == true){    //확인
		fnCheckedItem('chkI','chkI','eid');
		document.frmreg.mode.value = "itemdelete"
		document.frmreg.submit();
	}else{
	    return;
	}
}

// 상품명 추가 입력
function fnAddTextItems() {
	if (confirm("상품명 추가 입력을 하시겠습니까?") == true){    //확인
		fnCheckedItem('chkI','chkI','eid');
		fnCheckedItem('chkI','addtext1','addtext1');
		fnCheckedItem('chkI','addtext2','addtext2');
		document.frmreg.mode.value = "addsubtext"
		document.frmreg.submit();
	}else{
	    return;
	}
}

function popreport(mastercode,detailcode,idx) {
	var popreport = window.open('/admin/dataanalysis/report/simpleQry.asp?qryidx=236&detailcode='+detailcode+'&mastercode='+mastercode+'&idx='+idx,'dataanalysisreport','width=1280,height=800,scrollbars=yes,resizable=yes');
	popreport.focus();
}

</script>
<div class="content scrl" style="top:40px;">
	<div class="pad20">
		<!-- 상단 검색폼 시작 -->
		<div>
			<form name="refreshFrm" method="get">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>">
			<input type="hidden" name="page" value="">
			<table class="tbType1 listTb">
				<tr bgcolor="#FFFFFF">
					<th width="80" bgcolor="<%= adminColor("gray") %>">검색조건</th>
					<td style="text-align:left;">
						<% DrawMainPosCodeCombo "mastercode", mastercode ,"" %>
						<% if mastercode > 0 then %>
							<% DrawDetailSelectBox "detailcode" , detailcode , mastercode %>
						<% end if %>
						<select name="mdpick">
							<option value="">PICK여부</option>
							<option value="0" <% if mdpick = "0" then response.write " selected"%>>X</option>
							<option value="1" <% if mdpick = "1" then response.write " selected"%>>O</option>
						</select>
						<br><br>&nbsp;브랜드:
						<% drawSelectBoxDesignerwithName "ebrand", sBrand %>
					</td>
					<th>상품 코드</th>
					<td style="text-align:left;">
						<textarea rows="4" cols="80" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
					</td>
					<td width="50" bgcolor="<%= adminColor("gray") %>">
						<input type="button" class="button_s" value="검색" onclick="refreshFrm.submit();">
					</td>
				</tr>
			</table>
			</form>
		</div>
		<!-- 검색 끝 -->
		<!-- 액션 시작 -->
		<div class="tPad15">
			<form name="frmarr" method="post" action="">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>">
			<input type="hidden" name="mode" value="">
			<table class="tbType1 listTb">
			<tr>
				<th width="150">코드 관리</th>
				<td style="text-align:left;">
					<div style="float:left;">
						<input type="button" value="코드 관리" onclick="popGroupManage();" class="button">
					</div>
				</td>
			</tr>
			<% if mastercode > 0 then %>
			<tr>
				<th>이벤트 관리</th>
				<td style="text-align:left;">
					<input type="button" value="<%=getMasterCodeName(mastercode)%> 이벤트 관리" onclick="popEventManage('<%=mastercode%>');" class="button">
					<input type="button" value="<%=getMasterCodeName(mastercode)%> 기획전/이벤트 링크 관리" onclick="popEventLinkManage('<%=mastercode%>');" class="button">
					<input type="button" value="<%=getMasterCodeName(mastercode)%> 브랜드 관리" onclick="popBrandManage('<%=mastercode%>');" class="button">
				</td>
			</tr>
			<tr>
				<th>슬라이드 관리</th>
				<td style="text-align:left;">
					<input type="button" value="<%=getMasterCodeName(mastercode)%> 슬라이드 관리" onclick="popSlideManage('<%=mastercode%>','<%=menu%>');" class="button">
					<div style="float:right;">
						<strong>미리보기 : </strong>
						<input type="button" class="button" value="<%=getMasterCodeName(mastercode)%>" onclick="popSlideView('<%=mastercode%>','0','<%=menu%>')">&nbsp;
						<%=DrawDetailButtons(mastercode,"popSlideView",menu)%>
					</div>
				</td>
			</tr>
			<tr>
				<th>상품 순서 관리</th>
				<td style="text-align:left;">
					<div>
						<input type="button" value="<%=getMasterCodeName(mastercode)%> BEST PICK 순서관리" onclick="popMdpickSort('<%=mastercode%>','0');" class="button">
					</div>
					<div class="tPad15">
						<%=DrawDetailButtons(mastercode,"popMdpickSort","")%>
					</div>
				</td>
			</tr>
			<tr>
				<th>통계보기</th>
				<td style="text-align:left;">
					<div>
						<input type="button" value="통계보기" onclick="popreport('<%= mastercode %>','<%= detailcode %>','');" class="button">
					</div>
					<div class="tPad15">
					</div>
				</td>
			</tr>
			<% end if %>
			</table>
			</form>
			<form name="frmreg" method="post" action="/admin/exhibitionitems/lib/exhibition_proc.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="eid" value="">
				<input type="hidden" name="addtext1" value="">
				<input type="hidden" name="addtext2" value="">
				<input type="hidden" name="poscode" value="<%=poscode%>">
				<input type="hidden" name="page" value="<%=page%>">
				<input type="hidden" name="mastercode" value="<%=mastercode%>">
			</form>
		</div>
		<% if mastercode > 0 then %>
		<div style="padding-top:15px;padding-bottom:15px;">
			<div style="float:right;">
				<input type="button" value="선택상품 추가타이틀 입력" class="button" onclick="fnAddTextItems();">&nbsp;
				<input type="button" value="선택상품 삭제" class="button" onclick="fnDeleteItems();">&nbsp;
				<input type="button" value="선택상품 BEST PICK 저장" class="button" onClick="jsChangedAction();">
			</div>
		</div>
		<% end if %>
		<div class="tPad15">
			<!-- 리스트 시작 -->
			<table class="tbType1 listTb">
			<form name="fitem" method="post" style="margin:0px;">
				<tr height="25" bgcolor="FFFFFF">
					<td colspan="12" style="text-align:left;">
						<div>
							<div style="float:left;">
								검색결과 : <b><%= oExhibition.FTotalCount %></b>
								&nbsp;
								페이지 : <b><%= page %>/ <%= oExhibition.FTotalPage %></b>
							</div>
							<div style="float:right;">
								<input type="button" value="복수상품등록" class="button" onclick="popRegItems('<%=mastercode%>','<%=detailcode%>');">
								<input type="button" value="단일상품등록" class="button" onclick="popRegNew('<%=mastercode%>','<%=detailcode%>');">
							</div>
						</div>
					</td>
				</tr>
				<% IF oExhibition.FResultCount>0 Then %>
				<tr bgcolor="<%= adminColor("tabletop") %>">
					<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
					<td>번호</td>
					<td>구분 </td>
					<td>이미지 </td>
					<td>상품번호 </td>
					<td>상품명 </td>
					<td>BESTPICK 여부</td>
					<td>업체아이디 </td>
					<td>판매가</td>
					<td>마진</td>
					<td>계약구분 </td>
					<td>상품삭제 </td>
					<td>비고</td>
				</tr>
				<% For i =0 To  oExhibition.FResultCount -1 %>
				<tr bgcolor="#FFFFFF">
					<td >
						<input type="checkbox" name="chkI" onClick="AnCheckClick(this);" value="<%= oExhibition.FItemList(i).Fidx %>"/>
						<input type="hidden" name="checkflag" value="">
					</td>
					<td><a href="javascript:popRegModi(<%= oExhibition.FItemList(i).Fidx%>);"><%= oExhibition.FItemList(i).Fidx %> </a></td>
					<td style="text-align:left;padding:auto;">
						<%=getMasterCodeName(oExhibition.FItemList(i).Fmastercode)%>
						<% if oExhibition.FItemList(i).Fdetailcode > 0 then %>
						<br/><br/>
						ㄴ<%=getDetailCodeName(oExhibition.FItemList(i).Fmastercode,oExhibition.FItemList(i).Fdetailcode)%>
						<% end if %>
					</td>
					<td><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=oExhibition.FItemList(i).Fitemid%>" target="_blank"><img src="<%= db2html(oExhibition.FItemList(i).FImageList) %>" width="40" height="40" border="0" style="cursor:pointer"></a></td>
					<td><%= oExhibition.FItemList(i).Fitemid %> </a></td>
					<td><%= oExhibition.FItemList(i).fitemname %>
						<% if mastercode > 0 then %>
						<div>
							<ul>
								<li>추가(1) : <input type='text' name='addtext1' value='<%=oExhibition.FItemList(i).Faddtext1%>' size='40'></li>
								<li>추가(2) : <input type='text' name='addtext2' value='<%=oExhibition.FItemList(i).Faddtext2%>' size='40'></li>
							</ul>
						</div>
						<% end if %>
					</td>
					<td><%= chkiif(oExhibition.FItemList(i).Fpickitem = 1,"<span style='color:red'>Y</span>","N")%></td>
					<td><%= oExhibition.FItemList(i).fmakerid %> </td>
					<td>
						<%
						Response.Write FormatNumber(oExhibition.FItemList(i).Forgprice,0)
						'할인가
						if oExhibition.FItemList(i).Fsailyn="Y" then
							Response.Write "<br><font color=#F08050>("&CLng((oExhibition.FItemList(i).Forgprice-oExhibition.FItemList(i).Fsailprice)/oExhibition.FItemList(i).Forgprice*100) & "%할)" & FormatNumber(oExhibition.FItemList(i).Fsailprice,0) & "</font>"
						end if
						'쿠폰가
						if oExhibition.FItemList(i).FitemCouponYn="Y" then
							Select Case oExhibition.FItemList(i).FitemCouponType
								Case "1"
									Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oExhibition.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
								Case "2"
									Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oExhibition.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
							end Select
						end if
					%>
					</td><%'판매가%>
					<td>
						<%
						Response.Write fnPercent(oExhibition.FItemList(i).Forgsuplycash,oExhibition.FItemList(i).Forgprice,1)
						'할인가
						if oExhibition.FItemList(i).Fsailyn="Y" then
							Response.Write "<br><font color=#F08050>" & fnPercent(oExhibition.FItemList(i).Fsailsuplycash,oExhibition.FItemList(i).Fsailprice,1) & "</font>"
						end if
						'쿠폰가
						if oExhibition.FItemList(i).FitemCouponYn="Y" then
							Select Case oExhibition.FItemList(i).FitemCouponType
								Case "1"
									if oExhibition.FItemList(i).Fcouponbuyprice=0 or isNull(oExhibition.FItemList(i).Fcouponbuyprice) then
										Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fbuycash,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
									else
										Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fcouponbuyprice,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
									end if
								Case "2"
									if oExhibition.FItemList(i).Fcouponbuyprice=0 or isNull(oExhibition.FItemList(i).Fcouponbuyprice) then
										Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fbuycash,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
									else
										Response.Write "<br><font color=#5080F0>" & fnPercent(oExhibition.FItemList(i).Fcouponbuyprice,oExhibition.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
									end if
							end Select
						end if
						%>
					</td><%'마진%>
					<td><%=fnColor(oExhibition.FItemList(i).Fmwdiv,"mw")%><br/>
						<%
							If oExhibition.FItemList(i).Fdeliverytype = "1" Then
								response.write "텐배"
							ElseIf oExhibition.FItemList(i).Fdeliverytype = "2" Then
								response.write "무료"
							ElseIf oExhibition.FItemList(i).Fdeliverytype = "4" Then
								response.write "텐무"
							ElseIf oExhibition.FItemList(i).Fdeliverytype = "9" Then
								response.write "조건"
							ElseIf oExhibition.FItemList(i).Fdeliverytype = "7" Then
								response.write "착불"
							End If
						%>
					</td>
					<td><input type="button" value="삭제" onclick="fnDelItem(<%= oExhibition.FItemList(i).Fidx%>);"/></td>
					<td><input type="button" value="통계보기" onclick="popreport('<%= oExhibition.FItemList(i).Fmastercode %>','<%= oExhibition.FItemList(i).Fdetailcode %>','<%= oExhibition.FItemList(i).Fidx %>');" class="button"></td>
					<%'계약구분%>	
				</tr>
				<% Next %>
			<% else %>
				<tr bgcolor="#FFFFFF">
					<td colspan="12" class="page_link">[검색결과가 없습니다.]</td>
				</tr>
			<% End IF %>
				<tr bgcolor="#FFFFFF">
					<td colspan="12" align="center">
					<!-- 페이지 시작 -->
						<a href="?page=1&isusingbox=<%=isusing%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&menupos=<%=poscode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
						<% if oExhibition.HasPreScroll then %>
							<span class="list_link"><a href="?page=<%= oExhibition.StartScrollPage-1 %>&isusingbox=<%=isusing%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&menupos=<%=poscode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
						<% else %>
						&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
						<% end if %>
						<% for i = 0 + oExhibition.StartScrollPage to oExhibition.StartScrollPage + oExhibition.FScrollCount - 1 %>
							<% if (i > oExhibition.FTotalpage) then Exit for %>
							<% if CStr(i) = CStr(oExhibition.FCurrPage) then %>
							<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
							<% else %>
							<a href="?page=<%= i %>&isusingbox=<%=isusing%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&menupos=<%=poscode%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
							<% end if %>
						<% next %>
						<% if oExhibition.HasNextScroll then %>
							<span class="list_link"><a href="?page=<%= i %>&isusingbox=<%=isusing%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&menupos=<%=poscode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
						<% else %>
						&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
						<% end if %>
						<a href="?page=<%= oExhibition.FTotalpage %>&isusingbox=<%=isusing%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&menupos=<%=poscode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
					<!-- 페이지 끝 -->
					</td>
				</tr>
			</form>
			</table>
			<form name="itemdel" method="post" action="/admin/exhibitionitems/lib/exhibition_proc.asp">
			<input type="hidden" name="eidx" value=""/>
			<input type="hidden" name="mode" value="delitem" />
			<input type="hidden" name="poscode" value="<%=poscode%>"/>
			<input type="hidden" name="page" value="<%=page%>"/>
			</form>
			<!-- 리스트 끝 -->
		</div>
	</div>
</span>
<% Set oExhibition = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->