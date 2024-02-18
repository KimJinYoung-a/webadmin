<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결제요청서 리스트
' History : 2011.10.13 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<!-- #include virtual="/lib/classes/approval/innerOrdercls.asp"-->
<%

''내부거래관리 :
''
''
''
'' - 내부거래에 있어 부가세는 전부 제외된다.
''
'' - 내부거래물류입고는 모두 상품매입이다.(수수료매입의 사은품은 매입가가 0원이어야 한다.)
''
''
''=======================================================================
'' - 물류출고 : 온라인매입상품
''=======================================================================
''
''  - 물류->가맹점 : 온라인매출-오프라인본사매입(상품매입가)
''
''  - 물류->내부부서(직영 or 아이띵소 등) : 온라인매출-오프라인본사매입(상품매입가), 오프라인본사매출-내부부서매입(매장출고가)
''
''=======================================================================
'' - 물류출고 : 온라인매입이외 상품
''=======================================================================
''
''  - 물류->가맹점 : (내부거래 X)
''
''  - 물류->내부부서(직영 or 아이띵소 등) : 오프라인본사매출-내부부서매입(상품매입가)
''
''=======================================================================
'' - 업체개별배송 : 매장매입(출고시)
''=======================================================================
''
''   - 업체->가맹점 : (내부거래 X)
''
''   - 업체->내부부서(직영 or 아이띵소 등) : 오프라인본사매출-내부부서매입(매장출고가)
''
''=======================================================================
'' - 업체개별배송 : 업체위탁상품(판매시)
''=======================================================================
''
''   - 업체->가맹점 : (내부거래 X)
''
''   - 업체->내부부서(직영 or 아이띵소 등) : 오프라인본사매출-내부부서매입(매장출고가)
''
''=======================================================================
'' - 위탁판매 : 위탁상품(판매시)
''=======================================================================
''
''   - 물류->가맹점 : (내부거래 X)
''
''   - 물류->내부부서(직영 or 아이띵소 등) : 오프라인본사매출-내부부서매입(매장출고가)
''
''=======================================================================
'' - 온라인정산(내부부서 매입처)
''=======================================================================
''
''  - 내부부서(아이띵소 등)->온라인정산 : 내부부서매출-온라인매입(상품매입가)
''
''=======================================================================
'' - 오프라인정산(내부부서 매입처)
''=======================================================================
''
''  - 내부부서(아이띵소 등)->오프라인정산 : 내부부서매출-오프라인매입(상품매입가)

dim i, page, research
dim yyyy1,mm1,yyyy2,mm2
dim bizsection_cd
dim intLoop, tmpdate

dim groupingyn

page = requestCheckvar(Request("page"),32)
research = requestCheckvar(Request("research"),32)
groupingyn = requestCheckvar(Request("groupingyn"),32)

if (page = "") then
	page = 1
end if

yyyy1 = requestCheckvar(Request("yyyy1"),32)
mm1 = requestCheckvar(Request("mm1"),32)
yyyy2 = requestCheckvar(Request("yyyy2"),32)
mm2 = requestCheckvar(Request("mm2"),32)

bizsection_cd = requestCheckvar(Request("bizsection_cd"),32)

if yyyy1="" then
	tmpdate = CStr(Now)

	tmpdate = DateAdd("m", -1, tmpdate)

	yyyy1 = Left(tmpdate, 4)
	mm1 = Mid(tmpdate, 6, 2)

	yyyy2 = Left(tmpdate, 4)
	mm2 = Mid(tmpdate, 6, 2)
end if

'==============================================================================
dim oinnerorder
set oinnerorder = New CInnerOrder

oinnerorder.FCurrPage = page
oinnerorder.FPageSize = 100

oinnerorder.FRectStartYYYYMMDD = DateSerial(yyyy1, mm1, 1)

tmpdate = DateSerial(yyyy2, mm2, 1)
tmpdate = DateAdd("m", 1, tmpdate)
oinnerorder.FRectEndYYYYMMDD = tmpdate		'// 다음달 1일 이전까지

oinnerorder.FRectBizSection_CD = bizsection_cd

'// ocsmemo.FRectPhoneNumber = phonenumber

if (groupingyn = "Y") then
	oinnerorder.GetInnerOrderSummaryList
else
	oinnerorder.GetInnerOrderList
end if

%>




 <script language="javascript" src="/admin/approval/eapp/eapp.js"></script>
<script language="javascript">

function popRegInnerOrderByMonth() {
	var winR = window.open("popRegInnerOrderByMonth.asp","popRegInnerOrderByMonth","width=1000, height=600, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popRegInnerOrderMannualy() {
	var winR = window.open("popRegInnerOrderMannualy.asp","popRegInnerOrderMannualy","width=800, height=600, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popViewInnerOrder(idx) {
	var winR = window.open("popRegInnerOrderMannualy.asp?idx=" + idx,"popViewInnerOrder","width=800, height=600, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popViewInnerOrderDetail(masteridx) {
	if (masteridx < 0) {
		alert("합계보기 상태에서는 상세내역을 볼 수 없습니다.");
		return;
	}

	var winR = window.open("popViewInnerOrderDetail.asp?idx="+masteridx,"popViewInnerOrderDetail","width=600, height=700, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popViewOnlineInnerOrderDetail(masteridx) {
	if (masteridx < 0) {
		alert("합계보기 상태에서는 상세내역을 볼 수 없습니다.");
		return;
	}

	var winR = window.open("popViewOnlineInnerOrderDetail.asp?idx="+masteridx,"popViewOnlineInnerOrderDetail","width=1000, height=700, resizable=yes, scrollbars=yes");
	winR.focus();
}

function popViewInnerOrderDetailNew(masteridx) {
	if (masteridx < 0) {
		alert("합계보기 상태에서는 상세내역을 볼 수 없습니다.");
		return;
	}

	var winR = window.open("popViewInnerOrderDetailNew.asp?idx="+masteridx,"popViewInnerOrderDetailNew","width=1200, height=700, resizable=yes, scrollbars=yes");
	winR.focus();
}

function jsSearch(){
 document.frm.submit();
}

	// 페이지 이동
function jsGoPage(iCP)
{
	document.frm.iCP.value=iCP;
	document.frm.submit();
}

//수지항목 불러오기
function jsGetARAP(){
		var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp","popARAP","width=600,height=600,resizable=yes, scrollbars=yes");
		winARAP.focus();
}

function jsReSetARAP(){
		document.frm.iaidx.value = 0;
		document.frm.selarap.value = "";
}

//선택 수지항목 가져오기
function jsSetARAP(dAC, sANM,sACC,sACCNM){
	document.frm.iaidx.value = dAC;
	document.frm.selarap.value = sANM;
}

function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function checkThis(comp) {
	var frm = comp.form;

    AnCheckClick(comp);

    if (comp.checked != true) {
    	frm.chkAll.checked = false;
    }
}

function jsLinkERP(frm){
    var ischecked =false;

    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    ischecked = e.checked;
			if (ischecked) break;
		}
	}

	if (!ischecked){
	    alert('선택 내역이 없습니다.');
	    return;
	}

	if (confirm('선택 내역을 ERP로 전송하시겠습니까?')){
	    frm.LTp.value="A";
	    frm.submit();
	}
}

function jsReceiveERP(frm){
    if (confirm('결제 결과를 수신 하시겠습니까?')){
	    frm.LTp.value="R";
	    frm.submit();
	}
}

function popConfirmPayrequest(iridx,pidx){
    var iURI = '/admin/approval/eapp/confirmpayrequest.asp?iridx='+iridx+'&ipridx='+pidx+'&ias=1'; //ias 확인..
    var popwin = window.open(iURI,'popConfirmPayrequest','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popModPayDoc(iridx,pidx){
	 var iURI = '/admin/approval/eapp/modeappPayDoc.asp?iridx='+iridx+'&ipridx='+pidx ; //ias 확인..
    var popwin = window.open(iURI,'popConfirmPayrequest','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function jsDelSelected(frm) {

	var checkeditemfound = false;
	for (var i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];

		if (e.type == "checkbox") {
			if (e.name == "chk") {
				if (e.checked == true) {
					checkeditemfound = true;
					break;
				}
			}
		}
	}

	if (checkeditemfound == false) {
		alert("선택된 내역이 없습니다.");
		return;
	}

    if (confirm('선택 내역을 삭제하시겠습니까?') == true) {
	    frm.mode.value="delselectedarr";
	    frm.submit();
	}
}

</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="innerOrderList.asp">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="research" value="on">
			<input type="hidden" name="page" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="30" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					거래기간
					: <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
					&nbsp;&nbsp;
					사업부문:
					<%
					Dim clsBS, arrBizList
					Set clsBS = new CBizSection
                    	clsBS.FUSE_YN = "Y"
                    	clsBS.FOnlySub = "Y"
                    	arrBizList = clsBS.fnGetBizSectionList
                    Set clsBS = nothing
                    %>
                    <select name="bizsection_cd">
                    <option value="">--선택--</option>
                    <% For intLoop = 0 To UBound(arrBizList,2)	%>
                		<option value="<%=arrBizList(0,intLoop)%>" <%IF Cstr(bizsection_cd) = Cstr(arrBizList(0,intLoop)) THEN%> selected <%END IF%>><%=arrBizList(1,intLoop)%></option>
                	<% Next %>
                    </select>
				</td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
				</td>
			</tr>
			<tr align="center" bgcolor="#FFFFFF" >
				<td align="left" height="30">
				<input type=checkbox name=groupingyn value="Y" <% if (groupingyn = "Y") then %> checked<% end if %>> 합계보기
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<tr>
	<td>
		* 내부부서가 추가되는 경우<br><br>

		- 1. 사업부문 추가 : 전자결재>>자금관리부서<br>
		- 2. 내부부서 추가 : [경영]재무회계>>내부부서관리<br>
		- 3. 기본 매출부서 지정 : 브랜드리스트 > 기본 매출부서<br><br>

	    <table width="100%" cellspacing="0" cellpadding="0">
	    <tr>
	        <td align="left">
	        	<input type="button" class="button" value=" 내부거래 수기등록 " onClick="popRegInnerOrderMannualy();" disabled>
	        	<input type="button" class="button" value="내부거래 일괄생성" onClick="popRegInnerOrderByMonth();">
	        </td>
	        <td align="right">
	        	<input type="button" class="button" value="선택내역 [삭제]" onClick="jsDelSelected(frmAct);">
	        </td>
	    </tr>
	    </table>
	</td>
</tr>

<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a"   border="0">
		<Form name="frmAct" method="post" action="innerorder_process.asp">
		<input type="hidden" name="mode" value="">
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
				    <td width="20"><input type="checkbox" name="chkAll" onClick="CkeckAll(this)" <% if (groupingyn = "Y") then %>disabled<% end if %>></td>
					<td>IDX</td>
					<td width="80">거래일자</td>
					<td width="150">구분</td>
					<td width="80">계정</td>
					<td align=left>플러스(+)부서</td>
					<td align=left>마이너스(-)부서</td>
					<td>공급가</td>
					<td>부가세</td>
					<td>합계</td>
					<td>상세내역</td>
					<td>작성자</td>
					<td>작성일</td>
					<!--
					<td>ERP<br>연동상태</td>
					-->
				</tr>
				<%IF oinnerorder.FResultCount > 0 THEN %>
				<% for i = 0 to (oinnerorder.FResultCount - 1) %>
				<tr bgcolor="#FFFFFF" align="center">
				    <td><input type="checkbox" name="chk" value="<%= oinnerorder.FItemList(i).Fidx %>" onClick="checkThis(this)" <% if (groupingyn = "Y") then %>disabled<% end if %>></td>
					<td><a href="javascript:popViewInnerOrder(<%= oinnerorder.FItemList(i).Fidx %>);"><%= oinnerorder.FItemList(i).Fidx %></a></td>
					<td><a href="javascript:popViewInnerOrder(<%= oinnerorder.FItemList(i).Fidx %>);"><%= oinnerorder.FItemList(i).FappDate %></a></td>
					<td><font color="<%= oinnerorder.FItemList(i).GetDivcdColor %>"><%= oinnerorder.FItemList(i).GetDivcdName %></font></td>

					<td><%= oinnerorder.FItemList(i).Facc_nm %></td>

					<td align=left><%= oinnerorder.FItemList(i).FSELLBIZSECTION_NM %></td>
					<td align=left><%= oinnerorder.FItemList(i).FBUYBIZSECTION_NM %></td>

					<td align=right>
						<a href="javascript:popViewInnerOrderDetailNew(<%= oinnerorder.FItemList(i).Fidx %>)">
						<%= FormatNumber(oinnerorder.FItemList(i).FsupplySum, 0) %>
						</a>
					</td>
					<td align=right>
						<a href="javascript:popViewInnerOrderDetailNew(<%= oinnerorder.FItemList(i).Fidx %>)">
						<%= FormatNumber(oinnerorder.FItemList(i).FtaxSum, 0) %>
						</a>
					</td>
					<td align=right>
						<a href="javascript:popViewInnerOrderDetailNew(<%= oinnerorder.FItemList(i).Fidx %>)">
						<%= FormatNumber(oinnerorder.FItemList(i).FtotalSum, 0) %>
						</a>
					</td>

					<td><input type="button" class="button" value="조회" onClick="popViewInnerOrderDetailNew(<%= oinnerorder.FItemList(i).Fidx %>)"></td>
					<td><%= oinnerorder.FItemList(i).Freguserid %></td>
					<td><%= Left(oinnerorder.FItemList(i).Fregdate, 10) %></td>
					<!--
					<td></td>
					-->
				</tr>
				<%
					Next
				%>
				<%
					ELSE
				%>
				<tr bgcolor="#FFFFFF">
					<td colspan="13" align="center">등록된 내역이 없습니다.</td>
				</tr>
				<%END IF%>
				</table>
			</td>
		</tr>
        </form>
	    <tr align="center" bgcolor="#FFFFFF">
	        <td colspan="13">
	            <% if oinnerorder.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oinnerorder.StartScrollPage-1 %>')">[pre]</a>
	    		<% else %>
	    			[pre]
	    		<% end if %>

	    		<% for i=0 + oinnerorder.StartScrollPage to oinnerorder.FScrollCount + oinnerorder.StartScrollPage - 1 %>
	    			<% if i>oinnerorder.FTotalpage then Exit for %>
	    			<% if CStr(page)=CStr(i) then %>
	    			<font color="red">[<%= i %>]</font>
	    			<% else %>
	    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>

	    		<% if oinnerorder.HasNextScroll then %>
	    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
	        </td>
	    </tr>
		</table>
	</td>
</tr>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
