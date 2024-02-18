<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_cs_baljucls.asp"-->

<%

'// ===========================================================================
dim ix,iy
dim searchType, searchValue

searchType      = requestCheckVar(request("searchType"), 32)
searchValue     = requestCheckVar(request("searchValue"), 32)


'// ===========================================================================
dim page
dim ojumun

page = requestCheckVar(request("page"), 32)
if (page="") then page=1

set ojumun = new CCSJumunMaster

ojumun.FPageSize = 200
ojumun.FScrollCount = 10
ojumun.FCurrPage = page
ojumun.FRectDesignerID = session("ssBctID")
ojumun.FRectDivcd = ""
ojumun.FRectSearchType  = SearchType
ojumun.FRectSearchValue = SearchValue

ojumun.DesignerCS_BaljuMiBeasongList

%>
<script language='javascript'>
function chksubmit(){
    var frm = document.frm;

    if ((frm.searchType.value.length>0)&&(frm.searchValue.value.length<1)){
        alert('검색 내용을 입력하세요.');
        frm.searchValue.focus();
        return;
    }

    if ((frm.searchType.value=="orderserial")||(frm.searchType.value=="itemid")){
        if (!IsDigit(frm.searchValue.value)){
            alert('검색 내용은 숫자만 가능합니다.');
            frm.searchValue.focus();
            return;
        }
    }

    frm.submit();
}

function ShowOrderInfo(frm,orderserial){
	var props = "width=600, height=600, location=no, status=yes, resizable=no, scrollbars=yes";
	window.open("about:blank", "orderdetail", props);
    frm.target = "orderdetail";
    frm.orderserial.value = orderserial;
    frm.action="/designer/common/viewordermaster.asp";
	frm.submit();
}


function ViewItem(itemid){
    var popwin = window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"sample");
    popwin.focus();
}



function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}


function switchCheckBox(comp){
    var frm = comp.form;

	if(frm.chkidx.length>1){
		for(i=0;i<frm.chkidx.length;i++){
			frm.chkidx[i].checked = comp.checked;
			AnCheckClick(frm.chkidx[i]);
		}
	}else{
		frm.chkidx.checked = comp.checked;
		AnCheckClick(frm.chkidx);
	}
}

function BaljuReprint(){
    var frm = document.frmbalju;
	var pass = false;

    if(!frm.chkidx.length){
    	pass = frm.chkidx.checked;
    }else{
        for (var i=0;i<frm.chkidx.length;i++){
    	    pass = (pass||frm.chkidx[i].checked);
    	}
    }

	if (!pass) {
		alert("재출력할 내역을 선택하세요.");
		return;
	}else{
	    var popwin = window.open("about:blank","PopBaljuList","width=800,scrollbars=yes,resizable");
	    frm.target = "PopBaljuList";
	    frm.isall.value = "";
 		frm.action = "upchecs_reselectbaljulist.asp";
		frm.submit();
	}
}

function BaljuReprintAll(){
    var frm = document.frmbalju;

    if (confirm('미출고 내역 전체 발주서를 재출력 하시겠습니까?')){
        var popwin = window.open("about:blank","PopBaljuList","width=800,scrollbars=yes,resizable");
	    frm.target = "PopBaljuList";
	    frm.isall.value = "on";
 		frm.action = "upchecs_reselectbaljulist.asp";
		frm.submit();
    }
}

function trim(theString){
	var resultString = theString;

	if (theString.indexOf(" ") == 0) {
        resultString = theString.substring(1, theString.length);
	}

	if (resultString.lastIndexOf(" ") == resultString.length) {
        resultString = resultString.substring(1,theString.length-1);
	}

	return resultString
}

function ShowDateBox(comp){
    var frm = comp.form;
    var iid = comp.id;
    var idiv = eval("document.all.divipgodate" + iid);

    if ((comp.value=="03")||(comp.value=="02")){
        idiv.style.display = "inline";
    }else{
        idiv.style.display = "none";
    };

    if (!frm.chkidx.length){
        if (comp.id=="0"){
            frm.chkidx.checked = true;
            AnCheckClick(frm.chkidx);
        }
    }else{
        frm.chkidx[iid].checked = true;
        AnCheckClick(frm.chkidx[iid]);
    }
}

//
function popMisendInput(iidx){
    var popwin = window.open('popMisendInput.asp?idx=' + iidx,'popMisendInput','width=440,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//사용안함.
function MisendInput(){
    var frm = document.frmbalju;
	var pass = false;
    var today= new Date();
    var inputdate;
    var arrchkval = '';

    if(!frm.chkidx.length){
    	pass = frm.chkidx.checked;

    	if (frm.chkidx.checked){
	        if (frm.MisendReason.value==""){
	            alert('미출고 사유를 선택 하세요.');
	            frm.MisendReason.focus();
	            return;
	        }

	        //출고지연,주문제작
	        if ((frm.MisendReason.value=="03")||(frm.MisendReason.value=="02")){
	            var ipgodate = eval("frm.ipgodate0");
	            if (ipgodate.value.length!=10){
    	            alert('출고 예정일을 입력하세요.(YYYY-MM-DD)');
    	            ipgodate.focus();
    	            return;
    	        }

                inputdate = new Date(ipgodate.value.substr(0,4),ipgodate.value.substr(5,2)*1-1,ipgodate.value.substr(8,2));
    	        if (today>inputdate){
    	            alert('출고 예정일은 오늘 이후날짜로 설정이 가능합니다.');
    	            ipgodate.focus();
    	            return;
    	        }


	        }

	        arrchkval = "1";

	    }
    }else{
        for (var i=0;i<frm.chkidx.length;i++){
    	    pass = (pass||frm.chkidx[i].checked);

    	    if (frm.chkidx[i].checked){
    	        //if (!frm.MisendReason[i]){
    	        //    alert('D+1일 부터 미출고 입력 가능합니다.');
    	        //    frm.chkidx[i].focus();
    	        //    return;
    	        //}

    	        if (frm.MisendReason[i].value==""){
    	            alert('미출고 사유를 선택 하세요.');
    	            frm.MisendReason[i].focus();
    	            return;
    	        }

    	        //출고지연, 주문제작
    	        if ((frm.MisendReason[i].value=="03")||(frm.MisendReason[i].value=="02")){
    	            var ipgodate = eval("frm.ipgodate" + i);
    	            if (ipgodate.value.length!=10){
        	            alert('출고 예정일을 입력하세요.(YYYY-MM-DD)');
        	            ipgodate.focus();
        	            return;
        	        }

        	        inputdate = new Date(ipgodate.value.substr(0,4),ipgodate.value.substr(5,2)*1-1,ipgodate.value.substr(8,2));
        	        if (today>inputdate){
        	            alert('출고 예정일은 오늘 이후날짜로 설정이 가능합니다.');
        	            ipgodate.focus();
        	            return;
        	        }
    	        }

    	        if (arrchkval==""){
        	        arrchkval = (i*1+1);
        	    }else{
        	        arrchkval = arrchkval + "," + (i*1+1);
        	    }

    	    }

    	}
    }

	if (!pass) {
		alert("미출고 사유를 저장할 내역을 선택하세요.");
		return;
	}


	if (confirm('미출고 사유를 저장 하시겠습니까?')){
	    frm.target = "";
	    frm.ArrChkVal.value = arrchkval;
	    frm.action = "upchebeasong_Process.asp";
	    frm.mode.value   = "misendInput";
	    frm.submit();
	}
}

function ViewCSDetail(detailidx) {
    var popwin = window.open("/designer/jumunmaster/upchecsdetail.asp?idx=" + detailidx,"ViewCSDetail");
    popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="" onsubmit="chksubmit(); return false">
		<input type="hidden" name="page" value="1">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
			<td align="left" bgcolor="#FFFFFF">
				<select class="select" name="searchType" >
					<option value="">검색조건</option>
					<option value="orderserial" <%= ChkIIF(searchType="orderserial","selected","") %> >주문번호</option>
					<option value="itemid" <%= ChkIIF(searchType="itemid","selected","") %> >상품코드</option>
					<option value="buyname" <%= ChkIIF(searchType="buyname","selected","") %> >구매자</option>
					<option value="reqname" <%= ChkIIF(searchType="reqname","selected","") %> >수령인</option>
				</select>
				<input type="text" class="text" name="searchValue" value="<%= searchValue %>" size="13" maxlength="11">
				<br>
			</td>

			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
				<input type="button" class="button_s" value="검색" onClick="javascript:chksubmit();">
			</td>
		</tr>
	</form>
</table>

<p>

	<!-- 액션 시작 -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
		<tr align="center">
			<td align="left">
        		<input type="button" class="button" value="선택내역 CS출고발주서 재출력" onclick="javascript:BaljuReprint()">
				&nbsp;
        		<input type="button" class="button" value="미출고전체 CS출고발주서 재출력" onclick="javascript:BaljuReprintAll()">
			</td>
			<td align="right">
			</td>
		</tr>
	</table>
	<!-- 액션 끝 -->

	<p>

		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frmbalju" method="post" >
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="isall" value="">
				<input type="hidden" name="ArrChkVal" value="">
				<tr bgcolor="FFFFFF">
					<td height="25" colspan="15">
						검색결과 : <b><% = ojumun.FTotalCount %></b>
						&nbsp;
						페이지 : <b><%= page %> / <%= ojumun.FTotalpage %></b>
					</td>
				</tr>
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td width="30"><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
					<td width="80">원주문번호</td>
					<td width="55">고객명</td>
					<td width="55">수령인</td>

					<td>접수구분</td>
					<td>제목</td>
					<td>접수사유</td>

					<td width="50">상품코드</td>
					<td>상품명<font color="blue">&nbsp;[옵션]</font></td>
					<td width="30">수량</td>
					<td width="65">등록일</td>
				</tr>
				<% if ojumun.FresultCount<1 then %>
				<tr bgcolor="#FFFFFF">
					<td colspan="15" align="center">[검색결과가 없습니다.]</td>
				</tr>
				<% else %>
				<% for ix=0 to ojumun.FresultCount-1 %>
				<input type="hidden" name="detailidx" value="<%= ojumun.FMasterItemList(ix).Fidx %>">
				<tr align="center" class="a" bgcolor="#FFFFFF">
					<td><input type="checkbox" name="chkidx" value="<%= ojumun.FMasterItemList(ix).Fidx %>" onClick="AnCheckClick(this);"></td>
					<td>
						<%= ojumun.FMasterItemList(ix).FOrgOrderSerial %>
    					<% if (ojumun.FMasterItemList(ix).Forderserial <> ojumun.FMasterItemList(ix).Forgorderserial) then %>
    					+
    					<% end if %>
					</td>
					<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
					<td><%= ojumun.FMasterItemList(ix).FReqname %></td>

					<td><%= ojumun.FMasterItemList(ix).Fdivcdname %></td>
					<td><a href="javascript:ViewCSDetail(<%= ojumun.FMasterItemList(ix).Fmasteridx %>)"><%= ojumun.FMasterItemList(ix).Ftitle %></a></td>
					<td><%= ojumun.FMasterItemList(ix).Fgubun01name %> >> <%= ojumun.FMasterItemList(ix).Fgubun02name %></td>

					<td><%= ojumun.FMasterItemList(ix).FitemID %></td>
					<td align="left">
						<a href="javascript:ViewItem(<% =ojumun.FMasterItemList(ix).FItemid  %>)"><%= ojumun.FMasterItemList(ix).FItemname %></a>
						<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
						<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
						<% end if %>
					</td>
					<td><%= ojumun.FMasterItemList(ix).FItemcnt %></td>
					<td><acronym title="<%= ojumun.FMasterItemList(ix).Fregdate %>"><%= left(ojumun.FMasterItemList(ix).Fregdate,10) %></acronym></td>
				</tr>
				<% next %>
				<% end if %>
			</form>

			<tr height="25" bgcolor="FFFFFF">
				<td colspan="15" align="center">
					<% if ojumun.HasPreScroll then %>
					<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
					<% else %>
					[pre]
					<% end if %>
					<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
					<% if ix>ojumun.FTotalpage then Exit for %>
					<% if CStr(page)=CStr(ix) then %>
					<font color="red">[<%= ix %>]</font>
					<% else %>
					<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
					<% end if %>
					<% next %>

					<% if ojumun.HasNextScroll then %>
					<a href="javascript:NextPage('<%= ix %>')">[next]</a>
					<% else %>
					[next]
					<% end if %>
				</td>
			</tr>

		</table>

		<p>

			<%
			set ojumun = Nothing
			%>
			<form name="frmshow" method="post">
				<input type="hidden" name="orderserial" value="">

			</form>
			<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
			<!-- #include virtual="/lib/db/dbclose.asp" -->
