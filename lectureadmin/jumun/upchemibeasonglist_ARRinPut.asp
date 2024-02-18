<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_baljucls.asp"-->

<%


dim page
dim searchType, searchValue, MisendReason
dim ojumun

page = RequestCheckvar(request("page"),10)
searchType = RequestCheckvar(request("searchType"),16)
searchValue = RequestCheckvar(request("searchValue"),16)
MisendReason = RequestCheckvar(request("MisendReason"),10)

if (page="") then page=1

set ojumun = new CJumunMaster

ojumun.FPageSize = 50
ojumun.FCurrPage = page
ojumun.FRectSearchType  = SearchType
ojumun.FRectSearchValue = SearchValue
if (MisendReason="") then
    ojumun.FRectMisendReason = "AA"
else
    ojumun.FRectMisendReason = MisendReason
end if
ojumun.FRectDesignerID = session("ssBctID")
ojumun.DesignerDateBaljuinputlist

dim ix,iy

%>
<script language='javascript'>

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
 		frm.action = "reselectbaljulist.asp";
		frm.submit();
	}
}

function BaljuReprintAll(){
    var frm = document.frmbalju;

    if (confirm('미출고 내역 전체 발주서를 재출력 하시겠습니까?')){
        var popwin = window.open("about:blank","PopBaljuList","width=800,scrollbars=yes,resizable");
	    frm.target = "PopBaljuList";
	    frm.isall.value = "on";
 		frm.action = "reselectbaljulist.asp";
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

</script>



<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
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
			&nbsp;
			사유입력여부 :
			<select class="select" name="MisendReason">
				<option value="" >전체</option>
				<option value="03" <%= ChkIIF(MisendReason="03","selected","") %> >출고지연</option>
				<option value="05" <%= ChkIIF(MisendReason="05","selected","") %> >품절출고불가</option>
				<option>주문제작</option>
			</select>
			<br>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr align="center">
		<td align="left">
        	<input type="button" class="button" value="선택내역 발주서 재출력" onclick="javascript:BaljuReprint()">
			&nbsp;
        	<input type="button" class="button" value="미출고전체 발주서 재출력" onclick="javascript:BaljuReprintAll()">
        </td>
        <td align="right">
        	<input type="button" class="button" value="선택주문 미출고사유 저장" onclick="javascript:MisendInput()">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>
출고지연과 주문제작의 경우, 고객에게 SMS 및 안내메일 발송<br>
품절출고불가의 경우, 고객센터에서 처리

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
		<td width="70">주문번호</td>
		<td width="50">주문자</td>
		<td width="50">수령인</td>
		<td width="50">상품코드</td>
		<td>상품명<font color="blue">&nbsp;[옵션]</font></td>
		<td width="30">수량</td>
		<td width="65">입금확인일</td>
		<td width="65">출고기준일<!-- 주문통보일 --></td>
		<td width="65">주문확인일</td>
		<td width="40">경과일</td>
		<td width="100">미출고사유</td>
		<td width="120">출고예정일</td>
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
		<td height="25"><a href="javascript:ShowOrderInfo(frmshow,'<%= ojumun.FMasterItemList(ix).Forderserial %>')"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqname %></td>
		<td><%= ojumun.FMasterItemList(ix).FitemID %></td>
		<td align="left">
			<a href="javascript:ViewItem(<% =ojumun.FMasterItemList(ix).FItemid  %>)"><%= ojumun.FMasterItemList(ix).FItemname %></a>
			<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
			<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
			<% end if %>
		</td>
		<td><%= ojumun.FMasterItemList(ix).FItemcnt %></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).Fipkumdate %>"><%= left(ojumun.FMasterItemList(ix).Fipkumdate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).Fbaljudate %>"><%= left(ojumun.FMasterItemList(ix).Fbaljudate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).Fupcheconfirmdate %>"><%= left(ojumun.FMasterItemList(ix).Fupcheconfirmdate,10) %></acronym></td>
		<td>
		    <% if IsNULL(ojumun.FMasterItemList(ix).Fbaljudate) then %>
	        D+0
	        <% elseif datediff("d",(left(ojumun.FMasterItemList(ix).Fbaljudate,10)) , (left(now,10)) )>2 then %>
	        <font color="red"><b>D+<%= datediff("d",(left(ojumun.FMasterItemList(ix).Fbaljudate,10)) , (left(now,10)) ) %></b></font>
	        <% else %>
	        D+<%= datediff("d",(left(ojumun.FMasterItemList(ix).Fbaljudate,10)) , (left(now,10)) ) %>
	        <% end if %>
	    </td>
		<td>
			<% if TRUE or datediff("d",(left(ojumun.FMasterItemList(ix).Fbaljudate,10)) , (left(now,10)) )>1 then %>
			<select name="MisendReason" id="<%= ix %>" class="select" onChange="ShowDateBox(this);">
				<option value="">---------</option>
				<option value="03" <%= ChkIIF(ojumun.FMasterItemList(ix).FMisendReason="03","selected","") %> >출고지연</option>
				<option value="05" <%= ChkIIF(ojumun.FMasterItemList(ix).FMisendReason="05","selected","") %> >품절출고불가</option>
				<option value="02" <%= ChkIIF(ojumun.FMasterItemList(ix).FMisendReason="02","selected","") %> >주문제작</option>
				<!-- 텐바이텐배송 미출고사유와 통합했으면 합니다. -->
			</select>
			<% end if %>
		</td>
		<td>
		<div id="divipgodate<%= ix %>" name="divipgodate<%= ix %>" <%= ChkIIF(ojumun.FMasterItemList(ix).FMisendReason="03" or ojumun.FMasterItemList(ix).FMisendReason="02","style='display:inline'","style='display:none'") %>>
		    <input class="text" type="text" name="ipgodate<%= ix %>" value="<%= ojumun.FMasterItemList(ix).FMisendIpgodate %>" size="10" maxlength="10">
		    <a href="javascript:calendarOpen(frmbalju.ipgodate<%= ix %>);"><img src="/images/calicon.gif" border="0" align="top" height=20></a>
		</div>
	    </td>
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