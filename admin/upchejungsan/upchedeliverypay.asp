<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/jungsan/deliverpayjungsancls.asp" -->
<%
dim yyyy1,mm1
dim research, page, makerid
dim orderserial, userid, cknotreg, cksiteMM
yyyy1       = requestCheckvar(request("yyyy1"),4)
mm1         = requestCheckvar(request("mm1"),2)
research    = requestCheckvar(request("research"),10)
page        = requestCheckvar(request("page"),10)
makerid     = requestCheckvar(request("makerid"),32)
orderserial = requestCheckvar(request("orderserial"),11)
userid      = requestCheckvar(request("userid"),32)
cknotreg    = requestCheckvar(request("cknotreg"),10)
cksiteMM    = requestCheckvar(request("cksiteMM"),10)

if page="" then page=1

dim stdt

if (yyyy1="") then
	stdt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(stdt),4)
	mm1 = Mid(CStr(stdt),6,2)
end if

dim oDeliverPay
set oDeliverPay = new CUpcheDeliverPayJungsan
oDeliverPay.FCurrPage       = page
oDeliverPay.FPageSize       = 100
oDeliverPay.FRectYYYYMM     = yyyy1 + "-" + mm1
oDeliverPay.FRectMakerid    = makerid
oDeliverPay.FRectOrderserial = orderserial
oDeliverPay.FRectUserID     = userid
oDeliverPay.FRectOnlyNotReged = cknotreg
oDeliverPay.FRectcksiteMM = cksiteMM
oDeliverPay.GetMonthlyDeliverPayJungsanList

dim i
dim jungsanyyyy, jungsanmm
stdt = dateserial(year(Now),month(now)-1,1)
jungsanyyyy = Left(CStr(stdt),4)
jungsanmm = Mid(CStr(stdt),6,2)
%>
<script language='javascript'>
function goPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function AddBesongJungsan(asid){
    frmSubmit.asid.value = asid
    //var comp = eval("frmList.refunddeliverypay_" + asid);
    //frmSubmit.refunddeliverypay.value = comp.value
    //if (!IsDigit(frmSubmit.refunddeliverypay.value)){
    //    alert('숫자만 가능합니다.');
    //    comp.focus();
    //    return;
    //}
    if (confirm('등록하시겠습니까?')){
        frmSubmit.submit();
    }
}

function EditBesongJungsan(idetailid){
    var popwin = window.open('popbeasongpayedit.asp?detailid=' + idetailid,'EditBesongJungsan','width=1000,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddEtcJungsan(asid){
    //popetclistadd.asp
}

function AddEtcBesongJungsan(yyyy1,mm1){
    var popwin= window.open('popbeasongpayadd.asp?yyyy1=' + yyyy1 + '&mm1=' + mm1,'AddEtcBesongJungsan','width=1000,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function ckAll(comp){
    var frm=document.frmList;

	if(frm.ckidx.length>1)
	{
		for(i=0;i<frm.ckidx.length;i++)
		{
			if(comp.checked)
				frm.ckidx[i].checked=true;
			else
				frm.ckidx[i].checked=false;
		}
	}
	else
	{
		if(comp.checked)
			frm.ckidx.checked=true;
		else
			frm.ckidx.checked=false;
	}
}

function chkSubmit(){
    var chk = 0;
    var frm=document.frmList;
    var refunddeliverypayArr = "";
    var comp ;

	if(frm.ckidx.length>1){
		for(i=0;i<frm.ckidx.length;i++){
			if(frm.ckidx[i].checked){
				chk++;
				//comp = eval("frm.refunddeliverypay_" + frm.ckidx[i].value);
				//if (!IsDigit(comp.value)){
                //    alert('숫자만 가능합니다.');
                //    comp.focus();
                //    return;
                //}
				//refunddeliverypayArr = refunddeliverypayArr + comp.value + ",";
		    }
		}
	}else{
		if(frm.ckidx.checked){
			chk++;
			//comp = eval("frm.refunddeliverypay_" + frm.ckidx.value);
			//if (!IsDigit(comp.value)){
            //    alert('숫자만 가능합니다.');
            //    comp.focus();
            //    return;
            //}
			//refunddeliverypayArr = refunddeliverypayArr + eval("frm.refunddeliverypay_" + comp.value).value + ",";
		}
	}

	if(chk==0){
		alert("등록할 내역을 선택하세요.");
		return false;
	}else{
	    if (confirm('등록 하시겠습니까?')){
	        frm.refunddeliverypayArr.value = refunddeliverypayArr;
			frm.submit();
		}
	}
}

</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	기간검색 : <% DrawYMBox yyyy1,mm1 %> (정산 대상월 <%= jungsanmm %>-<%=jungsanmm %>)
	        	&nbsp;
	        	<input type="checkbox" name="cknotreg" value="on" <% if cknotreg="on" then response.write "checked" %> > 미등록내역만
	        	&nbsp;
	        	<input type="checkbox" name="cksiteMM" value="on" <% if cksiteMM="on" then response.write "checked" %> > lotteComM

	        	브랜드ID : <input type="text" name="makerid" value="<%= makerid %>" size="16" maxlength="32">
	        	&nbsp;
	        	주문번호 : <input type="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="11">
	        	&nbsp;
	        	고객ID : <input type="text" name="userid" value="<%= userid %>" size="12" maxlength="32">
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<h4>10월 정산시 6월부터 누락 내역 추가 할것('A001','A100','A002')</h4>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frmList" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="mode" value="beasongpayaddArr">
<input type="hidden" name="yyyy1" value="<%= jungsanyyyy %>">
<input type="hidden" name="mm1" value="<%= jungsanmm %>">
<input type="hidden" name="refunddeliverypayArr" value="">
<tr  bgcolor="#FFFFFF">
    <td colspan="5" >
	<input type="checkbox" name="itemnotax" >상품 면세로

    <input type="checkbox" name="notax" >면세로

    <input type="checkbox" name="jgubunMM"  >매입정산<!--disabled-->

    <input type="button" value="선택 내역 등록" onClick="chkSubmit()">

    <!-- input type="button" value="선택 내역 등록_면세" onClick="chkSubmit()" -->
    </td>
    <td colspan="3" ><input type="button" value="배송비 추가 등록" onClick="AddEtcBesongJungsan('<%= yyyy1 %>','<%= mm1 %>')"></td>
    <td colspan="8" align="right">Total : <%= oDeliverPay.FTotalCount %> page: <%= page %>/<%= oDeliverPay.FTotalPage %></td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
    <td width="20"><input type="checkbox" name="ckall" onclick="ckAll(this);"></td>
	<td width="50">접수ID</td>
	<td width="50">SITE</td>
	<td width="90">주문번호</td>
	<td width="60">고객명</td>
	<td width="70">고객ID</td>
	<td width="80">브랜드ID</td>
	<td width="60">등록자</td>
	<td width="60">처리자</td>
	<td width="180">접수내용</td>
	<td width="180">처리내용</td>
	<td width="60">초기배송비<br>환원<br></td>
	<td width="60">배송비<br>차감</td>
	<td width="60">기타<br>내역</td>
	<td width="60">정산<br>등록<br>내역</td>
	<td width="50">정산등록</td>
</tr>
<% for i=0 to oDeliverPay.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
    <td><% if Not oDeliverPay.FItemList(i).IsJungsanDataExists then %><input type="checkbox" name="ckidx" value="<%= oDeliverPay.FItemList(i).Fid %>"><% end if %></td>
    <td><%= oDeliverPay.FItemList(i).Fid %></td>
    <td><%= oDeliverPay.FItemList(i).Fsitename %></td>
    <td><%= oDeliverPay.FItemList(i).Forderserial %></td>
    <td><%= oDeliverPay.FItemList(i).Fcustomername %></td>
    <td><%= oDeliverPay.FItemList(i).Fuserid %></td>
    <td><%= oDeliverPay.FItemList(i).FMakerid %></td>
    <td ><%= oDeliverPay.FItemList(i).Fwriteuser %></td>
    <td ><%= oDeliverPay.FItemList(i).Ffinishuser %></td>
    <td ><textarea cols="20" rows="2" ><%= oDeliverPay.FItemList(i).Fcontents_jupsu %></textarea></td>
    <td ><textarea cols="20" rows="2" ><%= oDeliverPay.FItemList(i).Fcontents_finish %></textarea></td>

    <td ><%= oDeliverPay.FItemList(i).Frefundbeasongpay %></td>
    <td ><%= oDeliverPay.FItemList(i).Frefunddeliverypay %></td>

    <td ><%= oDeliverPay.FItemList(i).Fadd_upchejungsandeliverypay %></td>
    <td >
        <%= oDeliverPay.FItemList(i).FjungsanSuplycash %></td>
        <!-- <input type="text" class="text" name="refunddeliverypay_<%= oDeliverPay.FItemList(i).Fid %>" value="<%= oDeliverPay.FItemList(i).Frefunddeliverypay*-1 %>" size="5" maxlength="7"> -->
    </td>
    <td >
    <% if oDeliverPay.FItemList(i).IsJungsanDataExists then %>
    <input type="button" value="::수정::" onFocus="this.blur" onClick="EditBesongJungsan('<%= oDeliverPay.FItemList(i).FjungsanDetailId %>');">
    <% else %>
    <!-- <input type="button" value="등록" onFocus="this.blur" onClick="AddBesongJungsan('<%= oDeliverPay.FItemList(i).Fid %>');">  -->
    <% end if %>
    </td>
</tr>
<% next %>
<tr height="20">
    <td colspan="16" align="center" bgcolor="#FFFFFF">
        <% if oDeliverPay.HasPreScroll then %>
		<a href="javascript:goPage('<%= oDeliverPay.StarScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oDeliverPay.StarScrollPage to oDeliverPay.FScrollCount + oDeliverPay.StarScrollPage - 1 %>
    		<% if i>oDeliverPay.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oDeliverPay.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<form name="frmSubmit" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="mode" value="beasongpayadd">
<input type="hidden" name="refunddeliverypay" value="">
<input type="hidden" name="asid" value="">
</form>
<%
set oDeliverPay = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->