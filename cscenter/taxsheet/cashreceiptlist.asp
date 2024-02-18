<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_cashreceiptcls.asp"-->
<%
    '// 변수 선언 //
	dim Idx
	dim page, searchDiv, searchKey, searchString, param

	dim oreceipt, i,  bgcolor, strIsue, strConfirm
	dim ActType, useopt

    idx         = RequestCheckVar(request("idx"),9)
	page        = RequestCheckVar(request("page"),9)
	searchKey   = RequestCheckVar(request("searchKey"),20)
	searchString = RequestCheckVar(request("searchString"),20)
    searchDiv   = RequestCheckVar(request("searchDiv"),9)
    ActType     = RequestCheckVar(request("ActType"),9)
    useopt      = RequestCheckVar(request("useopt"),9)
	if page="" then
		page=1
		searchDiv = "N"
	end if

	if searchKey="" then searchKey="orderserial"
	if searchKey="reg_num" and searchString<>"" then searchString=replace(searchString,"-","")

    if ActType="" then ActType="R"

	If ActType = "F" Then
		searchDiv = "F"
	ElseIf ActType = "A" Then
		searchDiv = ""
	ELSE
	    searchDiv = "N"
	End If

	param = "&menupos=" & menupos & "&ActType="&ActType&"&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString

    set oreceipt = new CCashReceipt
    oreceipt.FCurrPage = page
	oreceipt.FPageSize = 24
	oreceipt.FsearchDiv = searchDiv
	oreceipt.FRectsearchKey = searchKey
	oreceipt.FRectsearchString = searchString
    oreceipt.FRectUseOpt = useopt
    if (ActType="C1") then
        oreceipt.FPageSize = 50
        oreceipt.getCancelRequireList
    elseif (ActType="C2") then
        oreceipt.GetMinusReceiptList()
    else
        oreceipt.GetReceiptList
    end if
%>
<script language='javascript'>
<!--
	function chk_form(){
		var frm = document.frm_search;
/*
		if(!frm.searchKey.value)
		{
			alert("검색 조건을 선택해주십시오.");
			frm.searchKey.focus();
			return;
		}
		else if(!frm.searchString.value)
		{
			alert("검색어를 입력해주십시오.");
			frm.searchString.focus();
			return;
		}
*/
		frm.submit();
	}

	function goPage(pg){
		var frm = document.frm_search;

		frm.page.value= pg;
		frm.submit();
	}

	function chgDiv(){
		var frm = document.frm_search;
		frm.page.value = "1";
		frm.submit();
	}

	function switchPrintBox(){
		var form=document.frm_list;

		if(form.chkPrint.length>1)
		{
			for(i=0;i<form.chkPrint.length;i++)
			{
				if ((form.switchPrint.checked)&&(!form.chkPrint[i].disabled))
					form.chkPrint[i].checked=true;
				else
					form.chkPrint[i].checked=false;
			}
		}
		else
		{
			if((form.switchPrint.checked)(!form.chkPrint.disabled))
				form.chkPrint.checked=true;
			else
				form.chkPrint.checked=false;
		}
	}

    function IssuSel(form){
		var chk = 0;

		if(form.chkPrint.length>1){
			for(i=0;i<form.chkPrint.length;i++){
				if(form.chkPrint[i].checked)
					chk++;
			}
		}else{
			if(form.chkPrint.checked)
				chk++;
		}

		if(chk==0){
			alert("발행을 원하시는 요청서를 선택해주십시요.");
			return false;
		}else{
		    if (confirm('발행 하시겠습니까?')){
    		    form.method="post";
    		    form.Atype.value="R";
    			form.action="receipt_process.asp";
    			form.submit();
			}
		}
	}


    function CancelSel(form){
        var chk = 0;

		if(form.chkPrint.length>1){
			for(i=0;i<form.chkPrint.length;i++){
				if(form.chkPrint[i].checked)
					chk++;
			}
		}else{
			if(form.chkPrint.checked)
				chk++;
		}

		if(chk==0){
			alert("취소를 원하시는 요청서를 선택해주십시요.");
			return false;
		}else{
		    if (confirm('취소 하시겠습니까?')){
    		    form.method="post";
    		    form.Atype.value="C1";
    			form.action="receipt_process.asp";
    			form.submit();
			}
		}
    }

    function cancelbyTid(form){
        if (confirm('취소 하시겠습니까?')){
		    form.method="post";
		    form.Atype.value="CH";
			form.action="receipt_process.asp";
			form.submit();
		}
    }

	//
	function popCashReceipt(idx)
	{
		var url = "popCashReceipt.asp?idx=" + idx;
		var popwin = window.open(url,"popCashReceipt","width=400,height=400");
		popwin.focus();
	}

	function popCashReceiptByOrderSerial(osn)
	{
		var url = "popCashReceipt.asp?orderserial=" + osn;
		var popwin = window.open(url,"popCashReceipt","width=400,height=400");
		popwin.focus();
	}

//-->
</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_search" method="GET" action="" onSubmit="return false">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
		    <input type="radio" name="ActType" value="R"  <%= chkIIF(ActType="R","checked","") %> >발행요망(미발행)
		    <input type="radio" name="ActType" value="C1" <%= chkIIF(ActType="C1","checked","") %> >취소요망
		    <input type="radio" name="ActType" value="C2" <%= chkIIF(ActType="C2","checked","") %> >마이너스발행요망
		    <input type="radio" name="ActType" value="F" <%= chkIIF(ActType="F","checked","") %> >발행실패
		    <input type="radio" name="ActType" value="A" <%= chkIIF(ActType="A","checked","") %>>전체


			<select class="select" name="useopt">
			<option value="">전체</option>
			<option value="0" <%= chkIIF(useopt="0","selected","") %>>소득공제</option>
			<option value="1" <%= chkIIF(useopt="1","selected","") %>>지출증빙</option>
			</select>

			&nbsp;
			(
			검색조건:
			<select class="select" name="searchKey">
				<option value="">선택</option>
				<option value="idx" <%= chkIIF(searchKey="idx","selected","") %>>번호</option>
				<option value="orderserial" <%= chkIIF(searchKey="orderserial","selected","") %>>주문번호</option>
				<option value="userid" <%= chkIIF(searchKey="userid","selected","") %>>회원아이디</option>
				<option value="reg_num" <%= chkIIF(searchKey="reg_num","selected","") %>>휴대폰/사업자번호</option>
			</select>
			<input type="text" class="text" name="searchString" size="20" value="<%= searchString %>">

			<script language="javascript">
				//document.frm_search.searchDiv.value="<%=searchDiv%>";
				//document.frm_search.searchKey.value="<%=searchKey%>";
			</script>
			)
			<!--
			&nbsp;&nbsp;
			발급여부:
			<select class="select" name="searchDiv" >
				<option value="">선택</option>
				<option value="Y">발급</option>
				<option value="N">미발급</option>
				<option value="F">실패</option>
			</select>
			-->
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="chk_form()">
		</td>
	</tr>
	</form>
</table>

<p>
<!-- 강제취소할 경우
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<form name="frm_can" method="get" action="">
	<input type="hidden" name="Atype" value="CH">
	<tr>
		<td align="left">
	    <input type="text" name="tid" value="" size="60">
	    <input type="button" value="취소" onClick="cancelbyTid(frm_can)">
	    </tr>
	</tr>
	</form>
</table>
<p>
	-->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<form name="frm_list" method="get" action="">
	<input type="hidden" name="Atype" value="">
	<tr>
		<td align="left">
		    <% if (ActType="R") then %>
			<input type="button" class="button" value="발행" onclick="IssuSel(frm_list);">
			<% elseif (ActType="C1") then %>
			<input type="button" class="button" value="취소" onclick="CancelSel(frm_list);">
			<% elseif (ActType="C2") then %>
			<input type="button" class="button" value="마이너스발행" onclick="MinusIssuSel(frm_list);">
			<% end if %>
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oreceipt.FTotalCount %>(<%= oreceipt.FResultCount %>)</b>
			&nbsp;
			페이지 : <b><%= page %> / <%= oreceipt.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="10"><% if oreceipt.FResultCount>0 then %><input type="checkbox" name="switchPrint" onClick="switchPrintBox()"><% end if %></td>
		<td width="70">주문번호</td>
		<td width="50">주문상태</td>
		<td width="40">주문<br>취소</td>
		<td width="60">결제금액</td>
		<td width="60">요청자</td>
		<td width="60">요청금액</td>
		<td width="80">아이디</td>
		<td width="60">승인번호</td>
		<td width="70">종류</td>
		<td width="40">결과<br>코드</td>
		<td>결과메세지</td>
		<td width="70">등록일</td>
		<td width="40">발급<br>여부</td>
	</tr>
	<%
		for i=0 to oreceipt.FResultCount - 1
			'발급여부
			if oreceipt.FItemList(i).Fresultcode="00" then
				strIsue = "<font color=darkblue>발행</font>"
			else
				strIsue = "<font color=darkred>미발행</font>"
			end if
	%>
	<tr align="center" height="25" bgcolor="#FFFFFF">
	    <% if (ActType="R") then %>
		<td><% if  IsNULL(oreceipt.FItemList(i).Fresultcode) or (oreceipt.FItemList(i).Fresultcode="R") then %><input type="checkbox" name="chkPrint" value="<%= oreceipt.FItemList(i).FIdx %>" <%= chkIIF(oreceipt.FItemList(i).Fipkumdiv<7,"disabled","") %> ><% end if %></td>
		<% else %>
		<td><% if  Not IsNULL(oreceipt.FItemList(i).Fresultcode) and (oreceipt.FItemList(i).Fresultcode="00") then %><input type="checkbox" name="chkPrint" value="<%= oreceipt.FItemList(i).FIdx %>" <%= chkIIF(oreceipt.FItemList(i).Fipkumdiv<3 or oreceipt.FItemList(i).Fidx<=1486270,"disabled","") %> ><% end if %></td>
		<% end if %>
		<td><%= oreceipt.FItemList(i).Forderserial %></td>
		<td><%= oreceipt.FItemList(i).Fipkumdiv %></td>
		<td>
		    <% if oreceipt.FItemList(i).FOrderCancelYn="N" then %>
		        정상
		    <% elseif oreceipt.FItemList(i).FOrderCancelYn="D" then %>
		        삭제
		    <% elseif oreceipt.FItemList(i).FOrderCancelYn="Y" then %>
		        취소
		    <% else %>
		        <%= oreceipt.FItemList(i).FOrderCancelYn %>
		    <% end if %>
		</td>
		<td>
		    <% if (oreceipt.FItemList(i).Fsubtotalprice<>oreceipt.FItemList(i).Fcr_price) then %>
		    <font color="red"><%= FormatNumber(oreceipt.FItemList(i).Fsubtotalprice,0) %></font>
		    <% else %>
		    <%= FormatNumber(oreceipt.FItemList(i).Fsubtotalprice,0) %>
		    <% end if %>
		</td>
		<td><%= (oreceipt.FItemList(i).Fbuyername) %></td>
		<td><%= CurrFormat(oreceipt.FItemList(i).Fcr_price) %></td>
		<td><%= printUserId(oreceipt.FItemList(i).Fuserid,2,"*") %></td>
		<td><%= oreceipt.FItemList(i).Fauthcode %></td>
		<td><%= oreceipt.FItemList(i).getReceiptType %></td>
		<td><%= oreceipt.FItemList(i).Fresultcode %></td>
		<td align="left"><%= oreceipt.FItemList(i).Fresultmsg %></td>
		<td><%= oreceipt.FItemList(i).Fregdate %></td>
		<td><a href="javascript:popCashReceiptByOrderSerial('<%= oreceipt.FItemList(i).Forderserial %>');"><%=strIsue%></a></td>
	</tr>
	<% next	%>
	</form>
</table>
<div align="center">
			<% sbDisplayPaging "page="&page, oreceipt.FTotalCount,oreceipt.FPageSize, 10%>
</div>


결과코드<br>
발행요청 : R ---> 주문시 요청했을 경우, 마이텐바이텐 증빙서류에서 재발행 불가하도록<br>
발행완료 : 00<br>
발행실패 : 01 ---> 마이텐바이텐 증빙서류 발행팝업에서 주민등록번호 재입력후 발행요청 하면, 다시 R로 변경<br>


<%
set oreceipt = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
