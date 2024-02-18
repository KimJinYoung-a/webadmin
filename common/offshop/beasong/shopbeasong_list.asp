<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.22 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim showshopselect, loginidshopormaker ,research, checkblock, reqhp
dim ojumun , i , orderno , ipkumdiv , page
	page = requestcheckvar(request("page"),10)
	orderno = requestcheckvar(request("orderno"),16)
	ipkumdiv = requestcheckvar(request("ipkumdiv"),2)
	research = requestcheckvar(request("research"),10)
	research = requestcheckvar(request("research"),10)
	reqhp = requestcheckvar(request("reqhp"),16)

if page = "" then page = 1
if (research = "") then
	ipkumdiv = 2	'"99"
end if

checkblock = false
showshopselect = false
loginidshopormaker = ""

if C_ADMIN_USER then
	showshopselect = true
	loginidshopormaker = request("shopid")
elseif (C_IS_SHOP) then
	'직영/가맹점
	loginidshopormaker = C_STREETSHOPID
else
	if (C_IS_Maker_Upche) then
		loginidshopormaker = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
			loginidshopormaker = "--"		'표시안한다. 에러.
		else
			showshopselect = true
			loginidshopormaker = request("shopid")
		end if
	end if
end if

set ojumun = new cupchebeasong_list
	ojumun.FPageSize = 50
	ojumun.FCurrPage = page
	ojumun.frectorderno = orderno
	ojumun.frectipkumdiv = ipkumdiv
	ojumun.frectshopid = loginidshopormaker
	ojumun.frectreqhp = reqhp
	ojumun.fbeagsong_list()

%>

<script type="text/javascript">

	//폼로드시 셀렉트
	function getOnload(){
	    frm.orderno.select();
	    frm.orderno.focus();
	}

	window.onload = getOnload;

	//폼전송
	function gosubmit(page){
		frm.page.value=page;
		frm.action='/common/offshop/beasong/shopbeasong_list.asp';
		frm.submit();
	}

	//주문수정
	function jumundetail(masteridx, orderno){
		//frmdetail.masteridx.value=masteridx;
		frmdetail.orderno.value=orderno;
		frmdetail.action='/common/offshop/beasong/shopbeasong_input.asp';
		frmdetail.submit();
	}

	//전체 배송 통보
	function beasonginput(upfrm){
		frminfo.masteridxarr.value='';
		frminfo.ordernoarr.value='';

		if (!CheckSelected()){
			alert('선택아이템이 없습니다.');
			return;
		}

		<% if C_ADMIN_AUTH and not(C_logics_Part) then %>
			if (confirm('[관리자권한]발주는 물류팀만 가능 합니다. 계속 진행하시겠습니까?')!=true){
				return;
			}

		<% elseif not(C_logics_Part) then %>
			alert('발주는 물류팀만 가능 합니다.');
			return;
		<% end if %>

		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.masteridxarr.value = upfrm.masteridxarr.value + frm.masteridx.value + "," ;
					upfrm.ordernoarr.value = upfrm.ordernoarr.value + frm.orderno.value + "," ;
				}
			}
		}

		if (confirm('배송 통보를 하시겠습니까?')){
			frminfo.mode.value='beasonginput';
			frminfo.action='/common/offshop/beasong/shopbeasong_process.asp';
			frminfo.submit();
		}
	}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page">
<input type="hidden" name="masteridx">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 매장ID :
		<% if (showshopselect = true) then %>
			<% 'drawSelectBoxOffShop "shopid",loginidshopormaker %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",loginidshopormaker, "21") %>
		<% else %>
			<%= loginidshopormaker %>
		<% end if %>
		&nbsp;&nbsp;
		* 주문번호 : <input type="text" name="orderno" value="<%= orderno %>" size="16" onKeyPress="if(window.event.keyCode==13) gosubmit('');">
		&nbsp;&nbsp;
		* 휴대폰번호 : <input type="text" name="reqhp" value="<%= reqhp %>" size=16 maxlength=16 onKeyPress="if(window.event.keyCode==13) gosubmit('');">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 배송상태 : <% drawshopIpkumDivName "ipkumdiv",ipkumdiv," onchange=gosubmit('');" %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>

<form name="frminfo" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="masteridxarr">
<input type="hidden" name="ordernoarr">
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" value="선택주문배송통보" class="button" onclick="beasonginput(frminfo);">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->
</form>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ojumun.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ojumun.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>IDX</td>
	<td>주문번호</td>
	<td>매장ID</td>
	<td>매장명</td>
	<td>수령인</td>
	<td width=110>수령인휴대폰번호</td>
	<td>배송요청일</td>
	<td>배송일</td>
	<td>배송상태</td>
	<td>비고</td>
</tr>
<% if ojumun.FresultCount>0 then %>
<%
for i=0 to ojumun.FresultCount-1

checkblock=false
' 배송지 입력완료 이전 상태 라면 통보가 되서는 안됨
if ojumun.FItemList(i).fipkumdiv < 2 then
	checkblock=true
end if

' 이미 배송통보 상태 라면 제낌
if ojumun.FItemList(i).fipkumdiv >= 5 then
	checkblock=true
end if
%>
<form action="" name="frmBuyPrc<%=i%>" method="get" style="margin:0px;">
<input type="hidden" name="orderno" value="<%= ojumun.FItemList(i).forderno %>">
<input type="hidden" name="masteridx" value="<%= ojumun.FItemList(i).fmasteridx %>">
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
	<td>
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if checkblock then %>disabled<% end if %> >
	</td>
	<td>
		<%= ojumun.FItemList(i).fmasteridx %>
	</td>
	<td>
		<%= ojumun.FItemList(i).forderno %>
	</td>
	<td>
		<%=ojumun.FItemList(i).fshopid %>
	</td>
	<td>
		<%=ojumun.FItemList(i).fshopname %>
	</td>
	<td>
		<%= ojumun.FItemList(i).freqname %>
	</td>
	<td>
		<%= ojumun.FItemList(i).freqhp %>
	</td>
	<td>
		<acronym title="<%= ojumun.FItemList(i).fregdate %>"><%= Left(ojumun.FItemList(i).fregdate, 10) %></acronym>
	</td>
	<td>
		<acronym title="<%= ojumun.FItemList(i).fbeadaldate %>"><%= Left(ojumun.FItemList(i).fbeadaldate, 10) %></acronym>
	</td>
	<td>
		<font color="<%= ojumun.FItemList(i).shopIpkumDivColor %>">
		<%= ojumun.FItemList(i).shopIpkumDivName %>
		</font>
	</td>
	<td>
		<input type="button" onclick="jumundetail('<%= ojumun.FItemList(i).fmasteridx %>','<%= ojumun.FItemList(i).forderno %>');" value="주문수정" class="button">
	</td>
</tr>
</form>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ojumun.HasPreScroll then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= ojumun.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ojumun.StartScrollPage to ojumun.StartScrollPage + ojumun.FScrollCount - 1 %>
			<% if (i > ojumun.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ojumun.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ojumun.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<form name="frmdetail" method="get" action="">
<input type="hidden" name="masteridx" value="">
<input type="hidden" name="orderno" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<%
set ojumun = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->