<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/incLotteiMallFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'''/// /admin/apps/outMallAutoJob.asp 동일 함수 존재 동시수정요망
function N_TenDlvCode2LotteDlvCode(imallname,itenCode)
    if (LCASE(imallname)="lottecom") then
        N_TenDlvCode2LotteDlvCode = TenDlvCode2LotteDlvCode(itenCode)
    elseif (LCASE(imallname)="lotteimall") then
        N_TenDlvCode2LotteDlvCode = TenDlvCode2LotteiMallDlvCode(itenCode)
    elseif (LCASE(imallname)="interpark") then
        N_TenDlvCode2LotteDlvCode = TenDlvCode2InterParkDlvCode(itenCode)
    end if
end function



dim sitename, research
dim matchState
Dim siteType

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim onemonthprevdate, nowdate,searchnextdate
dim startorderserial, endorderserial

sitename = RequestCheckVar(Request("sitename"),32)
research = RequestCheckVar(Request("research"),32)
siteType = RequestCheckVar(Request("siteType"),32)

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)
	onemonthprevdate = Left(DateAdd("m", -1, now()), 10)

	yyyy1 = Left(onemonthprevdate,4)
	mm1   = Mid(onemonthprevdate,6,2)
	dd1   = Mid(onemonthprevdate,9,2)

	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

startorderserial = Right(yyyy1, 2) & mm1 & dd1 & "00000"
endorderserial = Right(yyyy2, 2) & mm2 & dd2 & "99999"

Dim sqlStr
Dim ArrList, ArrList2
CONST MAXROWS = 200

sqlStr = " select top " + CStr(MAXROWS) + " "
sqlStr = sqlStr & " 	T.sellSite "
sqlStr = sqlStr & " 	, T.orderserial "
sqlStr = sqlStr & " 	, T.outmallorderserial "
sqlStr = sqlStr & " 	, T.OrgDetailKey  "
sqlStr = sqlStr & " 	, T.matchitemid "
sqlStr = sqlStr & " 	, T.matchitemoption "
sqlStr = sqlStr & " 	, T.orderitemname "
sqlStr = sqlStr & " 	, T.orderitemoptionname "
sqlStr = sqlStr & " 	, T.itemordercount "
sqlStr = sqlStr & " 	, d.itemno "
sqlStr = sqlStr & " 	, m.cancelyn "
sqlStr = sqlStr & " 	, d.cancelyn "
sqlStr = sqlStr & " 	, D.beasongdate "
sqlStr = sqlStr & " 	, D.currstate "
sqlStr = sqlStr & " 	, IsNULL(T.sendState,0) as sendState "
sqlStr = sqlStr & " from "
sqlStr = sqlStr & " 	db_temp.dbo.tbl_xSite_TMPOrder T "
sqlStr = sqlStr & "  	left Join db_order.dbo.tbl_order_detail D "
sqlStr = sqlStr & "  	on "
sqlStr = sqlStr & " 		T.orderserial=D.orderserial "
sqlStr = sqlStr & "  		and T.matchitemid=D.itemid "
sqlStr = sqlStr & "  		and T.matchitemoption=D.itemoption "
sqlStr = sqlStr & "  	left Join db_order.dbo.tbl_order_master M "
sqlStr = sqlStr & "  	on "
sqlStr = sqlStr & " 		D.orderserial=M.orderserial "
sqlStr = sqlStr & " where "
sqlStr = sqlStr & " 	1 = 1 "
sqlStr = sqlStr & " 	and m.cancelyn <> 'Y' "
sqlStr = sqlStr & " 	and m.ipkumdiv >= '7' "
sqlStr = sqlStr & " 	and T.orderserial >= '" + CStr(startorderserial) + "' "
sqlStr = sqlStr & " 	and T.orderserial < '" + CStr(endorderserial) + "' "
sqlStr = sqlStr & " 	and T.sellsite in ('lotteimall','lotteon','interpark', 'cjmall', 'gseshop', 'wmp', 'hmall1010', 'ezwel', 'auction1010', 'nvstorefarm', 'nvstoregift', 'Mylittlewhoopee', 'gmarket1010', '11st1010', 'coupang', 'ssg', 'skstoa', 'shintvshopping', 'kakaostore', 'lfmall')"
sqlStr = sqlStr & " 	and T.matchState not in ('D', 'A')"         ''취소 제외. 두가지 이상의 상품으로 나눠서 주문변경한 경우
sqlStr = sqlStr & " 	and T.changeitemid is NULL "				''변경상품 등록내역 제외
''sqlStr = sqlStr & " 	and IsNull(D.currstate, '0') <> '7' "
if (sitename<>"") then
    sqlStr = sqlStr & " and T.sellsite='"&sitename&"'"
end if
sqlStr = sqlStr & " 	and IsNULL(T.sendState,0)=0 "
sqlStr = sqlStr & " 	and ( "
sqlStr = sqlStr & " 		(d.orderserial is NULL) "
sqlStr = sqlStr & " 		or "
sqlStr = sqlStr & " 		((d.orderserial is not NULL) and (T.itemordercount <> 0 AND (T.itemordercount <> d.itemno))) "
sqlStr = sqlStr & " 		or "
sqlStr = sqlStr & " 		((d.orderserial is not NULL) and (d.cancelyn = 'Y')) "
sqlStr = sqlStr & " 	) "
sqlStr = sqlStr & " order by "
sqlStr = sqlStr & " 	T.orderserial desc "
''response.write sqlStr

rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
	ArrList2 = rsget.getRows
end if
rsget.Close

dim i
%>
<script language='javascript'>
function popEtcSiteOrderView(orderserial) {
    var popwin=window.open('popEtcSiteOrderView.asp?orderserial=' + orderserial,'popEtcSiteOrderView','width=1200,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<!-- 검색 시작 -->

<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">

	<tr align="center">
		<td width="50" bgcolor="<%= adminColor("gray") %>" class="td_br">검색<br>조건</td>
		<td align="left" class="td_br">
		    쇼핑몰 선택 :

		    <select class="select" name="sitename" >
		    <option value=""  >선택
    		<option value="lotteCom" <%= chkIIF(sitename="lotteCom","selected","") %> >롯데닷컴
    		<option value="lotteimall" <%= chkIIF(sitename="lotteimall","selected","") %> >롯데iMall
			<option value="lotteon" <%= chkIIF(sitename="lotteon","selected","") %> >롯데On
			<option value="shintvshopping" <%= chkIIF(sitename="shintvshopping","selected","") %> >신세계TV쇼핑
			<option value="skstoa" <%= chkIIF(sitename="skstoa","selected","") %> >SKSTOA
    		<option value="interpark" <%= chkIIF(sitename="interpark","selected","") %> >인터파크
			<option value="cjmall" <%= chkIIF(sitename="cjmall","selected","") %> >CJ몰
			<option value="gseshop" <%= chkIIF(sitename="gseshop","selected","") %> >GS샵
			<option value="homeplus" <%= chkIIF(sitename="homeplus","selected","") %> >홈플러스
			<option value="ezwel" <%= chkIIF(sitename="ezwel","selected","") %> >이지웰페어
			<option value="lfmall" <%= chkIIF(sitename="lfmall","selected","") %> >LFMall
			<option value="auction1010" <%= chkIIF(sitename="auction1010","selected","") %> >옥션
			<option value="nvstorefarm" <%= chkIIF(sitename="nvstorefarm","selected","") %> >스토어팜
			<option value="Mylittlewhoopee" <%= chkIIF(sitename="Mylittlewhoopee","selected","") %> >스토어팜 캣앤독
			<option value="nvstoregift" <%= chkIIF(sitename="nvstoregift","selected","") %> >스토어팜 선물하기
			<option value="gmarket1010" <%= chkIIF(sitename="gmarket1010","selected","") %> >G마켓
			<option value="11st1010" <%= chkIIF(sitename="11st1010","selected","") %> >11번가
			<option value="ssg" <%= chkIIF(sitename="ssg","selected","") %> >신세계몰(SSG)
			<option value="coupang" <%= chkIIF(sitename="coupang","selected","") %> >쿠팡
			<option value="hmall1010" <%= chkIIF(sitename="hmall1010","selected","") %> >Hmall
			<option value="wmp" <%= chkIIF(sitename="wmp","selected","") %> >위메프
			<option value="wmpfashion" <%= chkIIF(sitename="wmpfashion","selected","") %> >위메프W패션
			<option value="kakaostore" <%= chkIIF(sitename="kakaostore","selected","") %> >카카오톡스토어
    		</select>

            <!--
		    &nbsp;&nbsp;
		    처리상태 :
			<select class="select" name="matchState">
			<option value='' <%= chkIIF(matchState="","selected","") %> >전체</option>
	     	<option value='I' <%= chkIIF(matchState="I","selected","") %> >엑셀등록</option>
	     	<option value='O' <%= chkIIF(matchState="O","selected","") %> >주문입력완료</option>
	     	<option value='D' <%= chkIIF(matchState="D","selected","") %> >기입력삭제</option>
	     	</select>
	     	&nbsp;
            -->

			검색기간 :
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>

		<td width="50" bgcolor="<%= adminColor("gray") %>" class="td_br">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr height="25">
		<td align="left">
			<% if (IsArray(ArrList2)) THEN %>
            현재 검색건 : <%= UBound(ArrList2,2)+1 %>건 (MAX <%= MAXROWS %> )
            <% else %>
            현재 검색건 : 0 건
            <% end if %>
		<!--
			<input type="button" class="button" value="선택내역송장전송" onClick="SubmitSongjangInput(frmSvArr)">
		-->
		</td>
	</tr>
</table>
<!-- 액션 끝 -->
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl" >
	<tr align="center" class="tr_tablebar">
	    <td width="70" class="td_br">제휴사</td>
	    <td width="70" class="td_br">주문번호</td>
    	<td width="130" class="td_br">제휴주문번호</td>
      	<td class="td_br">제휴<br>상품코드</td>
      	<td width="50" class="td_br">상품코드</td>
      	<td width="50" class="td_br">옵션코드</td>
      	<td  class="td_br">상품명 <font color="blue">[옵션명]</font></td>
        <td width="30" class="td_br">주문<br>수량</td>
		<td width="30" class="td_br">배송<br>수량</td>
		<td width="30" class="td_br">취소<br>상태</td>
      	<td class="td_br">비고</td>
    </tr>
<% if (IsArray(ArrList2)) THEN %>
<%
Dim intRows2 : intRows2 = UBound(ArrList2,2)

for i=0 to intRows2
%>
<tr>
    <td class="td_br"><%= ArrList2(0,i) %></td>
    <td class="td_br"><%= ArrList2(1,i) %></td>
    <td class="td_br"><%= ArrList2(2,i) %></td>
    <td class="td_br"><%= ArrList2(3,i) %></td>
    <td class="td_br"><%= ArrList2(4,i) %></td>
	<td class="td_br"><%= ArrList2(5,i) %></td>
    <td class="td_br">
		<%= ArrList2(6,i) %>
		<% if (ArrList2(5,i) <> "0000") then %>
			<br><font color="blue">[<%= ArrList2(7,i) %>]</font>
		<% end if %>
	</td>
    <td class="td_br" align="center"><%= ArrList2(8,i) %></td>
    <td class="td_br" align="center">
		<% if (ArrList2(8,i) <> ArrList2(9,i)) then %><font color="red"><% end if %>
		<%= ArrList2(9,i) %>
	</td>
	<td class="td_br" align="center">
		<% if (ArrList2(10,i) = "Y") or (ArrList2(11,i) = "Y") then %>
			<b><font color="red">취소<font><b>
		<% end if %>
	</td>
    <td class="td_br" align="center"><input type="button" class="button" value="조회" onClick="popEtcSiteOrderView('<%= ArrList2(1,i) %>')"></td>
</tr>
<% next %>
<% ELSE %>
<tr>
    <td colspan="11" align="center">[검색 결과가 없습니다.]</td>
</tr>
<% end if %>

</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
