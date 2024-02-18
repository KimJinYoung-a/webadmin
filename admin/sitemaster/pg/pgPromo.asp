<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	: 2013.09.30 서동석 생성
'			  2022.07.04 한용민 수정(isms취약점수정)
'			  2023.01.30 원승현 수정(필요한 기능 정리 및 사용처리)
'	Description : 신용카드 프로모션 관리(결제단 무이자 display)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/pgPromotionCls.asp"-->
<%
dim i
Dim pgprogbn : pgprogbn = requestCheckVar(request("pgprogbn"),10)
Dim validgbn : validgbn = requestCheckVar(request("validgbn"),10)
Dim cardcd   : cardcd = requestCheckVar(request("cardcd"),10)
Dim iDt : iDt = requestCheckVar(request("iDt"),10)
Dim isusing : isusing = requestCheckVar(request("isusing"),10)
Dim page : page = requestCheckVar(getNumeric(request("page")),10)
Dim research : research = requestCheckVar(request("research"),10)

if (iDt="") then
    iDt=Left(CStr(now()),10)
end if

if (page="") then page=1
if (research="") and (isusing="") then isusing="Y"

Dim oCardPromo
SET oCardPromo= new CCardPromotion
oCardPromo.FRectpgprogbn = pgprogbn
oCardPromo.FRectCardCd = cardcd
oCardPromo.FRectMatchDate = iDt
oCardPromo.FRectIsusing=isusing
oCardPromo.FRectvalidgbn=validgbn
oCardPromo.getCardPromotionList

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>

function swChkTerm(comp){
   if (comp.value=="p"){
   $('#dispCal').show();
   }else{
   $('#dispCal').hide();
   }
}

function popNewCardPromotion(iidx){
	var popup_New = window.open("pop_CardSaleContentEdit.asp?idx="+iidx, "pop_RegPgPromotion", "width=1200,height=1000,scrollbars=yes,status=no,resizable=yes");
	popup_New.focus();
}

function popPreViewCurr(){
    var thatday = "";
    document.frmDumi.action="<%=wwwUrl%>/chtml/inipay/installmentMakePreView.asp?thatday="+thatday;
    document.frmDumi.submit();
}
</script>
<!-- 검색폼 시작 -->
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
        사용구분 :
        <select name="validgbn" class="select" onChange="swChkTerm(this);">
        <option value="">전체
        <option value="c" <%=CHKIIF(validgbn="c","selected","")%> >현재일기준
        <option value="p" <%=CHKIIF(validgbn="p","selected","")%> >특정일
        </select>

        <input type="radio" name="isusing" value=""  <%=CHKIIF(isusing="","checked","") %> >전체
        <input type="radio" name="isusing" value="Y" <%=CHKIIF(isusing="Y","checked","") %> >사용함
        <input type="radio" name="isusing" value="N" <%=CHKIIF(isusing="N","checked","") %> >사용안함

        <span id="dispCal" style="display:<%=CHKIIF(validgbn<>"p","none","") %>" >
        <input id="iDt" name="iDt" value="<%=iDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
        </span>
        <script language="javascript">
            var CAL_Start = new Calendar({
				inputField : "iDt", trigger    : "iDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					//CAL_End.args.min = date;
					//CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <!--<td align="left"><input type="button" value="미리보기" onclick="popPreViewCurr()" class="button"></td>-->
	<td align="right"><input type="button" value="새내용 추가" onclick="popNewCardPromotion('')" class="button"></td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="9">총&nbsp;검색수 : <%=oCardPromo.FTotalCount%> 건</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">번호</td>
	<td align="center">구분</td>
	<!--td align="center">카드사</td-->
	<!--td align="center">내용</td-->
	<!--td align="center">관련이미지</td-->
	<td align="center">기간</td>
	<td align="center">내용</td>
	<td align="center">사용여부</td>
</tr>
<% for i=0 to oCardPromo.FResultCount-1 %>
<tr  bgcolor="#FFFFFF" height="25" align="center">
    <td><%=oCardPromo.FItemList(i).FIdx%></td>
	<td>카드 무이자 할부</td>
    <td><%=Left(oCardPromo.FItemList(i).FSDt,10)%>~<%=Left(oCardPromo.FItemList(i).FeDt,10)%></td>	
	<td><input type="button" value="내용 확인 및 수정" onclick="window.open('pop_CardSaleContentEdit.asp?idx=<%=oCardPromo.FItemList(i).FIdx%>','popEditCardSaleCont','width=1200,height=1000');return false;" class="button"></td>
	<td><%=CHKIIF(isusing="Y" or isusing="","사용","사용안함")%></td>
	<!--
    <td><a href="javascript:popNewCardPromotion('<%=oCardPromo.FItemList(i).FIdx%>');"><%=getCdPromotionGubunName(oCardPromo.FItemList(i).Fpgprogbn)%></a></td>
    <td><%=getCardCd2Name(oCardPromo.FItemList(i).FCardCd)%></td>
    <td><%= ReplaceBracket(oCardPromo.FItemList(i).Fconts) %></td>
    <td>
        <% if oCardPromo.FItemList(i).getMayImageName<>"" then %>
            <img src="<%=oCardPromo.FItemList(i).getMayImageName%>" onClick="popNewCardPromotion('<%=oCardPromo.FItemList(i).FIdx%>');" style="cursor:pointer">
        <% end if %>
    </td>
    <td><%=oCardPromo.FItemList(i).getStateName%></td>
	-->
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">
	<% if oCardPromo.HasPreScroll then %>
		<a href="javascript:gotoPage(<%= oCardPromo.StarScrollPage-1 %>)">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oCardPromo.StarScrollPage to oCardPromo.FScrollCount + oCardPromo.StarScrollPage - 1 %>
		<% if i>oCardPromo.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:gotoPage(<%= i %>)">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oCardPromo.HasNextScroll then %>
		<a href="javascript:gotoPage(<%= i %>)">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<% set oCardPromo = Nothing %>
<form name="frmDumi" method="get" action="" target="_blank">

</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->