<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  출고지시서 작성
' History : 이상구 생성
'			2021.05.13 한용민 수정(한진택배 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbOpen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/tplbalju.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

DIM CBRAND_INEXCLUDE_USING : CBRAND_INEXCLUDE_USING = True
Dim FlushCount : FlushCount=100  ''2016/04/18 :: ASP 페이지를 실행하여 Response 버퍼의 구성된 제한이 초과되었습니다.

dim pagesize
dim notitemlist, itemlist
dim notitemlistinclude, itemlistinclude
dim notbrandlistinclude, brandlistinclude
dim research
dim yyyy1,mm1,dd1,yyyymmdd,nowdate
dim onlyOne,dcnt
dim danpumcheck
dim upbeaInclude
dim dcnt2
dim imsi, sagawa, ems, epostmilitary, bigitem, fewitem, kpack
dim searchtypestring
dim deliveryarea
dim onejumuntype
dim onejumuncount, onejumuncompare
dim tenbeaonly
dim tenbeamakeonorder
dim cn10x10
dim ecargo
dim extSiteName
dim stockLocationGubun
dim excMinusStock
dim presentOnly
dim show100
dim repeatOrderCnt,songjangdiv
dim excZipcode

'==============================================================================
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

pagesize = request("pagesize")

deliveryarea = request("deliveryarea")

bigitem = request("bigitem")
fewitem = request("fewitem")

upbeaInclude = request("upbeaInclude")

notitemlistinclude = request("notitemlistinclude")
itemlistinclude = request("itemlistinclude")

notbrandlistinclude = request("notbrandlistinclude")
brandlistinclude = request("brandlistinclude")

notitemlist = request("notitemlist")
itemlist = request("itemlist")

research = request("research")

onejumuntype = request("onejumuntype")
onejumuncount = request("onejumuncount")
onejumuncompare = request("onejumuncompare")

tenbeaonly = request("tenbeaonly")

tenbeamakeonorder = request("tenbeamakeonorder")

extSiteName = request("extSiteName")

stockLocationGubun = request("stockLocationGubun")
excMinusStock = request("excMinusStock")
presentOnly = request("presentOnly")
show100 = request("show100")
repeatOrderCnt = request("repeatOrderCnt")
excZipcode = request("excZipcode")

if (research = "") then
	''notitemlistinclude = "on"
	''if (CBRAND_INEXCLUDE_USING) then
	''    notbrandlistinclude = "on"
    ''end if
	''tenbeamakeonorder = "E"
	''extSiteName = "10x10"
	''presentOnly = "N"
	''show100 = "Y"
    'if Left(Now(), 10) >= "2021-06-15" then excZipcode = "Y"
end if

if (repeatOrderCnt = "") then
	repeatOrderCnt = "0"
end if

'dcnt = trim(request("dcnt"))
'dcnt2 = trim(request("dcnt2"))
'onlyOne = request("onlyOne")
'danpumcheck = request("danpumcheck")
'imsi  = request("imsi")
'sagawa= request("sagawa")
'ems   = request("ems")
'epostmilitary   = request("epostmilitary")



'==============================================================================
if yyyy1="" then
	nowdate = CStr(Now)
	nowdate = DateSerial(Left(nowdate,4), CLng(Mid(nowdate,6,2))-2,Mid(nowdate,9,2))
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
	dd1 = Mid(nowdate,9,2)
end if

if onejumuncount="" then
	onejumuncount = "1"
end if

if onejumuncompare="" then
	onejumuncompare = "less"
end if


'// ============================================================================
ems = ""
kpack = ""
epostmilitary = ""
cn10x10 = ""
ecargo = ""


'// ============================================================================
''임시..
'if (research="") then
'    notitemlist = "311341"
'    notitemlistinclude="on"
'end if

if research="" then
	'notitemlist = "45718"
	''if notitemlist="" then notitemlist="29002,29003,29004,29005,29006,29007,29008,29009,29010,29011,29012,29013,29014"
	''if itemlist="" then itemlist="29002,29003,29004,29005,29006,29007,29008,29009,29010,29011,29012,29013,29014"
	'if notitemlistinclude="" then notitemlistinclude="on"
end if

''쿠키 없앰 /2016/04/18
''if (pagesize="") then
''	pagesize = request.cookies("baljupagesize")
''end if

if (pagesize="") then pagesize=200
''if (pagesize>=2000) then pagesize=1000

''쿠키 없앰 /2016/04/18
''response.cookies("baljupagesize") = pagesize






dim page
dim ojumun

page = request("page")
if (page="") then page=1

set ojumun = new CTenBalju

''총 페이징의 2배 검색
''ojumun.FPageSize = pagesize * 3
ojumun.FPageSize = pagesize

if notitemlistinclude="on" then
	ojumun.FRectNotIncludeItem = "Y"
else
	ojumun.FRectNotIncludeItem = ""
end if

if itemlistinclude="on" then
	ojumun.FRectIncludeItem = "Y"
else
	ojumun.FRectIncludeItem = ""
end if

if notbrandlistinclude="on" then
	ojumun.FRectNotIncludebrand = "Y"
else
	ojumun.FRectNotIncludebrand = ""
end if

if brandlistinclude="on" then
	ojumun.FRectIncludebrand = "Y"
else
	ojumun.FRectIncludebrand = ""
end if

ojumun.FCurrPage = page

ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1

''업체배송 포함 주문건.
ojumun.FRectUpbeaInclude = upbeaInclude

''사가와 배송권역
ojumun.FRectOnlySagawaDeliverArea = sagawa

if deliveryarea<>"" then
	ojumun.FRectDeliveryArea = deliveryarea
end if

if fewitem<>"" then
	ojumun.FRectOnlyFewItem = fewitem
end if

if onejumuntype<>"" then
	ojumun.FRectOnlyOneJumun = "Y"

	ojumun.FRectOnlyOneJumunType = onejumuntype
	ojumun.FRectOnlyOneJumunCompare = onejumuncompare
	ojumun.FRectOnlyOneJumunCount = onejumuncount
end if

if tenbeaonly<>"" then
	ojumun.FRectTenbeaOnly = "Y"
end if

if tenbeamakeonorder <> "" then
	ojumun.FRectTenbeaMakeOnOrder = tenbeamakeonorder
end if


ojumun.FRectSiteGubun = extSiteName

ojumun.FRectStockLocationGubun = stockLocationGubun
ojumun.FRectExcMinusStock = excMinusStock
ojumun.FRectPresentOnly = presentOnly
ojumun.FRectRepeatOrderCnt = repeatOrderCnt
ojumun.FRectExcZipcode = excZipcode

ojumun.GetBaljuItemListProc
''ojumun.GetBaljuItemListNew


dim ix,iy
dim tenbaljucount
tenbaljucount =0

dim MaxTenBaljuCount : MaxTenBaljuCount = 100

%>
<script language='javascript'>
var tenBaljuCnt = 0;
function CheckNBalju(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 주문이 없습니다.');
		return;
	}

    if (document.all.groupform.songjangdiv.value.length<1){
		alert('출고 택배사를 선택 하세요.');
		document.all.groupform.songjangdiv.focus();
		return;
	}

	if (document.all.groupform.workgroup.value.length<1){
		alert('작업 그룹을 선택 하세요.');
		document.all.groupform.workgroup.focus();
		return;
	}

	var ret = confirm('선택 주문을 새 출고지시서로 저장하시겠습니까?');
	if (ret){
		upfrm.orderserial.value = "";
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderserial.value = upfrm.orderserial.value + "|" + frm.orderserial.value;
				}
			}
		}
		upfrm.songjangdiv.value = document.all.groupform.songjangdiv.value;
		upfrm.workgroup.value = document.all.groupform.workgroup.value;
		upfrm.extSiteName.value = "<%= extSiteName %>";

		if ((upfrm.songjangdiv.value == "91") || (upfrm.songjangdiv.value == "92") || (upfrm.songjangdiv.value == "93")) {
			upfrm.songjangdiv.value = "90";
		}

		upfrm.submit();
	}
}

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="viewordermaster.asp"
	frm.submit();

}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function EnableDiable(icomp){
	//return;
	var frm = document.frm;
	var ischecked = icomp.checked;
	if (ischecked){
		if (icomp.name=="notitemlistinclude"){
			frm.itemlistinclude.checked = !(ischecked);
		}else if (icomp.name=="itemlistinclude"){
			frm.notitemlistinclude.checked = !(ischecked);
		}

	}

	if (ischecked){
		if (icomp.name=="notbrandlistinclude"){
			frm.brandlistinclude.checked = !(ischecked);
		}else if (icomp.name=="brandlistinclude"){
			frm.notbrandlistinclude.checked = !(ischecked);
		}

	}

	if (icomp.name=="onlyOne"){
		frm.itemlistinclude.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.danpumcheck.checked = false;
	}


	if (icomp.name=="danpumcheck"){
		frm.itemlistinclude.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.onlyOne.checked = false;
	}
}

function poponeitem(){
	var popwin = window.open("/admin/etc/3pl/balju/poponeitem.asp","poponeitem","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function chkUpbea(){
    var frm;
    var checkedExists = false;
    for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.tenbeaexists.value!="Y"){
			    frm.cksel.checked = true;
			    AnCheckClick(frm.cksel);
			    checkedExists = true;
			}
		}
	}

	if (checkedExists){
	    document.groupform.songjangdiv.value="24";
	    document.groupform.workgroup.value="Z";
	    CheckNBalju();
	}
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
        		<tr height="35">
        			<td width="320"><b>기간 : <% DrawOneDateBox yyyy1,mm1,dd1 %> ~ 현재</td>
        			<td width="220">
        				&nbsp;&nbsp;|&nbsp;&nbsp;
        				<b>텐바이텐배송 건수</b> :
						<select name="pagesize" >
						<option value="10" <% if pagesize="10" then response.write "selected" %> >10</option>
						<option value="20" <% if pagesize="20" then response.write "selected" %> >20</option>
						<option value="50" <% if pagesize="50" then response.write "selected" %> >50</option>
						<option value="100" <% if pagesize="100" then response.write "selected" %> >100</option>
						<option value="120" <% if pagesize="120" then response.write "selected" %> >120</option>
						<option value="150" <% if pagesize="150" then response.write "selected" %> >150</option>
						<option value="200" <% if pagesize="200" then response.write "selected" %> >200</option>
						<option value="250" <% if pagesize="250" then response.write "selected" %> >250</option>
						<option value="300" <% if pagesize="300" then response.write "selected" %> >300</option>
						<option value="400" <% if pagesize="400" then response.write "selected" %> >400</option>
						<option value="500" <% if pagesize="500" then response.write "selected" %> >500</option>
						<option value="600" <% if pagesize="600" then response.write "selected" %> >600</option>
						<option value="800" <% if pagesize="800" then response.write "selected" %> >800</option>
						<option value="1000" <% if pagesize="1000" then response.write "selected" %> >1000</option>
						<!--
						<option value="2000" <% if pagesize="2000" then response.write "selected" %> >2000</option>
						-->
						</select>
        			</td>
        			<td width="250">
        				&nbsp;&nbsp;|&nbsp;&nbsp;
        				<b>주문사이트</b> :
						<% CALL drawPartner3plCompany("extSiteName",extSiteName,"") %>
        			</td>
        			<td width="150">
                        <input type="checkbox" name="excZipcode" value="Y" <% if (excZipcode = "Y") then %>checked<% end if %> > 배송불가지역 제외
        			</td>
        			<td width="200">

					</td>
					<td width="250">
					</td>
					<td>
					</td>
        		</tr>
        	</table>
        	<table border="0" cellpadding="0" cellspacing="0" class="a">
        		<tr height="35">
        			<td>
        				<b>단품주문</b> :
						<select name="onejumuntype" >
						<option value="" 	<% if onejumuntype="" then response.write "selected" %> ></option>
						<option value="all" <% if onejumuntype="all" then response.write "selected" %> >모든 단품주문</option>
						<option value="reg" <% if onejumuntype="reg" then response.write "selected" %> >설정된 단품주문</option>
						</select>

						<input type="text" name="onejumuncount" value="<%= onejumuncount %>" size=3>
						<select name="onejumuncompare" >
						<option value="less" 	<% if onejumuncompare="less" then response.write "selected" %> >개 이하</option>
						<option value="more" 	<% if onejumuncompare="more" then response.write "selected" %> >개 이상</option>
						<option value="equal" 	<% if onejumuncompare="equal" then response.write "selected" %> >개</option>
						</select>
        			</td>
        			<td>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<!--<input type="checkbox" name="danpumcheck" <% if danpumcheck="on" then response.write "checked" %> onclick="EnableDiable(this);">-->
						<input type="button" value="제외/포함/단품 상품설정" onclick="javascript:poponeitem();">
						&nbsp;&nbsp;
						<!--<input type="text" name="dcnt2" value="<%= dcnt2 %>" size=1> 개 (11 입력시 11개 이상, 0개 입력시 0개 이상)-->
						|
						&nbsp;&nbsp;
						<select class="select" name="fewitem">
							<option></option>
							<option value="2UP" <% if fewitem="2UP" then response.write "selected" %>>2 가지 이상</option>
							<option value="4UP" <% if fewitem="4UP" then response.write "selected" %>>4 가지 이상</option>
							<option value="5UP" <% if fewitem="5UP" then response.write "selected" %>>5 가지 이상</option>
							<option value="6UP" <% if fewitem="6UP" then response.write "selected" %>>6 가지 이상</option>
							<option value="10UP" <% if fewitem="10UP" then response.write "selected" %>>10 가지 이상</option>
							<option value="15UP" <% if fewitem="15UP" then response.write "selected" %>>15 가지 이상</option>
							<option value="20UP" <% if fewitem="20UP" then response.write "selected" %>>20 가지 이상</option>
							<option value="10DN" <% if fewitem="10DN" then response.write "selected" %>>10 가지 이하</option>
							<option value="3DN" <% if fewitem="3DN" then response.write "selected" %>>3 가지 이하</option>
							<option value="2DN" <% if fewitem="2DN" then response.write "selected" %>>2 가지 이하</option>
							<option value="1DN" <% if fewitem="1DN" then response.write "selected" %>>1 가지 이하</option>
							<option value="2UP,10DN" <% if fewitem="2UP,10DN" then response.write "selected" %>>2~10 가지</option>
						</select>
						주문만
        			</td>
        		</tr>

        	</table>
        </td>
        <td align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>

<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
	        총 미출고지시 건수 : <Font color="#3333FF"><b><%= FormatNumber(ojumun.FTotalCount,0) %></b></font>&nbsp;
			총 금액 : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotalsum,0) %></font>&nbsp;
			평균객단가 : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotalsum,0) %></font>
        </td>
        <td>&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="40" valign="center">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td width="200" align="left">
	        <div id="currsearchno">총 검색주문건수 : </div>
	        <div id="currtensearchno">텐바이텐배송 주문건수 : </div>
	        <!-- input type="checkbox" name="ck_upbea" onClick="chkUpbea();"> 업배출고지시 -->
        </td>
        <td align="right">
		<form name="groupform">
		    <!--
		    <select name="baljutype">
		        <option value="">일반
		        <option value="D">DAS
		        <option value="S">단품출고
		    </select>
		    -->
			<b>
			<%
			Select Case extSiteName
				Case "10x10"
					response.write "텐텐(제휴몰 제외) 주문"
				Case "extSiteAll"
					response.write "제휴몰전체(텐텐제외) 주문"
				Case "cjmall"
					response.write "제휴몰(cjmall) 주문"
				Case "interpark"
					response.write "제휴몰(interpark) 주문"
				Case "lotteCom"
					response.write "제휴몰(lotteCom) 주문"
				Case "lotteimall"
					response.write "제휴몰(lotteimall) 주문"
				Case "etcExtSite"
					response.write "기타제휴몰 주문"
				Case Else
					response.write "전체 주문"
			End Select
			%>
			</b>
			&nbsp;
		    <select name="songjangdiv">
		        <option value="">택배사선택</option>
				<option value="1" <% if songjangdiv="1" then response.write " selected" %> >한진택배</option>
				<option value="2" <% if songjangdiv="2" then response.write " selected" %> >롯데택배</option>
                <option value="4" <% if songjangdiv="4" then response.write " selected" %> >CJ택배</option>
				<option value="98" >퀵배송</option>
		    </select>
			<select name="workgroup">
			   	<option value="">작업그룹</option>
			   	<option value="3" >3 (3PL)</option>
				<option value="M" >M (3PL-단품출고)</option>
		   	</select>
			<input type="button" value="선택주문 출고지시서작성" onclick="CheckNBalju()">
		</form>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20" align="center"><input type="checkbox" name="cksel" onClick="AnSelectAllFrame(true)"></td>
	<td width="120">고객사</td>
	<td width="120">주문번호</td>
	<td width="120">Site</td>
	<td width="50">국가</td>
	<td width="120">UserID</td>
	<% if (FALSE) then %>
	<td width="120">구매자</td>
	<% end if %>
	<td width="120">수령인</td>
	<td width="60">결제금액</td>
	<td width="60">구매총액</td>
	<td width="80">결제방법</td>
	<td width="80">거래상태</td>
	<td width="110">주문일</td>
	<td width="60">상품<br />가지수</td>
	<td>
	    <% if upbeaInclude<>"" then %>
	    업배포함
	    <% else %>
	    텐배포함
	    <% end if %>
	    </td>
</tr>
<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="14" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<% if tenbaljucount < CLng(pagesize) and (tenbaljucount < MaxTenBaljuCount or show100 <> "Y")  then %>
	<% if (ix/FlushCount)=CLNG(ix/FlushCount) then response.write CLNG(ix/FlushCount): response.flush %>
<form name="frmBuyPrc_<%= ojumun.FItemList(ix).FOrderSerial %>" method="post" >
<input type="hidden" name="orderserial" value="<%= ojumun.FItemList(ix).FOrderSerial %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sitename" value="<%= ojumun.FItemList(ix).FSiteName %>">
<input type="hidden" name="dlvcontrycode" value="<%= ojumun.FItemList(ix).FDlvcountryCode %>">
<tr align="center" bgcolor="#FFFFFF">

<!-- !!! EMS 군부대 중국몰배송 체크는 클래스파일에서 한다. !!! -->

<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <%= CHKIIF(extSiteName<>ojumun.FItemList(ix).Ftplcompanyid,"disabled","") %> ></td>

<td><%= ojumun.FItemList(ix).Ftplcompanyid %></td>
<td><%= ojumun.FItemList(ix).FOrderSerial %></td>
<td><font color="<%= ojumun.FItemList(ix).SiteNameColor %>"><%= ojumun.FItemList(ix).FSitename %></font></td>
<td><%= ojumun.FItemList(ix).FDlvcountryCode %></td>
<td><%= printUserId(ojumun.FItemList(ix).FUserID,2,"*") %></td>
<% if (FALSE) then %>
<td><%= ojumun.FItemList(ix).FBuyName %></td>
<% end if %>
<td><%= ojumun.FItemList(ix).FReqName %></td>
<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></font></td>
<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FTotalSum,0) %></td>
<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
<td><%= Left(ojumun.FItemList(ix).FRegDate,16) %></td>
<td><%= ojumun.FItemList(ix).FTenbeaItemKindCnt %></td>
<td>
<% if ojumun.FItemList(ix).Ftenbeaexists then %>
<input type="hidden" name="tenbeaexists" value="Y">
<% tenbaljucount = tenbaljucount + 1 %>
√ <%= tenbaljucount %>
<% else %>
<input type="hidden" name="tenbeaexists" value="N">
<% end if %>
</td>
</tr>
</form>
	<% else %>
		<% exit for %>
	<% end if %>
	<% next %>
<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<form name="frmArrupdate" method="post" action="dobaljumaker_3PL.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="songjangdiv" value="">
<input type="hidden" name="workgroup" value="">
<input type="hidden" name="extSiteName" value="">
</form>
<%
set ojumun = Nothing
%>

<script language='javascript'>
document.all.currsearchno.innerHTML = "검색갯수 : <Font color='#3333FF'><%= ix %></font>";
document.all.currtensearchno.innerHTML = "텐바이텐배송 검색갯수 : <Font color='#3333FF'><%= tenbaljucount %></font>";
tenBaljuCnt = 1*<%= tenbaljucount %>;
<% if onlyOne<>"" then %>
EnableDiable(frm.onlyOne);
<% end if %>

<% if danpumcheck<>"" then %>
EnableDiable(frm.danpumcheck);
<% end if %>

</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->
