<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  브랜드 리스트
' History : 2012.08.21 서동석 생성
'			2012.08.22 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/util/base64unicode.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim isLocalIP : isLocalIP = fn_isDongSoongIP()
dim makerid ,catecode, groupid, offcatecode, offmduserid ,mrectTp ,i,page, pcuserdiv, readypartner   ''' partner'userdiv _ user_c'useridv
dim usingonly, research, userdiv, rect, crect, mrect, mduserid, companyno, itemid,socname_kr, Stdate, Eddate, purchasetype, qstring
dim jungsan_gubun, dispCate
	pcuserdiv   = RequestCheckVar(request("pcuserdiv"),32)
	makerid     = RequestCheckVar(request("makerid"),32)
	usingonly   = request("usingonly")
	research    = request("research")
	userdiv     = RequestCheckVar(request("userdiv"),32)
	rect        = RequestCheckVar(request("rect"),32)
	socname_kr  = requestCheckVar(request("socname_kr"),60)
	mduserid    = RequestCheckVar(request("mduserid"),32)
	catecode    = RequestCheckVar(request("catecode"),32)
	crect       = RequestCheckVar(request("crect"),32)
	mrect       = RequestCheckVar(request("mrect"),64)
	companyno   = RequestCheckVar(request("companyno"),32)
	itemid		= RequestCheckVar(request("itemid"),32)
	groupid     = RequestCheckVar(request("groupid"),32)
	offcatecode = RequestCheckVar(request("offcatecode"),32)
	offmduserid = RequestCheckVar(request("offmduserid"),32)
	mrectTp     = RequestCheckVar(request("mrectTp"),32)
	page        = request("page")
	Stdate     = RequestCheckVar(request("Stdate"),10)
	Eddate     = RequestCheckVar(request("Eddate"),10)
	purchasetype     = RequestCheckVar(request("purchasetype"),10)
	readypartner     = RequestCheckVar(request("readypartner"),2)
	jungsan_gubun     = RequestCheckVar(request("jungsan_gubun"),10)
	dispCate	= RequestCheckVar(request("dispCate"),3)

'####### 20110905 최맑은소리 사용중인것을 전체로 바꿔달라고 함.
'''if ((research="") and (usingonly="")) then usingonly="all" ''디폴트 빈값.
if page="" then page=1

dim opartner
set opartner = new CPartnerUser
	opartner.FCurrpage = page
	opartner.FPageSize = 50
	opartner.FRectPCuserDiv = pcuserdiv
	opartner.FRectGroupid = groupid
	opartner.FRectDesignerID = makerid
	opartner.FrectIsUsing = usingonly
	opartner.FRectDesignerDiv = userdiv
	opartner.FRectMdUserID = mduserid
	opartner.FRectInitial = Replace(rect,"'","''")
	opartner.FRectSOCName  = socname_kr
	opartner.FRectCompanyname = crect

	if mrectTp = "dname" then
		opartner.FRectManagerName = mrect
	elseif mrectTp = "demail" then
		opartner.FRectManageremail = mrect
	elseif mrectTp = "dphone" then
		opartner.FRectManagerhp = mrect
	end if
	if jungsan_gubun<>"" then
		opartner.FRectJungsanGubun = jungsan_gubun
	end if
	opartner.FRectCatecode = catecode
	opartner.Fitemid = itemid
	opartner.FRectCompanyNo = replace(companyno,"-","")
	opartner.FRectoffcatecode = offcatecode
	opartner.FRectoffmduserid = offmduserid
	opartner.FRectStdate = Stdate
	opartner.FRectEddate = Eddate
	opartner.FRectpurchasetype = purchasetype
	opartner.FRectReadyPartner = readypartner
	opartner.FRectDispCate = dispCate
	opartner.GetPartnerSearch
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="JavaScript" src="/js/jquery-1.7.2.min.js"></script>
<script language='javascript'>

function NextPage(page){
	frm.page.value = page;
	frm.submit();
}

function research(frm,order){
	frm.rectorder.value = order;
	frm.submit();
}

function AddNewBrand(){
	var popwin = window.open("/admin/member/addnewbrand.asp","addnewbrand","width=1400 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function AddNewBrand2(){
	var popwin = window.open("/admin/member/addnewbrand_step1.asp","addnewbrand2","width=1400 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function AddNewUpcheReg(qs){
	var popwin = window.open("/common/partner/companyinfo.asp?qs="+qs,"addnewbrand2","width=1200 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
}

function AnCheckNSongjangView(){
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
		alert('선택 내역이 없습니다.');
		return;
	}

	var ret = confirm('선택 내역으로 품절확인 및 SMS발송을 하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idarr.value = upfrm.idarr.value + frm.id.value + ",";
				}
			}
		}
		//alert(upfrm.idarr.value);
		upfrm.submit();
	}
}
function onlyNumberInput()
{
	var code = window.event.keyCode;
	if ((code > 34 && code < 41) || (code > 47 && code < 58) || (code > 95 && code < 106) || code == 8 || code == 9 || code == 13 || code == 46) {
		window.event.returnValue = true;
		return;
	}
	window.event.returnValue = false;
}

function checkform(frm)
{
    var chr1;
    for (var i=0; i<frm.itemid.value.length; i++){
        chr1 = frm.itemid.value.charAt(i);
        if(!(chr1 >= '0' && chr1 <= '9')) {
            alert("상품번호를 숫자만 입력하세요.");
            frm.itemid.focus();
            return false;
        }
    }

	if (frm.Stdate.value != "") {
		if (frm.Stdate.value.length != 10) {
			alert('잘못된 날짜입니다.');
            frm.Stdate.focus();
            return false;
		}
	}

	if (frm.Eddate.value != "") {
		if (frm.Eddate.value.length != 10) {
			alert('잘못된 날짜입니다.');
            frm.Eddate.focus();
            return false;
		}
	}

	frm.submit();
}

function popSimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popShopInfo(ishopid){
	var popwin = window.open("/admin/lib/popoffshopinfo.asp?shopid=" + ishopid + "&menupos=277","popoffshopinfo",'width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popSendJoinInfo(brandid,qs){
	var popwin = window.open("/admin/member/reSendJoinInfo.asp?brandid=" + brandid + "&qs=" + qs,"popjoinpage",'width=500,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();

}
</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<form name="frm" method="get" action="" onSubmit="return checkform(this);">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="rectorder" value="">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="3" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
    <td align="left">
    	<input type="radio" name="pcuserdiv" value="" <% if pcuserdiv="" then response.write "checked" %> >전체
        <input type="radio" name="pcuserdiv" value="9999_02" <% if pcuserdiv="9999_02" then response.write "checked" %> >매입처(일반)
        <input type="radio" name="pcuserdiv" value="9999_14" <% if pcuserdiv="9999_14" then response.write "checked" %> >매입처(강사)
        <% if (FALSE) then %>
        <input type="radio" name="pcuserdiv" value="9999_15" <% if pcuserdiv="9999_15" then response.write "checked" %> >매입처(Fingers)
        <% end if %>
        &nbsp;|&nbsp;
        <input type="radio" name="pcuserdiv" value="999_50"  <% if pcuserdiv="999_50" then response.write "checked" %> >제휴사(온라인)
        <input type="radio" name="pcuserdiv" value="501_21"  <% if pcuserdiv="501_21" then response.write "checked" %> >직영점
		<input type="radio" name="pcuserdiv" value="502_21"  <% if pcuserdiv="502_21" then response.write "checked" %> >가맹점
        <input type="radio" name="pcuserdiv" value="503_21"  <% if pcuserdiv="503_21" then response.write "checked" %> >도매처
        <input type="radio" name="pcuserdiv" value="900_21"  <% if pcuserdiv="900_21" then response.write "checked" %> >출고처(기타)
		<input type="radio" name="pcuserdiv" value="901_21"  <% if pcuserdiv="901_21" then response.write "checked" %> >웹에이전시
		<input type="radio" name="pcuserdiv" value="902_21"  <% if pcuserdiv="902_21" then response.write "checked" %> >협력업체
		<input type="radio" name="pcuserdiv" value="903_21"  <% if pcuserdiv="903_21" then response.write "checked" %> >3PL(대표)
        &nbsp;&nbsp;&nbsp;
        <input type="checkbox" name="usingonly" value="on" <%= CHKIIF(usingonly="on","checked","") %> > 사용브랜드만 보기
		&nbsp;&nbsp;&nbsp;
		<input type="checkbox" name="readypartner" value="on" <%= CHKIIF(readypartner="on","checked","") %> > 입점 진행중인 업체만 보기
	</td>
	<td rowspan="3" width="50" bgcolor="#EEEEEE"><input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td align="left" >
		전시카테고리: <%= fnStandardDispCateSelectBox(1,"", "dispCate", dispCate, "")%>
		&nbsp;
		카테고리ON : <% SelectBoxBrandCategory "catecode", catecode %>
		&nbsp;
		카테고리OFF : <% SelectBoxBrandCategory "offcatecode", offcatecode %>
		&nbsp;
		담당자ON : <% drawSelectBoxCoWorker_OnOff "mduserid", mduserid, "on" %>
		&nbsp;
		담당자OFF : <% drawSelectBoxCoWorker_OnOff "offmduserid", offmduserid, "off" %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" >
    <td align="left" >
        브랜드ID <input type="text" name="rect" value="<%= rect %>" Maxlength="32" size="14">
        &nbsp;
		그룹코드 <input type="text" name="groupid" value="<%= groupid %>" Maxlength="32" size="7">
		&nbsp;
		스트리트명(한글) : <input type="text" class="text" name="socname_kr" value="<%= socname_kr %>" Maxlength="32" size="16">
		&nbsp;
		회사명 <input type="text" name="crect" value="<%= crect %>" Maxlength="32" size="12">
		&nbsp;
		사업자번호 <input type="text" name="companyno" value="<%=companyno %>" Maxlength="32" size="12">
		<br>
		<select name="mrectTp">
			<option value="dname"  <%=CHKIIF(mrectTp="dname","selected","") %> >담당자명
			<option value="demail" <%=CHKIIF(mrectTp="demail","selected","") %> >담당자Email
			<option value="dphone" <%=CHKIIF(mrectTp="dphone","selected","") %> >담당자연락처
		</select>
		<input type="text" name="mrect" value="<%= mrect %>" Maxlength="32" size="10">
		&nbsp;
		상품번호 <input type="text" name="itemid" value="<%=itemid%>" size="8" />
		&nbsp;
		등록일 :
		<input id="Stdate" name="Stdate" value="<%=Stdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="Stdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="Eddate" name="Eddate" value="<%=Eddate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="Eddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "Stdate", trigger    : "Stdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "Eddate", trigger    : "Eddate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		&nbsp;
		구매유형 :
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",purchaseType,"" %>
		&nbsp;
		과세 구분 :
		<select name="jungsan_gubun" class="select" onchange="fnJungsanGubunChanged()">
			<option value="" <% if jungsan_gubun="" then response.write "selected" %>>전체</option>
			<option value="일반과세" <% if jungsan_gubun="일반과세" then response.write "selected" %>>일반과세</option>
			<option value="간이과세" <% if jungsan_gubun="간이과세" then response.write "selected" %>>간이과세</option>
			<option value="원천징수" <% if jungsan_gubun="원천징수" then response.write "selected" %>>원천징수</option>
			<option value="면세" <% if jungsan_gubun="면세" then response.write "selected" %>>면세</option>
			<option value="영세(해외)" <% if jungsan_gubun="영세(해외)" then response.write "selected" %>>영세(해외)</option>
		</select>
    </td>
</tr>
</table>
<!-- 표 상단바 끝-->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1" class="a" >
<tr>
	<td align="left">

	</td>
	<td align="right">
		<input type=button value="입점간소화등록" onclick="AddNewBrand2();" class="button"> <input type=button value="신규업체등록" onclick="AddNewBrand();" class="button">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		검색결과 : <b><%= opartner.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= opartner.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan=2>구분</td>
	<td rowspan=2>브랜드ID</td>
	<td rowspan=2>브랜드명(한글)<br>브랜드명(영문)</td>
	<td rowspan=2>그룹코드<br>사업자번호</td>
	<td rowspan=2>회사명</td>
	<td rowspan=2>구매유형</td>
	<td rowspan=2>등록일</td>
	<td rowspan=2>담당자</td>
	<td width="90" rowspan=2>전화번호<br>핸드폰번호</td>
	<td width="40" rowspan=2>이메일</td>
	<td width="70" colspan=3>사용여부</td>
	<td rowspan=2>업체어드민<br>오픈여부</td>
	<td rowspan=2>브랜드<br>추가정보</td>
	<td rowspan=2>기타정보</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="35">텐바이텐<br>ON</td>
	<td width="35">텐바이텐<br>OFF</td>
	<td width="35">제휴몰</td>
</tr>
<% if opartner.FresultCount > 0 then %>
<% for i=0 to opartner.FresultCount-1 %>
<% qstring = Server.UrlEncode(TBTEncryptUrl(Cstr(opartner.FPartnerList(i).FID) + "|" +Cstr(opartner.FPartnerList(i).FpcUserDiv))) %>
<% if opartner.FPartnerList(i).Fisusing="Y"	then %>
<tr bgcolor="#FFFFFF">
<% else %>
<tr bgcolor="#EEEEEE">
<% end if %>
	<td align="center"><%= opartner.FPartnerList(i).GetUserDivName %></a></td>
	<td><a href="<%=vwwwUrl%>/street/street_brand.asp?makerid=<%= opartner.FPartnerList(i).FID %>" title="브랜드 스트리트 보기" target="_blank"><%= opartner.FPartnerList(i).FID %></a></td>
	<td>
		<a href="javascript:PopBrandInfoEdit('<%= opartner.FPartnerList(i).FID %>')" title="브랜드 정보 수정">
			<%= opartner.FPartnerList(i).FSocName_Kor %><br>
			<%= opartner.FPartnerList(i).FSocName %>
		</a>
	</td>
	<td <%= CHKIIF((Trim(opartner.FPartnerList(i).FGroupId)="") or isNULL(opartner.FPartnerList(i).FGroupId),"bgcolor='#EEEEEE'","") %> >
		<% if opartner.FPartnerList(i).FGroupId="" or IsNull(opartner.FPartnerList(i).FGroupId) then %>
		<a href="javascript:AddNewUpcheReg('<%=qstring%>')"><font color="red">그룹코드 생성</font> </a><br>
        <% else %>
        <%= opartner.FPartnerList(i).FGroupId %><br>
		<% end if %>
		<%= socialnoBlank(opartner.FPartnerList(i).Fcompany_no) %>
	</td>
	<td><a href="javascript:PopUpcheInfoEdit('<%= opartner.FPartnerList(i).FGroupID %>')"><%= opartner.FPartnerList(i).Fcompany_name %></a></td>
	<td align="center">
		<%= opartner.FPartnerList(i).fpurchasetypename %>
	</td>
	<td align="center"><%= Left(opartner.FPartnerList(i).Fregdate,10) %></td>
	<td align="center"><%= opartner.FPartnerList(i).Fmanager_name %></td>
	<td>
	    <% if (isLocalIP) then '' md쪽 요청(희란) //2016/02/26
	    %>
	        <%= (opartner.FPartnerList(i).Ftel) %><br>
		    <%= (opartner.FPartnerList(i).Fmanager_hp) %>
	    <% else %>
		    <%= GetTelWithAsterisk(opartner.FPartnerList(i).Ftel) %><br>
		    <%= GetTelWithAsterisk(opartner.FPartnerList(i).Fmanager_hp) %>
	    <% end if %>
	</td>
	<td align="center">
	     <% if (isLocalIP) then %>
	<%= opartner.FPartnerList(i).Femail %>
	<%end if%>
		<% if opartner.FPartnerList(i).Femail<>"" then %>
		&nbsp;<a href="mailto:<%= opartner.FPartnerList(i).Femail %>"><img src="/images/icon_search.jpg" width="16" border="0" alt="<%= opartner.FPartnerList(i).Femail %>"></a>
		<% else %>
		&nbsp;
		<% end if %>
	</td>
	<td align=center>
		<% if opartner.FPartnerList(i).Fisusing="Y" then %>
		O
		<% else %>
		X
		<% end if %>
	</td>
	<td align=center>
		<% if opartner.FPartnerList(i).Fisoffusing="Y"	then %>
		O
		<% else %>
		X
		<% end if %>
	</td>
	<td align=center>
		<% if opartner.FPartnerList(i).Fisextusing="Y"	then %>
		O
		<% else %>
		X
		<% end if %>
	</td>
	<td align=center>
		<a href="javascript:PopBrandAdminUsingChange('<%= opartner.FPartnerList(i).FID %>');">
		<% if opartner.FPartnerList(i).Fpartnerusing="Y" then %>
			<% if opartner.FPartnerList(i).Fisusing="N" then %>
			<font color="red"><b>O</b></font>
			<% else %>
			O
			<% end if %>
		<% elseif IsNULL(opartner.FPartnerList(i).Fpartnerusing) then %>
		<font color="red">없음</font>
		<% else %>
		<font color="red">X</font>
		<% end if %>
		</a>
	</td>
	<td align=center>
	<% if (opartner.FPartnerList(i).isbuyingPartner) then %>
	<a href="javascript:popSimpleBrandInfo('<%= opartner.FPartnerList(i).FID %>')">[보기]</a>
	<% elseif (opartner.FPartnerList(i).isShopPartner) then %>
	<a href="javascript:popShopInfo('<%= opartner.FPartnerList(i).FID %>')">[보기]</a>
	<% end if %>
	</td>
	<td>
		<% if opartner.FPartnerList(i).FGroupId="" or IsNull(opartner.FPartnerList(i).FGroupId) then %>
		<a href="javascript:AddNewUpcheReg('<%=qstring%>')">입점정보저장</a><br>
		<a href="javascript:popSendJoinInfo('<%= opartner.FPartnerList(i).FID %>','<%=qstring%>')">입점정보 재발송</a>
		<% end if %>
	</td>
</tr>
<% next %>

<tr bgcolor="FFFFFF">
	<td colspan="16" align="center">
    	<% if opartner.HasPreScroll then %>
		<a href="javascript:NextPage('<%= opartner.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + opartner.StartScrollPage to opartner.FScrollCount + opartner.StartScrollPage - 1 %>
			<% if i>opartner.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if opartner.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="16">
		검색 결과가 없습니다.
	</td>
</tr>
<% end if %>
</table>

<form name="frmArrupdate" method="post" action="soldout_comparison_ok.asp">
	<input type="hidden" name="idarr" value="">
</form>
<%
set opartner = Nothing

function ereg(strOriginalString, strPattern, varIgnoreCase)
    ' Function matches pattern, returns true or false
    ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
    dim objRegExp : set objRegExp = new RegExp
    with objRegExp
        .Pattern = strPattern
        .IgnoreCase = varIgnoreCase
        .Global = True
    end with
    ereg = objRegExp.test(strOriginalString)
    set objRegExp = nothing
end Function

function ereg_replace(strOriginalString, strPattern, strReplacement, varIgnoreCase)
    ' Function replaces pattern with replacement
    ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
    dim objRegExp : set objRegExp = new RegExp
    with objRegExp
        .Pattern = strPattern
        .IgnoreCase = varIgnoreCase
        .Global = True
    end with
    ereg_replace = objRegExp.replace(strOriginalString, strReplacement)
    set objRegExp = nothing
end function

''TODO : 마이너스 없는 전화번호 처리 안함.
''(0101112222, 021112222, 0312223333)
function GetTelWithAsterisk(telNo)
	dim resultStr, tmpArr, i

	resultStr = telNo

	if IsNull(telno) then
		GetTelWithAsterisk = resultStr
		Exit Function
	end if

	tmpArr = Split(telNo, "-")

	Select Case UBound(tmpArr)
		Case 1
			resultStr = ereg_replace(tmpArr(0), ".", "*", True) & "-" & tmpArr(0)
		Case 2
			resultStr = tmpArr(0) & "-" & ereg_replace(tmpArr(1), ".", "*", True) & "-" & tmpArr(2)
		Case Else
			resultStr = "ERR"
	End Select

	GetTelWithAsterisk = resultStr
end Function
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
