<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [LOG]입출고관리>>출고리스트
' Hieditor : 이상구 생성
'			 2017.03.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim code, alinkcode, page,designer, statecd, research, itemid, tplgubun,pcuserdiv,rstate, Chargename
dim chulgocheck, yyyy1,yyyy2,mm1,mm2,dd1,dd2, fromDate, toDate, PrcGbn, notalinkcode
dim totalsellcash,totalsuply,totalbuycash,totalsuply_plus,totalsuply_minus, totalitemno, i, comment
	page = requestCheckvar(request("page"),32)
	designer = requestCheckvar(request("designer"),32)
	statecd = requestCheckvar(request("statecd"),32)
	code = requestCheckvar(request("code"),640)
	alinkcode = requestCheckvar(request("alinkcode"),640)
	research = requestCheckvar(request("research"),32)
	itemid = requestCheckvar(request("itemid"),32)
	tplgubun = requestCheckvar(request("tplgubun"),32)
	pcuserdiv = requestCheckvar(request("pcuserdiv"),32)
	rstate= requestCheckvar(request("rstate"),32)
	Chargename= requestCheckvar(request("Chargename"),32)
	chulgocheck = requestCheckvar(request("chulgocheck"),32)
	yyyy1 = requestCheckvar(request("yyyy1"),32)
	yyyy2 = requestCheckvar(request("yyyy2"),32)
	mm1	  = requestCheckvar(request("mm1"),32)
	mm2	  = requestCheckvar(request("mm2"),32)
	dd1	  = requestCheckvar(request("dd1"),32)
	dd2	  = requestCheckvar(request("dd2"),32)
	PrcGbn	  = requestCheckvar(request("PrcGbn"),32)
	notalinkcode	  = requestCheckvar(request("notalinkcode"),2)

    comment	  = requestCheckvar(request("comment"),32)
    comment = Replace(comment, "'", "")

if page="" then page=1
''if (statecd="") and (research="") then statecd="1"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

code = Trim(code)
alinkcode = Trim(alinkcode)

if code <> "" then
	code = RemoveLastCariageReturn(code)
end if

if alinkcode <> "" then
	alinkcode = RemoveLastCariageReturn(alinkcode)
end if

dim oipchul
set oipchul = new CIpChulStorage
	oipchul.FCurrPage = page
	oipchul.Fpagesize=50
	oipchul.FRectCode = code
	oipchul.FRectALinkCode = alinkcode
	oipchul.FRectItemID = itemid
	oipchul.FtplGubun = tplgubun
	oipchul.FRectReportState = rstate
	oipchul.FRectPCuserDiv = pcuserdiv
	oipchul.FRectChargename = Chargename
	oipchul.FRectPrcGbn = PrcGbn
	oipchul.FRectNotalinkcode = notalinkcode

    if chulgocheck="on" or (statecd = "0" or statecd = "1") then
        oipchul.FRectComment = comment
    end if

	if code="" then
		oipchul.FRectCodeGubun = "SO"  ''출고
		oipchul.FRectSocID = designer
		oipchul.FRectChulgoState = statecd
	end if

	if chulgocheck="on" then
		oipchul.FRectExecuteDtStart = fromDate
		oipchul.FRectExecuteDtEnd   = toDate
	end if

	if oipchul.FRectItemID<>"" then
		oipchul.GetIpChulgoByItemID
	else
		oipchul.GetIpChulgoList
	end if

totalsellcash = 0
totalsuply	  = 0
totalbuycash  = 0
totalsuply_plus = 0
totalsuply_minus = 0

%>

<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script src="/js/jquery.placeholder.min.js"></script>
<script type="text/javascript">

function PopUpcheBrandInfoEdit(v){
	window.open("/admin/member/popupchebrandinfo.asp?designer=" + v,"PopUpcheBrandInfoEdit","width=640,height=580,scrollbars=yes,resizable=yes");
}

function ChulgoInput(){
	location.href="/admin/newstorage/chulgoinput.asp?menupos=<%= menupos %>";
}
function PopChulgoSheet(v,itype){
	var popwin;
	popwin = window.open('popchulgosheetNew.asp?idx=' + v + '&itype=' + itype,'popchulgosheetNew','width=760,height=600,scrollbars=yes,status=no,resizable=yes');
	popwin.focus();
}

function ExcelSheet(v,itype){
	window.open('popchulgosheetNew.asp?idx=' + v + '&itype=' + itype + '&xl=on');
}

function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
}

function NextPage(page){
	ClearPlaceHolder();
	document.frm.page.value = page;
	document.frm.submit();
}

function trim(value) {
	return value.replace(/^\s+|\s+$/g,"");
}

// 상품코드 체크
function isUInt(val) {
	var re = /^[0-9]+$/;
	return re.test(val);
}

function SubmitFrm(frm) {
	frm.itemid.value = trim(frm.itemid.value);

	if (frm.itemid.value.length > 0) {
		if (isUInt(frm.itemid.value) != true) {
			alert("상품코드는 숫자만 가능합니다.");
			return;
		}
	}

	if ((frm.PrcGbn.value === "50000") && (frm.designer.value !== "itemgift")) {
		alert("출고처가 itemgift 일때만 선택가능한 검색조건(금액구분) 입니다.");
		return;
	}

    if (frm.comment.value != '') {
        if ((frm.chulgocheck.checked != true) && (frm.statecd.value != '0') && (frm.statecd.value != '1')) {
            alert('출고일이 지정되거나,\n출고상태가 주문접수 또는 작성중 상태로 검색해야 합니다.');
            return;
        }
    }

	ClearPlaceHolder();

	if (frm.code.value.length > 0) {
		if (frm.code.value.substring(0,2) != "SO") {
			alert("출고코드가 아닙니다.");
			return;
		}
	}

	frm.submit();
}

function popXL(fromDate, toDate) {
	<% if chulgocheck<>"on" then %>
	alert("먼저 출고일을 지정하세요.");
	return;
	<% end if %>

	var popwin = window.open("/admin/newstorage/pop_chulgolist_xl_download.asp?fromDate=" + fromDate + "&toDate=" + toDate,"popXL","width=100,height=100 scrollbars=yes resizable=yes");
	popwin.focus();
}

//전자결재 품의서 등록
function jsRegEapp(scmidx,executedt){
	var BasicMonth ="<%= CStr(DateSerial(Year(now()),Month(now())-1,1))%>";
 	if ( executedt=="" ){
		alert("이미 출고처리 하였습니다.");
		return;
	}

	if (executedt.length<1){
		alert('출고일을 입력하세요.');
		calendarOpen(frm.executedt);
		return;
	}
	<% if Not (C_ADMIN_AUTH) then %>
		if ((executedt!='')&&(executedt< BasicMonth)){
			alert('출고일이 두달 지난 날짜로는 수정 불가 합니다.');
			return;
		}
	<% end if %>

	var winEapp = window.open("/admin/approval/eapp/regeapp.asp","popE","width=1000,height=600,scrollbars=yes,resizable=yes");
	document.frmEapp.iSL.value = scmidx;
	document.frmEapp.tC.value = document.all.divEapp.innerHTML.replace(/\r|\n/g,"");
	document.frmEapp.target = "popE";
	document.frmEapp.submit();
	winEapp.focus();
}

//전자결재 품의서 내용보기
function jsViewEapp(reportidx,reportstate){
	var winEapp = window.open("/admin/approval/eapp/popIndex.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}

function ClearPlaceHolder() {
	var frm = document.frm;
	frm.code.value = $('#code').val();
	frm.alinkcode.value = $('#alinkcode').val();
}

$( document ).ready(function() {
    $('textarea').placeholder();
});

function popMakeReturn(masteridx, mastercode, socid) {
    <% if Not C_ADMIN_AUTH then %>
    alert('관리자만 사용가능합니다.');
    return;
    <% else %>
    alert('관리자 권한');
    var pop = window.open("popMakeReturn.asp?idx=" +masteridx + '&code=' + mastercode + '&socid=' + socid, "popMakeReturn" , 'width=400,height=350,scrollbars=yes,resizable=yes');
	pop.focus();
    <% end if %>
}

function jsModiChulgoPrice() {
    <% if Not C_ADMIN_AUTH then %>
    alert('관리자만 사용가능합니다.');
    return;
    <% else %>
    alert('관리자 권한');
    var pop = window.open("popModiChulgoPrice.asp", "popModiChulgoPrice" , 'width=600,height=800,scrollbars=yes,resizable=yes');
	pop.focus();
    <% end if %>
}

</script>

<style>
textarea:-webkit-input-placeholder {color:#acacac;}
textarea:-moz-placeholder {color:#acacac;}
textarea:-ms-input-placeholder {color:#acacac;}
.placeholder { color: #acacac; }
</style>

<!-- 표 상단바 시작-->

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="3" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
    <td align="left">
    	출고코드 :
		<textarea class="textarea" id="code" name="code" cols="12" rows="1" placeholder="최대50개"><%= code %></textarea>
		&nbsp;
    	주문코드 :
		<textarea class="textarea" id="alinkcode" name="alinkcode" cols="12" rows="1" placeholder="최대50개"><%= alinkcode %></textarea>
		&nbsp;
		 출고처 :
		<% drawSelectBoxOffShopNotUsingAll "designer",designer %>
		<!--
		<% drawSelectBoxChulgo "designer", designer %>
		-->
    </td>
    <td rowspan="3" width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="javascript:SubmitFrm(document.frm);">
	</td>
</tr>
 <tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size=8 maxlength=12>
		&nbsp;
		등록자 : <input type="text" class="text" name="Chargename" value="<%= Chargename %>" size=8 maxlength=12>
        &nbsp;
		기타사항 : <input type="text" class="text" name="comment" value="<%= comment %>" size=10 maxlength=20>
		&nbsp;
		출고구분:
		 <input type="radio" name="pcuserdiv" value="501_21"  <% if pcuserdiv="501_21" then response.write "checked" %> >직영점
         <input type="radio" name="pcuserdiv" value="503_21"  <% if pcuserdiv="503_21" then response.write "checked" %> >기타매장
         <input type="radio" name="pcuserdiv" value="900_21"  <% if pcuserdiv="900_21" then response.write "checked" %> >출고처(기타)
		&nbsp;
		<input type="checkbox" name="notalinkcode" <% if notalinkcode="on" then  response.write "checked" %>>주문코드미연결
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
		<input type=checkbox name="chulgocheck" <% if chulgocheck="on" then  response.write "checked" %> onclick="EnDisabledDateBox(this)">출고일
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	 	&nbsp;
		출고상태 :
		<select class="select" name="statecd" >
		<option value="">전체</option>
		<option value="0" <% if statecd="0" then response.write "selected" %> >주문서작성</option>
		<option value="1" <% if statecd="1" then response.write "selected" %> >주문접수</option>
		<option value="7" <% if statecd="7" then response.write "selected" %> >출고완료</option>
		</select>
		&nbsp;
		품의상태 :
		<select class="select" name="rstate" >
		<option value="">전체</option>
		<option value="0" <% if rstate="0" then response.write "selected" %> >품의작성전</option>
		<option value="1" <% if rstate="1" then response.write "selected" %> >품의진행중 </option>
		<option value="5" <% if rstate="5" then response.write "selected" %> >품의반려 </option>
		<option value="7" <% if rstate="7" then response.write "selected" %> >품의완료</option>
		</select>
		&nbsp;
		3PL구분 : <% Call drawSelectBoxTPLGubun("tplgubun", tplgubun) %>
		&nbsp;
		금액구분 :
		<select class="select" name="PrcGbn">
			<option value=""></option>
			<option value="50000" <%= CHKIIF(PrcGbn="50000", "selected", "")%> >5만원초과(단가)</option>
		</select>
    </td>
</tr>
</table>
</form>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="5" class="a" >
<tr height="30">
    <td align="left">* 금액마이너스 정상출고&nbsp;<font color="#EE3333">*금액플러스 출고반품</font></td>
    <td align="right">
		검색결과 : <b><%= oipchul.FTotalCount %></b> <%= page %>/<%= oipchul.FTotalPage %>
		&nbsp;
        <input type="button" class="button" value=" 출고가 일괄수정" onclick="jsModiChulgoPrice();">
        &nbsp;
		<input type="button" class="button" value=" 엑셀받기 " onclick="popXL('<%= fromDate %>', '<%= toDate %>');">
		&nbsp;
		<input type="button" class="button" value=" 출고입력" onclick="ChulgoInput();">
	</td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">출고코드</td>
	<td width="60">주문코드</td>
	<td>출고처ID</td>
	<td>출고처명</td>
	<td width="60">등록자</td>
	<td width="60">처리자</td>
	<td width="60">출고상태</td>
	<td width="60">품의상태</td>
	<td width="70">요청일</td>
	<td width="70">출고일</td>
	<td width="70">판매가</td>
	<td width="70">출고가</td>
	<td width="70">매입가</td>
	<td width="60">수량</td>
	<td width="80">구분</td>
	<td width="40">할인율</td>
	<td width="40">수익</td>
    <td width="40">기능</td>
	<td width="50">내역서</td>
</tr>

<% if oipchul.FResultCount >0 then %>
	<% for i=0 to oipchul.FResultcount-1 %>
	<%
	totalsellcash = totalsellcash + oipchul.FItemList(i).Ftotalsellcash
	totalsuply	  = totalsuply + oipchul.FItemList(i).Ftotalsuplycash
	totalbuycash  = totalbuycash + oipchul.FItemList(i).Ftotalbuycash
	totalitemno = totalitemno + oipchul.FItemList(i).ftotalitemno

	if  oipchul.FItemList(i).Ftotalsuplycash>0 then
	totalsuply_plus = totalsuply_plus + oipchul.FItemList(i).Ftotalsuplycash
	else
	totalsuply_minus = totalsuply_minus + oipchul.FItemList(i).Ftotalsuplycash
	end if
	%>
	<tr bgcolor="#FFFFFF" height=24>
		<td align=center>
			<a href="/admin/newstorage/chulgodetail.asp?idx=<%= oipchul.FItemList(i).Fid %>&opage=<%= page %>&menupos=<%=menupos%>">
		  	<%= oipchul.FItemList(i).Fcode %></a>
		</td>
		<td align=center>
			<a href="/admin/fran/jumunlist.asp?menupos=520&baljucode=<%= oipchul.FItemList(i).Falinkcode %>" target="_blank">
			<%= oipchul.FItemList(i).Falinkcode %></a>
		</td>
		<td align=left><b><a href="javascript:PopUpcheBrandInfoEdit('<%= oipchul.FItemList(i).Fsocid %>');"><%= oipchul.FItemList(i).Fsocid %></a></b></td>
		<td align=left><%= oipchul.FItemList(i).Fsocname %></b></td>
		<td align=center><%= oipchul.FItemList(i).Fchargename %></td>
		<td align=center><%= oipchul.FItemList(i).Ffinishname %></td>
		<td align=center>
		    <% IF oipchul.FItemList(i).Fstatecd = "7" or oipchul.FItemList(i).Fexecutedt <> "" or not isnull(oipchul.FItemList(i).Fexecutedt) THEN  %>
		    	출고완료
		    <%elseif oipchul.FItemList(i).Fstatecd = "1" then%>
		    	주문접수
		    <%ELSE%>
		    	주문서작성
		    <%END IF%>
		</td>
		<td align=center>
			<%if oipchul.FItemList(i).Freportidx <> "" and not isNUll( oipchul.FItemList(i).Freportidx ) then%>
				<a href="javascript:jsViewEapp('<%=oipchul.FItemList(i).Freportidx%>','<%= oipchul.FItemList(i).Freportstate %>');">
				<%if oipchul.FItemList(i).Freportstate = "7" then %>
					품의완료
				<%elseif oipchul.FItemList(i).Freportstate = "5" then %>
					품의반려
				<%else%>
					진행중
				<%end if%>
				</a>
			<% end if%>
		</td>
		<td align=center><font color="#777777"><%= Left(oipchul.FItemList(i).Fscheduledt,10) %></font></td>
		<td align=center><%= Left(oipchul.FItemList(i).Fexecutedt,10) %></td>
		<td align=right>
			<% if oipchul.FItemList(i).Ftotalsellcash>0 then %>
				<font color="#EE3333"><%= FormatNumber(oipchul.FItemList(i).Ftotalsellcash,0) %></font>
			<% else %>
				<%= FormatNumber(oipchul.FItemList(i).Ftotalsellcash,0) %>
			<% end if %>
		</td>
		<td align="right">
			<% if oipchul.FItemList(i).Ftotalsellcash>0 then %>
				<font color="#EE3333"><%= FormatNumber(oipchul.FItemList(i).Ftotalsuplycash,0) %></font>
			<% else %>
				<%= FormatNumber(oipchul.FItemList(i).Ftotalsuplycash,0) %>
			<% end if %>
		</td>
		<td align="right">
			<% if oipchul.FItemList(i).Ftotalsellcash>0 then %>
				<font color="#EE3333"><%= FormatNumber(oipchul.FItemList(i).Ftotalbuycash,0) %></font>
			<% else %>
				<%= FormatNumber(oipchul.FItemList(i).Ftotalbuycash,0) %>
			<% end if %>
		</td>
		<td align=right>
			<% if oipchul.FItemList(i).ftotalitemno>0 then %>
				<font color="#EE3333"><%= FormatNumber(oipchul.FItemList(i).ftotalitemno,0) %></font>
			<% else %>
				<%= FormatNumber(oipchul.FItemList(i).ftotalitemno,0) %>
			<% end if %>
		</td>
		<td align=center><font color="<%= oipchul.FItemList(i).GetDivCodeColor %>"><%= oipchul.FItemList(i).GetDivCodeName %></font></td>
		<td align=right>
			<% if oipchul.FItemList(i).Ftotalsellcash<>0 then %>
				<%= 100-CLng(oipchul.FItemList(i).Ftotalsuplycash/oipchul.FItemList(i).Ftotalsellcash*100*100)/100 %>%
			<% end if %>
		</td>
		<td align=right>
			<% if oipchul.FItemList(i).Ftotalsuplycash<>0 then %>
				<%= round((100-CLng(oipchul.FItemList(i).Ftotalbuycash/oipchul.FItemList(i).Ftotalsuplycash*100*100)/100),2) %>%
			<% end if %>
		</td>
        <td align=center>
            <a href="javascript:popMakeReturn('<%= oipchul.FItemList(i).Fid %>', '<%= oipchul.FItemList(i).Fcode %>', '<%= oipchul.FItemList(i).Fsocid %>');">반품</a>
        </td>
		<td align=center>
	          <a href="javascript:PopChulgoSheet('<%= oipchul.FItemList(i).Fid %>','');"><img src="/images/iexplorer.gif" width=21 border=0></a> <a href="javascript:ExcelSheet('<%= oipchul.FItemList(i).Fid %>','');"><img src="/images/iexcel.gif" width=21 border=0></a>
	    </td>
	</tr>
	<% next %>

	<tr bgcolor="#FFFFFF">
		<td align="center">총계</td>
		<td colspan=9></td>
		<td align=right><%= formatNumber(totalsellcash,0) %></td>
		<td align=right>
		<%= formatNumber(totalsuply,0) %>
	<!--
		<br>
		(<%= formatNumber(totalsuply_plus,0) %>)
		<br>
		(<%= formatNumber(totalsuply_minus,0) %>)
	-->
		</td>
		<td align="right"><%= formatNumber(totalbuycash,0) %></td>
		<td align="right"><%= formatNumber(totalitemno,0) %></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
        <td></td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="19" align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr valign="bottom" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
    	<% if oipchul.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oipchul.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oipchul.StartScrollPage to oipchul.FScrollCount + oipchul.StartScrollPage - 1 %>
			<% if i>oipchul.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oipchul.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->

<%
set oipchul = Nothing
%>

<script type="text/javascript">
	EnDisabledDateBox(document.frm.chulgocheck);
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
