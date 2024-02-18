<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 장비 자산 리스트
' History : 2008년 06월 27일 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/bscclass/equipmentcls.asp"-->

<%
dim page, jangbi, sayoug, idx, user, usingIp , code ,equip_name ,manufacture_company
dim ipgocheck, yyyy1, yyyy2, mm1, mm2, dd1, dd2, fromDate, toDate, ipcheck, useip
dim totalcurrsum,totaljasan, Alltotaljasan, getAllCurrentValue ,equip_gubun, part_code
	page = request("page")
	if page="" then page=1
	jangbi = Request("jangbi")		'장비검색에 필요한변수
	sayoug = Request("sayoug")		'사용구분에 필요한 변수
	user = Request("user")			'사용자 검색에 필요한변수
	idx = Request("idx")			'페이지 인덱스 저장
	code = request("code")			'장비코드 검색에 필요한 변수
	ipcheck = request("ipcheck")		'ip검색에 필요한 변수
	ipgocheck = request("ipgocheck")
	yyyy1 = request("yyyy1")
	yyyy2 = request("yyyy2")
	mm1	  = request("mm1")
	mm2	  = request("mm2")
	dd1	  = request("dd1")
	dd2	  = request("dd2")
	equip_gubun = Request("equip_gubun")
	part_code = Request("part_code")
	equip_name = Request("equip_name")
	manufacture_company = Request("manufacture_company")
		
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (ipcheck <> "") then ipcheck = "on"

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim oequip,i
set oequip = new CEquipment
	oequip.FPageSize = 50
	oequip.FCurrPage = page
	oequip.FRectJangbi = jangbi
	oequip.FRectSayoug = sayoug
	oequip.FRectUser = user
	oequip.FRectIp = ipcheck
	oequip.Fequip_code = code
	oequip.frectequip_name = equip_name
	oequip.frectmanufacture_company = manufacture_company
	
	if ipgocheck = "on" then
		oequip.FRectBuyDateDtStart = fromDate
		oequip.FRectBuyDateDtEnd = toDate
	end if
	
	oequip.getEquipmentList

totalcurrsum = 0
totaljasan	 = 0
Alltotaljasan = 0
%>

<script language='javascript'>

//신규등록
function regEquipment(idx){
	var popwin = window.open('/common/pop_equipmentreg.asp?idx=' + idx,'regEquipment','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function NextPage(page){
	frm.page.value= page;
	frm.submit();
}

//구매일 체크
function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
}

function UseIpCheck(comp){
	//document.frm.ipcheck.disabled = comp.checked;
}

//삭제
function DelMe(frm){
	var ret = confirm('정말 삭제하시겠습니까? (삭제하신 장비내역은 수행하신분 성명과 함께 로그테이블에 저장 됩니다.)');

	if (ret){
		frm.submit();
	}
}

//현재페이지 인쇄
function popprint(page,jangbi,sayoug,user,idx,code){
	var popprint = window.open('/admin/newreport/equipment_print.asp?page='+page+'&jangbi='+jangbi+'&sayoug='+sayoug+'&user='+user+'&idx='+idx+'&code='+code,'popprint','width=1024,height=768,scrollbars=yes,resizable=yes');
	popprint.focus();
}

//현재페이지엑셀출력
function pageexcelsheet(page,jangbi,sayoug,user,idx,code){
	var pageexcelsheet = window.open('/admin/newreport/equipment_excel.asp?page='+page+'&jangbi='+jangbi+'&sayoug='+sayoug+'&user='+user+'&idx='+idx+'&code='+code,'pageexcelsheet','width=400,height=300,scrollbars=yes,resizable=yes');
	pageexcelsheet.focus();
}

//라벨프린터출력 시작
function ExcelSheet(idx1){
	var ExcelSheet = window.open('/common/popexcel_equipment.asp?idx=' + idx1,'ExcelSheet','width=400,height=300,scrollbars=yes,resizable=yes');
	ExcelSheet.focus();
}

//바코드 출력 팝업
function barcode(barcode){
	var barcode = window.open('/common/barcode/barcode_image.asp?barcode='+barcode+'&image=3&barcodetype=23&height=30&barwidth=1','barcode','width=600,height=400,scrollbars=yes,resizable=yes');
	barcode.focus();
}

//삭제내역보기
function poplog(){
	var poplog = window.open('/admin/newreport/equipment_loglist.asp','poplog','width=1024,height=768,scrollbars=yes,resizable=yes');
	poplog.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<input type=checkbox name="ipgocheck" value="on" <% if ipgocheck="on" then  response.write "checked" %> onclick="EnDisabledDateBox(this)">	
		구매일<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %><br>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="NextPage('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		사용자 : <% drawpartneruser "user", user ,"" %>
		장비구분 : <% DrawEquipMentGubun "10","jangbi",jangbi ,""%>
		사용구분 : <% DrawEquipMentGubun "20","sayoug",sayoug ,"" %>
	</td>	
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!--<input type="checkbox" name="ipcheck" value="on" <%if ipcheck="on" then response.write "checked" %>>사용 IP-->
		장비코드 : <input type="text" name="code" value="<%=code%>">
		제품명 : <input type="text" name="equip_name" value="<%=equip_name%>">
		제조사 : <input type="text" name="manufacture_company" value="<%=manufacture_company%>">
	</td>	
</tr>
</form>
</table>
<!-- 검색 끝 -->

<Br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" class="button" onclick="popprint('<%=page%>','<%=jangbi%>','<%=sayoug%>','<%=user%>','<%=idx%>','<%=code%>');" value="현재페이지인쇄">		
		<input type="button" class="button" onclick="pageexcelsheet('<%=page%>','<%=jangbi%>','<%=sayoug%>','<%=user%>','<%=idx%>','<%=code%>');" value="현재페이지엑셀출력">	
	</td>
	<td align="right">
		<input type="button" class="button" onclick="poplog();" value="삭제내역">
		<input type="button" class="button" onclick="regEquipment('');" value="신규등록">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oequip.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oequip.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>장비코드</td>
	<td>사용자</td>
	<td>사양</td>
	<td>장비<br>구분</td>
	<td>사용<br>구분</td>
	<td>제품명</td>
	<td>구매일</td>
	<td>구매<br>원가</td>    	
	<td>구매가</td>
	<td>자산<br>가치</td>
	<td>비고</td>
	<td>라벨<br>출력</td>
	<td>바코드출력</td>	
</tr>
<% if oequip.FResultCount > 0 then %>
<% for i=0 to oequip.FResultCount - 1 %>
<%
totalcurrsum = totalcurrsum + oequip.FItemList(i).Fbuy_sum
totaljasan	 = totaljasan + oequip.FItemList(i).GetCurrentvalue
%>
<form name=frm_<%= i %> method="post" action="frmdel.asp">	<!-- for문 안에서 i값을 가지고 실행-->
<input type="hidden" name="idx" value="<%= oequip.FItemList(i).Fidx %>">
<input type="hidden" name="ssBctId" value="<%= session("ssBctId")%>">
<% 
if oequip.FItemList(i).fstatediv <> "Y" then
'/or (oequip.FItemList(i).Fusinguserid = "" and oequip.FItemList(i).Fpart_code <>"10") 
%>
	<tr align="center" bgcolor="eeeeee" onMouseOver= this.style.background='f1f1f1'; onMouseOut=this.style.background='eeeeee';>	
<% else %>
	<tr align="center" bgcolor="#FFFFFF" onMouseOver= this.style.background='f1f1f1'; onMouseOut=this.style.background='#ffffff';>
<% end if %>
	<td width=130>
		<a href="javascript:regEquipment('<%= oequip.FItemList(i).Fidx %>');" onfocus="this.blur()">
		<%= oequip.FItemList(i).Fequip_code %></a>
	</td>
	<td width=100>
		<%= oequip.FItemList(i).FusinguserName %>
		<% if oequip.FItemList(i).fstatediv <> "Y" then %>
			<font color="red">[퇴사]</font>
		<% end if %>
		
		<% if oequip.FItemList(i).Fusinguserid <> "" then %>
			<Br><%= oequip.FItemList(i).Fusinguserid %>
		<% end if %>
	</td>
	<td>
		<%= oequip.FItemList(i).Fdetail_quality1 %><br><%= oequip.FItemList(i).Fdetail_quality2 %>
	</td>
	<td width=100>
		<%= oequip.FItemList(i).Fequip_gubun_name %>
	</td>
	<td width=100>
		<%= oequip.FItemList(i).Fpart_code_name %>
	</td>
	<td>
		<%= oequip.FItemList(i).Fequip_name %>
	</td>
	<td width=80>
		<%= oequip.FItemList(i).Fbuy_date %>
	</td>
	<td align="right" width=70>
		<%= formatNumber(oequip.FItemList(i).fwonga_cost,0) %>
	</td>		
	<td align="right" width=70>
		<%= formatNumber(oequip.FItemList(i).Fbuy_sum,0) %>
	</td>
	<td align="right" width=70>
		<% if oequip.FItemList(i).getCurrentValue<>"" then %>
			<font color="#EE3333"><%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%></font>
		<% else %>
			<%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%>
		<% end if %>
	</td>
	<td align="center" width=60>
		<input type="button" class="button" value="삭제" onclick="DelMe(frm_<%= i %>);">
	</td>
	<td width=30>
		<a href="javascript:ExcelSheet('<%= oequip.FItemList(i).Fidx %>');">
		<img src="images/iexcel.gif" border="0"></a>
	</td>
	<td align="center" width=250>
		<Br>
		<a href="javascript:barcode('<%= oequip.FItemList(i).Fequip_code %>');" onfocus="this.blur()">
		<img src="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=23&data=<%= trim(oequip.FItemList(i).Fequip_code) %>&height=30&barwidth=1&Size=7" border=0></a>
		<Br>
	</td>			
</tr>   
</form>
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan=7>총계</td>
	<td align="right"><!-- <%= formatNumber(oequip.FItemList(0).Getallcurrentvalue,0) %> --></td>
	<td align="right"><%= formatNumber(totalcurrsum,0) %></td>
	<td align="right"><font color="#EE3333"><%= formatNumber(totaljasan,0) %></font></td>
	<td align="right" colspan=3><!-- 구분별 Total : <%= formatNumber(oequip.FTotalSum,0) %> --></td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
    	<% if oequip.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oequip.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oequip.StartScrollPage to oequip.FScrollCount + oequip.StartScrollPage - 1 %>
			<% if i>oequip.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oequip.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<%
	set oequip = Nothing
%>

<script language='javascript'>
	EnDisabledDateBox(document.frm.ipgocheck);
	//UseIpCheck(document.frm.ipcheck);
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->