<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	2010년 01월 06일 한용민 생성
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/bscclass/equipmentcls.asp"-->
<%
' 변수 선언
dim page, jangbi, sayoug, idx, user, usingIp , code 
dim ipgocheck, yyyy1, yyyy2, mm1, mm2, dd1, dd2, fromDate, toDate, ipcheck, useip
	page = request("page")
	if page="" then page=1
	jangbi = Request("jangbi")		'장비검색에 필요한변수
	sayoug = Request("sayoug")		'사용구분에 필요한 변수
	user = Request("user")			'사용자 검색에 필요한변수
	idx = Request("idx")			'페이지 인덱스 저장
	code = request("code")			'장비코드 검색에 필요한 변수
	ipcheck = request("ipcheck")		'ip검색에 필요한 변수

	' 입고일 검색에 필요한 변수 대입
	ipgocheck = request("ipgocheck")
	yyyy1 = request("yyyy1")
	yyyy2 = request("yyyy2")
	mm1	  = request("mm1")
	mm2	  = request("mm2")
	dd1	  = request("dd1")
	dd2	  = request("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (ipcheck <> "") then ipcheck = "on"

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim oequip,i				'class 선언
set oequip = new CEquipment		'class변수 지정
	oequip.FPageSize = 50
	oequip.FCurrPage = page
	oequip.FRectJangbi = jangbi
	oequip.FRectSayoug = sayoug
	oequip.FRectUser = user
	oequip.FRectIp = ipcheck
	oequip.Fequip_code = code
		if ipgocheck = "on" then
			oequip.FRectBuyDateDtStart = fromDate
			oequip.FRectBuyDateDtEnd = toDate
		end if
	
	oequip.getEquipmentList		'class 함수 실행

'변수 선언
Dim equip_gubun, part_code
	equip_gubun = Request("equip_gubun")	'장비구분
	part_code = Request("part_code")		'사용구분

dim totalcurrsum,totaljasan, Alltotaljasan, getAllCurrentValue
	totalcurrsum = 0	'현재 페이지의 구매가를 합계내기 위한 변수.
	totaljasan	 = 0	'현재 페이지의 자산가치를 합계내기 위한 변수.
	Alltotaljasan = 0
%>

<script language="javascript">

	window.onload = function regprint(){
		window.print();
		self.close();
	}
	
</script>

<!-- 리스트 시작 -->
<table width="100%" align="center" border=1 cellspacing="1" bordercolor="black">
	<% if oequip.FResultCount > 0 then %>
	<tr height="25" >
		<td colspan="10">
			검색결과 : <b><%= oequip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oequip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" >
    	<td >장비<br>코드</td>
    	<td >사용자<br>이름</td>
    	<td >사양</td>
    	<td >장비<br>구분</td>
    	<td >사용<br>구분</td>
    	<td >제품명</td>
    	<td >구매일</td>
    	<td >구매<br>원가</td>    	
    	<td >구매가</td>
    	<td >자산<br>가치</td>
    </tr>
<% for i=0 to oequip.FResultCount - 1 %>
	<%
	totalcurrsum = totalcurrsum + oequip.FItemList(i).Fbuy_sum 		'페이지당 합계를 구하기 위해서 현재 페이지의 구매가(Fbuy_sum)를 모두 (fot~next로 loop를 돌려서)더해서 totalcurrsum 변수에 저장
	totaljasan	 = totaljasan + oequip.FItemList(i).GetCurrentvalue	'페이지당 자산가치를 구하기 위해 현재 페이지의 자산가치(GetCurrentvalue)를 모두 for 문으로 돌려서 더하고 totaljasan 변수에 저장
	%>
	<form name=frm_<%= i %> method="post" action="frmdel.asp">	<!-- for문 안에서 i값을 가지고 실행-->
	<input type="hidden" name="idx" value="<%= oequip.FItemList(i).Fidx %>">
	<input type="hidden" name="ssBctId" value="<%= session("ssBctId")%>">
    <tr align="center" >
		<td><%= oequip.FItemList(i).Fequip_code %></td>
		<td><%= oequip.FItemList(i).FusinguserName %>&nbsp;&nbsp;<%= (oequip.FItemList(i).Fusinguserid) %></td>
		<td><%= oequip.FItemList(i).Fdetail_quality1 %><br><%= oequip.FItemList(i).Fdetail_quality2 %></td>
		<td><%= oequip.FItemList(i).Fequip_gubun_name %></td>
		<td><%= oequip.FItemList(i).Fpart_code_name %></td>
		<td><%= oequip.FItemList(i).Fequip_name %></td>
		<td align="center"><%= oequip.FItemList(i).Fbuy_date %></td>
		<td align="right"><%= formatNumber(oequip.FItemList(i).fwonga_cost,0) %></td>		
		<td align="right"><%= formatNumber(oequip.FItemList(i).Fbuy_sum,0) %></td>
		<td align="right">
			<% if oequip.FItemList(i).getCurrentValue<>"" then %>
				<font color="#EE3333"><%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%></font>
			<% else %>
				<%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%>
			<% end if %>
		</td>				
    </tr>   
	</form>
	<% next %>
	
	<tr >
		<td align="center" colspan=7>총계</td>
		<td align="right"><!-- <%= formatNumber(oequip.FItemList(0).Getallcurrentvalue,0) %> --></td>
		<td align="right"><%= formatNumber(totalcurrsum,0) %></td>
		<td align="right"><font color="#EE3333"><%= formatNumber(totaljasan,0) %></font></td>
	
	</tr>
		

	<% else %>
		<tr >
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
</table>
</body>
</html>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->