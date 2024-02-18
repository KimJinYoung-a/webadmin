<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim idx
idx = request("idx")

dim ojumunmaster
set ojumunmaster = new COrderSheet
ojumunmaster.FRectIdx = idx
ojumunmaster.GetOneOrderSheetMaster

dim ojumundetail
set ojumundetail= new COrderSheet
ojumundetail.FRectIdx = idx
ojumundetail.GetOrderSheetDetail

'// 상품 매입속성 체크
dim maeipItemExist : maeipItemExist = False
dim witakItemExist : witakItemExist = False
for i=0 to ojumundetail.FResultCount-1
	if ojumundetail.FItemList(i).FItemDefaultMwDiv = "M" then
		maeipItemExist = True
	end if

	if ojumundetail.FItemList(i).FItemDefaultMwDiv = "W" then
		witakItemExist = True
	end if
next


dim oipchul
set oipchul = new CIpChulStorage
oipchul.FCurrPage = 1
oipchul.Fpagesize=5
oipchul.FRectCodeGubun = "ST"  ''입고
oipchul.FRectSocID = ojumunmaster.FOneItem.Ftargetid
oipchul.GetIpChulgoList

dim i

dim BasicMonth, IsExpireEdit
BasicMonth = CStr(DateSerial(Year(now()),Month(now())-1,1))

%>
<script language='javascript'>
function ipgoproc(frm){

//	if ((!frm.mode[0].checked)&&(!frm.mode[1].checked)){
//		alert('구분을 선택하세요.');
//		frm.mode[0].focus();
//		return;
//	}

	if ((!frm.divcode[0].checked)&&(!frm.divcode[1].checked)){
		alert('매입 구분을 선택하세요.');
		frm.divcode[0].focus();
		return;
	}

	<% if (maeipItemExist = True) then %>
	if (frm.divcode[1].checked == true) {
		if (confirm("--------------------------\n!!! 매입속성 불일치 !!!\n--------------------------\n\n상품중에 매입속성 상품이 있습니다.\n\n계속 진행하시겠습니까?") != true) {
			return;
		}
	}
	<% end if %>

	<% if (witakItemExist = True) then %>
	if (frm.divcode[0].checked == true) {
		if (confirm("--------------------------\n!!! 매입속성 불일치 !!!\n--------------------------\n\n상품중에 위탁속성 상품이 있습니다.\n\n계속 진행하시겠습니까?") != true) {
			return;
		}
	}
	<% end if %>

	if (frm.scheduledt.value.length<1){
		alert('거래명세서일자를 입력하세요.');
		frm.scheduledt.focus();
		return;
	}

	if (frm.ipgodate.value.length<1){
		alert('입고일자 입력하세요.');
		frm.ipgodate.focus();
		return;
	}

	if (frm.ipgodate.value<'<%= BasicMonth %>'){
		alert('입고일이 두달 이전 날짜로는 추가/수정/삭제 불가 합니다.');
		return;
	}

	var ret = confirm('진행 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

	<tr height="25" bgcolor="<%= adminColor("gray") %>">
		<td colspan="15">
			<img src="/images/icon_star.gif" align="absbottom">
		    <font color="red"><strong>입고확정</strong></font>
		</td>
	</tr>

    <form name=frm method=post action="shopjumun_process.asp">
    <input type="hidden" name="masteridx" value="<%= ojumunmaster.FOneItem.Fidx %>">
	<input type="hidden" name="ojbaljucode" value="<%= ojumunmaster.FOneItem.Fbaljucode %>">
    <input type="hidden" name="finishuser" value="<%= session("ssBctId") %>">
    <input type="hidden" name="finishname" value="<%= session("ssBctCname") %>">
    <input type="hidden" name="targetid" value="<%= ojumunmaster.FOneItem.Ftargetid %>">
    <input type="hidden" name="targetname" value="<%= ojumunmaster.FOneItem.Ftargetname %>">
    <input type="hidden" name="checkusersn" value="<%= ojumunmaster.FOneItem.Fcheckusersn %>">
    <input type="hidden" name="rackipgousersn" value="<%= ojumunmaster.FOneItem.Frackipgousersn %>">
    <tr bgcolor="#FFFFFF">
    	<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
    	<td colspan="3"><%= ojumunmaster.FOneItem.Ftargetid %>(<%= ojumunmaster.FOneItem.Ftargetname %>)</td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td align="center" bgcolor="<%= adminColor("tabletop") %>">매입구분</td>
    	<td>
    	    <% if ojumunmaster.FOneItem.Fdivcode="301" then %>
        		<input type="radio" name="divcode" value="001" checked >매입
        		<input type="radio" name="divcode" value="002">위탁
        	<% elseif ojumunmaster.FOneItem.Fdivcode="302" then %>
        		<input type="radio" name="divcode" value="001">매입
        		<input type="radio" name="divcode" value="002" checked >위탁
        	<% end if %>
    	</td>
    	<td align="center" bgcolor="<%= adminColor("tabletop") %>">구분</td>
    	<td>
    	    <input type="radio" name="mode" value="savennext" checked ><b>입고내역에 입력후 상태변경</b>
		    <!-- <input type="radio" name="mode" value="justnext"><font color="#CCCCCC">상태만 변경</font> -->
    	</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td align="center" bgcolor="<%= adminColor("tabletop") %>">거래명세서일자</td>
    	<td colspan="3">
    	    <!--
        	<input type=text class="text" name="scheduledt" value="<%= Left(ojumunmaster.FOneItem.getScheduledate,10) %>" size=11 readonly ><a href="javascript:calendarOpen(frm.scheduledt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (거래명세서 날짜를 입력하세요.)
        	-->
        	<input type=text class="text" name="scheduledt" value="" size=11 readonly ><a href="javascript:calendarOpen(frm.scheduledt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (거래명세서 날짜를 입력하세요.)
    	</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td align="center" bgcolor="<%= adminColor("tabletop") %>">입고일자</td>
    	<td colspan="3">
        	<input type=text class="text" name="ipgodate" value="<%= Left(now,10) %>" size=11 readonly ><a href="javascript:calendarOpen(frm.ipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (실 입고 날짜를 입력하세요. - 재고와 관련)
    	</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td align="center" bgcolor="<%= adminColor("tabletop") %>">총 소비가</td>
    	<td><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsellcash,0) %></td>
    	<td align="center" bgcolor="<%= adminColor("tabletop") %>">총 매입가</td>
    	<td><%= FormatNumber(ojumunmaster.FOneItem.Ftotalbuycash,0) %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td align="center" bgcolor="<%= adminColor("tabletop") %>">비고</td>
    	<td colspan="3">
    		<textarea name=comment cols=60 rows=4></textarea>
    	</td>
    </tr>

    <tr height="25" bgcolor="<%= adminColor("gray") %>">
		<td colspan="15" align="center">
			<input type="button" class="button" value="입고 확정" onClick="ipgoproc(frm)">
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

	<tr height="25" bgcolor="<%= adminColor("gray") %>">
		<td colspan="15">
			<img src="/images/icon_star.gif" align="absbottom">
		    <b><%= ojumunmaster.FOneItem.Ftargetid %>(<%= ojumunmaster.FOneItem.Ftargetname %>)의 최근 5건 입고 리스트</b>
		</td>
	</tr>

    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="80">입고코드</td>
    	<td width="60">처리자</td>
    	<td>예정일</td>
    	<td>입고일</td>
    	<td width="100">소비자가</td>
    	<td width="100">매입가</td>
    	<td width="50">구분</td>
    	<td width="50">마진</td>
    </tr>
    <% if oipchul.FResultCount >0 then %>
    <% for i=0 to oipchul.FResultcount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= oipchul.FItemList(i).Fcode %></a></td>
    	<td><%= oipchul.FItemList(i).Fchargename %></td>
    	<td><font color="#777777"><%= Left(oipchul.FItemList(i).Fscheduledt,10) %></font></td>
    	<td><%= Left(oipchul.FItemList(i).Fexecutedt,10) %></td>
    	<td align="right">
        	<% if ojumunmaster.FOneItem.Ftotalsellcash=oipchul.FItemList(i).Ftotalsellcash then %>
        		<b><%= FormatNumber(oipchul.FItemList(i).Ftotalsellcash,0) %></b>
        	<% else %>
        		<%= FormatNumber(oipchul.FItemList(i).Ftotalsellcash,0) %>
        	<% end if %>
    	</td>
    	<td align="right">
        	<% if ojumunmaster.FOneItem.Ftotalbuycash=oipchul.FItemList(i).Ftotalsuplycash then %>
        		<b><%= FormatNumber(oipchul.FItemList(i).Ftotalsuplycash,0) %></b>
        	<% else %>
        		<%= FormatNumber(oipchul.FItemList(i).Ftotalsuplycash,0) %>
        	<% end if %>
    	</td>
    	<td><font color="<%= oipchul.FItemList(i).GetDivCodeColor %>"><%= oipchul.FItemList(i).GetDivCodeName %></font></td>
    	<td align="right">
        	<% if oipchul.FItemList(i).Ftotalsellcash<>0 then %>
        	  <%= 100-CLng(oipchul.FItemList(i).Ftotalsuplycash/oipchul.FItemList(i).Ftotalsellcash*100*100)/100 %>%
        	<% end if %>
    	</td>
    </tr>
    <% next %>
    <% end if %>
</table>


<%
set ojumunmaster = Nothing
set ojumundetail = Nothing
set oipchul = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
