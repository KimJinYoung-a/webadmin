<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터 년별 마일리지소멸
' History : 서동석 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_mileagecls.asp" -->

<%
dim userid, yyyymmdd
userid = request("userid")
yyyymmdd = requestCheckvar(request("yyyymmdd"),10)

if (yyyymmdd="") then
    yyyymmdd=Left(now(),4) & "-12-31"
end if

''현재 마일리지
dim myMileage
set myMileage = new CCSCenterMileage
myMileage.FRectUserID = userid
if (userid<>"") then
    myMileage.getUserCurrentMileage
end if

''만료예정 마일리지 년도별 리스트
dim oExpireMile
set oExpireMile = new CCSCenterMileage
oExpireMile.FRectUserid = userid
''''해당Expire 내역만 보여줄 경우
''oExpireMile.FRectExpireDate = yyyymmdd

if (userid<>"") then
    oExpireMile.getNextExpireMileageYearList
end if


''만료예정  마일리지 합계
dim oExpireMileTotal
set oExpireMileTotal = new CCSCenterMileage
oExpireMileTotal.FRectUserid = userid
oExpireMileTotal.FRectExpireDate = yyyymmdd
if (userid<>"") then
    oExpireMileTotal.getNextExpireMileageSum
end if

dim i
dim Tot_GainMileage, Tot_YearMaySpendMileage, Tot_YearMayRemainMileage, Tot_realExpiredMileage
%>
<style>
.black12px {font-family: 굴림; FONT-SIZE: 12px; COLOR: #000000;  TEXT-DECORATION: none; font-weight: bold;}
</style>
<script language='javascript'>
function research(frm){
    if (frm.userid.value.length<1){
        alert('아이디를 입력하세요.');
        frm.userid.focus();
        return;
    }
    frm.submit();
}
</script>
<table width="600" border="0" cellpadding="2" cellspacing="1" bgcolor="d3d3d3" class="a">
<form name="frmresearch" >
<input type="hidden" name="yyyymmdd" value="<%= yyyymmdd %>">
<tr bgcolor="#FFFFFF"  height="26">
    <td>아이디 : <input type="text" name="userid" value="<%= userid %>" size="16" maxlength="32" class="text">
    <input type="button" value="재검색" onclick="research(frmresearch);" class="button">
    </td>
</tr>
</form>
</table>
<table width="600" border="0" cellpadding="2" cellspacing="1" bgcolor="d3d3d3" class="a">
<tr bgcolor="#DDDDFF" align="center" height="26">
    <td width="120">적립년도</td>
    <td width="120">적립마일리지</td>
    <td width="120">사용</td>
    <td width="120">소멸</td>
    <td width="120">잔여</td>
    <td width="120">소멸예정일</td>
</tr>
<% if oExpireMile.FResultCount>0 then %>
    <% for i=0 to oExpireMile.FResultCount-1 %>
    <%
        Tot_GainMileage           = Tot_GainMileage + oExpireMile.FItemList(i).getGainMileage
        Tot_YearMaySpendMileage   = Tot_YearMaySpendMileage + oExpireMile.FItemList(i).getYearMaySpendMileage
        Tot_YearMayRemainMileage  = Tot_YearMayRemainMileage + oExpireMile.FItemList(i).getYearMayRemainMileage
        Tot_realExpiredMileage    = Tot_realExpiredMileage + oExpireMile.FItemList(i).FrealExpiredMileage
    %>
    <tr bgcolor="#FFFFFF" align="center" height="26" >
        <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> ><%= oExpireMile.FItemList(i).FRegYear %></td>
        <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> ><%= FormatNumber(oExpireMile.FItemList(i).getGainMileage ,0) %></td>
        <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> ><%= FormatNumber(oExpireMile.FItemList(i).getYearMaySpendMileage,0) %></td>
        <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> ><%= FormatNumber(oExpireMile.FItemList(i).FrealExpiredMileage,0) %></td>
        <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> ><%= FormatNumber(oExpireMile.FItemList(i).getYearMayRemainMileage,0) %></td>
        <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> ><%= oExpireMile.FItemList(i).FExpiredate %></td>
    </tr>
    <% next %>

    <%
    '' 현재 마일리지에서 역으로 계산.
    if (oExpireMile.FResultCount>0) and (oExpireMile.FRectExpireDate="") then %>
    <tr bgcolor="#FFFFFF" align="center" height="26">
        <td><%= CStr(Year(Now) - 4)&" ~ " %></td>
        <td><%= FormatNumber(myMileage.FOneItem.FJumunMileage + myMileage.FOneItem.FFlowerJumunmileage + myMileage.FOneItem.FAcademymileage + myMileage.FOneItem.FBonusMileage - Tot_GainMileage ,0) %>
        <td><%= FormatNumber(myMileage.FOneItem.FSpendMileage - Tot_YearMaySpendMileage,0) %></td>
        <!-- <td></td> -->
		<td><%= FormatNumber(myMileage.FOneItem.FrealExpiredMileage - Tot_realExpiredMileage,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.getCurrentMileage - Tot_YearMayRemainMileage,0) %></td>
        <td></td>
    </tr>
    <tr height="1" bgcolor="#FFFFFF">
        <td colspan="6"></td>
    </tr>
    <tr bgcolor="#FFFFFF" align="center" height="26">
        <td>합계</td>
        <td><%= FormatNumber(myMileage.FOneItem.FJumunMileage + myMileage.FOneItem.FFlowerJumunmileage + myMileage.FOneItem.FAcademymileage + myMileage.FOneItem.FBonusMileage,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.FSpendMileage ,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.FrealExpiredMileage ,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.getCurrentMileage,0) %></td>
        <td>&nbsp;</td>
    </tr>
    <% end if %>
<% else %>
    <tr bgcolor="#FFFFFF" align="center" height="26">
        <td colspan="6" align="center">[ <%= yyyymmdd %> 소멸 대상 내역이 없습니다.]</td>
    </tr>
<% end if %>
</table>
<br>
<% if myMileage.FResultCount>0 then %>
<table width="600" border="0" cellpadding="2" cellspacing="1" bgcolor="d3d3d3" class="a">
<tr bgcolor="#FFFFFF" align="center" height="26">
    <td width="120">현재</td>
    <td width="120"><%= FormatNumber(myMileage.FOneItem.FJumunMileage + myMileage.FOneItem.FFlowerJumunmileage + myMileage.FOneItem.FAcademymileage + myMileage.FOneItem.FBonusmileage,0) %></td>
    <td width="120"><%= FormatNumber(myMileage.FOneItem.FSpendMileage ,0) %></td>
    <td width="120"><%= FormatNumber(myMileage.FOneItem.getCurrentMileage,0) %></td>
    <td width="120">&nbsp;</td>
</tr>
<% end if %>
</table>
<br>
<table width="600" border="0" cellpadding="2" cellspacing="1" bgcolor="d3d3d3" class="a">
<tr bgcolor="#FFFFFF" align="center" height="26">
    <td align="center" >
    <font style="font-family: 돋움; COLOR: #333333; FONT-SIZE: 13px; font-weight: bold;">
    <%= oExpireMileTotal.FOneItem.getKorExpireDateStr %> 소멸 대상 마일리지 : <font color="red"><%= FormatNumber(oExpireMileTotal.FOneItem.getMayExpireTotal,0) %> </font> Point
    </font>
    </td>
</tr>
</table>
<br>
<table width="600" border="0" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="a">
<tr>
    <td>
        ※ 마일리지는 부여된 순서로 사용되며 적립 후 60개월 내 미사용 시 60개월이 되는 당해년도 12월 31일 자동 소멸됩니다.<br>
    </td>
</tr>
<tr>
    <td>예) <%= CLng(Left(yyyymmdd,4)-5) %> 년 적립 마일리지 4500 / 사용 마일리지 4000 / 잔여 마일리지 500 인 경우 500 포인트는 <%= Left(yyyymmdd,4) %>년 12월 31일 자동 소멸됩니다.</td>
</tr>
</table>
<%
set myMileage = Nothing
set oExpireMile = Nothing
set oExpireMileTotal = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
