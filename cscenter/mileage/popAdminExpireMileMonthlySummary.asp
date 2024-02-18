<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터 월별 마일리지소멸
' History : 서동석 생성(년별 마일리지 소멸)
'           2023.07.21 한용민 수정(월별 마일리지 소멸로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_mileagecls.asp" -->
<%
dim userid, yyyymmdd, myMileage, oExpireMile, oExpireMileTotal, currentdate, i, menupos
dim Tot_GainMileage, Tot_MonthMaySpendMileage, Tot_MonthMayRemainMileage, Tot_realExpiredMileage
    userid = requestcheckvar(request("userid"),32)
    yyyymmdd = requestCheckvar(request("yyyymmdd"),10)
    menupos = requestCheckvar(getNumeric(request("menupos")),10)

currentdate=date()

if (yyyymmdd="") then
    ' 이번달말일
    yyyymmdd=dateadd("d",-1,dateserial(year(dateadd("m",+1,currentdate)),month(dateadd("m",+1,currentdate)),"01"))
end if

' 현재 마일리지
set myMileage = new CCSCenterMileage
    myMileage.FRectUserID = userid
    if (userid<>"") then
        myMileage.getUserCurrentMileage
    end if

' 만료예정 마일리지 년도별 리스트
set oExpireMile = new CCSCenterMileage
    oExpireMile.FRectUserid = userid
    ' 해당Expire 내역만 보여줄 경우
    ' oExpireMile.FRectExpireDate = yyyymmdd

    if (userid<>"") then
        oExpireMile.getNextExpireMileageMonthlyList
    end if

''만료예정  마일리지 합계
set oExpireMileTotal = new CCSCenterMileage
    oExpireMileTotal.FRectUserid = userid
    oExpireMileTotal.FRectExpireDate = yyyymmdd

    if (userid<>"") then
        oExpireMileTotal.getNextExpireMileageMonthlySum
    end if

%>
<style>
.black12px {font-family: 굴림; FONT-SIZE: 12px; COLOR: #000000;  TEXT-DECORATION: none; font-weight: bold;}
</style>
<script type='text/javascript'>

function research(frm){
    if (frm.userid.value.length<1){
        alert('아이디를 입력하세요.');
        frm.userid.focus();
        return;
    }
    frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frmresearch" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="yyyymmdd" value="<%= yyyymmdd %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
    <td align="left">
        아이디 : <input type="text" name="userid" value="<%= userid %>" size="16" maxlength="32" class="text">
    </td>
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="검색" onClick="research(frmresearch);">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">

    </td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
        ※ 마일리지는 부여된 순서로 사용되며 적립 후, 60개월 내 미사용 시 60개월이 되는 월 말일에 자동 소멸됩니다.
        <br>예) <%= left(dateadd("m",-60,DateSerial(Year(yyyymmdd), month(yyyymmdd), day(yyyymmdd))),7) %>월 
        적립 마일리지 4500 / 사용 마일리지 4000 / 잔여 마일리지 500 인 경우 500 포인트는 <%= yyyymmdd %>일 자동 소멸됩니다.
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>적립날짜</td>
    <td>적립마일리지</td>
    <td>사용</td>
    <td>소멸</td>
    <td>잔여</td>
    <td>소멸예정일</td>
</tr>
<% if oExpireMile.FResultCount>0 then %>
<%
for i=0 to oExpireMile.FResultCount-1

Tot_GainMileage           = Tot_GainMileage + oExpireMile.FItemList(i).getGainMileage
Tot_MonthMaySpendMileage   = Tot_MonthMaySpendMileage + oExpireMile.FItemList(i).getMonthlyMaySpendMileage
Tot_MonthMayRemainMileage  = Tot_MonthMayRemainMileage + oExpireMile.FItemList(i).getMonthlyMayRemainMileage
Tot_realExpiredMileage    = Tot_realExpiredMileage + oExpireMile.FItemList(i).FrealExpiredMileage
%>
<tr align="center" bgcolor="#FFFFFF">
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%
        ' 2018년8월 마일리지 부터 월단위 마일리지 소멸을 시작함. 기존 데이터는 년단위 데이터임.
        if datediff("d",oExpireMile.FItemList(i).Fregmonth&"-01","2018-08-01")>0 then
        %>
            <%= left(oExpireMile.FItemList(i).Fregmonth,4) %>
        <% else %>
            <%= oExpireMile.FItemList(i).Fregmonth %>
        <% end if %>
    </td>
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%= FormatNumber(oExpireMile.FItemList(i).getGainMileage ,0) %>
    </td>
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%= FormatNumber(oExpireMile.FItemList(i).getMonthlyMaySpendMileage,0) %>
    </td>
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%= FormatNumber(oExpireMile.FItemList(i).FrealExpiredMileage,0) %>
    </td>
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%= FormatNumber(oExpireMile.FItemList(i).getMonthlyMayRemainMileage,0) %>
    </td>
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%= oExpireMile.FItemList(i).FExpiredate %>
    </td>
</tr>   
<% next %>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="16" align="center" class="page_link">[ <%= yyyymmdd %> 소멸 대상 내역이 없습니다.]</td>
    </tr>
<% end if %>

<%
' 현재 마일리지에서 역으로 계산.
if (oExpireMile.FResultCount>0) and (oExpireMile.FRectExpireDate="") then
%>
    <tr bgcolor="#FFFFFF" align="center" height="26">
        <td><%= left(dateadd("m",-59,DateSerial(Year(yyyymmdd), month(yyyymmdd), day(yyyymmdd))),7) & "월~" %></td>
        <td><%= FormatNumber(myMileage.FOneItem.FJumunMileage + myMileage.FOneItem.FFlowerJumunmileage + myMileage.FOneItem.FAcademymileage + myMileage.FOneItem.FBonusMileage - Tot_GainMileage ,0) %>
        <td><%= FormatNumber(myMileage.FOneItem.FSpendMileage - Tot_MonthMaySpendMileage,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.FrealExpiredMileage - Tot_realExpiredMileage,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.getCurrentMileage - Tot_MonthMayRemainMileage,0) %></td>
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
    <% if myMileage.FResultCount>0 then %>
        <tr height="1" bgcolor="#FFFFFF">
            <td colspan="6"></td>
        </tr>
        <tr bgcolor="#FFFFFF" align="center" height="26">
            <td>현재</td>
            <td><%= FormatNumber(myMileage.FOneItem.FJumunMileage + myMileage.FOneItem.FFlowerJumunmileage + myMileage.FOneItem.FAcademymileage + myMileage.FOneItem.FBonusmileage,0) %></td>
            <td><%= FormatNumber(myMileage.FOneItem.FSpendMileage ,0) %></td>
            <td>&nbsp;</td>
            <td><%= FormatNumber(myMileage.FOneItem.getCurrentMileage,0) %></td>
            <td>&nbsp;</td>
        </tr>
    <% end if %>
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

<%
set myMileage = Nothing
set oExpireMile = Nothing
set oExpireMileTotal = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
