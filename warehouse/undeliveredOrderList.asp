<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [LOG]물류센터>>[CENTER]메인 > 미배송 주문 목록
' History : 2020.11.20 허진원
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/undeliveredOrderCls.asp"-->
<%
Dim yyyymmdd, yyyymmdd2, mode, ordstat, danjong, i

yyyymmdd	= requestCheckvar(request("yyyymmdd"),10)
yyyymmdd2	= requestCheckvar(request("yyyymmdd2"),10)
mode	    = requestCheckvar(request("mode"),2)
ordstat     = requestCheckvar(request("ordstat"),1)
danjong     = requestCheckvar(request("danjong"),1)
if yyyymmdd 	= "" then yyyymmdd=dateadd("d",-3,date())     'D+3
if yyyymmdd2 	= "" then yyyymmdd2=dateadd("d",-3,date())     'D+3
if mode 		= "" then mode="OD"   'OD:주문일기준, IT:상품기준

dim oOrder
set oOrder = new COrder
    oOrder.FRectDate = yyyymmdd
    oOrder.FRectDate2 = yyyymmdd2
    oOrder.FRectMode = mode
    oOrder.FRectStat = ordstat
    oOrder.FRectDanjong = danjong
    oOrder.OrderList()
%>
<script type="text/javascript" src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
th {
    position: sticky; top: 0; background:<%= adminColor("tabletop") %>;
    border-bottom:1px solid <%= adminColor("tablebg") %>;
}
.txtct {text-align:center;}
.txtrt {text-align:right;}
</style>
<script type="text/javascript">
function SubmitFrm() {
    var frm = document.frm;
    frm.submit();
}
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
    	출고기준일 :
        <input id="baseDate" name="yyyymmdd" value="<%=yyyymmdd%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="baseDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
        ~
        <input id="baseDate2" name="yyyymmdd2" value="<%=yyyymmdd2%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="baseDate2_trigger" border="0" style="cursor:pointer" align="absmiddle" />
        <script type="text/javascript">
            var CAL_baseDate = new Calendar({
                inputField : "baseDate",
                trigger    : "baseDate_trigger",
                bottomBar: true,
                dateFormat: "%Y-%m-%d",
                onSelect: function() {
                    this.hide();
                }
            });
            var CAL_baseDate2 = new Calendar({
                inputField : "baseDate2",
                trigger    : "baseDate2_trigger",
                bottomBar: true,
                dateFormat: "%Y-%m-%d",
                onSelect: function() {
                    this.hide();
                }
            });
        </script>
        /
        표시방법 :
        <label><input type="radio" name="mode" value="OD" <%=chkIIF(mode="OD","checked","")%> /> 주문번호 포함</label>
        <label><input type="radio" name="mode" value="IT" <%=chkIIF(mode="IT","checked","")%> /> 상품 기준</label>
        /
        출고상태 :
        <select name="ordstat">
            <option value="" <%=chkIIF(ordstat="","selected","")%>>전체</option>
            <option value="B" <%=chkIIF(ordstat="B","selected","")%>>미출고지시</option>
            <option value="C" <%=chkIIF(ordstat="C","selected","")%>>미배송</option>
        </select>
        /
        단종여부 :
        <select name="danjong">
            <option value="" <%=chkIIF(danjong="","selected","")%>>전체</option>
            <option value="N" <%=chkIIF(danjong="N","selected","")%>>생산중</option>
            <option value="S" <%=chkIIF(danjong="S","selected","")%>>일시품절</option>
            <option value="M" <%=chkIIF(danjong="M","selected","")%>>MD품절</option>
            <option value="Y" <%=chkIIF(danjong="Y","selected","")%>>단종</option>
        </select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="SubmitFrm();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br />

* 최대 3천건까지만 검색됩니다.

<br />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<thead>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="<%=chkIIF(mode="OD","14","12")%>">
		검색결과 : <b><%= oOrder.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <% if mode="OD" then %>
    <th>주문번호</th>
    <th>결제일</th>
    <% end if %>
    <th>브랜드ID</th>
    <th>상품코드</th>
    <th>옵션코드</th>
    <th>상품명</th>
    <th>주문수량</th>
    <th>재고</th>
    <th>판매가</th>
    <th>매입가</th>
    <th>기발주수량</th>
    <th>단종여부</th>
    <th>랙코드</th>
    <th>출고상태</th>
</tr>
</thead>
<tbody>
<%
    if oOrder.FResultCount>0 then
        for i=0 to oOrder.FResultCount - 1
%>
<tr bgcolor="#FFFFFF">
    <% if mode="OD" then %>
    <td class="txtct"><%= oOrder.FItemList(i).Forderserial %></td>
    <td class="txtct"><%= left(oOrder.FItemList(i).Fipkumdate,10) %></td>
    <% end if %>
    <td><%= oOrder.FItemList(i).Fmakerid %></td>
    <td class="txtct"><%= oOrder.FItemList(i).Fitemid %></td>
    <td class="txtct"><%= oOrder.FItemList(i).Fitemoption %></td>
    <td><%= oOrder.FItemList(i).Fitemname & chkIIF(oOrder.FItemList(i).Foptionname<>""," ("&oOrder.FItemList(i).Foptionname&")","") %></td>
    <td class="txtrt"><%= oOrder.FItemList(i).Ficnt %></td>
    <td class="txtrt"><%= oOrder.FItemList(i).Frealstock %></td>
    <td class="txtrt"><%= FormatNumber(oOrder.FItemList(i).Fsellcash,0) %></td>
    <td class="txtrt"><%= FormatNumber(oOrder.FItemList(i).Fbuycash,0) %></td>
    <td class="txtrt"><%= oOrder.FItemList(i).Fpreordernofix %></td>
    <td class="txtct"><%= oOrder.FItemList(i).FisDanjong %></td>
    <td class="txtct"><%= oOrder.FItemList(i).FrackcodeByOption & chkIIF(oOrder.FItemList(i).FsubRackcodeByOption<>"","<br />("&oOrder.FItemList(i).FsubRackcodeByOption&")","") %></td>
    <td class="txtct"><%= oOrder.FItemList(i).Fgubun %></td>
</tr>
<%
            '버퍼 플러싱
            if ((i+1) mod 500)=0 then
                response.Flush()
            end if
        next
    else
%>
<tr bgcolor="#FFFFFF">
    <td colspan="<%=chkIIF(mode="OD","14","12")%>" style="text-align:center;">검색결과가 없습니다.</td>
</tr>
<%  end if %>
</tbody>
</table>
<%
    set oOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
