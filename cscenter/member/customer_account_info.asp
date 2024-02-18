<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터 >>[CS]고객환불계좌정보
' History : 2020.12.01 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/member/customercls.asp"-->

<%
Dim oaccount,i,page,userid
	userid = requestcheckvar(request("userid"),32)
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
    page = requestcheckvar(getNumeric(request("page")),10)

if page = "" then page = 1

set oaccount = new CUserInfo
	oaccount.FPageSize = 20
	oaccount.FCurrPage = page
	oaccount.frectuserid = userid

    if userid<>"" then
	    oaccount.GetUser_accountinfo_List()
    end if
%>

<script type="text/javascript">

function getsubmit(page){
    frm.page.value=page;
    frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
    <td align="left">
        고객아이디(비회원은 주문번호) : <input type="text" name="userid" value="<%= userid%>" size="12" onKeyPress="if(window.event.keyCode==13) getsubmit('1');"> 			
    </td>	
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="검색" onClick="getsubmit('1');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
	
    </td>
</tr>
</table>
<!-- 검색 끝 -->

<br>
		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
        <% if userid="" then %>
            <font color="red">고객아이디(비회원은 주문번호)를 입력해 주세요.</font>
        <% end if %>
    </td>
    <td align="right">	
    </td>
</tr>
</table>
</form>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="5">
        검색결과 : <b><%= oaccount.FTotalCount %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>고객아이디<br>(비회원은 주문번호)</td>
    <td>은행명</td>	
    <td>계좌번호</td>	
    <td>계좌명</td>
    <td>비고</td>
</tr>
<% if oaccount.FresultCount>0 then %>
<% for i=0 to oaccount.FresultCount-1 %>

<tr align="center" bgcolor="#FFFFFF">
    <td>
        <%= oaccount.FItemList(i).fuserid %>
    </td>		
    <td>
        <%= oaccount.FItemList(i).frebankname %>
    </td>	
    <td align="left">
        <%= oaccount.FItemList(i).fencaccount %>
    </td>
    <td>
        <%= oaccount.FItemList(i).frebankownername %>
    </td>
    <td></td>
</tr>   

<% next %>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="5" align="center" class="page_link">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>

</table>

<%
set oaccount = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

