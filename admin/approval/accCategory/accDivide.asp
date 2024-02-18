<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 계정 카테고리 리스트
' History : 2012.08.07 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/accCategoryCls.asp"-->
<%
Dim clsAcc
Dim acccd
dim accnm, issale10x10, issalepartner, sdivide, sdividedesc
acccd =  requestCheckvar(Request("acccd"),15)
Set clsAcc = new CAccCategory
clsAcc.FACCCD = acccd
clsAcc.fnGetACCDivData
accnm = clsAcc.FACCNM
issale10x10 = clsAcc.FSale10x10
issalepartner = clsAcc.FSalePartner
sdivide = clsAcc.FDivide
sdividedesc = clsAcc.FDividedesc
Set clsAcc = nothing
%>
<form name="frmDiv" method="post" action="procCategory.asp">
<input type="hidden" name="hidM" value="C">
<input type="hidden" name="hidacc" value="<%=acccd%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
    <td bgcolor="<%= adminColor("tabletop") %>">계정과목</td>
    <td bgcolor="#ffffff"><%=accnm%></td>
</tr>
<tr>
    <td bgcolor="<%= adminColor("tabletop") %>">매출처</td>
    <td bgcolor="#ffffff">
    <input type="checkbox" name="isS10" value="1" <%if issale10x10 then%>checked<%end if%>> 10x10
    <input type="checkbox" name="isSP" value="1" <%if issalepartner then%>checked<%end if%>> 제휴
    </td>
</tr>
<tr>
    <td bgcolor="<%= adminColor("tabletop") %>">안분기준</td>
    <td bgcolor="#ffffff">
        <select class="select" name="sdivide">
        <option value="주문번호" <%if sdivide ="주문번호" then%>selected<%end if%>>주문번호 (상품코드 기준 카테고리로 안분)</option>
        <option value="배송건수" <%if sdivide ="배송건수" then%>selected<%end if%>>배송건수</option>
        </select>
    </td>
</tr>
<tr>
    <td bgcolor="<%= adminColor("tabletop") %>">안분기준 설명</td>
    <td bgcolor="#ffffff">
    <textarea name="sdividedesc" style="width:400px;height:50px;" > <%=sdividedesc%></textarea>
   
    </td>
</tr>
</table>
<div style="text-align:center;padding:10px"><input type="button" class="button" value="등록" onClick="javascript:document.frmDiv.submit();" style="width:200px;height:30px;"></div>
</form>