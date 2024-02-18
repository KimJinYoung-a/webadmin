<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_ordercls.asp"-->
<%

dim orderserial, oordermaster, oorderdetail

orderserial = RequestCheckvar(request("orderserial"),16)

%>
<script language='javascript'>
location.replace("/academy/lecture/lec_request_detail.asp?orderserial=<%= orderserial %>");
</script>

<%
dbget.close()	:	response.End

set oordermaster = new CLectureFingerOrder
oordermaster.FRectOrderSerial = orderserial
oordermaster.GetFingerOrderList

set oorderdetail = new CLectureFingerOrder
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.GetFingerOrderDetail

dim i, ix

%>
<body topmargin=10 leftmargin=10 marginwidth=0 marginheight=0>

<% if (oordermaster.FItemList(0).IsLectureOrder) then %>
<table width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
    <tr>
        <td>
            <table width="100%" border=0 cellspacing=0 cellpadding=2 class=a bgcolor="FFFFFF">
                <tr align="center">
                    <td width="30">상태</td>
                    <td width="30">진행<br>상태</td>
                    <td width="40">상품<br>코드</td>
                    <td width="40">옵션<br>코드</td>
                   	<td width="50">이미지</td>
                	<td width="90">브랜드ID</td>
                	<td width="250" align="left">상품명(옵션)</font></td>
                	<td width="30">갯수</td>
                	<td width="50" align="right">단가</td>
                	<td width="70">수강생</td>
                	<td width="70">수강생<br>연락처</td>
                	<td width="70">취소<br>구분</td>
                	<td></td>
                </tr>
                <tr>
            		<td height="1" colspan="15" bgcolor="#CCCCCC"></td>
            	</tr>
                <% for ix=0 to oorderdetail.FResultCount-1 %>
                <% if oorderdetail.FItemList(ix).Fitemid <>0 then %>
                <% if oorderdetail.FItemList(ix).CancelStateStr = "취소" then %>
                <tr align="center" bgcolor="#EEEEEE" class="gray">
                <% else %>
                <tr align="center">
                <% end if %>
                    <td><%= oorderdetail.FItemList(ix).CancelStateStr %></td>
                    <td><font color="<%= oorderdetail.FItemList(ix).GetStateColor %>"><%= oorderdetail.FItemList(ix).GetStateName %></font></td>
                	<% if oorderdetail.FItemList(ix).Fisupchebeasong="Y" then %>
                	<td>
                	    <a href="http://thefingers.co.kr/lecture/lecturedetail.asp?lec_idx=<%= oorderdetail.FItemList(ix).Fitemid %>" target="_blank"><font color="red"><%= oorderdetail.FItemList(ix).Fitemid %></font></a>
                    </td>
                    <% else %>
                    <td>
                        <a href="http://thefingers.co.kr/lecture/lecturedetail.asp?lec_idx=<%= oorderdetail.FItemList(ix).Fitemid %>" target="_blank"><%= oorderdetail.FItemList(ix).Fitemid %></a>
                    </td>
                    <% end if %>
                    <td align="center"><%= oorderdetail.FItemList(ix).Fitemoption %></td>
                    <td align="center"><a href="http://thefingers.co.kr/lecture/lecturedetail.asp?lec_idx=<%= oorderdetail.FItemList(ix).Fitemid %>" target="_blank"><img src="<%= oorderdetail.FItemList(ix).FImageSmall %>" width="50" height="50" border="0"></a></td>
                    <td width="90" align="left"><acronym title="<%= oorderdetail.FItemList(ix).Fmakerid %>"><%= Left(oorderdetail.FItemList(ix).Fmakerid,12) %></acronym></td>
                	<td width="250" align="left">
                	    <acronym title="<%= oorderdetail.FItemList(ix).FItemName %>"><%= Left(oorderdetail.FItemList(ix).FItemName,35) %></acronym>
                	    <br>
                	    <font color="blue"><%= oorderdetail.FItemList(ix).FItemoptionName %></font>
                	</td>

                	<% if oorderdetail.FItemList(ix).FItemNo > 1 then %>
                	<td><b><font color="blue"><%= oorderdetail.FItemList(ix).FItemNo %></font></b></td>
                	<% elseif oorderdetail.FItemList(ix).FItemNo < 1 then %>
                	<td><b><font color="red"><%= oorderdetail.FItemList(ix).FItemNo %></font></b></td>
                	<% else %>
                	<td><font color="blue"><%= oorderdetail.FItemList(ix).FItemNo %></font></td>
                	<% end if %>

                   	<% if oorderdetail.FItemList(ix).FItemNo < 1 then %>
                   	<td align="right"><font color="red"><%= FormatNumber(oorderdetail.FItemList(ix).Fitemcost,0) %></font></td>
                   	<% else %>
                   	<td align="right"><font color="blue"><%= FormatNumber(oorderdetail.FItemList(ix).Fitemcost,0) %></font></td>
                   	<% end if %>


                	<td><%= oorderdetail.FItemList(ix).Fentryname %></td>
                	<td><%= oorderdetail.FItemList(ix).Fentryhp %></td>
                	<td><%= oorderdetail.FItemList(ix).Fcancelyn %></td>
                	<td></td>
                </tr>
                <tr>
            		<td height="1" colspan="15" bgcolor="#CCCCCC"></td>
            	</tr>
                <% end if %>
                <% next %>
            </table>
        </td>
    </tr>
</table>
<% else %>
not Lecture Order - 관리자 문의 요망
<% end if %>

<%
set oordermaster = Nothing
set oorderdetail = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->