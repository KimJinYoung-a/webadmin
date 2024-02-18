<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<%
dim mode, research
dim supplyCtrtSeq, strDate, endDate
dim iParkURL

mode        = request("mode")
research    = request("research")
supplyCtrtSeq = request("supplyCtrtSeq")

IF (application("Svr_Info")="Dev") THEN
    iParkURL = "http://sptest.interpark.com"
ELSE
    iParkURL = "http://www.interpark.com"
END IF

iParkURL = iParkURL + "/order/OrderClmExInterface.do"

dim iParams

iParams = "_method=" & mode & "&entrId=GODO&supplyEntrNo=3000010614&supplyCtrtSeq=" & supplyCtrtSeq & "&strDate=" & strDate & "000000&endDate=" & endDate & "235959"

response.write iParkURL & "?" & iParams
dim replyXML
if (mode<>"") and (supplyCtrtSeq<>"") then
    replyXML = SendReq(iParkURL, iParams)
elseif (research<>"") then
    response.write "<script language='javascript'>alert('검색구분, 몰구분을 선택 하세요.');</script>"
end if
%>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr >
		<td class="a">
		    조회 구분 :
    		<select name="mode">
    		<option value="">선택
    		<option value="orderListEntrNoForComm" <%= ChkIIF(mode="orderListEntrNoForComm","selected","") %> >주문조회(발주 미확인)
    		<option value="orderListDelvByEntrNoForComm" <%= ChkIIF(mode="orderListDelvByEntrNoForComm","selected","") %> >주문조회(발주 확인)
    		</select>
    		&nbsp;&nbsp;
    		몰 구분 : 
    		<select name="supplyCtrtSeq">
    		<option value="">선택
    		<option value="2" <%= ChkIIF(supplyCtrtSeq="2","selected","") %>>리빙
    		<option value="3" <%= ChkIIF(supplyCtrtSeq="3","selected","") %>>잡화
    		<option value="4" <%= ChkIIF(supplyCtrtSeq="4","selected","") %>>의류
    		</select>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr>
    <td>
        <textarea cols="60" rows="10"><%= replyXML %></textarea>
    </td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->