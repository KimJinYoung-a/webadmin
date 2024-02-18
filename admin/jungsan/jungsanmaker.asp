<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim refer
request.ServerVariables("HTTP_REFERER")

dim extsitename, commission
dim startdate, enddate
dim orderserial, userid, buyname
dim totalsum, deasangsum, beasongpay, jungsansum

dim tot_totalsum, tot_deasangsum, tot_beasongpay, tot_jungsansum

extsitename = request("extsitename")
commission = request("commission")
startdate = request("startdate")
enddate = request("enddate")
orderserial = request("orderserial")
userid = request("userid")
buyname = request("buyname")
totalsum = request("totalsum")
deasangsum = request("deasangsum")
beasongpay = request("beasongpay")
jungsansum = request("jungsansum")


dim i
dim totalcount
dim arr_totalsum, arr_deasangsum, arr_beasongpay
dim arr_jungsansum , iorderseriallist

iorderseriallist = Mid(orderserial,2,10240)
iorderseriallist = replace(orderserial,"|","','")

''#########중복체크##########################
dim sqlStr
sqlStr = " select top 1 orderserial from [db_jungsan].[dbo].tbl_etcsite_jungsandetail"
sqlStr = sqlStr + " where orderserial in ('" + iorderseriallist + "')"
rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
%>
	<script >alert('이미정산된 데이터-<%= rsget("orderserial") %>가 있습니다.');</script>
	<script >location.replace('<%= refer %>');</script>
<%
	dbget.close()	:	response.End
end if
rsget.Close
''##########################################

arr_totalsum = split(totalsum,"|")
arr_deasangsum = split(deasangsum,"|")
arr_beasongpay = split(beasongpay,"|")
arr_jungsansum = split(jungsansum,"|")

totalcount = UBound(arr_jungsansum)
for i=1 to totalcount
	tot_totalsum   = tot_totalsum + (arr_totalsum(i)*1.0)
	tot_deasangsum = tot_deasangsum + (arr_deasangsum(i)*1.0)
	tot_beasongpay = tot_beasongpay + (arr_beasongpay(i)*1.0)
	tot_jungsansum = tot_jungsansum + (arr_jungsansum(i)*1.0) ''CLng
next

''2016/01/04 수정.
tot_totalsum   = CLNG(tot_totalsum)
tot_deasangsum   = CLNG(tot_deasangsum)
tot_beasongpay   = CLNG(tot_beasongpay)
tot_jungsansum = CLNG(tot_jungsansum)

%>
<script language='javascript'>
function checkNSubmit(frm){
	var ret;
	ret = confirm('저장하시겠습니까?');

	if (ret) {
		frm.submit();
	}
}
</script>
<table width="760" border="1" cellpadding="0" cellspacing="0" class="a">
<form name="frm" method="post" action="dojungsanmaker.asp">
<input type="hidden" name="extsitename" value="<%= extsitename %>">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="buyname" value="<%= buyname %>">
<input type="hidden" name="totalsum" value="<%= totalsum %>">
<input type="hidden" name="deasangsum" value="<%= deasangsum %>">
<input type="hidden" name="beasongpay" value="<%= beasongpay %>">
<input type="hidden" name="jungsansum" value="<%= jungsansum %>">
<tr>
	<td width="140">사이트명</td>
	<td ><%= extsitename %></td>
</tr>
<tr>
	<td width="140">기간</td>
	<td >
		<input type="text" name="startdate" value="<%= startdate %>">
	 	~
	 	<input type="text" name="enddate" value="<%= enddate %>">
	</td>
</tr>
<tr>
	<td width="140">커미션</td>
	<td ><input type="text" name="commission" value="<%= commission %>"></td>
</tr>
<tr>
	<td width="140">건수</td>
	<td ><input type="text" name="totalcount" value="<%= totalcount %>"></td>
</tr>
<tr>
	<td width="140">총결제금액</td>
	<td ><input type="text" name="tot_totalsum" value="<%= tot_totalsum %>"></td>
</tr>
<tr>
	<td width="140">총배송.포장비</td>
	<td ><input type="text" name="tot_beasongpay" value="<%= tot_beasongpay %>"></td>
</tr>
<tr>
	<td width="140">정산대상금액</td>
	<td ><input type="text" name="tot_deasangsum" value="<%= tot_deasangsum %>"></td>
</tr>
<tr>
	<td width="140">정산금액</td>
	<td ><input type="text" name="tot_jungsansum" value="<%= tot_jungsansum %>"></td>
</tr>
<tr>
	<td width="140">기타사항</td>
	<td ><textarea name="txetc" cols="40" rows="7"></textarea></td>
</tr>
<tr>
	<td colspan="2" align="center">
	<input type="button" value="저장" onclick="checkNSubmit(frm)"> &nbsp;&nbsp;
	<input type="button" value="취소" onclick="history.back();">
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->