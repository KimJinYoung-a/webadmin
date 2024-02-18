<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  송장목록 다운로드
' History : 이상구 생성
'			2021.05.25 한용민 수정(한진택배 추가)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
dim yyyymmdd, strSql, arrRS, i, songjangno, songjangdiv,printsongjangdivname
	songjangdiv = requestCheckvar(getNumeric(request("songjangdiv")),30)
	yyyymmdd = requestCheckvar(request("yyyymmdd"),10)

songjangdiv=trim(songjangdiv)

if (yyyymmdd <> "") then
	strSql = " select SONGJANGNO, div_cd as songjangdivname"
	strSql = strSql + " from [db_aLogistics].[dbo].[tbl_Logistics_songjang_log] with (nolock)"
	strSql = strSql + " where DateDiff(d, REGDATE, '" & yyyymmdd & "') = 0 "

	if songjangdiv<>"" and not(isnull(songjangdiv)) then
		strSql = strSql & " and div_cd='"& songjangdiv &"'"
	end if

	strSql = strSql + " order by idx desc "

	'response.write strSql & "<br>"
	rsget_Logistics.CursorLocation = adUseClient
	rsget_Logistics.Open strSql, dbget_Logistics, adOpenForwardOnly, adLockReadOnly
	if  not rsget_Logistics.EOF  then
		arrRS = rsget_Logistics.GetRows()
	end if
	rsget_Logistics.Close

	if Not IsArray(arrRS) then
		response.write "출력할 송장이 없습니다."
		dbget.close() : dbget_Logistics.close() : response.end
	end if

	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=SONGJANG_" & yyyymmdd & ".xls"

	response.clear

	response.write "<meta http-equiv=""content-type"" content=""text/html; charset=euc-kr"">" & vbCrLf
	response.write "<table border=1>" & vbCrLf
	response.write "<tr><td>택배사</td><td>운송장번호</td></tr>" & vbCrLf

    For i = LBound(arrRS, 2) To UBound(arrRS, 2)
        songjangno = arrRS(0, i)
		printsongjangdivname = getsongjangdivname(arrRS(1, i))
		response.write "<tr><td>" & printsongjangdivname & "</td><td style=""mso-number-format:'\@';"">" & songjangno & "</td></tr>" & vbCrLf
        if i mod 3000 = 0 then
            Response.Flush		' 버퍼리플래쉬
        end if
    Next

	response.write "</table>"

	dbget.close() : dbget_Logistics.close() : response.end

end if


if (yyyymmdd = "") then
	yyyymmdd = Left(Now(), 10)
end if

%>
<script type='text/javascript'>
function Research(frm){
	if (frm.yyyymmdd.value.length != 10) {
		alert('날짜를 입력하세요.' + frm.yyyymmdd.value.length);
		return;
	}
	frm.submit();
}
</script>

<form name="frmbar" method="get" style="margin:0px;">
  <table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="3">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red">&nbsp;<strong>송장목록 엑셀받기</strong></font>
				    </td>
				    <td align="right">
						택배사 : 
						<input type="hidden" name="songjangdiv" value="<%= songjangdiv %>">
						<%= getsongjangdivname(songjangdiv) %>
						<input type="text" class="text"  name="yyyymmdd" value="<%= yyyymmdd %>" size=14 maxlength=14 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">
        				<input type="button" class="button" value="엑셀받기" onclick="Research(frmbar)" >
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- 상단바 끝 -->
</table>
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->