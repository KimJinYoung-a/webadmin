<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<%
dim odataarr, dataarr, bufarr, bufstr
Dim gubuncd
odataarr = request("dataarr")
dataarr = request("dataarr")
gubuncd = request("gubuncd")

dim i, sqlStr
dim ErrStr

if (dataarr<>"") then
	'response.write dataarr
	dataarr = split(dataarr,vbcrlf)
	for i=LBound(dataarr) to UBound(dataarr)
		bufarr = split(dataarr(i),chr(9))
		if UBound(bufarr)>9 then
            if (Trim(bufarr(2))="") or (Trim(bufarr(3))="") or (Trim(bufarr(4))="") or (Len(Trim(bufarr(5)))<5) or (Trim(bufarr(6))="") or (Trim(bufarr(7))="") or (Trim(bufarr(8))="") then 
                'skip
                ErrStr = ErrStr + CStr(i+1) + "행 " + bufarr(0) + " 등록오류 \n"
            else
    			sqlStr = "insert into  [db_sitemaster].[dbo].tbl_etc_songjang"    + VbCrlf
    			sqlStr = sqlStr + " (userid, username, reqname, reqphone, reqhp, reqzipcode, reqaddress1," + VbCrlf
    			sqlStr = sqlStr + " reqaddress2, gubuncd, gubunname, prizetitle, reqetc, inputdate, reqdeliverdate)" + VbCrlf
    			sqlStr = sqlStr + " values(" + VbCrlf
    			sqlStr = sqlStr + " '" + html2db(LeftB(replace(Trim(bufarr(0)),"'",""),64)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(1)),"'",""),64)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(2)),"'",""),64)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(3)),"'",""),32)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(4)),"'",""),32)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(5)),"'",""),14)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(6)),"'",""),255)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(7)),"'",""),255)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,'"&gubuncd&"'" + VbCrlf
    			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(8)),"'",""),255)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(9)),"'",""),255)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(10)),"'",""),255)) + "'" + VbCrlf
    			sqlStr = sqlStr + " ,getdate()" + VbCrlf
    			sqlStr = sqlStr + " ,convert(varchar(10),getdate(),21)" + VbCrlf
    			sqlStr = sqlStr + " )"

    			rsget.Open sqlStr,dbget,1
            end if
		end if
	next
	'bufstr = Left(bufstr,Len(bufstr)-1)

	'response.write sqlStr + "<br>"
end if

%>
<script language='javascript'>
function saveClick(){
	var frm = document.frm;
	
	if(frm.gubuncd.value == ""){
		alert('구분을 선택하세요');
		frm.gubuncd.focus();
		return;
	}

	if (confirm('저장하시겠습니까?')){
		frm.submit();
	}
}
</script>
<table border=0 cellspacing=0 cellpadding=0 class="a">
<form name=frm method=post>
<tr>
	<td colspan=2><font color="red">탭으로 분리</font><br>
	아이디,이름,받는분,전화(-),핸드폰(-),우편번호(-),주소1,주소2,구분명(이벤트명),상품명,기타사항<br>
	userid, username, reqname, reqphone, reqhp, reqzipcode, reqaddress1,
	reqaddress2, gubunname, prizetitle, reqetc<br>
	<font color="red">우편번호(-),주소1,주소2에 빈값이 있으면 등록이 안됩니다.</font>
	
	</td>
</tr>
<tr>
	<td>
		<select name="gubuncd" class="select">
			<option value="">전체</option>
<!--
			<option value="96" <%=chkiif(gubuncd = "96","selected","")%> >29cm용</option>
			<option value="97" <%=chkiif(gubuncd = "97","selected","")%>>고객</option>
-->
			<option value="98" <%=chkiif(gubuncd = "98","selected","")%> >판촉</option>
			<option value="99" <%=chkiif(gubuncd = "99","selected","")%>>기타</option>
		</select>
	</td>
	<td align="right" valign="bottom">
		<a href="FORM.xlsx" target="_blank">[엑셀FORM]</a>
	</td>
</tr>
<tr>
	<td colspan=2>
	<textarea name="dataarr" cols=230 rows=8><%= odataarr %></textarea>
	</td>
</tr>
<tr>
	<td>
	<input type= button value=clear onclick="frm.dataarr.value=''; frm.pbrandid.value=''">
	</td>
	<td><input type= button value="저장" onclick="saveClick()"></td>
</tr>
</form>
</table>
<%
if odataarr<>"" then
%>
<script language='javascript'>
<% if ErrStr<>"" then %>
    alert('<%= ErrStr %>');
    opener.location.reload();
    window.close();
<% else %>
    alert('ok');
    opener.location.reload();
    window.close();
<% end if %>


</script>
<%
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->