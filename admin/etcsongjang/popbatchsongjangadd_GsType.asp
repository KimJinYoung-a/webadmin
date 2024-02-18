<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<%
dim odataarr, dataarr, bufarr, bufstr
odataarr = request("dataarr")
dataarr = request("dataarr")

dim i, sqlStr, assignedRow
dim ErrStr

assignedRow = 0

dim addrAll, buf
dim username, reqaddress1, reqaddress2
dim gubunname, prizetitle, reqetc
dim etcKey, etcBaljuNo, etcKeyExists
gubunname  = "세이브더칠드런"

if (dataarr<>"") then
	'response.write dataarr
	dataarr = split(dataarr,vbcrlf)
	for i=LBound(dataarr) to UBound(dataarr)
	  if (Trim(dataarr(i))<>"") then
		bufarr = split(dataarr(i),chr(9))
		
		'response.write UBound(bufarr)&"<br>"
		if UBound(bufarr)>=16 then
		    username    = ""
		    reqaddress1 = ""
            reqaddress2 = ""
            prizetitle  = ""
            reqetc      = ""
            etcKey      = ""
            etcBaljuNo  = ""        ''송장일괄등록시 필요
            etcKeyExists= false
            
            username = replace(Trim(bufarr(10)),"'","")
            if (Instr(username,"/")>0) then
                username = Left(username,Instr(username,"/")-1)
            end if

            addrAll = LeftB(replace(Trim(bufarr(9)),"'",""),255)
            buf = split(addrAll," ")
            
            if Ubound(buf)>3 then
                reqaddress1 = buf(0) + " " + buf(1) + " " + buf(2)
                reqaddress2 = Mid(addrAll,Len(reqaddress1)+2,255)
            elseif Ubound(buf)>2 then
                reqaddress1 = buf(0) + " " + buf(1) 
                reqaddress2 = Mid(addrAll,Len(reqaddress1)+2,255)
            elseif Ubound(buf)>1 then
                reqaddress1 = buf(0)
                reqaddress2 = Mid(addrAll,Len(reqaddress1)+2,255)
            else
                
            end if
            
            prizetitle = LeftB(replace(Trim(bufarr(16)),"'",""),255)
            if (prizetitle="3935344") or (prizetitle="""3935344") then prizetitle = "신생아살리기 모자뜨기 Kit"
            
            if (Trim(bufarr(14))<>"1") then
                prizetitle = prizetitle  + " [" + Trim(bufarr(14)) + "] 개"
            end if
            reqetc = LeftB(replace(Trim(bufarr(15)),"'",""),255)
            etcKey = replace(Trim(bufarr(3)),"'","")
            etcBaljuNo = replace(Trim(bufarr(2)),"'","")
                
		    if (Left(Trim(bufarr(0)),7)="세이브더칠드런") or (Trim(bufarr(0))="No.") or (Trim(bufarr(0))="") then
		        ''skip 
            elseif (etcKey="") or (etcBaljuNo="") or (prizetitle="") or (reqaddress1="") or (reqaddress2="") or (Trim(bufarr(10))="") then 
                'skip
                ErrStr = ErrStr + CStr(i+1) + "열 " + bufarr(0) + " 등록오류 \n"
            else
                
                
                
                ''주문번호 중복 체크
                sqlStr = "select count(*) as CNT"
                sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_etc_songjang"
                sqlStr = sqlStr + " where etcKey='" + etcKey + "'"
                sqlStr = sqlStr + " and etcBaljuNo='" + etcBaljuNo + "'"
                sqlStr = sqlStr + " and gubuncd='91'"
                sqlStr = sqlStr + " and deleteyn='N'"
                
                rsget.Open sqlStr,dbget,1
                    etcKeyExists = rsget("CNT")>0
                rsget.Close
                
                if (etcKeyExists) then
                    ErrStr = ErrStr + CStr(i+1) + "열 등록오류 - 기입력된 주문번호 : "& etcKey & " , 발주번호" &etcBaljuNo & "\n"
                else
        			sqlStr = "insert into  [db_sitemaster].[dbo].tbl_etc_songjang"    + VbCrlf
        			sqlStr = sqlStr + " (userid, username, reqname, reqphone, reqhp, reqzipcode, reqaddress1," + VbCrlf
        			sqlStr = sqlStr + " reqaddress2, gubuncd, gubunname, prizetitle, reqetc, inputdate, reqdeliverdate, etcKey, etcBaljuNo)" + VbCrlf
        			sqlStr = sqlStr + " values(" + VbCrlf
        			sqlStr = sqlStr + " ''" + VbCrlf
        			sqlStr = sqlStr + " ,'" + html2db(username) + "'" + VbCrlf
        			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(5)),"'",""),64)) + "'" + VbCrlf
        			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(6)),"'",""),32)) + "'" + VbCrlf
        			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(7)),"'",""),32)) + "'" + VbCrlf
        			sqlStr = sqlStr + " ,'" + html2db(LeftB(replace(Trim(bufarr(8)),"'",""),14)) + "'" + VbCrlf
        			sqlStr = sqlStr + " ,'" + html2db(reqaddress1) + "'" + VbCrlf
        			sqlStr = sqlStr + " ,'" + html2db(reqaddress2) + "'" + VbCrlf
        			sqlStr = sqlStr + " ,'91'" + VbCrlf
        			sqlStr = sqlStr + " ,'" + html2db(gubunname) + "'" + VbCrlf
        			sqlStr = sqlStr + " ,'" + html2db(prizetitle) + "'" + VbCrlf
        			sqlStr = sqlStr + " ,'" + html2db(reqetc) + "'" + VbCrlf
        			sqlStr = sqlStr + " ,getdate()" + VbCrlf
        			sqlStr = sqlStr + " ,convert(varchar(10),getdate(),21)" + VbCrlf
        			sqlStr = sqlStr + " ,'" + etcKey + "'" + VbCrlf
        			sqlStr = sqlStr + " ,'" + etcBaljuNo + "'" + VbCrlf
        			sqlStr = sqlStr + " )"
    
        			dbget.Execute sqlStr
        			
        			assignedRow = assignedRow + 1
    			end if
            end if
        else
            ErrStr = ErrStr + CStr(i+1) + "열 등록오류 - 필드수 부족\n"
		end if
      end if
	next
	'bufstr = Left(bufstr,Len(bufstr)-1)

	response.write bufstr + "<br>"
end if

%>
<script language='javascript'>
function saveClick(){
	if (confirm('저장하시겠습니까?')){
		frm.submit();
	}
}
</script>
<table border=0 cellspacing=0 cellpadding=0 class="a">
<form name=frm method=post>
<tr>
	<td colspan=2><font color="red">탭으로 분리</font><br>
	</td>
</tr>
<tr>
	<td colspan=2>
	<textarea name="dataarr" cols=90 rows=8><%= odataarr %></textarea>
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
    alert('ok\n\n<%= assignedRow %> 건 등록 완료 ');
    opener.location.reload();
    window.close();
<% end if %>


</script>
<%
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->