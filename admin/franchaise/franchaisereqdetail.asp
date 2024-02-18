<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/offshop/franchaisereqcls.asp" -->
<%
dim idx, gubun, onlymifinish, mode, finishflag, admincomment
idx = request("id")
gubun = request("gubun")
onlymifinish = request("onlymifinish")
mode = request("mode")
finishflag  = request("finishflag")
admincomment = html2db(request("admincomment"))

dim sqlStr
if (mode="edit") then
	sqlStr = "update [db_cs].[dbo].tbl_franchaise" + vbCrlf
	sqlStr = sqlStr + " set finishflag='" + finishflag + "'" + vbCrlf
	sqlStr = sqlStr + " , admincomment='" + admincomment + "'" + vbCrlf
	sqlStr = sqlStr + " where idx=" + idx
	rsget.Open sqlStr,dbget,1

end if

dim ofran
set ofran = new CFranChaiseReqList
ofran.FRectIdx = idx

ofran.GetReqList
%>
<script language='javascript'>
function submitFrm(frm){
	var ret=confirm('저장하시겠습니까?');
	if (ret){
		frm.submit();
	}
}

function NewWindow(v){
	// var p = (v);
	 var w = window.open("http://www.thefingers.co.kr/myfingers/showimage.asp?img=" + v, "imageView", "left=10px,top=10px, width=560,height=400,status=no,resizable=yes,scrollbars=yes");
	// w.focus();
}
function backtopage(){
	location.href='franchaisereq.asp';
}
</script>
<table width="760" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr>
	<td bgcolor="#DDDDFF" width=100>문의자 성함</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fusername %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>점포유무</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fshop_exists %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>현재 직업</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fshop_maymonthgain %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>이메일</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fuseremail %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>휴대폰</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fhphone %></td>
<tr>
	<td bgcolor="#DDDDFF" width=100>연락처</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fuserphone %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>거주지</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Faddress %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>개설희망지역</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fshop_mayarea %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>예상오픈시기</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Ffran_open %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>예상실평수</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fshop_maypyng %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>예상투자비</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Finvest_money %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>운영주체</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Getshop_opertypeName %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>알게된 경로</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).GetKyungro %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>첨부파일</td>
	<td bgcolor="#FFFFFF" >
	<%
		if ofran.FItemList(0).Fetcfile<>"" then
			'파일종류에 따른 보기 옵션 추가
			Select Case getFileExtention(ofran.FItemList(0).Fetcfile)
				Case "jpg", "gif", "png"
					Response.Write "<span onClick=""NewWindow('" & staticImgUrl & ofran.upfolder & ofran.FItemList(0).Fetcfile & "')"" style='cursor:pointer'>" & ofran.FItemList(0).Fetcfile & "</span>"
				Case Else
					Response.Write "<a href='" & staticImgUrl & "/linkweb/download.asp?filepath=" & Server.URLencode(ofran.upfolder) & "&filename=" & Server.URLencode(ofran.FItemList(0).Fetcfile) & "'>" & ofran.FItemList(0).Fetcfile & "</a>"
			end Select
		end if
	%>
		</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>기타문의사항</td>
	<td bgcolor="#FFFFFF" ><%= nl2br(ofran.FItemList(0).Finvest_etc) %></td>
</tr>
<!--
<tr>
	<td bgcolor="#DDDDFF" width=100>상담종류</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).GetconsulttypeName %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>UserID</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fuserid %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>주민번호</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fjuminno %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" colspan=2 height="3"></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>기대수익율</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fgain_percent %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>기대수익</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fgain_money %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>예상투자기간</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Getinvest_yearName %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" colspan=2 height="3"></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>현재업종</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fshop_upjong %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>영업지역</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fshop_currarea %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>매장실평수</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fshop_pyng %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>예상개설자금</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fshop_mayfund %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>기타문의사항</td>
	<td bgcolor="#FFFFFF" ><%= nl2br(ofran.FItemList(0).Fshop_etc) %></td>
</tr>
-->
<tr>
	<td bgcolor="#DDDDFF" width=100>등록일</td>
	<td bgcolor="#FFFFFF" ><%= ofran.FItemList(0).Fregdate %></td>
</tr>
<form name="frmsubmit" method="post" action="">
<input type="hidden" name="id" value="<%= ofran.FItemList(0).FIdx %>">
<input type="hidden" name="mode" value="edit">
<tr>
	<td bgcolor="#DDDDFF" width=100>상태</td>
	<td bgcolor="#FFFFFF" >
	<input type="radio" name="finishflag" value="0" <% if (ofran.FItemList(0).Ffinishflag="0") then response.write "checked" %> >접수
	<input type="radio" name="finishflag" value="3" <% if (ofran.FItemList(0).Ffinishflag="3") then response.write "checked" %> >진행중
	<input type="radio" name="finishflag" value="7" <% if (ofran.FItemList(0).Ffinishflag="7") then response.write "checked" %> >완료
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>관리자 코멘트</td>
	<td bgcolor="#FFFFFF" >
	<textarea name=admincomment cols=80 rows=7><%= ofran.FItemList(0).Fadmincomment %></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" align="center">
	<input type="button" value="목록으로 이동" onclick="backtopage();">
	</td>
	<td bgcolor="#FFFFFF" align="center">
	<input type="button" value="저장" onclick="submitFrm(frmsubmit)">
	</td>
</tr>
</form>
</table>
<br><br>
<%
set ofran = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->