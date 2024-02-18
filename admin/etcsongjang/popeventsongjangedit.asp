<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 당첨자
' History : 2009.04.17 최초생성자 모름
'			2016.06.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<%
Dim id
	id = requestCheckvar(request("id"),10)

Dim ibeasong
Set ibeasong = new CEventsBeasong
	ibeasong.FRectId = id
	ibeasong.GetOneWinnerItem

If ibeasong.FResultCount < 1 Then
	response.write "<script>alert('검색된 내역이 없습니다.');</script>"
	response.write "<script>history.back();</script>"
	dbget.close()	:	response.End
End If

Dim i
Dim hpArr,hp1,hp2,hp3
Dim phoneArr,phone1,phone2,phone3

If IsNULL(ibeasong.FOneItem.Freqphone) then ibeasong.FOneItem.Freqphone=""
If IsNULL(ibeasong.FOneItem.Freqhp) then ibeasong.FOneItem.Freqhp=""
If IsNULL(ibeasong.FOneItem.Freqzipcode) then ibeasong.FOneItem.Freqzipcode=""

phoneArr = split(ibeasong.FOneItem.Freqphone,"-")
hpArr = split(ibeasong.FOneItem.Freqhp,"-")

if UBound(hpArr)>=0 then hp1 = hpArr(0)
if UBound(hpArr)>=1 then hp2 = hpArr(1)
if UBound(hpArr)>=2 then hp3 = hpArr(2)

if UBound(phoneArr)>=0 then phone1 = phoneArr(0)
if UBound(phoneArr)>=1 then phone2 = phoneArr(1)
if UBound(phoneArr)>=2 then phone3 = phoneArr(2)

%>
<script type="text/javascript">

function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".zipcode").value = post1 + "-" + post2;
    
    eval(frmname + ".addr1").value = addr;
    eval(frmname + ".addr2").value = dong;
}

function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function delThis(){
    var frm = document.infoform;

    if (confirm('삭제 하시겠습니까?')){
        if (confirm('정말로 삭제 하시겠습니까?')){
            frm.mode.value="del";
    		frm.submit();
		}
	}

}

function gotowrite(){
    var frm = document.infoform;
	if(frm.username.value == ""){
		alert("당첨자성함을 입력해주세요.");
	    frm.username.focus();
	    return;
	}

    if(frm.reqname.value == ""){
		alert("받으시는 분의 이름을 입력해주세요.");
	    frm.reqname.focus();
	    return;
	}

	if(frm.reqphone1.value == "" || frm.reqphone2.value == "" || frm.reqphone3.value == ""){
		alert("받으시는 분의 전화번호를 입력해주세요.");
	    frm.reqphone1.focus();
	    return;
	}

	if(frm.reqhp1.value == "" || frm.reqhp2.value == "" || frm.reqhp3.value == ""){
		alert("받으시는 분의 핸드폰 번호를 입력해주세요.");
	    frm.reqphone1.focus();
	    return;
	}

	if(frm.zipcode.value == ""){
		alert("받으시는 분의 주소를 입력해주세요.");
	    frm.zipcode.focus();
	    return;
	}

	if(frm.addr2.value == ""){
		alert("받으시는 분의 나머지주소를 입력해주세요.");
	    frm.addr2.focus();
	    return;
	}

	if (frm.reqdeliverdate.value.length!=10){
	    alert('출고 요청일을 입력하세요.');
	    frm.reqdeliverdate.focus();
	    return;
	}

	if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
	    alert('배송 구분을 선택 하세요.');
	    frm.isupchebeasong[0].focus();
	    return;
	}
	if(frm.isupchebeasong[1].checked&&(frm.jungsan.checked)&&((frm.jungsanValue.value=="")||(frm.jungsanValue.value=="0"))){
	    alert('정산 함 인 경우 정산액(매입가)를 입력하세요');
	    frm.jungsanValue.focus();
	    return;
	}
	if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
	    alert('업체 배송인 경우 브랜드 아이디를  선택 하세요.');
	    frm.makerid.focus();
	    return;
	}


    if (frm.issended.value=="Y"){
        if (frm.songjangdiv.value.length<1){
            alert("택배사를 선택하세요.");
    	    frm.songjangdiv.focus();
    	    return;
        }

        if (frm.songjangno.value.length<1){
            alert("송장번호를 입력하세요.");
    	    frm.songjangno.focus();
    	    return;
        }
    }

    //발송완료로 변경 안하는경우 Check
    if ((frm.isupchebeasong[0].checked)&&(frm.songjangdiv.value.length>0)&&(frm.songjangno.value.length>0)&&(frm.issended.value=="N")){
        alert('발송 완료인경우 발송완료로 변경해주셔야 합니다.');
        frm.issended.focus();
        return;
        //if (!confirm("발송 완료인경우 발송완료로 변경해주셔야 합니다. \n계속 하시겠습니까?")){
        //    return;
        //}

    }


	if (confirm('입력 내용이 정확합니까?')){
	    frm.mode.value="";
		frm.submit();
	}

}

function disabledBox(comp){
    var frm = comp.form;
    if (comp.value=="Y"){
        frm.makerid.disabled = false;
        frm.jungsan.disabled = false;

		frm.jungsanValue.disabled = false;
        frm.jungsan.checked = true;
    }else{
        //frm.makerid.selectedIndex = 0;
       // frm.makerid.value = '';
		frm.makerid.disabled = true;
		frm.jungsan.disabled = true;

        //frm.jungsanValue.value = '';
        frm.jungsanValue.disabled = true;
        frm.jungsan.checked = false;
    }
}

function jungsanYN(){
	var frm = document.infoform;
	if(frm.jungsan.checked==true){
		frm.jungsanValue.disabled = false;
	}else{
		frm.jungsanValue.value = '';
		frm.jungsanValue.disabled = true;
	}
}

function checkover1(obj) {
	var val = obj.value;
	if (val) {
		if (val.match(/^\d+$/gi) == null) {
			alert("숫자만 넣으세요!");
			document.infoform.jungsanValue.value = '';
			obj.select();
			return;
		}
	}
}

</script>
<!--
<table width="600" border="0" cellpadding="0" cellspacing="0" height="50">
  <tr valign="middle">
    <td width="8"><img src="http://fiximage.10x10.co.kr/images/my10x10/myeventmaster_popup_title.gif" width="580" height="50" hspace="10" vspace="10" ></td>
  </tr>
</table>
-->
<table width="100%" border="0" cellpadding="0" cellspacing=0 class="a">
<form name="infoform" method="post" action="/admin/etcsongjang/lib/doeventbeasonginfo.asp">
<input type="hidden" name="id" value="<%= id %>">
<input type="hidden" name="mode" value="">
<tr>
	<td align="center">
		<table width="98%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr height="30">
			<td height="2" colspan="2" >* 이벤트 및 기타출고 배송정보 입력/ 수정</td>
		</tr>
		<tr height="2">
			<td height="2" colspan="2" bgcolor="#AAAAAA"></td>
		</tr>
	<!--
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">이벤트<br>PrizeCode</td>
			<td style="padding-left:7"></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
	-->

	    <tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">구분</td>
			<td style="padding-left:7"><%= ibeasong.FOneItem.getEventKind %></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">이벤트명(구분명)</td>
			<td style="padding-left:7"><%= ibeasong.FOneItem.Fgubunname %></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">아이디</td>
			<td style="padding-left:7"><%= ibeasong.FOneItem.fuserid %></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">당첨상품</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="prizetitle" size="40" maxlength="64" value="<%= ibeasong.FOneItem.getPrizeTitle %>" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">당첨자성함</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="username" size="20" maxlength="20" value="<%= ibeasong.FOneItem.Fusername %>" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">수령인성함</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="reqname" size="20" maxlength="20" value="<%= ibeasong.FOneItem.Freqname %>" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">연락처</td>
			<td class="verdana_s" style="padding-left:7">
				<input type="text" class="text" name="reqphone1" size="3" class="verdana_s" maxlength="3" value="<%= phone1 %>">
				-
				<input type="text" class="text" name="reqphone2" size="4" class="verdana_s" maxlength="4" value="<%= phone2 %>">
				-
				<input type="text" class="text" name="reqphone3" size="4" class="verdana_s" maxlength="4" value="<%= phone3 %>">
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">핸드폰</td>
			<td class="verdana_s" style="padding-left:7">
				<input type="text" class="text" name="reqhp1" size="3" class="verdana_s"  maxlength="3" value="<%= hp1 %>">
				-
				<input type="text" class="text" name="reqhp2" size="4" class="verdana_s"  maxlength="4" value="<%= hp2 %>">
				-
				<input type="text" class="text" name="reqhp3" size="4" class="verdana_s"  maxlength="4" value="<%= hp3 %>">
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">수령인 주소</td>
			<td class="verdana_s" style="padding:5 0 5 7">
				<input type="text" class="text_ro" name="zipcode" size="7" class="verdana_s" readOnly value="<%= ibeasong.FOneItem.Freqzipcode %>">
				<input type="button" class="button" value="검색" onClick="FnFindZipNew('infoform','E')">
				<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('infoform','E')">
				<% '<input type="button" value="검색(구)" class="button" onclick="PopSearchZipcode('infoform');" onFocus="this.blur();"> %>
				<br>
				<input type="text" class="text_ro" name="addr1" size="16" maxlength="64"  readOnly value="<%= ibeasong.FOneItem.Freqaddress1 %>" ><br>
				<input type="text" class="text" name="addr2" size="40" maxlength="64" value="<%= ibeasong.FOneItem.Freqaddress2 %>" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">기타요청사항</td>
			<td class="verdana_s" style="padding:5 0 5 7"><textarea class="text" name="reqetc" class="textarea" style="width:350px;height:40px;"><%= ibeasong.FOneItem.Freqetc %></textarea></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		</table>
		<p>
		<table width="98%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
		    <td colspan="4" >* 사은품 정보</td>
		</tr>
		
		<tr height="1">
			<td height="1" colspan="4" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">배송구분</td>
			<td style="padding-left:7" colspan="3" >
			<% If IsNULL(ibeasong.FOneItem.Fisupchebeasong) or (Not (ibeasong.FOneItem.Fisupchebeasong="Y")) Then %>
				<input type=radio name=isupchebeasong value="N" checked onClick="disabledBox(this);">텐바이텐배송
				<input type=radio name=isupchebeasong value="Y" onClick="disabledBox(this);">업체직접배송
			<% Else %>
				<input type=radio name=isupchebeasong value="N" onClick="disabledBox(this);">텐바이텐배송
				<input type=radio name=isupchebeasong value="Y" checked onClick="disabledBox(this);">업체직접배송
			<% End If %>
			&nbsp;
			<% drawSelectBoxDesignerwithName "makerid",ibeasong.FOneItem.Fdelivermakerid %>
            
			<% If IsNULL(ibeasong.FOneItem.Fisupchebeasong) or (Not (ibeasong.FOneItem.Fisupchebeasong="Y")) Then %>
			<script language='javascript'>
				document.infoform.makerid.disabled=true;
			</script>
			<% End If %>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="4" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">사은품코드</td>
			<td style="padding-left:7" width="30%"><%= ibeasong.FOneItem.Fgift_code %>
			<% if Not isNULL(ibeasong.FOneItem.Fgift_itemid) then %>
			    (상품코드:<%=ibeasong.FOneItem.Fgift_itemid%>)
			<% end if %>
			</td>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">사은품명</td>
			<td style="padding-left:7" width="30%"><%= ibeasong.FOneItem.Fgiftkind_name %></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="4" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">정산여부</td>
			<td style="padding-left:7" width="30%">
				<input type="checkbox" class="checkbox" name="jungsan" id="jungsan" onclick="javascript:jungsanYN();" <%=CHKIIF(ibeasong.FOneItem.FjungsanYN="Y","checked","")%> >정산함&nbsp;&nbsp;
			</td>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">정산액(매입가)</td>
			<td style="padding-left:7" width="30%">
			<input type="text" size="9" style="text-align:right" class="text" id="jungsanValue" name="jungsanValue" value="<%=ibeasong.FOneItem.Fjungsan%>" onkeyup="checkover1(this)" <%=chkiif(IsNULL(ibeasong.FOneItem.Fjungsan) = True,"disabled","")%>>원
			</td>
			
		</tr>
		<tr height="1">
			<td height="1" colspan="4" bgcolor="#DDDDDD"></td>
		</tr>
		</table>
		<p>
		<table width="98%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
		    <td colspan="2" >* 출고정보</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">출고요청일</td>
			<td class="verdana_s" style="padding:5 0 5 7">
			<input type="text" class="text_ro" name="reqdeliverdate" size="10" maxlength="10"  value="<%= ibeasong.FOneItem.FreqDeliverDate %>" >
			<a href="javascript:jsPopCal('reqdeliverdate');"><img src="/images/calicon.gif" border="0" align="absmiddle"></a>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">발송상태 / 출고일</td>
			<td style="padding-left:7">
				<select name="issended" >
				<% If ibeasong.FOneItem.Fissended="Y" Then %>
					<option value="N">미발송
					<option value="Y" selected >발송완료
				<% Else %>
					<option value="N" selected >미발송
					<option value="Y">발송완료
				<% End If %>
				</select>
				/ <%= ibeasong.FOneItem.Fsenddate %>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">송장</td>
			<td style="padding-left:7">
				<% drawSelectBoxDeliverCompany "songjangdiv",ibeasong.FOneItem.Fsongjangdiv %>
				<input type="text" class="text" name="songjangno" size="14" maxlength="20" value="<%= ibeasong.FOneItem.Fsongjangno %>">
			</td>
		</tr>
		<tr height="2">
			<td height="2" colspan="2" bgcolor="#AAAAAA"></td>
		</tr>
		<tr height="30">
			<td colspan="2" align="center">
		<% If (ibeasong.FOneItem.IsSended) Then %>
			<input type="button" class="button" value=" 저 장 " onClick="if (confirm('이미 발송된 내역 입니다. 수정 하시겠습니까?')) { gotowrite(); };" onfocus="this.blur();">
		<% Else %>
			<input type="button" class="button" value=" 저 장 " onClick="gotowrite();" onfocus="this.blur();">
			&nbsp;&nbsp;&nbsp;
			<input type="button" class="button" value=" 삭 제 " onClick="delThis();" onfocus="this.blur();">
		<% End If %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<% Set ibeasong = Nothing %>
<!-- #include virtual="/admin/lib/poptail.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->