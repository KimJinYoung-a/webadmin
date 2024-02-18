<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/jungsan_function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim id
id = RequestCheckVar(request("id"),9)
dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectId = id
ojungsan.JungsanMasterList

dim rd_state
rd_state = ojungsan.FItemList(0).Ffinishflag
%>
<script language='javascript'>
function savestate(frm){
 //   if ((frm.idx.value=="294398")||(frm.idx.value=="312521")||(frm.idx.value=="314361")){ alert('해당건은 변경불가합니다.'); return; } //2016/09/29희란 요청
       if ((frm.idx.value=="294398")){ alert('해당건은 변경불가합니다.'); return; } //2016/12/01희란 요청 314361 오픈
    //  if ((frm.idx.value=="348027")){ alert('해당건은 변경불가합니다.'); return; } //2017/06/05희란 요청
    if((frm.idx.value=="354608") || (frm.idx.value=="354186")){ alert('해당건은 변경불가합니다.'); return; } //2017/07/07
     
      
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="statechange";
		frm.submit();
	}
}

function savetaxReg(frm){
    if (frm.taxregdate.value.length<1){
        alert('발행일이 지정되지 않았습니다. ');
        return;
    }

    if (frm.billsiteCode.value.length<1){
        alert('발행 업체가 지정되지 않았습니다. 계속 하시겠습니까?');
        return;
    }


	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="taxregchange";
		frm.submit();
	}
}

function saveipkumReg(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="ipkumregchange";
		frm.submit();
	}
}

function savetaxtype(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="taxtypechange";
		frm.submit();
	}
}

function savedifferencekey(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="differencekeychange";
		frm.submit();
	}
}

function saveGroupid(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="editGroupid";
		frm.submit();
	}
}

function saveJacct(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="editJAcctCd";
		frm.submit();
	}
}

function saveAvailNeo(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="editAvailNeo";
		frm.submit();
	}
}

function delTaxInfo(frm){
	var ret = confirm('계산서 발행정보를 삭제 하시겠습니까?');
	if (ret){
		frm.mode.value="delTaxInfo";
		frm.submit();
	}
}

function popSearchGroupID(frmname,compname){
    var popwin = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname + "&compname=" + compname,"popSearchGroupID","width=800 height=680 scrollbars=yes resizable=yes");
    popwin.focus();
}


function jsGetTax(ibizNo, itotSum){
	var sSearchText = ibizNo;
	var itotSum = itotSum;
	var winTax = window.open("/admin/tax/popSetEseroTax.asp?sST="+sSearchText+"&totSum="+itotSum+"&tgType=NRM","popGetTaxInfo","width=1200, height=800, resizable=yes, scrollbars=yes");
	winTax.focus();
}

function fillTaxInfo(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP){
    var frm = document.statefrm;
    frm.taxregdate.value = dID;
    frm.eseroEvalSeq.value = eTax;

    //발행업체 지정
    var mayApCd = eTax.substring(8,16);
    if (mayApCd=="10000000"){
        //국세청
        frm.billsiteCode.value = 'E';
    }else if(mayApCd=="10000966"){
        //빌365
        frm.billsiteCode.value = 'B';
    }else{
        //기타
        frm.billsiteCode.value = 'Y';
    }
}

function jsNewRegXML(){
    var winD = window.open("/admin/tax/popRegfileXML.asp","popDXML","width=600, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}


function jsNewRegHand(){
    var winD = window.open("/admin/tax/popRegfileHand.asp","popDHand","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

</script>
<br>
<table width="760" cellspacing="0" class="a">
<tr>
  <td align="right"><a href="nowjungsanlist.asp?menupos=130">목록</a></td>
</tr>
</table>

<div class="a">1.기준정보</div>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<form name="statefrm" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="mode" value="statechange">
<input type="hidden" name="idx" value="<%= ojungsan.FItemList(0).FId %>">
    <tr >
      <td width="100" bgcolor="#DDDDFF">브랜드ID</td>
      <td bgcolor="#FFFFFF"><%= ojungsan.FItemList(0).Fdesignerid %></td>
    </tr>
    <tr >
      <td width="100" bgcolor="#DDDDFF">정산대상년월</td>
      <td bgcolor="#FFFFFF"><%= ojungsan.FItemList(0).FYYYYMM %>&nbsp;(<%= ojungsan.FItemList(0).Fdifferencekey %>차)</td>
    </tr>
    <tr >
      <td width="100" bgcolor="#DDDDFF">정산방식</td>
      <td bgcolor="#FFFFFF"><%= ojungsan.FItemList(0).getJGubunName %></td>
    </tr>
    <tr >
      <td width="100" bgcolor="#DDDDFF">현재상태</td>
      <td bgcolor="#FFFFFF">
      	<input type="radio" name="rd_state" value="0" <% if rd_state="0" then response.write "checked" %> >수정중
		<input type="radio" name="rd_state" value="1" <% if rd_state="1" then response.write "checked" %> >업체확인중
		<input type="radio" name="rd_state" value="2" <% if rd_state="2" then response.write "checked" %> >업체확인완료
		<input type="radio" name="rd_state" value="3" <% if rd_state="3" then response.write "checked" %> >정산확정
		<input type="radio" name="rd_state" value="7" <% if rd_state="7" then response.write "checked" %> >입금완료
		<input type="button" value="상태변경" onclick="savestate(statefrm);">
      </td>
    </tr>
    <tr>
    	<td width="100" bgcolor="#DDDDFF">계산서구분</td>
      	<td bgcolor="#FFFFFF">
      	<select name="taxtype" >
      	<option value="" <% if IsNULL(ojungsan.FItemList(0).Ftaxtype) or (ojungsan.FItemList(0).Ftaxtype="") then response.write "selected" %> >
      	<option value="01" <% if (ojungsan.FItemList(0).Ftaxtype="01") then response.write "selected" %> >세금계산서
      	<option value="02" <% if (ojungsan.FItemList(0).Ftaxtype="02") then response.write "selected" %> >계산서
      	<option value="03" <% if (ojungsan.FItemList(0).Ftaxtype="03") then response.write "selected" %> >원천징수
      	</select>
      	<input type="button" value="저장" onclick="savetaxtype(statefrm);">
      	</td>
    </tr>
    <tr>
    	<td width="100" bgcolor="#DDDDFF">차수</td>
      	<td bgcolor="#FFFFFF">
      	<input type="text" name="differencekey" value="<%= ojungsan.FItemList(0).Fdifferencekey %>" size="2" maxlength="2">
      	<input type="button" value="저장" onclick="savedifferencekey(statefrm);">
      	(숫자로 입력)
      	</td>
    </tr>
    <tr>
      <td width="100" bgcolor="#DDDDFF">세금계산서발행일</td>
      <td bgcolor="#FFFFFF">
      	<% if (rd_state="1") or (rd_state="3") or (rd_state="7") then %>
      	<input type="text" name="taxregdate" value="<%= ojungsan.FItemList(0).Ftaxregdate %>" size="10">
      	<a href="javascript:calendarOpen(statefrm.taxregdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
      	<input type="button" value="계산서정보저장" onclick="savetaxReg(statefrm);">

      	    <% If ISNULL(ojungsan.FItemList(0).Ftaxlinkidx) then %>
          	&nbsp;&nbsp;&nbsp;&nbsp;
          	<input type="button" value="선택입력" onClick="jsGetTax('<%= REplace(ojungsan.FItemList(0).Fcompany_no,"-","") %>','<%= ojungsan.FItemList(0).GetTotalSuplycash %>');">
          	<input type="button" value="XML" onClick="jsNewRegXML();">
          	<input type="button" value="종이계산서입력" onClick="jsNewRegHand();">
          	<% end if %>
      	<br>
      	<input type="hidden" name="taxlinkidx" value="<%= ojungsan.FItemList(0).Ftaxlinkidx %>">
      	<% if isNULL(ojungsan.FItemList(0).Ftaxlinkidx) then %>
            <% call DrawBillSiteCombo("billsiteCode",ojungsan.FItemList(0).FbillsiteCode) %>
        <% else %>
            <input type="hidden" name="billsiteCode" value="<%= ojungsan.FItemList(0).FbillsiteCode %>">
            <%= ojungsan.FItemList(0).FbillSiteName %>
        <% end if %>
      	<input type="text" name="neotaxno" value="<%= ojungsan.FItemList(0).Fneotaxno %>" size="20" maxlength="32" <%= CHKIIF(ISNULL(ojungsan.FItemList(0).Ftaxlinkidx),"","class='text_ro' READONLY") %>>(TAXNO)
      	<br>
      	<input type="text" name="eseroEvalSeq" value="<%= ojungsan.FItemList(0).FeseroEvalSeq %>" size="30" maxlength="24" <%= CHKIIF(ISNULL(ojungsan.FItemList(0).Ftaxlinkidx),"","class='text_ro' READONLY") %> >(이세로 승인번호 '-' 빼고입력 24자리)

      	<% end if %>

      	<% If ISNULL(ojungsan.FItemList(0).Ftaxlinkidx) then %>
      	    <% if (ojungsan.FItemList(0).Ffinishflag="0" or ojungsan.FItemList(0).Ffinishflag="1" or ojungsan.FItemList(0).Ffinishflag="2") then %>
          	<br><input type="button" value="계산서발행정보삭제" onClick="delTaxInfo(statefrm)">
          	<% end if %>
        <% end if %>
      </td>
    </tr>
    <tr>
      <td width="100" bgcolor="#DDDDFF">입금일</td>
      <td bgcolor="#FFFFFF">
      	<% if rd_state="7" then %>
      	<input type="text" name="ipkumregdate" value="<%= ojungsan.FItemList(0).Fipkumdate %>" size="10"> (예 2002-09-12)
      	<input type="button" value="저장" onclick="saveipkumReg(statefrm);">
      	<% end if %>
      </td>
    </tr>
    <tr>
      <td width="100" bgcolor="#DDDDFF">그룹코드</td>
      <td bgcolor="#FFFFFF">
      	<input type="text" class="text" name="groupid" value="<%= ojungsan.FItemList(0).Fgroupid %>" size="10" >
      	<input type="button" class="button" value="Code검색" onclick="popSearchGroupID(this.form.name,'groupid');" >
      	<input type="button" value="저장" onclick="saveGroupid(statefrm);" <%= chkIIF(rd_state>1,"disabled","") %> >
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6A6A6" >계정과목코드</td>
        <td>
            
            <input type="text" class="text_ro" value="<%= ojungsan.FItemList(0).Fjacc_nm %>" size="10" readonly >
            <input type="text" class="text" name="jacctcd" value="<%= ojungsan.FItemList(0).Fjacctcd %>" size="7" >
      	    <input type="button" value="저장" onclick="saveJacct(statefrm);" >
      	    <!-- 기본 계정과목(미 입력시)은 [매입-상품매출원가,매출-] -->
        </td>
    </tr>
    <!--
    <tr>
      <td width="100" bgcolor="#DDDDFF">네오포트발행</td>
      <td bgcolor="#FFFFFF">
      	<input type="checkbox" name="availneoport" <%= CHKIIF(ojungsan.FItemList(0).Favailneo=1,"checked","") %>>가능
      	<input type="button" value="저장" onclick="saveAvailNeo(statefrm);" <%= chkIIF(rd_state>=3,"disabled","") %> >
      </td>
    </tr>
    -->
</form>
</table>

<br>
<div class="a"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=upche">2.정산내역</a></div>
<table width="760" cellspacing="1" cellpadding=2 class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td width=100 align=left>구분</td>
	<td width=100>총주문건수</td>
	<td width=100>소비자가총액</td>
	<td width=100>공급가총액</td>
	<td width=70>마진</td>
	<td>기타</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=upche">업체배송</a></td>
	<td align=right><%= ojungsan.FItemList(0).Fub_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fub_totalsellcash<>0 then %>
	<td align=center><%= ojungsan.FItemList(0).Fub_totalsuplycash/ojungsan.FItemList(0).Fub_totalsellcash * 100 %> %</td>
	<% else %>
	<td align=center></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fub_comment) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=maeip">매입내역</a></td>
	<td align=right><%= ojungsan.FItemList(0).Fme_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fme_totalsellcash<>0 then %>
	<td align=center><%= ojungsan.FItemList(0).Fme_totalsuplycash/ojungsan.FItemList(0).Fme_totalsellcash * 100 %> %</td>
	<% else %>
	<td align=center></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fme_comment) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=witaksell">위탁온라인내역</a></td>
	<td align=right><%= ojungsan.FItemList(0).Fwi_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fwi_totalsellcash<>0 then %>
	<td align=center><%= ojungsan.FItemList(0).Fwi_totalsuplycash/ojungsan.FItemList(0).Fwi_totalsellcash * 100 %> %</td>
	<% else %>
	<td align=center></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fwi_comment) %></td>
</tr>
<!--
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">위탁 오프라인</td>
	<td><%= ojungsan.FItemList(0).Fsh_cnt %></td>
	<td><%= FormatNumber(ojungsan.FItemList(0).Fsh_totalsellcash,0) %></td>
	<td><%= FormatNumber(ojungsan.FItemList(0).Fsh_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fsh_totalsellcash<>0 then %>
	<td><%= ojungsan.FItemList(0).Fsh_totalsuplycash/ojungsan.FItemList(0).Fsh_totalsellcash * 100 %> %</td>
	<% else %>
	<td>?</td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fsh_comment) %></td>
</tr>
-->
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">위탁 기타</td>
	<td align=right><%= ojungsan.FItemList(0).Fet_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fet_totalsellcash<>0 then %>
	<td align=center><%= ojungsan.FItemList(0).Fet_totalsuplycash/ojungsan.FItemList(0).Fet_totalsellcash * 100 %> %</td>
	<% else %>
	<td align=right></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fet_comment) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">총계</td>
	<td></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).GetTotalSellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).GetTotalSuplycash,0) %></td>
	<% if ojungsan.FItemList(0).GetTotalSellcash<>0 then %>
	<td align=center><%= ojungsan.FItemList(0).GetTotalSuplycash/ojungsan.FItemList(0).GetTotalSellcash * 100 %> %</td>
	<% else %>
	<td align=right></td>
	<% end if %>
	<td></td>
</tr>
</table>
<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->