<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  브랜드별매출
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
const CUNDERMargin = 10

dim yyyy1,mm1,designer,rectorder, groupid, finishflag, taxtype, chkmargin, vPurchaseType, jgubun
dim research, page, ix, targetGbn, jacctcd, differencekey, companynoYN
dim searchType, searchText, jungsanGubun

designer = requestCheckVar(request("designer"),32)
yyyy1    = requestCheckVar(request("yyyy1"),4)
mm1      = requestCheckVar(request("mm1"),2)
rectorder = requestCheckVar(request("rectorder"),32)
groupid  = requestCheckVar(request("groupid"),32)
research = requestCheckVar(request("research"),32)
finishflag = requestCheckVar(request("finishflag"),10)
page     = requestCheckVar(request("page"),10)
taxtype  = requestCheckVar(request("taxtype"),10)
chkmargin= requestCheckVar(request("chkmargin"),10)
vPurchaseType = requestCheckVar(request("purchasetype"),2)
jgubun   = requestCheckVar(request("jgubun"),10)
targetGbn = requestCheckVar(request("targetGbn"),10)
jacctcd = requestCheckVar(request("jacctcd"),10)
differencekey = requestCheckVar(request("differencekey"),10)
searchType = requestCheckVar(request("searchType"), 32)
searchText = requestCheckVar(request("searchText"), 32)
companynoYN = requestCheckVar(request("companynoYN"), 1)
jungsanGubun = requestCheckVar(request("jungsanGubun"), 12)

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

if page="" then page=1

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FPageSize  = 100
ojungsan.FCurrPage  = page
ojungsan.FRectYYYYMM = yyyy1 + "-" + mm1
ojungsan.FRectDesigner = designer
ojungsan.FRectGroupID = groupid
ojungsan.FrectOrder = rectorder
ojungsan.Frectfinishflag = finishflag
ojungsan.FRectTaxType = taxtype
ojungsan.FRectPurchaseType = vPurchaseType
ojungsan.FRectJGubun = jgubun
ojungsan.FRecttargetGbn = targetGbn
ojungsan.FRectjacctcd = jacctcd
ojungsan.FRectdifferencekey = differencekey
ojungsan.FRectSearchType = searchType
ojungsan.FRectSearchText = searchText
ojungsan.FRectCompanynoYN = companynoYN
ojungsan.FRectJungsanGubun = jungsanGubun


IF (chkmargin="on") then
    ojungsan.FRectUnderMargin = CStr(CUNDERMargin)
end if

if (research<>"") then
    ojungsan.JungsanMasterList
end if

dim i
dim tot1,tot2,tot3,tot4,tot5, totcom, totdlv,totsum
tot1 = 0
tot2 = 0
tot3 = 0
tot4 = 0
tot5 = 0
totsum = 0
%>
<script language='javascript'>
function NextPage(ipage){
    document.frm.page.value=ipage;
    document.frm.submit();
}

function research(frm,order){
	frm.rectorder.value = order;
	frm.submit();
}

function PopUpchebrandInfo(v){
	var popwin = window.open("/admin/lib/popupchebrandinfo.asp?designer=" + v,"popupchebrandinfo","width=640 height=680 scrollbars=yes resizable=yes");
    popwin.focus();
}

function popSearchGroupID(frmname,compname){
    var popwin = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname + "&compname=" + compname,"popSearchGroupID","width=800 height=680 scrollbars=yes resizable=yes");
    popwin.focus();
}

function popDetail(v){
	window.open('popdetail.asp?id=' + v );
}

function dellThis(v){
	var upfrm = document.frmarr;
	var ret = confirm('모든 정산 데이터를 삭제 하시겠습니까?');
	if (ret){
		upfrm.idx.value = v;
		upfrm.mode.value = "dellall";
		upfrm.submit();
	}
}

function NextStep(idx){
	<%if groupid = "G02856" or groupid = "g02856" then %>
	alert('해당건은 변경불가합니다.'); return;
	<%end if%>
 //   if ((idx=="294398")||(idx=="312521")||(idx=="314361")){ alert('해당건은 변경불가합니다.'); return; } //2016/09/29희란 요청
    if ((idx=="294398")){ alert('해당건은 변경불가합니다.'); return; } //2016/12/01희란 요청(312461 해제)
      if((idx=="354608") || (idx=="354186") || (idx=="380557")){ alert('해당건은 변경불가합니다.'); return; } //2017/07/07
	var upfrm = document.frmarr;
	upfrm.mode.value= "statechange";
	upfrm.idx.value= idx;
	upfrm.rd_state.value="1";

	var ret = confirm('확인 대기 상태로 진행 하시겠습니까?<%=groupid%>');
	if (ret){
		upfrm.submit();
	}
}

function MakeBrandBatchJungsan(frm){
    if (frm.jgubun.value.length<1){
        alert('정산 방식 구분을 선택 하세요.');
        frm.jgubun.focus();
        return;
    }

    if (frm.differencekey.value.length<1){
        alert('차수 구분을 선택 하세요.');
        frm.differencekey.focus();
        return;
    }

    if (frm.itemvatYN.value.length<1){
        alert('상품 과세 구분을 선택 하세요.');
        frm.itemvatYN.focus();
        return;
    }

    if (confirm('정산내역을 작성 하시겠습니까?')){
        var queryurl = 'dodesignerjungsan.asp?mode=brandbatchprocess&jgubun='+frm.jgubun.value+'&designer=' + frm.makerid.value + '&yyyy1=' + frm.yyyy.value + '&mm1=' + frm.mm.value + '&differencekey=' + frm.differencekey.value + '&itemvatYN=' + frm.itemvatYN.value+'&ipchulArr='+frm.ipchulArr.value;

        var popwin = window.open(queryurl ,'on_jungsan_process','width=200, height=200, scrollbars=yes, resizable=yes');
    }
}

//전체 선택
function jsChkAll(){
var frm;
frm = document.frmList;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
	   	   	if(frm.chkitem.disabled==false){
		   	 	frm.chkitem.checked = true;
		   	}
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					 	if(frm.chkitem[i].disabled==false){
					frm.chkitem[i].checked = true;
				}
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}
		}
	  }

	}

}

 //다중 선택 상태변경
function jsMultiStateChange(){
	var frm = document.frmList;
	if(typeof(frm.chkitem) !="undefined"){
	 	if(!frm.chkitem.length){
	 		if(!frm.chkitem.checked){
	 			alert("선택한 정산 대상이 없습니다. 선택해 주세요");
	 			return;
	 		}
	 	}
        else{
            for(i=0;i<frm.chkitem.length;i++){
                if(frm.chkitem[i].checked) {
                    frm.idxarr.value = frm.idxarr.value + frm.chkitem[i].value + ",";
                }
            }
	 		if(frm.idxarr.value==""){
	 			alert("선택한 정산 대상이 없습니다. 선택해 주세요");
	 			return;
	 		}else{
                //alert(frm.idxarr.value);
                frm.submit();
            }
        }
	}
}
</script>


<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rectorder" value="<%=rectorder%>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">

   	<tr align="center" bgcolor="#FFFFFF" >
        <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
        <td align="left">
	        	정산대상년월 : <% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;
				브랜드ID : <% drawSelectBoxDesignerwithName "designer",designer  %>&nbsp;&nbsp;
				업체(그룹코드) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >
				<input type="button" class="button" value="Code검색" onclick="popSearchGroupID(this.form.name,'groupid');" >&nbsp;&nbsp;
                계정과목코드 : <input type="text" class="text" name="jacctcd" value="<%= jacctcd %>" size="7" >
	        </td>
        <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
    		<a href="javascript:document.frm.rectorder.value='';document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
    	</td>
    </tr>
	<tr>
        <td bgcolor="#FFFFFF" >
        	구매유형 : 
            <% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
			&nbsp;&nbsp;
			상태
			<select name="finishflag" >
			<option value="">전체
			<option value="0" <%= CHKIIF(finishflag="0","selected","") %> >수정중
			<option value="1" <%= CHKIIF(finishflag="1","selected","") %> >업체확인대기
			<option value="2" <%= CHKIIF(finishflag="2","selected","") %> >업체확인완료
			<option value="3" <%= CHKIIF(finishflag="3","selected","") %> >정산확정
			<option value="7" <%= CHKIIF(finishflag="7","selected","") %> >입금완료
			</select>
			&nbsp;&nbsp;
			계산서과세구분
			<select name="taxtype" >
			<option value="">전체
			<option value="01" <%= CHKIIF(taxtype="01","selected","") %> >과세
			<option value="02" <%= CHKIIF(taxtype="02","selected","") %> >면세
			<option value="03" <%= CHKIIF(taxtype="03","selected","") %> >원천
			</select>
			&nbsp;&nbsp;
			<input type="checkbox" name="chkmargin" <%= CHKIIF(chkmargin="on","checked","") %>> 마진 <%= CUNDERMargin %> % 미만
			&nbsp;&nbsp;
			검색조건:
			<select class="select" name="searchType">
				<option></option>
				<option value="socname" <% if (searchType = "socname") then %>selected<% end if %> >업체명</option>
				<option value="socno" <% if (searchType = "socno") then %>selected<% end if %> >사업자번호</option>
			</select>
			&nbsp;
			<input type="text" class="text" name=searchText value="<%= searchText %>" size="15" maxlength="20">
        </td>
    </tr>
    <tr>
        <td bgcolor="#FFFFFF" >
        정산방식구분 :
        <% drawSelectBoxJGubun "jgubun",jgubun %>
        &nbsp;&nbsp;
        ON/AC 구분 :
        <select name="targetGbn" >
		<option value="">전체
		<option value="ON" <%= CHKIIF(targetGbn="ON","selected","") %> >ON
		<option value="AC" <%= CHKIIF(targetGbn="AC","selected","") %> >AC
		</select>
		&nbsp;&nbsp;
		차수
		<input type="text" class="text" name="differencekey" value="<%= differencekey %>" size="2" >
		&nbsp;&nbsp;
		* 텐바이텐 사업자 여부 : 
        <select name="companynoYN" class="select">
			<option value="">전체
			<option value="Y" <%= CHKIIF(companynoYN="Y","selected","") %> >사업자만
			<option value="N" <%= CHKIIF(companynoYN="N","selected","") %> >사업자제외
		</select>
		&nbsp;&nbsp;
		* 업체과세구분 : 
		<select name="jungsanGubun" class="select">
			<option value="" <% if jungsanGubun="" then response.write "selected" %>>전체</option>
			<option value="일반과세" <% if jungsanGubun="일반과세" then response.write "selected" %>>일반과세</option>
			<option value="간이과세" <% if jungsanGubun="간이과세" then response.write "selected" %>>간이과세</option>
			<option value="원천징수" <% if jungsanGubun="원천징수" then response.write "selected" %>>원천징수</option>
			<option value="면세" <% if jungsanGubun="면세" then response.write "selected" %>>면세</option>
			<option value="영세(해외)" <% if jungsanGubun="영세(해외)" then response.write "selected" %>>영세(해외)</option>
		</select>
        </td>
    </tr>
	</form>
</table>
<!-- 표 상단바 끝-->
<p>
<% if (designer<>"") and (yyyy1<>"") and (mm1<>"") then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="brandbatch" >
<input type="hidden" name="makerid" value="<%= designer %>">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
<tr bgcolor="#FFFFFF">
    <td>
        <select name="jgubun">
        <option value="">정산 방식 선택
        <option value="MM">매입
        <option value="CC">수수료
        <option value="CE">기타매출
        </select>
        <select name="differencekey">
        <option value="">차수 선택
        <option value="0">0차
        <option value="1">1차
        <option value="2">2차
        <option value="3">3차
        <option value="4">4차
        <option value="5">5차
        <option value="6">6차
        <option value="7">7차
        <option value="8">8차
        <option value="9">9차
        </select>
        <select name="itemvatYN">
        <option value="">상품 과세 구분 선택
        <option value="Y">과세
        <option value="N">면세
        </select>
        <input type="hidden" name="ipchulArr" value="">
        <input type="button" value=" <%= designer %> &nbsp;&nbsp;<%= yyyy1 %>년 <%= mm1 %>월 정산 작성 " onClick="MakeBrandBatchJungsan(document.brandbatch);">
    </td>
</form>
</tr>
</table>
<% end if %>
<% if taxtype="03" then %>
<input type="button" value="선택 정산확정" onclick="jsMultiStateChange();">
<% end if %>
<form name="frmList" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="idxarr" value="">
<input type="hidden" name="mode" value="multistatechange">
<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor=#BABABA>
    <tr bgcolor="#FFFFFF">
      <td colspan="30" >
      <%= page %>/<%= ojungsan.FTotalPage %> page 총 <%=ojungsan.FTotalCount %>건
      </td>
    </tr>
    <tr align="center" bgcolor="#DDDDFF">
      <% if taxtype="03" then %>
      <td width="70"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
      <% end if %>
      <td width="70">정산월</td>
      <td width="40">구분</td>
      <td width="50">정산<br>방식</td>
      <td width="50">계정<br>과목</td>
      <td width="30">차수</td>
      <td width="30">과세<br>(계산서)</td>
      <td width="30">과세<br>(상품)</td>
      <td width="90"><a href="javascript:research(frm,'designer')">브랜드ID</a></td>
      <td>회사명</td>
      <td width="60">업체배송</td>
      <td width="30">마진</td>
      <td width="60">매입총액</td>
      <td width="30">마진</td>
      <td width="60">위탁총액</td>
      <td width="30">마진</td>
      <td width="60">기타판매</td>
      <td width="30">마진</td>
      <td width="70">총수수료</td>
      <td width="70">배송비/기타</td>
      <td width="80">총정산액</td>
      <td width="80"><a href="javascript:research(frm,'state')">상태</a></td>
      <td width="70">세금계산서<br>등록일</td>
      <td width="70"><a href="javascript:research(frm,'segum')">세금발행일</a></td>
      <td width="70">입금일</td>
      <td width="20">E</td>
      <td width="20">S</td>
      <td width="50"><a href="javascript:research(frm,'tax')">과세구분</a></td>
      <td width="30">정산</td>
      <td width="30">비고</td>
    </tr>
<% if ojungsan.FResultCount<1 then %>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="30" align="center" height="30">
        <% if research="" then %>
            [검색 버튼을 눌러주세요.]
        <% else %>
            [검색 결과가 없습니다.]
        <% end if %>
        </td>
    </tr>
<% else %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    	tot1 = tot1 + ojungsan.FItemList(i).Fub_totalsuplycash
    	tot2 = tot2 + ojungsan.FItemList(i).Fme_totalsuplycash
    	tot3 = tot3 + ojungsan.FItemList(i).Fwi_totalsuplycash
    	tot4 = tot4 + ojungsan.FItemList(i).Fet_totalsuplycash
    	tot5 = tot5 + ojungsan.FItemList(i).Fsh_totalsuplycash

    	totcom = totcom + ojungsan.FItemList(i).Ftotalcommission
    	totdlv = totdlv + ojungsan.FItemList(i).Fdlv_totalsuplycash

    %>
   <tr align="center" bgcolor="#FFFFFF">
      <% if taxtype="03" then %>
      <td ><input type="checkbox" name="chkitem" value="<%= ojungsan.FItemList(i).FId %>"<% if ojungsan.FItemList(i).Ffinishflag="1" or ojungsan.FItemList(i).Ffinishflag="2" then %><% else %> disabled<% end if %>></td>
      <% end if %>
      <td ><a target=_blank href="nowjungsanmasteredit.asp?id=<%= ojungsan.FItemList(i).FId %>"><%= ojungsan.FItemList(i).Fyyyymm %>&nbsp;<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a></td>
      <td ><%= ojungsan.FItemList(i).FtargetGbn %></td>
      <td ><%= ojungsan.FItemList(i).getJGubunName %></td>
      <td ><%= ojungsan.FItemList(i).Fjacc_nm %></td>
      <td ><%= ojungsan.FItemList(i).Fdifferencekey %></td>
      <td ><font color="<%= ojungsan.FItemList(i).GetTaxtypeNameColor %>"><%= ojungsan.FItemList(i).GetSimpleTaxtypeName %></font></td>
      <td ><%= ojungsan.FItemList(i).GetItemVatTypeName %></td>
      <td ><a href="javascript:PopBrandInfoEdit('<%= ojungsan.FItemList(i).Fdesignerid %>')"><%= ojungsan.FItemList(i).Fdesignerid %></a></td>
      <td align="left"><a href="javascript:PopUpcheInfoEdit('<%= ojungsan.FItemList(i).FGroupID %>')"><%= ojungsan.FItemList(i).Fcompany_name %></a></td>
      <td align="right"><a target=_blank href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=upche"><%= FormatNumber(ojungsan.FItemList(i).Fub_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fub_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fub_totalsuplycash/ojungsan.FItemList(i).Fub_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>
      <td align="right"><a target=_blank href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=maeip"><%= FormatNumber(ojungsan.FItemList(i).Fme_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fme_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fme_totalsuplycash/ojungsan.FItemList(i).Fme_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>
      <td align="right"><a target=_blank href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=witaksell"><%= FormatNumber(ojungsan.FItemList(i).Fwi_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fwi_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fwi_totalsuplycash/ojungsan.FItemList(i).Fwi_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>
      <td align="right"><a target=_blank href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=witakchulgo"><%= FormatNumber(ojungsan.FItemList(i).Fet_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fet_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fet_totalsuplycash/ojungsan.FItemList(i).Fet_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Ftotalcommission,0) %></td>
      <td align="right"><a target=_blank href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=DL"><%= FormatNumber(ojungsan.FItemList(i).Fdlv_totalsuplycash,0) %></a></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
      <td ><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font>
	  <% if ojungsan.FItemList(i).Ffinishflag="0" then %>
      <a href="javascript:NextStep('<%= ojungsan.FItemList(i).FId %>');">
     <img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom">
      </a>
      <% end if %>
      </td>
	    <% if IsNULL(ojungsan.FItemList(i).Ftaxinputdate) then %>
	    <td ></td>
  	    <% else %>
 	    <td ><%= Left(Cstr(ojungsan.FItemList(i).Ftaxinputdate),10) %></td>
  	    <% end if %>
      <% if isNull(ojungsan.FItemList(i).Ftaxregdate) then %>
      <td ></td>
      <% else %>
      <td ><%= Left(Cstr(ojungsan.FItemList(i).Ftaxregdate),10) %></td>
      <% end if %>
      <% if isNull(ojungsan.FItemList(i).Fipkumdate) then %>
      <td ></td>
      <% else %>
      <td ><%= Left(Cstr(ojungsan.FItemList(i).Fipkumdate),10) %></td>
      <% end if %>
      <td ><a href="javascript:PopCSMailSend('<%= ojungsan.FItemList(i).FDesignerEmail %>','','');"><% if ojungsan.FItemList(i).FDesignerEmail<>"" then response.write "E" %></a></td>
      <td ><a href="javascript:PopCSSMSSend('<%= ojungsan.FItemList(i).Fjungsan_hp %>','','','');"><% if ojungsan.FItemList(i).Fjungsan_hp<>"" then response.write "S" %></a></td>
      <td ><%= ojungsan.FItemList(i).Fjungsan_gubun %></td>
      <td ><%= ojungsan.FItemList(i).Fjungsan_date %></td>
      <% if ojungsan.FItemList(i).Ffinishflag="0" then %>
      	<td ><a href="javascript:dellThis('<%= ojungsan.FItemList(i).FId %>')">x</a></td>
      <% else %>
        <td >
            <% if Not IsNULL(ojungsan.FItemList(i).FTaxLinkidx) then %>
      	        <img src="/images/icon_print02.gif" width="14" height="14" border=0 onclick="window.open('http://www.bill36524.com/popupBillTax.jsp?NO_TAX=<%= ojungsan.FItemList(i).Fneotaxno %>&NO_BIZ_NO=2118700620')" style="cursor:hand">
      	   <% else %>
      	        <%= ojungsan.FItemList(i).Fbillsitecode %>
      	    <% end if %>

      	    <a href="/admin/upchejungsan/monthjungsanAdm.asp?makerid=<%= ojungsan.FItemList(i).Fdesignerid %>&yyyy1=<%= LEFT(ojungsan.FItemList(i).Fyyyymm,4) %>&mm1=<%= right(ojungsan.FItemList(i).Fyyyymm,2) %>" target="_blank">POP</a>
        </td>
      <% end if %>
    </tr>
    <% next %>
<% end if %>
    <% totsum = totsum + tot1 + tot2 + tot3 + tot4 + tot5 +totdlv %>
    <tr bgcolor="#FFFFFF" align="right">
      <% if taxtype="03" then %>
      <td></td>
      <% end if %>
      <td>합계</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td><%= FormatNumber(tot1,0) %></td>
      <td></td>
      <td><%= FormatNumber(tot2,0) %></td>
      <td></td>
      <td><%= FormatNumber(tot3,0) %></td>
      <td></td>
      <td><%= FormatNumber(tot4,0) %></td>
      <td></td>
      <td><%= FormatNumber(totcom,0) %></td>
      <td><%= FormatNumber(totdlv,0) %></td>
      <td><%= FormatNumber(totsum,0) %></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr bgcolor="#FFFFFF" >
        <td colspan="30" align="center">
            <% if ojungsan.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojungsan.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for ix=0 + ojungsan.StarScrollPage to ojungsan.FScrollCount + ojungsan.StarScrollPage - 1 %>
				<% if ix>ojungsan.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
			<% next %>

			<% if ojungsan.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
    </tr>
</table>
</form>
<form name="frmarr" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="idx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="rd_state" value="">
</form>
<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
