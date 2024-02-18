<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 수수료 세금계산서 발행 배치
' Hieditor : 2019.04.04 서동석 생성
'            2021.02.03 한용민 수정(세금계산서 발행 빌36524 api -> 위하고 api 로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/util/htmllib_utf8.asp"-->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls_utf8.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<%
dim jungsan_gubun
Dim yyyy1 : yyyy1 = requestCheckvar(request("yyyy1"),4)
Dim mm1   : mm1   = requestCheckvar(request("mm1"),2)
Dim targetGbn : targetGbn = requestCheckVar(request("targetGbn"),10)
Dim jungsan_date : jungsan_date = requestCheckvar(request("jungsan_date"),10)
Dim itopn : itopn   = requestCheckvar(request("itopn"),10)
    jungsan_gubun = requestCheckvar(request("jungsan_gubun"),10)

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

Dim yyyymm : yyyymm = yyyy1&"-"&mm1


Dim research : research = requestCheckvar(request("research"),32)
Dim iMaxpageSize : iMaxpageSize=10000   '' 최대 허용 검색수 1만건

if (itopn="") then itopn=1000   '기본값 1천건
''dim ArrRows : ArrRows = getEvalCCBatchTargetList(yyyymm,targetGbn,jungsan_date,iMaxpageSize)

dim ojungsanTax
set ojungsanTax = new CUpcheJungsanTax
ojungsanTax.FCurrPage = 1
ojungsanTax.FPageSize = itopn
ojungsanTax.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTax.FRecttargetGbn = targetGbn
ojungsanTax.FRectJungsanDate = jungsan_date
ojungsanTax.FRectJungsanException = jungsan_gubun
ojungsanTax.getEvalJungsanTaxTargetListAdm


dim i, ttlCnt : ttlCnt = 0
ttlCnt = ojungsanTax.FResultCount

Dim rowErr : rowErr=0
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<script type='text/javascript'>
var batchstarted = false;
var nextid = 0;
var evalWinNm ="_evalWin" + (new Date()).getTime();
function EvalBatchStart(comp){
    if (batchstarted) return;

    comp.style="background-color: border: 1px solid #999999;#cccccc;color: #888888;"
    comp.value="발행중 ....."
    comp.disabled = true;
    batchstarted = true;

    addNotiLog("start");

    var popwin = window.open("" ,evalWinNm,"width=1200,height=768,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwin.focus();

    fnNextEvalProc();
}

function addNotiLog(ilog){
    document.getElementById("disp1").innerHTML += ilog+"<br>";
}

function addResultLog(orderSeq,ilog){
    document.getElementById("oseq_"+orderSeq).innerHTML = ilog;
}

function fnNextEvalProc(){
    var frm = document.frmBatchArr;
    if (!frm.ix){
        addNotiLog('내역이 없습니다.')
        return;
    }

    var ix = -1;
    if (!frm.ix.length){
        ix = frm.ix.value*1;
    }else{
        if (frm.ix.length>nextid){
            ix = frm.ix[nextid].value*1;
        }
    }

    if (nextid>ix){
        addNotiLog('finished !');
        setTimeout(function(){ alert('finished'); }, 100);  
        return;
    }

    if (nextid><%=iMaxpageSize%>){
       ddNotiLog('oops !');
       return;     
    }

    setTimeout(function(){
        oneEvalProc(ix);
    }, 100);  

    
}

function oneEvalProc(iidx){
    nextid = iidx+1;
    var arrfrm = document.frmBatchArr;

    if (!arrfrm.ix.length){
        if (arrfrm.rowErrNo.value*1>0){
            addResultLog(arrfrm.jidx.value,"skip");
            fnNextEvalProc();    
        }else{
            document.frmBatch.makerid.value = arrfrm.makerid.value;
            document.frmBatch.yyyy1.value = arrfrm.yyyy1.value;
            document.frmBatch.mm1.value = arrfrm.mm1.value;
            document.frmBatch.onoffGubun.value = arrfrm.onoffGubun.value;
            document.frmBatch.jidx.value = arrfrm.jidx.value;
            
            addNotiLog(document.frmBatch.onoffGubun.value + ":" + document.frmBatch.jidx.value);
            document.frmBatch.target=evalWinNm;
            document.frmBatch.submit();
        }
    }else{
        if (arrfrm.rowErrNo[iidx].value*1>0){
            addResultLog(arrfrm.jidx[iidx].value,"skip");
            fnNextEvalProc();
        }else{
            document.frmBatch.makerid.value = arrfrm.makerid[iidx].value;
            document.frmBatch.yyyy1.value = arrfrm.yyyy1[iidx].value;
            document.frmBatch.mm1.value = arrfrm.mm1[iidx].value;
            document.frmBatch.onoffGubun.value = arrfrm.onoffGubun[iidx].value;
            document.frmBatch.jidx.value = arrfrm.jidx[iidx].value;

            addNotiLog(document.frmBatch.onoffGubun.value + ":" + document.frmBatch.jidx.value);
            document.frmBatch.target=evalWinNm;
            document.frmBatch.submit();
        }
    }
    

    
}
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    * 정산 대상 년월 : <% DrawYMBox yyyy1,mm1 %>
        &nbsp;
        ON/AC/OF 구분 :
        <select name="targetGbn" >
		<option value="">전체
		<option value="ON" <%= CHKIIF(targetGbn="ON","selected","") %> >ON
		<option value="OF" <%= CHKIIF(targetGbn="OF","selected","") %> >OF
		<option value="AC" <%= CHKIIF(targetGbn="AC","selected","") %> >AC
		</select>

        &nbsp;
        정산일 :
        <select name="jungsan_date">
        <option value="" <% if jungsan_date="" then response.write "selected" %> >선택
        <option value="15일" <% if jungsan_date="15일" then response.write "selected" %> >15일
        <option value="말일" <% if jungsan_date="말일" then response.write "selected" %> >말일
        <option value="수시" <% if jungsan_date="수시" then response.write "selected" %> >수시
        </select>

        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        검색수 : 
        <select name="itopn">
        <option value="1" <% if itopn="1" then response.write "selected" %> >1
        <option value="10" <% if itopn="10" then response.write "selected" %> >10
        <option value="100" <% if itopn="100" then response.write "selected" %> >100
        <option value="1000" <% if itopn="1000" then response.write "selected" %> >1000
        <option value="3000" <% if itopn="3000" then response.write "selected" %> >3000
        <option value="5000" <% if itopn="5000" then response.write "selected" %> >5000
        <option value="10000" <% if itopn="10000" then response.write "selected" %> >10000
        </select>
        &nbsp;
        <input type="checkbox" name="jungsan_gubun"<% if jungsan_gubun="on" then response.write " checked"%>> 과세구분 영세(해외) 제외
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
    <td>
    <div id="disp1" style="overflow: auto; width:100%; height: 50px;"></div>
    </td>
    <td width="300">
    <iframe name="xLink3" id="xLink3" frameborder="0" width="100%" height="50"></iframe>
    </td>
</tr>
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="right">
        <input type="button" class="button" value="일괄발행" onClick="EvalBatchStart(this);">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<form name="frmBatchArr" style="margin:0px;" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		검색결과 : <b><%= ttlCnt %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60"></td>
    <td width="90">정산월</td>
    <td width="90">매출처</td>
    <td width="80">정산방식</td>
    <td width="60">계산서종류</td>

    <td width="60">그룹코드</td>
    <td width="100">브랜드ID</td>
    <td >내용</td>
    <td width="50">발행일</td>
    <td width="70">공급가</td>
    <td width="70">부가세</td>
    <td width="70">합계</td>
    <td width="60">상태</td>
    <td >비고</td>
 </tr>
 <% if ojungsanTax.FResultCount>0 then %>
 <% For i =0 To ojungsanTax.FResultCount-1 %>
 <%
 rowErr = 0
 %>
 <input type="hidden" name="ix" value="<%=i%>">
 <input type="hidden" name="yyyy1" value="<%=LEFT(ojungsanTax.FItemList(i).FYYYYMM,4)%>">
 <input type="hidden" name="mm1" value="<%=Right(ojungsanTax.FItemList(i).FYYYYMM,2)%>">
 <input type="hidden" name="onoffGubun" value="<%= ojungsanTax.FItemList(i).FtargetGbn %>">
 <input type="hidden" name="jidx" value="<%= ojungsanTax.FItemList(i).Fid %>">
 <input type="hidden" name="makerid" value="<%= ojungsanTax.FItemList(i).Fmakerid %>">
 <input type="hidden" name="rowErrNo" value="<%=rowErr%>">

 <tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
    <td>
        <%= ojungsanTax.FItemList(i).Fid %>
    </td>
    <td><%=ojungsanTax.FItemList(i).FYYYYMM%></td>
    <td><%=ojungsanTax.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTax.FItemList(i).getTaxJungsanGubun%></td>
    <td><%=ojungsanTax.FItemList(i).getTaxTypeStrUpcheView%></td>
    <td><%=ojungsanTax.FItemList(i).Fgroupid%></td>
    <td><%=ojungsanTax.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Ftitle%></td>
    <td><%= ojungsanTax.FItemList(i).Ftaxregdate %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxSuply,0) %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxVat,0) %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxSum,0) %></td>
    <td><%= ojungsanTax.FItemList(i).GetTaxEvalStateName %></td>
    <td>
        <div id="oseq_<%=ojungsanTax.FItemList(i).Fid%>"></div>
    </td>
</tr>
<%
        '출력버퍼 중간 정리
        if ((i+1) mod 500) = 0 then
            response.flush
        end if
    Next
%>

<% else %>
<tr align="center" bgcolor="FFFFFF" >
    <td colspan="16">검색결과가 없습니다.</td>
</tr>
<% end if %>
</table>
</form>

<form name="frmBatch" method="get" action="/admin/upchejungsan/popUWehagotaxregapi.asp" style="margin:0px;" >
<input type="hidden" name="makerid" value="">
<input type="hidden" name="yyyy1" value="">
<input type="hidden" name="mm1" value="">
<input type="hidden" name="onoffGubun" value="">
<input type="hidden" name="jidx" value="">

<input type="hidden" name="isauto" value="on">
<input type="hidden" name="autotype" value="V2">
</form>
<%
public function getEvalCCBatchTargetList(yyyymm,targetGbn,JungsanDate,iPageSize)
    Dim sqlStr
    sqlStr = "db_jungsan.dbo.[usp_TEN_JungsanBatchEvalTarget] '"&yyyymm&"','"&targetGbn&"','"&JungsanDate&"',"&iPageSize&""
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        getEvalCCBatchTargetList = rsget.GetRows()
    end if
    rsget.close

end function

set ojungsanTax = Nothing

session.codePage = 949
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
