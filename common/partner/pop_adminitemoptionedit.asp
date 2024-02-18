<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 
' History : 최초생성자모름
'			2017.04.10 한용민 수정(보안관련처리)
'           2019.04.23 정태훈 옵션 삭제 못하도록 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<%
dim mode
dim itemid, itemoption
dim oitem, oitemoption, oOptionMultipleType, oOptionMultiple

itemid = requestCheckVar(request("itemid"),10)
if itemid="" then itemid=0
mode= requestCheckVar(request("mode"),32)
itemoption= requestCheckVar(request("itemoption"),4)

dim sqlStr
dim ErrStr

set oitem = new CItem
oitem.FRectItemID = itemid
if (C_IS_Maker_Upche) then
    oitem.FRectMakerid = session("ssBctid")
end if

if itemid<>"" then
	oitem.GetOneItem
end if

if (oitem.FResultCount<1) then 
    response.write "권한이 없습니다."
    dbget.close()	:	response.End
end if

set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

set oOptionMultipleType = new CItemOptionMultiple
oOptionMultipleType.FRectItemID = itemid 
if itemid<>"" then
    oOptionMultipleType.GetOptionTypeInfo
end if

set oOptionMultiple = new CitemOptionMultiple
oOptionMultiple.FRectItemID = itemid
if itemid<>"" then
    oOptionMultiple.GetOptionMultipleInfo
end if
        

dim i, j, k, TrFlag, pp
TrFlag = false
pp=0

dim maxcustomoptionno
maxcustomoptionno = 11
for i=0 to oitemoption.FResultCount - 1
    if IsNumeric(oitemoption.FItemlist(i).Fitemoption) then
        if (CInt(oitemoption.FItemlist(i).Fitemoption) < 100) then
            if (CInt(oitemoption.FItemlist(i).Fitemoption) > maxcustomoptionno) then
                maxcustomoptionno = CInt(oitemoption.FItemlist(i).Fitemoption)
            end if
        end if
    end if
next

dim ItemDefaultMargin
if oitem.FOneItem.Fsellcash>0 then
	ItemDefaultMargin = 100-CLng(oitem.FOneItem.FBuycash/oitem.FOneItem.Fsellcash*100*100)/100
else
	ItemDefaultMargin = 0
end if

''20091126수정 : 텐바이텐배송, 옵션이 없는상품이 재고디비에 있을경우 옵션추가 불가
'20150821 추가수정: 텐바이텐배송 재고 있거나 판매내역 있는 경우 옵션명 수정 관리자만 가능하도록 
dim OptionAddDisable : OptionAddDisable = false
dim OptionModDisable : OptionModDisable = false
 
if (oitem.FOneItem.FMwDiv<>"U") then
    sqlStr =  " select isNull(sum(CNT),0) as CNT, isNull(sum(TotCnt),0) as TotCnt "
    sqlStr = sqlStr & " from ( "
    sqlStr = sqlStr & "     select case when itemoption = '0000' then count(itemid) else 0 end as CNT"
    sqlStr = sqlStr & " , count(itemid) as TotCnt "
    sqlStr = sqlStr & " from db_summary.dbo.tbl_current_logisstock_summary"
    sqlStr = sqlStr & " where itemgubun='10'"
    sqlStr = sqlStr & " and itemid="&itemid
    sqlStr = sqlStr & "  group by itemid, itemoption "
    sqlStr = sqlStr & " ) as T "  
    rsget.Open sqlStr,dbget,1
        OptionAddDisable = rsget("CNT")>0
        OptionModDisable = rsget("TotCnt")>0 
        if  C_ADMIN_AUTH then 
         OptionModDisable = False
        end if  
         
    rsget.Close
    
end if

'/임시 분기 처리
if itemid="1521739" then OptionAddDisable = false
%>

<script type='text/javascript'>
var VItemDefaultMargin = <%= ItemDefaultMargin %>;
function EditOptionInfo(){
    var frm = document.frmEdit;
    var optAddpriceExists = false;
    
    if (frm.mode.value=="editOptionMultiple"){
        //이중옵션
        if (!frm.optionTypename.length){
            if (frm.optionTypename.value.length<1){
                alert('옵션 구분명을 입력하세요.');
                frm.optionTypename.focus();
                return;
            }
        }else{
            for (var i=0;i<frm.optionTypename.length;i++){
                if (frm.optionTypename[i].value.length<1){
                    alert('옵션 구분명을 입력하세요.');
                    frm.optionTypename[i].focus();
                    return;
                }
                
                //옵션구분명이 중복되는지 체크.
                for (var j=0;j<frm.optionTypename.length;j++){
                    if ((i!=j)&&(fnTrim(frm.optionTypename[i].value)==fnTrim(frm.optionTypename[j].value))){
                        alert('옵션 구분명을 중복하여 사용할 수 없습니다. - [' + frm.optionTypename[j].value + ']');
                        frm.optionTypename[j].focus();
                        return;
                    }
                }
            }
        }
        /*
        옵션명 변경 불가 처리 작업 (2019.04.29 정태훈)
        */
        if (!frm.optionName.length){
            if (frm.optionName.value.length<1){
                alert('옵션명을 입력하세요.');
                frm.optionName.focus();
                return;
            }
        }else{
            for (var i=0;i<frm.optionName.length;i++){
                if (frm.optionName[i].value.length<1){
                    alert('옵션명을 입력하세요.');
                    frm.optionName[i].focus();
                    return;
                }
                
                //옵션명이 중복되는지 체크.(이중옵션일때 옵션상세명 중복가능하므로 제외 : (frm.TypeSeq[i].value==frm.TypeSeq[j].value) 조건추가)
                for (var j=0;j<frm.optionName.length;j++){
                    if ((i!=j)&&(fnTrim(frm.optionName[i].value)==fnTrim(frm.optionName[j].value))&&(frm.TypeSeq[i].value==frm.TypeSeq[j].value)){
                        alert('옵션명을 중복하여 사용할 수 없습니다. - [' + frm.optionName[j].value + ']');
                        frm.optionName[j].focus();
                        return;
                    }
                }
            }
        }

        //추가금액
        if (!frm.optaddprice.length){
            if (frm.optaddprice.value.length<1){
                alert('추가금액을 입력하세요. (추가금액이 없으면 0)');
                frm.optaddprice.focus();
                return;
            }
            
            if (!IsDigit(frm.optaddprice.value)){
                alert('추가금액은 숫자만 가능합니다.');
                frm.optaddprice.focus();
                return;
            }
            
            if ((frm.optaddbuyprice.value*1)>(frm.optaddprice.value*1)) {
                alert('공급가가 매입가 보다 클 수 없습니다.');
                frm.optaddbuyprice.focus();
                return;
            }
            
            if ((frm.optaddprice.value*1>0) && (frm.optaddbuyprice.value*1!=parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100))) {
                if (!confirm('옵션 추가 금액에 대한 공급 금액이 상품 기본 마진 (<%= ItemDefaultMargin %>) 공급액(' + parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100) + '원) 과 일치 하지 않습니다. 계속 하시겠습니까?')){
                    frm.optaddbuyprice.focus();
                    return;
                }
            }
            
            optAddpriceExists = (optAddpriceExists||(frm.optaddprice.value*1>0));
        }else{
            for (var i=0;i<frm.optaddprice.length;i++){
                if (frm.optaddprice[i].value.length<1){
                    alert('추가금액을 입력하세요. (추가금액이 없으면 0)');
                    frm.optaddprice[i].focus();
                    return;
                }
                
                if (!IsDigit(frm.optaddprice[i].value)){
                    alert('추가금액은 숫자만 가능합니다.');
                    frm.optaddprice[i].focus();
                    return;
                }
                
                if ((frm.optaddbuyprice[i].value*1)>(frm.optaddprice[i].value*1)) {
                    alert('공급가가 매입가 보다 클 수 없습니다.');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
                
                if ((frm.optaddprice[i].value*1>0) && (frm.optaddbuyprice[i].value*1!=parseInt(frm.optaddprice[i].value*1*(100-VItemDefaultMargin)/100))) {
                    if (!confirm('옵션 추가 금액에 대한 공급 금액이 상품 기본 마진 (<%= ItemDefaultMargin %>) 공급액(' + parseInt(frm.optaddprice[i].value*1*(100-VItemDefaultMargin)/100) + '원) 과 일치 하지 않습니다. 계속 하시겠습니까?')){
                        frm.optaddbuyprice[i].focus();
                        return;
                    }
                }
                
                optAddpriceExists = (optAddpriceExists||(frm.optaddprice[i].value*1>0));
            }
        }
        
        //추가금액-공급가
        if (!frm.optaddbuyprice.length){
            if (frm.optaddbuyprice.value.length<1){
                alert('추가금액 공급가를 입력하세요. (추가금액이 없으면 0)');
                frm.optaddbuyprice.focus();
                return;
            }
            
            if (!IsDigit(frm.optaddbuyprice.value)){
                alert('추가금액 공급가는 숫자만 가능합니다.');
                frm.optaddbuyprice.focus();
                return;
            }
        }else{
            for (var i=0;i<frm.optaddbuyprice.length;i++){
                if (frm.optaddbuyprice[i].value.length<1){
                    alert('추가금액 공급가를 입력하세요. (추가금액이 없으면 0)');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
                
                if (!IsDigit(frm.optaddbuyprice[i].value)){
                    alert('추가금액 공급가는 숫자만 가능합니다.');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
            }
        }

        //추가금액 없는 기본옵션 존재여부 확인
        if (!frm.optaddbuyprice.length){
            if (frm.optaddprice.value>0){
                alert('옵션구분 내에 기본옵션이 필요합니다.\n추가금액이 없는(0원) 기본 옵션을 추가해주세요.');
                return;
            }
        }else{
            var chkPreTseq, chkBsOpt = false;
            for (var i=0;i<frm.optaddbuyprice.length;i++){
                if(chkPreTseq != frm.TypeSeq[i].value) chkBsOpt = false;
                chkPreTseq = frm.TypeSeq[i].value
                if (frm.optaddprice[i].value==0){
                    chkBsOpt = true;
                }
            }

            if(!chkBsOpt) {
                alert('옵션구분 내에 기본옵션이 필요합니다.\n추가금액이 없는(0원) 기본 옵션을 추가해주세요.');
                return;
            }
        }
    }else{
        //단일옵션
        if (frm.optionTypename.value.length<1){
            alert('옵션 구분명을 입력하세요.');
            frm.optionTypename.focus();
            return;
        }
        /*
        옵션명 변경 불가 처리 작업 (2019.04.29 정태훈)
        */
        if (!frm.optionName.length){
            if (frm.optionName.value.length<1){
                alert('옵션명을 입력하세요.');
                frm.optionName.focus();
                return;
            }
        }else{
            for (var i=0;i<frm.optionName.length;i++){
                if (frm.optionName[i].value.length<1){
                    alert('옵션명을 입력하세요.');
                    frm.optionName[i].focus();
                    return;
                }
                
                //옵션명이 중복되는지 체크.
                for (var j=0;j<frm.optionName.length;j++){
                    if ((i!=j)&&(frm.optionName[i].value==frm.optionName[j].value)){
                        alert('옵션명을 중복하여 사용할 수 없습니다. - [' + frm.optionName[j].value + ']');
                        frm.optionName[j].focus();
                        return;
                    }
                }
                
            }
        }

        //추가금액
        if (!frm.optaddprice.length){
            if (frm.optaddprice.value.length<1){
                alert('추가금액을 입력하세요. (추가금액이 없으면 0)');
                frm.optaddprice.focus();
                return;
            }
            
            if (!IsDigit(frm.optaddprice.value)){
                alert('추가금액은 숫자만 가능합니다.');
                frm.optaddprice.focus();
                return;
            }
            
            if ((frm.optaddbuyprice.value*1)>(frm.optaddprice.value*1)) {
                alert('공급가가 매입가 보다 클 수 없습니다.');
                frm.optaddbuyprice.focus();
                return;
            }
            
            if ((frm.optaddprice.value*1>0) && (frm.optaddbuyprice.value*1!=parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100))) {
                if (!confirm('옵션 추가 금액에 대한 공급 금액이 상품 기본 마진 (<%= ItemDefaultMargin %>) 공급액(' + parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100) + '원) 과 일치 하지 않습니다. 계속 하시겠습니까?')){
                    frm.optaddbuyprice.focus();
                    return;
                }
            }
            
            optAddpriceExists = (optAddpriceExists||(frm.optaddprice.value*1>0));
        }else{
            for (var i=0;i<frm.optaddprice.length;i++){
                if (frm.optaddprice[i].value.length<1){
                    alert('추가금액을 입력하세요. (추가금액이 없으면 0)');
                    frm.optaddprice[i].focus();
                    return;
                }
                
                if (!IsDigit(frm.optaddprice[i].value)){
                    alert('추가금액은 숫자만 가능합니다.');
                    frm.optaddprice[i].focus();
                    return;
                }
                
                if ((frm.optaddbuyprice[i].value*1)>(frm.optaddprice[i].value*1)) {
                    alert('공급가가 매입가 보다 클 수 없습니다.');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
                
                if ((frm.optaddprice[i].value*1>0) && (frm.optaddbuyprice[i].value*1!=parseInt(frm.optaddprice[i].value*1*(100-VItemDefaultMargin)/100))) {
                    if (!confirm('옵션 추가 금액에 대한 공급 금액이 상품 기본 마진 (<%= ItemDefaultMargin %>) 공급액(' + parseInt(frm.optaddprice[i].value*1*(100-VItemDefaultMargin)/100) + '원) 과 일치 하지 않습니다. 계속 하시겠습니까?')){
                        frm.optaddbuyprice[i].focus();
                        return;
                    }
                }
                
                optAddpriceExists = (optAddpriceExists||(frm.optaddprice[i].value*1>0));
            }
        }
        
        //추가금액-공급가
        if (!frm.optaddbuyprice.length){
            if (frm.optaddbuyprice.value.length<1){
                alert('추가금액 공급가를 입력하세요. (추가금액이 없으면 0)');
                frm.optaddbuyprice.focus();
                return;
            }
            
            if (!IsDigit(frm.optaddbuyprice.value)){
                alert('추가금액 공급가는 숫자만 가능합니다.');
                frm.optaddbuyprice.focus();
                return;
            }
        }else{
            for (var i=0;i<frm.optaddbuyprice.length;i++){
                if (frm.optaddbuyprice[i].value.length<1){
                    alert('추가금액 공급가를 입력하세요. (추가금액이 없으면 0)');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
                
                if (!IsDigit(frm.optaddbuyprice[i].value)){
                    alert('추가금액 공급가는 숫자만 가능합니다.');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
            }
        }

        //추가금액 없는 기본옵션 존재여부 확인
        if (!frm.optaddbuyprice.length){
            if (frm.optaddprice.value>0){
                alert('기본옵션이 필요합니다.\n추가금액이 없는(0원) 기본 옵션을 추가해주세요.');
                return;
            }
        }else{
            var chkBsOpt = false;
            for (var i=0;i<frm.optaddbuyprice.length;i++){
                if (frm.optaddprice[i].value==0){
                    chkBsOpt = true;
                }
            }

            if(!chkBsOpt) {
                alert('기본옵션이 필요합니다.\n추가금액이 없는(0원) 기본 옵션을 추가해주세요.');
                return;
            }
        }
    }
    
    <% if (oitem.FOneItem.FMwDiv<>"U") then %>
    //텐배송은 옵션 추가금액 사용불가 하게 20120326
    <% if NOT ((session("ssBctID")="icommang") or (session("ssBctID")="hrkang97")) then %>  //201509/01
    if (optAddpriceExists){
        alert('텐바이텐 배송의 경우 옵션 추가금액을 사용할 수 없습니다.');
        return;
    }
    <% else %>
    if (optAddpriceExists){
        alert('관리자 수정 MODE.');
        
    }
    <% end if %>
    <% end if %>
    
     if (optAddpriceExists){
    	var isOversea = "<%=oitem.FOneItem.FdeliverOverseas%>";
    	if (isOversea=="Y"){
    		   alert('해외배송을 하는 경우 옵션 추가금액을 사용할 수 없습니다.');
        return;
    	}
    }
    
    if (confirm('수정 하시겠습니까?')){
        frm.submit();
    }
}


function SaveOption(){
	var frm;
	var upfrm = document.frmarr;

	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;
					upfrm.itemoption.value = upfrm.itemoption.value + "|" + frm.itemoption.value;
					if (frm.isusing[0].checked==true){
						upfrm.isusing.value = upfrm.isusing.value + "|" + "Y";
					}else{
						upfrm.isusing.value = upfrm.isusing.value + "|" + "N";
					}
			}
		}
		upfrm.mode.value = "modiitemoptionarr";
		upfrm.submit();
	}
}

function DelItemOption(itemid,itemoption){
    var frm = document.frmOption;
    
	if (confirm('상품 구성이 변경되지 않는한 삭제하지 마시기 바랍니다. \n\n정말 삭제 하시겠습니까?')){
		frm.mode.value = "deleteoption";
		frm.itemid.value = itemid;
		frm.itemoption.value = itemoption;
		//frm.submit();
	}
}


function DelItemOptionMultiple(itemid,typeseq,kindseq){
    var frm = document.frmOption;

	//마지막 구분인지 확인
	if($("input[name='TypeSeq']").last().val()>typeseq) {
		if($("input[name='TypeSeq'][value='"+typeseq+"']").length<=1) {
			alert("뒷 순번이 존재하는 앞 순번의 옵션구분은 삭제할 수 없습니다.");
			return;
		} else {
			alert("앞 순번의 옵션구분은 삭제할 수 없으니 참고하세요.");
		}
	}

    if (confirm('상품 구성이 변경되지 않는한 삭제하지 마시기 바랍니다. \n\n정말 삭제 하시겠습니까?')){
		frm.mode.value = "deleteMultipleOption";
		frm.itemid.value = itemid;
		frm.typeseq.value = typeseq;
		frm.kindseq.value = kindseq;
		//frm.submit();
	}
}

function AutoCalcuBuyPrice(comp,j){
    var frm = document.frmEdit;
    
    if (!frm.optaddbuyprice.length){
        frm.optaddbuyprice.value = parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100);
    }else{
        frm.optaddbuyprice[j].value = parseInt(frm.optaddprice[j].value*1*(100-VItemDefaultMargin)/100);
    }

}

// ============================================================================


function AddOptionPop(iitemid){
    <% if (OptionAddDisable) then %>
    	alert('텐바이텐 배송 옵션 없는상품이 재고에 있으므로 옵션추가 불가합니다.');
    	return;
    <% else %>
	    var popwin = window.open('pop_optionAdd.asp?itemid=' + iitemid,'pop_optionAdd','width=1024,height=700,scrollbars=yes,resizable=yes');
	    popwin.focus();
    <% end if %>
}
</script>
</head>
<body>
<div class="popupWrap">
	<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="창닫기" onclick="window.close();" /></p>
	</div>
	<div class="popContent scrl">
		<div class="contTit bgNone"><!-- for dev msg : 타이틀 영역하단에 searchWrap이 올 경우엔 bgNone 클래스 삭제 -->
			<h2>옵션수정</h2>  
			<ul class="txtList">
				<li><span style="color:red">옵션은 삭제할 수 없습니다.(사용안함 으로 수정하세요)</span></li>
				<li>판매/입고/출고된 내역이 있는 옵션은 삭제가 불가능합니다.(사용안함 으로 수정하세요)</li>
				<%if (oitem.FOneItem.FMwDiv<>"U") then %>
				<li>판매/입고/출고된 내역이 있는 옵션의 옵션명 수정은 담당엠디에게 문의주세요</li>
				<%end if%>
				<li>추가금액이 있을경우 자동으로 표시됩니다.(옵션 명에 추가금액을 넣지 마세요)</li>
			</ul>
 		</div>
		<div class="cont">  
		  <table class="tbType1 writeTb tMar10">
					<colgroup>
						<col width="15%" /><col width="" />
					</colgroup>
					<tbody>
					<tr>
						<th><div>상품코드</div></th>
						<td><%= itemid %></td>
						<th><div>옵션 선택 미리보기</div></th>
				</tr>
				<tr>
					<th><div>상품명</div></th>
					<td><%= oitem.FOneItem.Fitemname %></td>
					<td  rowspan="2" align="center" style="background-color:#FFF2E6">
						<%= getOptionBoxHTML_FrontType(itemid) %>
					</td>
				</tr>
				<tr>
					<th><div>브랜드</div></th>
					<td><%= oitem.FOneItem.Fmakerid %> (<%= oitem.FOneItem.FBrandName %>)</td>
				</tr>
			</table> 
		<div class="tPad20">		
	   <div class="overHidden">
				<div class="ftLt"><h3>등록된 옵션 리스트</h3></div>
		    <div class="ftRt"><input type="button" class="btn" value="옵션추가 +" onClick="AddOptionPop('<%= itemid %>');"></div>
		  </div>  
		</div>
		  <div class="tPad10">
		  	<form name="frmEdit" method="post" action="do_adminitemoptionedit.asp">
				<input type="hidden" name="itemid" value="<%= itemid %>">
				<% if (oitemoption.IsMultipleOption) then %>
				<input type="hidden" name="mode" value="editOptionMultiple">
				<% else %>
				<input type="hidden" name="mode" value="editOption">
				<% end if %> 
		   <table class="tbType1 listTb">
	<% if oitemoption.FResultCount<1 then %>
    		<tr >
	    		<td colspan="8" >등록된 옵션이 없습니다.</td>
    		</tr>
    <% else %>
        <% if (oitemoption.IsMultipleOption) then %>
        <!-- 이중옵션 -->
        <tr>
        	<th><div>순번</div></th>
        	<th><div>옵션구분명</div></th>
        	<th><div>옵션상세명</div></th> 
        	<th><div>추가가격</div></th> 
        	<th><div>공급가</div></th>
        </tr>
        <% for i=0 to oOptionMultipleType.FResultCount-1 %>
    	<tr >  
    	    <input type="hidden" name="TypeSeqTmp" value="<%= oOptionMultipleType.FItemList(i).FTypeSeq %>">
        	<td rowspan="<%= oOptionMultipleType.FItemList(i).FOptionCount %>" width="30"><%= i+1 %></td>
        	<td rowspan="<%= oOptionMultipleType.FItemList(i).FOptionCount %>"> 
        	    <input type="text" name="optionTypename" value="<%= oOptionMultipleType.FItemList(i).FoptionTypename %>" size="20" maxlength="20" <%if OptionModDisable  then%>readonly class="formTxt readonly" <%else%>class="formTxt" <%end if%>>
        	</td>
            <% TrFlag = false %>
        	<% for k=0 to oOptionMultiple.FResultCount -1 %>
        	<% if (oOptionMultipleType.FItemList(i).FoptionTypename=oOptionMultiple.FItemList(k).FoptionTypename) and (oOptionMultipleType.FItemList(i).FTypeSeq=oOptionMultiple.FItemList(k).FTypeSeq) then %>
        	<% if (TrFlag) then %>
        </tr>
        <tr >
            <% end if %>
            <input type="hidden" name="TypeSeq" value="<%= oOptionMultiple.FItemList(k).FTypeSeq %>">
            <input type="hidden" name="KindSeq" value="<%= oOptionMultiple.FItemList(k).FKindSeq %>">
        	<td><input type="text"  name="optionName" value="<%= oOptionMultiple.FItemList(k).FoptionKindName %>" size="20" maxlength="20" <%if OptionModDisable then%>readonly class="formTxt readonly" <%else%>class="formTxt" <%end if%>></td>
        	<!-- <td></td> -->
        	<td><input type="text" class="formTxt" name="optaddprice" value="<%= oOptionMultiple.FItemList(k).Foptaddprice %>" size="9" maxlength="9" style="text-align:right" onKeyUp="AutoCalcuBuyPrice(this,'<%= pp %>');"></td>
        	<td><input type="text" class="formTxt" name="optaddbuyprice" value="<%= oOptionMultiple.FItemList(k).Foptaddbuyprice %>" size="9" maxlength="9" style="text-align:right"></td>
        </tr>
            <% pp = pp + 1 %>
            <% TrFlag = true %>
        	<% end if %>
        	<% next %>
    	<% next %>
	    <% else %>
	    <!-- 단일옵션  -->
	    <tr>
        	<th><div>옵션구분명</div></th>
        	<th><div>옵션상세명</div></th>
        	<th><div>사용<br>여부</div></th>
        	<th><div>품절<br>여부</div></th>
        	<th><div>추가가격</div></th>
        	<th><div>공급가</div></th>
        </tr>
	    <tr>
        	<td rowspan="<%= oitemoption.FResultCount %>">
        	    <input type="formTxt"  name="optionTypename" value="<%= oitemoption.FItemList(0).FoptionTypename %>" size="20" maxlength="20" <%if OptionModDisable  then%>readonly class="formTxt readonly" <%else%>class="formTxt" <%end if%>>
        	</td>
        	<% TrFlag = false %>
        	<% for k=0 to oitemoption.FResultCount -1 %>
        	<% if (TrFlag) then %>
        </tr>
        <tr align="center" bgcolor="<%= ChkIIF(oitemoption.FItemList(k).Foptisusing="Y","#FFFFFF","#DDDDDD") %>">
            <% end if %>
            <input type="hidden" name="itemoption" value="<%= oitemoption.FItemList(k).FItemOption %>">
        	<td><input type="text"  name="optionName" value="<%= oitemoption.FItemList(k).FoptionName %>" size="20" maxlength="20" <%if OptionModDisable then%>readonly class="formTxt readonly" <%else%>class="formTxt" <%end if%>></td>
        	<td><font color="<%= ChkIIF(oitemoption.FItemList(k).Foptisusing="Y","#000000","#FF0000") %>"><%= oitemoption.FItemList(k).Foptisusing %></font></td>
        	<td><% if oitemoption.FItemList(k).IsOptionSoldOut then %><font color="red">품절</font><% end if %></td>
        	<td><input type="text" class="formTxt" name="optaddprice" value="<%= oitemoption.FItemList(k).Foptaddprice %>" size="9" maxlength="9" style="text-align:right" onKeyUp="AutoCalcuBuyPrice(this,'<%= pp %>');"></td>
        	<td><input type="text" class="formTxt" name="optaddbuyprice" value="<%= oitemoption.FItemList(k).Foptaddbuyprice %>" size="9" maxlength="9" style="text-align:right"></td>
            <% pp = pp + 1 %>
        </tr>
            <% TrFlag = true %>
        	<% next %>
        </tr>
    	<% end if %>
	<% end if %>
	</table>
</form>
</div>

<p>
<% if oitemoption.FResultCount>0 then %>
			<div class="tPad15 ct"> 
					<input type="button" value=" 옵션 내용 수정 " onclick="EditOptionInfo()" class="btn3 btnRd" />
			</div>  
<% end if %>

<form name="frmOption" method="post" action="do_adminitemoptionedit.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid">
<input type="hidden" name="itemoption">
<input type="hidden" name="typeseq">
<input type="hidden" name="kindseq">
</form>
<%
set oitem = Nothing
set oOptionMultipleType = Nothing
set oOptionMultiple = Nothing
set oitemoption = Nothing
%>
	</div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->