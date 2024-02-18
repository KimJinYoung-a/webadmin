<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  보너스 쿠폰
' History : 서동석 생성
'			2022.07.04 한용민 수정(isms취약점수정)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/newcouponcls.asp" -->
<%
dim dispCate : dispCate = requestCheckvar(request("dispCate"),32)
dim idx, ocoupon
	idx = requestCheckvar(request("idx"),10)
	if idx="" then idx=0

set ocoupon = new CCouponMaster
	ocoupon.FRectIdx = idx

	if idx<>0 then
		ocoupon.GetOneCouponMaster ''GetCouponMasterList
	else
		set ocoupon.FOneItem = new CCouponMasterItem
	end if
%>

<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.auto.min.js"></script>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script lanuage='javscript'>
    let changeCouponImageFlag = false;

    $(function(){
        showTip();
         $("[name=isopenlistcoupon]").click(function(){
            showTip();
         })
    })

    function showTip(){
        if(document.frm.isopenlistcoupon.value == "N"){
            $("#tip").css("display","");
        }else{
            $("#tip").css("display","none");
        }
    }

    function submitForm(frm){
        if (frm.couponname.value.length<1){
            alert('쿠폰명을 입력하세요.');
            frm.couponname.focus();
            return;
        }

        //무료배송쿠폰 체크
        if (frm.isfreebeasongcoupon.checked){
    //        if ((!frm.coupontype[1].checked)||(frm.couponvalue.value*1!=2000)||(frm.minbuyprice.value*1!=0)){
    //            alert('무료배송 쿠폰일 경우 할인타입 - 원 할인금액 2000, 최소구매금액 0으로 설정 됩니다.');
    //
    //            frm.coupontype[1].checked = true;
    //            frm.couponvalue.value = 2000;
    //            frm.minbuyprice.value = 0;
    //            return;
    //        }
        }else{
            if ((!frm.coupontype[0].checked)&&(!frm.coupontype[1].checked)){
                alert('쿠폰 타입을 선택하세요.');
                frm.coupontype[0].focus();
                return;
            }

            if (frm.couponvalue.value.length<1){
                alert('할인 금액이나 %를 입력하세요.');
                frm.couponvalue.focus();
                return;
            }

            if (frm.minbuyprice.value.length<1){
                alert('최소 구매금액을 입력하세요.');
                frm.minbuyprice.focus();
                return;
            }
        }

        //카테고리/브랜드쿠폰의 경우 %쿠폰 선택불가 (잠정.)
        //2019-11-13 MD팀의 블랙프라이데이 이벤트 진행으로 인한 브랜드,카테고리 % 쿠폰 허용으로 변경
        if ((frm.targetcpntype.value=="B")||(frm.targetcpntype.value=="C")){
            if (frm.isfreebeasongcoupon.checked){
                alert('카테고리,브랜드 쿠폰은 무료배송 쿠폰을 선택할 수 없습니다.');
                return;
            }

            if (frm.coupontype[0].checked){
                if (!confirm('카테고리,브랜드 % 쿠폰적용은 상품마진에 따라 적용이 안될 수 있습니다.\n진행하시겠습니까?')){
                    return;
                }
                <% ' 프론트에 플러스세일일경우 브랜드쿠폰 정률쿠폰은 기능이 없음. 절대 열지 말것. %>
                alert('카테고리,브랜드 쿠폰은 % 쿠폰을 선택할 수 없습니다. 금액쿠폰을 사용하세요.');
                return;
            }
        }
        if ((frm.targetcpntype.value=="B")){
            if (frm.brandShareValue.value>50){
                alert('브랜드쿠폰의 업체 분담율은 50%를 넘을수 없습니다.');
                return;
            }
        }
        if ((frm.coupontype[0].checked)&&(frm.mxCpnDiscount.value.length<1)){
            alert('최대할인금액을 입력하세요.');
            frm.startdate.focus();
            return;
        }
        if (frm.startdate.value.length<1){
            alert('유효기간 시작일을 입력하세요.');
            frm.startdate.focus();
            return;
        }

        if (frm.expiredate.value.length<1){
            alert('유효기간 만료일을 입력하세요.');
            frm.expiredate.focus();
            return;
        }

        if (frm.openfinishdate.value.length<1){
            alert('쿠폰 발급 마감일을 입력하세요.');
            frm.openfinishdate.focus();
            return;
        }

        if ((frm.validsitename.value=="academy")||(frm.validsitename.value=="diyitem")){
            if (frm.isfreebeasongcoupon.checked){
                alert('핑거스 아카데미 쿠폰인경우 무료배송 쿠폰 체크 불가');
                frm.isfreebeasongcoupon.focus();
                return;
            }

            if (!confirm('핑거스 아카데미 쿠폰으로 선택하셨습니다. 계속하시겠습니까?')){
                return;
            }
        }

        var ret = confirm('저장 하시겠습니까?');

        if (ret){
            save_image().then(function(data){
               frm.submit();
            });
        }
    }

    function EnableBox(comp){
        if (comp.checked){
            frm.targetitemlist.disabled = false;
            frm.couponmeaipprice.disabled = false;

            frm.targetitemlist.style.backgroundColor = "#FFFFFF";
            frm.couponmeaipprice.style.backgroundColor = "#FFFFFF";
        }else{
            frm.targetitemlist.disabled = true;
            frm.couponmeaipprice.disabled = true;

            frm.targetitemlist.style.backgroundColor = "#E6E6E6";
            frm.couponmeaipprice.style.backgroundColor = "#E6E6E6";
        }

    }

    function disableType(comp){
        var frm = comp.form;
        if (comp.name=="isfreebeasongcoupon"){
            frm.couponvalue.disabled = comp.checked;
            frm.coupontype[0].disabled = comp.checked;
            frm.coupontype[1].disabled = comp.checked;
            //frm.minbuyprice.disabled = comp.checked;
            frm.mxCpnDiscount.disabled = comp.checked;

        }else if (comp.name=="targetcpntype"){
            if (comp.value=="C"){
                document.getElementById("brandSBtn").style.display = "none";
                document.getElementById("cateSelBtn").style.display = "block";
                frm.isfreebeasongcoupon.disabled = true;
            }else if (comp.value=="B"){
                document.getElementById("brandSBtn").style.display = "block";
                document.getElementById("cateSelBtn").style.display = "none";
                frm.isfreebeasongcoupon.disabled = true;
            }else{
                document.getElementById("brandSBtn").style.display = "none";
                document.getElementById("cateSelBtn").style.display = "none";
                frm.isfreebeasongcoupon.disabled = false;
            }
        }
        chkCpnType(frm);
    }

    function jsSearchDispCate(frmname,targetcompname, targetcpndtlnm){
        var dispCate = eval(frmname+'.'+targetcompname).value;
        var uri = '/common/module/popDispCateSelect.asp?dispCate='+dispCate+'&frmname='+frmname+'&targetcompname='+targetcompname+'&targetcpndtlnm='+targetcpndtlnm;
        var popwin = window.open(uri,'popDispCateSelect','width=800, height=400, scrollbars=yes, resizable=yes');
        popwin.focus();
    }

    function chkCpnType(o){
        var dctp = o.coupontype;
        var tgtp = o.targetcpntype;
        if (dctp.value=="1"&&tgtp.value==""){
            document.getElementById("imxcpndiscount_tr").style.display = "";
        }else{
            frm.mxCpnDiscount.value=0;
            document.getElementById("imxcpndiscount_tr").style.display = "none";
        }
    }

    $(document).ready(function(){
        $("#couponImageFile").change(function(event){
            const file = event.target.files[0];

            if (!file.type.match("image.*")) {
                $("#couponImageFile").val("");
                return alert("only image");
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);

            reader.onload = function(e){
                $("#couponImage").attr("src", e.target.result);
                $("#couponImageDiv").css("display", "block");
                $("#delete_image_button").css("display", "inline");
            }

            changeCouponImageFlag = true;
        });

        <% IF ocoupon.FOneItem.Fcouponimage <> "" THEN %>
            $("#couponImageDiv").css("display", "block");
            $("#delete_image_button").css("display", "block");
        <% END IF %>
    });

    function delete_image(){
        $("#couponImageFile").val("");
        $("#usercouponimage").val("");
        $("#couponImageDiv").css("display", "none");
        $("#delete_image_button").css("display", "none");
    }

    function save_image(){
        return new Promise(function (resolve, reject) {
            if("<%= ocoupon.FOneItem.FIdx %>" == "0"){
                alert("쿠폰코드가 없습니다. 쿠폰 등록 후 이미지를 올려주세요.");
                return reject();
            }

            if(changeCouponImageFlag){
                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = 'https://upload.10x10.co.kr';
                } else {
                    api_url = 'http://testupload.10x10.co.kr';
                }

                const imgData = new FormData();
                imgData.append('coupon_image', document.getElementById("couponImageFile").files[0]);
                imgData.append("coupon_code", "<%= ocoupon.FOneItem.FIdx %>");
                imgData.append("reg_year", "<%= LEFT(ocoupon.FOneItem.Fregdate, 4) %>");

                $.ajax({
                    url: api_url + "/linkweb/coupon/coupon_admin_imgreg_json.asp"
                    , type: "POST"
                    , processData: false
                    , contentType: false
                    , data: imgData
                    , crossDomain: true
                    , success: function (data) {
                        //console.log(data);
                        const response = JSON.parse(data);

                        $("input[name=usercouponimage]").val(response.coupon_image);

                        return resolve();
                    }
                    , error : function (request,status,error){
                        console.log("code", request.status);
                        console.log("message", request.responseText);
                        console.log("error", error);

                        return reject();
                    }
                });
            }else{
                return resolve();
            }
        });
    }
</script>

<form name="frm" method="post" action="/admin/sitemaster/docoupon.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="idx" value="<%= ocoupon.FOneItem.FIdx %>">
<table width="900" border="0" cellpadding="3" cellspacing="1" bgcolor=#3d3d3d class=a>
<tr>
	<td bgcolor="#DDDDFF" width="120">IDx</td>
	<td bgcolor="#FFFFFF"><%= ocoupon.FOneItem.FIdx %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">쿠폰명</td>
	<td bgcolor="#FFFFFF"><input type=text name=couponname value="<%= ocoupon.FOneItem.Fcouponname %>" maxlength="100" size=80>
	<br>(ex 텐바이텐 주말 쿠폰)</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">쿠폰 이미지(선택)</td>
	<td bgcolor="#FFFFFF">
        <%
            IF ocoupon.FOneItem.FIdx > 0 THEN
        %>
            <input type="file" id="couponImageFile" value="" />
            <input id="delete_image_button" type="button" value="삭제" onclick="delete_image()" style="display: none;"/>
            <div id="couponImageDiv" class="thumbnail-area" style="display: none;">
                <img id="couponImage" src="<%=ocoupon.FOneItem.FcouponimageUrl%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
            </div>
        <%
            ELSE
        %>
            <b><font color="red">※ 쿠폰이미지는 등록 후 수정 페이지에서 등록 가능합니다.</font></b>
        <%
            END IF
        %>

        <input type="text" name="usercouponimage" id="usercouponimage" value="<%=ocoupon.FOneItem.Fcouponimage%>" style="display: none"/>
    </td>
</tr>
<!-- 2018/01/18 추가 -->
<tr>
	<td bgcolor="#DDDDFF">쿠폰타입II</td>
	<td bgcolor="#FFFFFF">
	    <label style="margin-right:5px;"><input type="radio" name="targetcpntype" value=""  <%= CHKIIF(ocoupon.FOneItem.Ftargetcpntype="","checked","") %> onClick="disableType(this);">일반</label>
	    <label style="margin-right:5px;"><input type="radio" name="targetcpntype" value="B" <%= CHKIIF(ocoupon.FOneItem.Ftargetcpntype="B","checked","") %> onClick="disableType(this);" >브랜드쿠폰</label>
	    <label style="margin-right:6px;"><input type="radio" name="targetcpntype" value="C" <%= CHKIIF(ocoupon.FOneItem.Ftargetcpntype="C","checked","") %> onClick="disableType(this);" >카테고리쿠폰</label>

	    <div id="brandSBtn" <%= CHKIIF(ocoupon.FOneItem.Ftargetcpntype="B","style='display:block'","style='display:none'") %> >
	        <p>대상 브랜드 :
			<input type="text" name="targetcpnsourcebrand" value="<%=ocoupon.FOneItem.Ftargetcpnsource%>" size="20" maxlength="32" readonly >
	        <input type="button" class="button" value="브랜드검색" onclick="jsSearchBrandID(this.form.name,'targetcpnsourcebrand');">
			</p>
			<p>업체 분담율 :
				<input type="text" name="brandShareValue" value="<%=chkIIF(ocoupon.FOneItem.FbrandShareValue="","0",ocoupon.FOneItem.FbrandShareValue)%>" size="3" style="text-align:right;" /> %
			</p>
	    </div>
	    <div id="cateSelBtn" <%= CHKIIF(ocoupon.FOneItem.Ftargetcpntype="C","style='display:block'","style='display:none'") %> >
	        <input type="text" name="targetcpnsourcecate" value="<%=ocoupon.FOneItem.Ftargetcpnsource%>" size="20" maxlength="32" readonly >
	        <input type="text" name="targetcpndtlnm" value="<%=ocoupon.FOneItem.getTargetCateName%>" size="40"  readonly>
	        <input type="button" class="button" value="카테고리선택" onclick="jsSearchDispCate(this.form.name,'targetcpnsourcecate','targetcpndtlnm');" >
	    </div>
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">무료배송쿠폰<br>(텐바이텐배송)</td>
	<td bgcolor="#FFFFFF">
		<%
		''//[Fingers]사이트관리>>보너스쿠폰프로모션 페이지에서는 핑거스 쿠폰은 텐바이텐 무료 배송 안함. '현재 all 업배
		if menupos <> "1224" and menupos <> "1216" then
		%>
	    	<input type="checkbox" name="isfreebeasongcoupon" value="Y" <% if ocoupon.FOneItem.IsFreedeliverCoupon then response.write "checked" %> onClick="disableType(this);"> 무료배송쿠폰
	    <% else %>
	    	<input type="checkbox" name="isfreebeasongcoupon" value="Y" disabled <% if ocoupon.FOneItem.IsFreedeliverCoupon then response.write "checked" %> onClick="disableType(this);"> 무료배송쿠폰
	    <% end if %>
	    <!--
	    <br>
	    <input type="checkbox" name="isweekendcoupon" value="Y" <% if ocoupon.FOneItem.IsWeekendCoupon then response.write "checked" %> > 주말 쿠폰
        -->
	</td>
</tr>

<tr>
	<td bgcolor="#DDDDFF">할인타입</td>
	<td bgcolor="#FFFFFF">
		<input type=text name=couponvalue value="<%= ocoupon.FOneItem.Fcouponvalue %>" maxlength=7 size=10 <% if ocoupon.FOneItem.IsFreedeliverCoupon then response.write "disabled" %> >
	    <label style="margin-right:5px;"><input type="radio" name="coupontype" value="1" <%=chkIIF(ocoupon.FOneItem.IsFreedeliverCoupon,"disabled","")%> <%=chkIIF(ocoupon.FOneItem.Fcoupontype="1","checked","")%> onClick="chkCpnType(this.form)" />%할인</label>
	    <label style="margin-right:5px;"><input type="radio" name="coupontype" value="2" <%=chkIIF(ocoupon.FOneItem.IsFreedeliverCoupon,"disabled","")%> <%=chkIIF(ocoupon.FOneItem.Fcoupontype="2" or ocoupon.FOneItem.Fcoupontype="","checked","")%> onClick="chkCpnType(this.form)" />원할인</label>
		(금액 또는 % 할인)
	</td>
</tr>
<!--
<% if (FALSE) then %>
<tr>
	<td bgcolor="#DDDDFF" width="100">특정상품쿠폰</td>
	<% if ocoupon.FOneItem.IsTargetItemCoupon then %>
		<td bgcolor="#FFFFFF">
		특정상품 쿠폰 사용함: <input type=checkbox name=targetitemusing onclick="EnableBox(this)" checked ><br>
		상품번호: <input type=text name=targetitemlist value="<%= ocoupon.FOneItem.Ftargetitemlist %>" size=9 maxlength=9  >(특정 상품만 할인됨)
		&nbsp;&nbsp;
		쿠폰적용시 매입가: <input type=text name=couponmeaipprice value="<%= ocoupon.FOneItem.Fcouponmeaipprice %>" size=7 maxlength=9  >(업체부담할 경우 매입가 지정)
		</td>
	<% else %>
		<td bgcolor="#FFFFFF">
		특정상품 쿠폰 사용함: <input type=checkbox name=targetitemusing onclick="EnableBox(this)"><br>
		상품번호: <input type=text name=targetitemlist value="<%= ocoupon.FOneItem.Ftargetitemlist %>" size=9 maxlength=9 disabled style='background-color:#E6E6E6;'>(특정 상품만 할인됨)
		&nbsp;&nbsp;
		쿠폰적용시 매입가: <input type=text name=couponmeaipprice value="<%= ocoupon.FOneItem.Fcouponmeaipprice %>" size=7 maxlength=9 disabled style='background-color:#E6E6E6;'>(업체부담할 경우 매입가 지정)
		</td>
	<% end if %>
</tr>
<% end if %>
-->
<tr>
	<td bgcolor="#DDDDFF">최소구매금액</td>
	<td bgcolor="#FFFFFF"><input type="text" name="minbuyprice" value="<%= ocoupon.FOneItem.Fminbuyprice %>" maxlength="7" size="10" />원 이상 구매시 사용가능(숫자)</td>
</tr>
<tr id="imxcpndiscount_tr" <%=CHKIIF((ocoupon.FOneItem.Fcoupontype="1" and ocoupon.FOneItem.Ftargetcpntype="") or ocoupon.FOneItem.FmxCpnDiscount>0,"style='display:'","style='display:none'")%>>
	<td bgcolor="#DDDDFF">최대할인금액</td>
	<td bgcolor="#FFFFFF"><input type="text" name="mxCpnDiscount" value="<%= ocoupon.FOneItem.FmxCpnDiscount %>" maxlength="7" size="10" />원 할인(숫자)(ex 5% 시 10000 / 10%시 20000 / 무제한 0 입력)</td>
</tr>

<tr>
	<td bgcolor="#DDDDFF">유효기간</td>
	<td bgcolor="#FFFFFF">
	    <input type=text name=startdate value="<%= ocoupon.FOneItem.Fstartdate %>" maxlength=19 size=20>~<input type=text name=expiredate value="<%= ocoupon.FOneItem.Fexpiredate %>" maxlength=19 size=20>
	    (<%= Left(now(),10) %> 00:00:00 ~ <%= Left(now(),10) %> 23:59:59)
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">쿠폰발급마감일</td>
	<td bgcolor="#FFFFFF"><input type=text name="openfinishdate" value="<%= ocoupon.FOneItem.Fopenfinishdate %>" maxlength=19 size=20>(2004-04-31 23:59:59)</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">사용처</td>
	<td bgcolor="#FFFFFF">
		<label style="margin-right:5px;"><input type="radio" name="validsitename" value="" <%= CHKIIF(ocoupon.FOneItem.Fvalidsitename="","checked","") %> >텐바이텐 보너스 쿠폰</label>
		<!-- 중지
		<label><input type="radio" name="validsitename" value="academy" <%= CHKIIF(ocoupon.FOneItem.Fvalidsitename="academy","checked","") %> >핑거스 아카데미 강좌 쿠폰</label>
		<label><input type="radio" name="validsitename" value="diyitem" <%= CHKIIF(ocoupon.FOneItem.Fvalidsitename="diyitem","checked","") %> >핑거스 아카데미 상품 쿠폰</label>
		-->
		<label style="margin-right:5px;"><input type="radio" name="validsitename" value="mobile" <%= CHKIIF(ocoupon.FOneItem.Fvalidsitename="mobile","checked","") %> >모바일 보너스 쿠폰</label>
		<label style="margin-right:5px;"><input type="radio" name="validsitename" value="app" <%= CHKIIF(ocoupon.FOneItem.Fvalidsitename="app","checked","") %> >APP 보너스 쿠폰</label>
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">기타코멘트</td>
	<td bgcolor="#FFFFFF"><textarea name="etcstr" cols=80 rows=8><%= ReplaceBracket(ocoupon.FOneItem.Fetcstr) %></textarea></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">전체쿠폰여부</td>
	<td bgcolor="#FFFFFF">
		<label style="margin-right:5px;"><input type="radio" name="isopenlistcoupon" value="N" <% if ocoupon.FOneItem.Fisopenlistcoupon="N" then Response.Write "checked" %> />전체고객</label>
		<label style="margin-right:5px;"><input type="radio" name="isopenlistcoupon" value="Y" <% if ocoupon.FOneItem.Fisopenlistcoupon="Y" or ocoupon.FOneItem.Fisopenlistcoupon="" then Response.Write "checked" %> />선택고객(지정고객)</label>
		<label style="margin-right:5px;"><input type="radio" name="isopenlistcoupon" value="J" <% if ocoupon.FOneItem.Fisopenlistcoupon="J" then Response.Write "checked" %> />회원가입쿠폰</label>
		<label style="margin-right:5px;"><input type="radio" name="isopenlistcoupon" value="M" <% if ocoupon.FOneItem.Fisopenlistcoupon="M" then Response.Write "checked" %> />모바일 발행용</label>
		<div id="tip" style="color:red;display:none">**전체고객 선택시 해당 쿠폰은 유효기간 내 로그인시 무조건 발급됩니다.</div>
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">사용여부</td>
	<td bgcolor="#FFFFFF">
		<label style="margin-right:5px;"><input type="radio" name="isusing" value="Y" <%=chkIIF(ocoupon.FOneItem.FIsUsing="Y","checked","")%> />Y</label>
		<label style="margin-right:5px;"><input type="radio" name="isusing" value="N" <%=chkIIF(ocoupon.FOneItem.FIsUsing="N","checked","")%> />N</label>
	</td>
</tr>
<tr>
	<td colspan="2" align=center bgcolor="#FFFFFF"><input type=button value="저장" onClick="submitForm(frm);" class="button"></td>
</tr>
</table>
</form>
<%
set ocoupon = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->