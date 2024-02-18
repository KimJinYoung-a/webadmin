<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<%
    dim imageUrl : imageUrl = requestCheckvar(request("imageUrl"),200)
    dim xPosition : xPosition = requestCheckVar(Request("xPo"),4)
    dim yPosition : yPosition = requestCheckVar(Request("yPo"),4)
    dim imageNumber : imageNumber = requestCheckVar(Request("imageNumber"),4)
    dim itemid : itemid = requestCheckVar(Request("itemid"),9)
    dim info : info = Request("info")

    Dim strSql, orgPrice, finalPrice, salePercent, brandName
    strSql =    "SELECT" + vbcrlf _
                    & "  I.orgprice" + vbcrlf _
                    & ", CASE WHEN I.itemcouponyn = 'Y' THEN" + vbcrlf _
                    & "     CASE WHEN I.itemcoupontype = 1 THEN I.sellcash - (I.sellcash * I.itemcouponvalue / 100)" + vbcrlf _
                    & "     WHEN I.itemcoupontype = 2 THEN I.sellcash - I.itemcouponvalue" + vbcrlf _
                    & "     ELSE I.sellcash END" + vbcrlf _
                    & "  ELSE I.sellcash END as finalprice" + vbcrlf _
                    & ", B.socname as brandname" + vbcrlf _
                & "FROM db_item.dbo.tbl_item I" + vbcrlf _
                & "LEFT JOIN db_user.dbo.tbl_user_c B ON I.makerid = B.userid" + vbcrlf _
                & "WHERE I.itemid = '" & itemid & "'"
    rsget.Open strSql,dbget
        IF not rsget.EOF THEN
            orgPrice = rsget("orgprice")
            finalPrice = rsget("finalprice")
            salePercent = FormatNumber((orgPrice - finalPrice)/orgPrice * 100, 0)
            brandName = rsget("brandname")
        End IF
    rsget.Close

    if xPosition = "" then xPosition = ""
    if yPosition = "" then yPosition = ""
%>
<link rel="stylesheet" type="text/css" href="/webfonts/CoreSansC.css">
<style>
.map-wrap {position:relative; overflow:hidden; display:inline-block; cursor:pointer;}
.map-wrap .mark {position:absolute; left:100%; top:100%; width:10px; transform:translate(-50%,-50%);}
#xPosition, #yPosition {width: 50px; padding: 5px; text-align: center;}
</style>
<script>
    // 기본 노출 정보들(pop_event_customgroupItem2.asp Line 162 ~ 179)
    const info = JSON.parse(decodeURI(decodeURIComponent('<%=info%>')));
    window.addEventListener('load', function() {
        document.getElementById('eventProduct').classList.add('type' + info.basic.type);
    });
</script>
<div class="popV19">
	<div class="popHeadV19">
		<h1>이미지 맵 링크</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col>
			</colgroup>
			<tbody>
                <tr>
                    <td style="padding-left:5px">
                        <div id="eventProduct" class="map-wrap evt-itemV20">
                            <img src="<%=imageUrl%>"  id="myImg">
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>
                        X : <input type="text" name="xPosition" id="xPosition" value="<%=xPosition%>"/> %
                        Y : <input type="text" name="yPosition" id="yPosition" value="<%=yPosition%>"/> %
                    </td>
                </tr>
			</tbody>
        </table>
    </div>
	<div class="popBtnWrapV19">
        <button class="btn4 btnBlue1" onClick="saveMapLocation();return false;">저장</button>
	</div>
</div>
<script>
$(document).ready(function () {
    $('.map-wrap').append(getPriceHtml('<%=xPosition%>', '<%=yPosition%>'));

    // 이미지 클릭 시 좌표값 등록
    $('#myImg').click(function(e) {   
        e.preventDefault();
        var posX = Math.round(e.offsetX / $(this).width() * 100);
        var posY = Math.round(e.offsetY / $(this).height() * 100) ;

        $('.price-info').hide();
        $('.map-wrap').append(getPriceHtml(posX, posY));

        $('#xPosition').val(posX);
        $('#yPosition').val(posY);
    });
});
function getPriceHtml(xPosition, yPosition) {
    // 브랜드명 HTML
    const brandHtml = info.show.brandname === 'Y' ? 
        `<span class='brand' ${info.color.item_and_brand_name !== '' ? 'style="color:' + info.color.item_and_brand_name + ';"' : ''}><%=brandName%></span>` : ``;
    // 상품명 HTML
    const itemNameHtml = info.show.itemname === 'Y' ?
        `<p class='name' ${info.color.item_and_brand_name !== '' ? 'style="color:' + info.color.item_and_brand_name + ';"' : ''}>${info.basic.itemname}</p>` : ``;
    // 가격 HTML
    const priceHtml = info.show.price === 'Y' ?
        `<div class='price'>
            <p class='origin-price' ${info.color.org_price !== '' ? 'style="color:' + info.color.org_price + ';"' : ''}><%=FormatNumber(orgPrice,0)%></p>
            <b class='discount' ${info.color.sale_percent !== '' ? 'style="color:' + info.color.sale_percent + ';"' : ''}><%=salePercent%>%</b>
            <span class='sum' ${info.color.price !== '' ? 'style="color:' + info.color.price + ';"' : ''}><%=FormatNumber(finalPrice,0)%>원</span>
        </div>` : ``;

    return `
        <div class='desc price-info' style='left:${xPosition}%;top:${yPosition}%;position: absolute;'>
            ${brandHtml}
            ${itemNameHtml}
            ${priceHtml}
        </div>
    `;
}

function saveMapLocation(){
    var xPo = "xPosition<%=imageNumber%>"
    var yPo = "yPosition<%=imageNumber%>"

    window.document.domain = "10x10.co.kr";
    window.opener.document.getElementById(xPo).value = $('#xPosition').val();
    window.opener.document.getElementById(yPo).value = $('#yPosition').val();

    self.close();
}

function popupAutoResize() {
   var thisX = parseInt(document.body.scrollWidth);
   var thisY = parseInt(document.body.scrollHeight);
   var maxThisX = screen.width - 50;
   var maxThisY = screen.height - 50;
   var marginY = 0;
 

   if (navigator.userAgent.indexOf("MSIE 6") > 0) marginY = 45;        // IE 6.x
   else if(navigator.userAgent.indexOf("Firefox") > 0) marginY = 50;   // FF
   else if(navigator.userAgent.indexOf("Opera") > 0) marginY = 30;     // Opera
   else if(navigator.userAgent.indexOf("Netscape") > 0) marginY = -2;  // Netscape 
   else marginY = 70;

   if (navigator.userAgent.indexOf("MSIE 6") > 0) marginX = 40;        // IE 6.x
   else if (navigator.userAgent.indexOf("MSIE 7") > 0) marginX = 40;        // IE 7.x
   else marginX = 20;

   if (thisX > maxThisX) {
       window.document.body.scroll = "yes";
       thisX = maxThisX;
   }

   if (thisY > maxThisY - marginY) {
       window.document.body.scroll = "yes";
       thisX += 19;
       thisY = maxThisY - marginY;
   }

   window.resizeTo(thisX+marginX, thisY+marginY);
}

popupAutoResize();
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->