<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/event/imageLinkCls.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<%
Dim masterIdx, i
	masterIdx = requestCheckvar(request("idx"),16)

    If masterIdx = "" Then
        response.write "<script>alert('정상적인 경로로 접근해 주세요');history.back();</script>"
        response.end
    End If

	dim oLinkContents
		set oLinkContents = new CimageLink
		oLinkContents.FRectIdx = masterIdx
		oLinkContents.GetMasterContents

    dim oLinkDetailContents
        set oLinkDetailContents = new CimageLink
        oLinkDetailContents.FRectMasterIdx = masterIdx
        oLinkDetailContents.GetLinkListContents()

%>
<style type="text/css">
.map-wrap {position:relative; overflow:hidden; display:inline-block; height:50vh; cursor:pointer;}
.map-wrap .mark {position:absolute; left:100%; top:100%; width:10px; transform:translate(-50%,-50%);}
</style>
<script type="text/javascript" src="http://davidlynch.org/projects/maphilight/jquery.maphilight.min.js"></script>
<script>
var _ClickCnt=1;
var _X1value=0;
var _Y1value=0;
_ClickCnt=1;
function getLoc(){
    var x = event.offsetX;
    var y = event.offsetY;
    $('.map-wrap').append('<span class="mark" id="markposition_'+x+y+'" style="left:' + x + 'px; top:' + y + 'px;"><img src="http://fiximage.10x10.co.kr/web2019/common/img_red_dot.png" alt=""></span>');
    if(_ClickCnt==1){
        _X1value=x;
        _Y1value=y;
    }
    if(_ClickCnt==2){
        winImg = window.open('pop_event_imagemapSet.asp?mode=reg&masterIdx=<%=masterIdx%>&x1='+_X1value+'&y1='+_Y1value+'&x2='+x+'&y2='+y,'popImg','width=500,height=300');
		winImg.focus();
        _ClickCnt=1;
        _X1value=0;
        _Y1value=0;
    }
    else{
        _ClickCnt=2;
    }
}

function fnMapContentsEdit(didx,x1,y1,x2,y2){
    winImg = window.open('pop_event_imagemapSet.asp?mode=edit&masterIdx=<%=masterIdx%>&didx='+didx+'&x1='+x1+'&y1='+y1+'&x2='+x2+'&y2='+y2,'popImg','width=500,height=300');
    winImg.focus();
}

function fnLinkInfoView(){
    
}

$(function(){
    $('#imgmap').maphilight();
});
</script>
<div class="popV19">
	<div class="popHeadV19">
		<h1>이미지 맵 링크</h1>
        <h3> - 링크를 삽입할 상품위에 커서를 놓고 클릭해주세요.</h3>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <td>
                        <div class="map-wrap">
                            <img src="<%=oLinkContents.FOneItem.Fimage%>" onClick="javascript:getLoc()" id="imgmap" usemap="#plan01">
                        </div>
                        <map name="plan01" id="plan01">
                            <% for i=0 to oLinkDetailContents.FResultCount-1 %>
                            <area shape="rect" coords="<%=oLinkDetailContents.FItemList(i).FXValue%>,<%=oLinkDetailContents.FItemList(i).FYValue%>,<%=oLinkDetailContents.FItemList(i).FWValue%>,<%=oLinkDetailContents.FItemList(i).FHValue%>" href="javascript:fnMapContentsEdit(<%=oLinkDetailContents.FItemList(i).Fidx%>,<%=oLinkDetailContents.FItemList(i).FXValue%>,<%=oLinkDetailContents.FItemList(i).FYValue%>,<%=oLinkDetailContents.FItemList(i).FWValue%>,<%=oLinkDetailContents.FItemList(i).FHValue%>);" title="<%=oLinkDetailContents.FItemList(i).FLinkURL%>" />
                            <% Next %>
                        </map>
                    </td>
                </tr>
			</tbody>
        </table>
    </div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
	</div>
</div>

<%
Set oLinkContents = Nothing
Set oLinkDetailContents = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->