<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/imageLinkCls.asp" -->
<%
Dim masterIdx
	masterIdx = requestCheckvar(request("idx"),16)

    If masterIdx = "" Then
        response.write "<script>alert('정상적인 경로로 접근해 주세요');history.back();</script>"
        response.end
    End If

	dim oLinkContents
		set oLinkContents = new CimageLink
		oLinkContents.FRectIdx = masterIdx
		oLinkContents.GetOneContents

    dim oLinkDetailContents
        set oLinkDetailContents = new CimageLink
        oLinkDetailContents.FRectMasterIdx = masterIdx
        oLinkDetailContents.GetLinkListContents()

%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<style type="text/css">
.map-wrap {position:relative; overflow:hidden; display:inline-block; height:50vh; cursor:pointer;}
.map-wrap > img {width:auto; max-height:100%;}
.map-wrap .mark {position:absolute; left:50%; top:50%; width:30px; transform:translate(-50%,-50%);}
</style>
<script>
$(function(){
	$(".map-wrap").click(function(e){
        setTimeout(() => {
            var posX = Math.round( e.offsetX / $('.map-wrap').width() * 100 );
            var posY = Math.round( e.offsetY / $('.map-wrap').height() * 100 );
            $('.map-wrap').append('<span class="mark" id="markposition_'+posX+posY+'" style="left:' + posX + '%; top:' + posY + '%;"><img src="http://fiximage.10x10.co.kr/web2019/diary2020/ico_mark.png" alt=""></span>');

            if(confirm("이곳에 등록하시겠습니까?")) {
                window.open('popimagelinkdetailedit.asp?masteridx=<%=masterIdx%>&posX='+posX+'&posY='+posY,'imagelinkedit','width=800,height=600,scrollbars=yes,resizable=yes');return false;
            } else {
                $("#markposition_"+posX+posY).remove();
            }
        }, 50);
		//$('.mark').css({"left": posX + "%", "top": posY + "%"});
		//alert(posX +","+ posY);
	});
});
</script>
</head>
<body>
<h2>이미지 상품링크 등록</h2>
<h3> - 링크를 삽입할 상품위에 커서를 놓고 클릭해주세요.</h3>
<div>
    <div class="map-wrap">
        <img src="<%=oLinkContents.FOneItem.Fimage%>" alt="">
        <% 
            Dim i
            for i=0 to oLinkDetailContents.FResultCount-1 
        %>
            <span class="mark" data-value="<%=oLinkDetailContents.FItemList(i).FIdx%>" style="left:<%=oLinkDetailContents.FItemList(i).FXValue%>%; top:<%=oLinkDetailContents.FItemList(i).FYValue%>%;" onclick="window.open('popimagelinkdetailedit.asp?idx=<%=oLinkDetailContents.FItemList(i).FIdx%>&masteridx=<%=masterIdx%>','imagelinkedit','width=800,height=600,scrollbars=yes,resizable=yes');return false;"><img src="http://fiximage.10x10.co.kr/web2019/diary2020/ico_mark.png" alt=""></span>
        <% Next %>
    </div>
    <div style="left:50%"><input type="button" value="목록으로" onclick="location.href='/admin/sitemaster/ImageLinkMap/'"></div>
</div>
</body>
</html>
<%
    Set oLinkContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->