//$(function() {
//    // 차트를 동적으로 로딩하기 때문에 ready 가 호출된 이후부터 차트를 사용할수 있음.
//    readyChart(function() {
//        // 여기 부터 차트 함수들을 사용할 수 있음.
//        drawSampleChart($('#inputApiUrl').val());
//    });
//
//    function hook(row) {
//        row[0] = new Date(row[0]);
//        return row;
//    }
//
//    function drawSampleChart(url) {
//        var containers = ['#container1', '#container2', '#container3'];
//        _.each(containers, function(c) {
//            $(c).html('');
//        });
//
//        loadWithLocalProxy(url, function(data, error) {
//            if ( data.response == 'error' ) {
//                alert(data.errmsg);
//                return;
//            }
//            if ( data ) {
//                // converter로 데이터를 차트에 맞는 데이터로 변경함.
//                // 구글 차트의 특성에 맞게 chart가 삽입될 엘리먼트 id를 넘김.
//                // 데이터 조작을 하기 때문에 딥카피를 해서 데이터를 다뤄야 함.
//                drawRaw(data, 'container1');
//                drawGoogleChartLine(convertDataForGoogleChartLine(data, [0, 1, 3, 5]).dataTable, 'container2');
//                drawGoogleChartTable(convertDataForGoogleChartTable(data).dataTable, 'container3');
//                drawGoogleChartPie(convertDataForGoogleChartPie(data, [0, 1]).dataTable, 'container4');
//            } else {
//                console.log(error);
//            }
//        });
//    }
//    var api = {
//        gadata: 'http://wapi.10x10.co.kr/anal/getque.asp?kind=gadata&startdate=2015-11-01&enddate=2016-01-14&dimensions=date&param1=33744458&param2=all',
//        getque: 'http://wapi.10x10.co.kr/anal/getque.asp?kind=bestseller&startdate=2015-11-01&enddate=2016-01-14&dimensions=date',
//        pweek: 'http://wapi.10x10.co.kr/anal/getque.asp?kind=gadata&startdate=2015-11-01&enddate=2016-01-14&dimensions=date&param1=33744458&param2=sessions&pretype=pweek'
//    };
//
//    // test 용
//    function onChangedSelApiUrl() {
//        var selApiUrl = $('#selApiUrl').val();
//        $('#inputApiUrl').val(api[selApiUrl]);
//    }
//
//
//    $('#selApiUrl').change(function() {
//        onChangedSelApiUrl();
//    });
//
//    $('#btnOk').click(function() {
//        drawSampleChart($('#inputApiUrl').val());
//    });
//
//    onChangedSelApiUrl();
//});
