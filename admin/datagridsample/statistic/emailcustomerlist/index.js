$(function(){
    $("#gridContainer").dxDataGrid({
        dataSource: "datatojson.asp",
        showColumnLines: true, // 컬럼 라인
        showRowLines: true, // 로우 라인
        rowAlternationEnabled: true, // 로우별 회색 색상
        showBorders: true, // 전체 보더
        columnChooser: { // 화면에 보여주는 컬럼 선택
            enabled: true,
            mode: "select" // or "dragAndDrop"
        },
        selection: { // 로우 선택
            mode: "multiple", // or "single" | "none"
            showCheckBoxesMode : "always" // or "onClick" | "onLongTap" | "always"
        },
        "export": { // 엑셀 다운로드 관련
            enabled: true,
            fileName: "EmailCustomerList",
            allowExportSelectedData: true
        },
        filterRow: { // 로우별 검색 사용 
            visible: true,
            applyFilter: "auto"
        },
        headerFilter: { // 컬럼명 깔대기 검색 
            visible: true
        },
        columnAutoWidth: true,
        columns: [{
                    caption : "발송 이름", // 컬럼명
                    dataField : "sendName", // 바인딩 이름
                    width : 270, // 컬럼 너비
                    alignment : "center", // 정렬
                    fixed: true,
                }, {
                    caption : "메일 제목",
                    dataField : "mailTitle",
                    alignment : "center"
                }, {
                    caption : "총 대상자수",
                    dataField : "totalSendUserCount",
                    alignment : "right",
                    format: { // 데이터 표기 방법 
                        type: 'fixedPoint',
                        precision: 0
                    }
                }, {
                    caption : "성공 발송수",
                    dataField : "successSendCount",
                    alignment : "right",
                    format: {
                        type: 'fixedPoint',
                        precision: 0
                    }
                }, {
                    caption : "오픈 통수",
                    dataField : "emailOpenCount",
                    alignment : "right",
                    format: {
                        type: 'fixedPoint',
                        precision: 0
                    }
                }, {
                    caption : "클릭 통수",
                    dataField : "emailClickCount",
                    alignment : "right",
                    format: {
                        type: 'fixedPoint',
                        precision: 0
                    }
                }, {
                    caption : "발송 시간",
                    dataField : "emailSendDate",
                    alignment : "center",
                    dataType : "datetime",
                    format : "yyyy-MM-ddTHH:mm:ss",
                    
                }, {
                    caption : "완료 시간",
                    dataField : "emailSendCompleteDate",
                    alignment : "center",
                    dataType: "datetime",
                    format: "yyyy-MM-ddTHH:mm:ss",
                }, {
                    caption : "ETC",
                    alignment : "center",
                    dataField: "idx",
                    width : 120,
                    allowFiltering : false ,
                    allowSorting : false , 
                    cellTemplate: function(element, dataField) { // sell 커스텀 탬플릿 
                        var url = '/admin/report/mailing_data_reg.asp?idx='+ dataField.value + '&mode=edit';
                        element.append("<div><a href="+ url +">상세내용보기</a></div>")
                               .css("color", "blue");
                    }
                }],
        paging: { enabled: false } // 기본 페이징 유무 (api 유무 상관 없이 default로 분리 됨)
    });
});