$(function(){
    var orders = new DevExpress.data.CustomStore({
        load: function (loadOptions) {
            var deferred = $.Deferred(),
                args = {};
    
            if (loadOptions.sort) {
                args.orderby = loadOptions.sort[0].selector;
                if (loadOptions.sort[0].desc)
                    args.orderby += " desc";
            }
    
            args.skip = loadOptions.skip;
            args.take = loadOptions.take;
    
            $.ajax({
                url: "datatojson.asp",
                dataType: "json", 
                data: args,
                success: function(result) {
                    deferred.resolve(result.items, { totalCount: result.totalCount });
                },
                error: function() {
                    deferred.reject("Data Loading Error");
                },
                timeout: 5000
            });
    
            return deferred.promise();
        },
        insert: function (values) {
            return $.ajax({
                url: "datatojson.asp",
                dataType: "json",
                method : "POST",
                data: JSON.stringify({
                    "mode" : "POST",
                    "part_name" : values.part_name,
                    "part_sort" : values.part_sort,
                })
            })
        },
        update : function (key, values) {
            return $.ajax({
                url : "datatojson.asp",
                dataType: 'json',
                method : "POST",
                data : JSON.stringify({
                    "mode" : "PUT",
                    "part_sn" : key.part_sn,
                    "part_name" : values.part_name,
                    "part_sort" : values.part_sort,
                }),
            })
        },
        remove : function(key) {
            return $.ajax({
                url : "datatojson.asp",
                dataType: 'json',
                method : "POST",
                data : JSON.stringify({
                    "mode" : "DELETE",
                    "part_sn" : key.part_sn,
                }),
            })
        }

    });
    
    $("#gridContainer").dxDataGrid({
        dataSource: {
            store: orders
        },
        showColumnLines: true, // 컬럼 라인
        showRowLines: true, // 로우 라인
        rowAlternationEnabled: true, // 로우별 회색 색상
        showBorders: true, // 전체 보더
        filterRow: { // 로우별 검색 사용 
            visible: true,
            applyFilter: "auto"
        },
        headerFilter: { // 컬럼명 깔대기 검색 
            visible: true
        },
        scrolling: {
            mode: "virtual"
        },
        height: 600,
        columnAutoWidth: true,
        //remoteOperations: false, //데이터 한번이 읽어 온다는 뜻 - 전체 불러오는 쿼리가 필요함
        // remoteOperations: true, //자체 페이징 그루핑 사용 유무 true로 바꿀 경우 데이터를 계속 읽어옴 아닌경우 가지고온 데이터 내부에서 처리
        remoteOperations: {
            sorting: true, // 소팅 사용 -- 소팅용 쿼리를 만들어서 날려야됨
            paging: true, // 페이징 사용 -- 페이징용 쿼리를 만들어서 날려야됨
            //groupPaging: true // 그룹 페이징 사용 -- 그룹용 쿼리를 만들어서 날려야됨 : false // 페이징 쿼리를 사용 안하므로 전체 불러오는 쿼리가 따로 있어야됨
        },
        editing: { // 수정
            allowAdding: true,
            allowUpdating: true,
            allowDeleting: true,
            texts : { // 수정 텍스트 
                editRow : "Edit", // 수정 Row
                deleteRow : "UseYN", // 삭제 Row
                confirmDeleteMessage : "사용 여부를 수정 하시겠습니까?", // 삭제 Row message
            }
        },
        paging : {
            pageSize : 20 // paging size 지정 가능 
        },
        columns: [{
                    caption : "부서 번호", // 컬럼명
                    dataField : "part_sn", // 바인딩 이름
                    width : 270, // 컬럼 너비
                    alignment : "center", // 정렬
                    allowEditing : false, // 열 입력 수정 안되도록 막는 코드 (insert 할 항목이 아닐 경우)
                }, {
                    caption : "부서명",
                    dataField : "part_name",
                    alignment : "center",
                    validationRules: [{ 
                        type : "custom",
                        validationCallback: validationCallback,
                        message : "부서명 입력이 필요 합니다."
                    }]
                }, {
                    caption : "정렬번호",
                    dataField : "part_sort",
                    alignment : "right",
                    validationRules: [{ 
                        type : "custom",
                        validationCallback: validationCallback,
                        message : "정렬번호 입력이 필요 합니다."
                    }]
                }, {
                    caption : "사용여부",
                    dataField : "part_isDel",
                    alignment : "right",
                    allowEditing : false, // 열 입력 수정 안되도록 막는 코드 (insert 할 항목이 아닐 경우)
                }
        ],
    }).dxDataGrid("instance");
});

function validationCallback(e) {
    return e.value;
}
