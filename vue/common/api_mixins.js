const api_mixin = Vue.mixin({
    data() {return {
        href : unescape(location.href),
    }},
    computed : {
        isDevelop() {
            return this.href.includes('//localhost') || this.href.includes('//testwebadmin') || this.href.includes('//localwebadmin');
        },
        isStaging() {
            return this.href.includes('//stgwebadmin');
        },
        isProduction() {
            return this.href.includes('//webadmin');
        },
        apiUrl() {
            let apiUrl;
            if( this.isDevelop )
                apiUrl = '//localfapi.10x10.co.kr:8080'
                //apiUrl = '//testfapi.10x10.co.kr'
            else if( this.isStaging )
                apiUrl = '//stgfapi.10x10.co.kr'
            else
                apiUrl = '//fapi.10x10.co.kr'
            return apiUrl + '/api/admin';
        },
        parameters() {
            if( location.search.length === 0 )
                return {};
            
            const parameters = {};
            const parameterArr = location.search.substr(1).split('&');
            parameterArr.forEach(p => {
                const arr = p.split('=');
                parameters[arr[0]] = arr[1];
            });
            return parameters;
        },
    },
    methods : {
        //region callApi Api 호출
        callApi(v, type, uri, data, success_callback, error_callback) {
            if (error_callback === undefined) {
                error_callback = function (xhr) {
                    console.log(xhr);
                    try {
                        let message = JSON.parse(xhr.responseText).message;
                        alert(message ? message : '에러가 발생했습니다.');
                    } catch(e) {
                        alert('에러가 발생했습니다.');
                    }
                };
            }

            $.ajax({
                type: type,
                url: `${this.apiUrl}/v${v}${uri}`,
                data: data,
                crossDomain: true,
                xhrFields: {
                    withCredentials: true,
                },
                success: success_callback,
                error: error_callback,
            });
        },
        //endregion
        //region getLocalDateTimeFormat Localdatetime 포맷 -> 입력한 포맷으로 수정
        getLocalDateTimeFormat(date, format) {
            if (!date.valueOf()) return "";
            const d = new Date(date);

            const weekName = ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"];

            String.prototype.string = function(len){let s = '', i = 0; while (i++ < len) { s += this; } return s;};
            String.prototype.zf = function(len){return "0".string(len - this.length) + this;};
            Number.prototype.zf = function(len){return this.toString().zf(len);};

            return format.replace(/(yyyy|yy|MM|dd|E|hh|mm|ss|a\/p)/gi, function($1) {
                switch ($1) {
                    case "yyyy": return d.getFullYear();
                    case "yy": return (d.getFullYear() % 1000).zf(2);
                    case "MM": return (d.getMonth() + 1).zf(2);
                    case "dd": return d.getDate().zf(2);
                    case "E": return weekName[d.getDay()];
                    case "HH": return d.getHours().zf(2);
                    case "hh": return ((h = d.getHours() % 12) ? h : 12).zf(2);
                    case "mm": return d.getMinutes().zf(2);
                    case "ss": return d.getSeconds().zf(2);
                    case "a/p": return d.getHours() < 12 ? "오전" : "오후";
                    default: return $1;
                }
            });
        },
        //endregion
        //region decodeBase64 Base64 디코딩
        decodeBase64(str) {
            if( str == null ) return null;
            return atob(str.replace(/_/g, '/').replace(/-/g, '+'));
        },
        //endregion
    }
});