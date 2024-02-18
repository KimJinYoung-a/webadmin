Vue.component('ITEM-NAME-PREFIX-RESULT-ITEM', {
    template : `
        <tr>
            <td>{{prefix.prefixIdx}}</td>
            <td><a @click="$emit('updatePrefix', prefix)">{{prefix.prefixWord}}</a></td>
            <td>{{getLocalDateTimeFormat(prefix.startDate, 'yyyy-MM-dd HH:mm:ss')}}</td>
            <td>{{getLocalDateTimeFormat(prefix.endDate, 'yyyy-MM-dd HH:mm:ss')}}</td>
            <td>{{prefix.state}}</td>
            <td>{{numberFormat(prefix.itemCount)}}</td>
            <td>{{getLocalDateTimeFormat(prefix.regDate, 'yyyy-MM-dd HH:mm')}}</td>
            <td>{{prefix.regAdminName}}</td>
            <td><button @click="$emit('manageProduct', prefix.prefixIdx)" class="btn">惑前包府</button></td>
        </tr>
    `,
    props : {
        //region prefix 富赣府
        prefix : {
            prefixIdx : { type:Number, default:0 },
            prefixWord : { type:String, default:'' },
            startDate : { type:String, default:'' },
            endDate : { type:String, default:'' },
            use : { type:String, default:'' },
            state : { type:String, default:'' },
            itemCount : { type:Number, default:0 },
            regDate : { type:String, default:'' },
            regAdminId : { type:String, default:'' },
            regAdminName : { type:String, default:'' }
        },
        //endregion
    },
    methods : {
        //region numberFormat 箭磊 玫磊府 (,) format
        numberFormat(num){
            if( num == null )
                return '';

            num = num.toString();
            return num.replace(/(\d)(?=(?:\d{3})+(?!\d))/g,'$1,');
        },
        //endregion
    }
});