

require.config({
    paths: {
        "jquery": "jquery/jquery-1.9.0.min",
        "base": "ligerUI/js/core/base",
        "ligerui": "ligerUI/js/ligerui.min",
        "ligerGrid": "ligerUI/js/plugins/ligerGrid",
        "ligerResizable": "ligerUI/js/plugins/ligerResizable",
        "ligerDrag": "ligerUI/js/plugins/ligerDrag",
        "table2excel": "excel/table2excel",
        "FileSaver": "excel/FileSaver.min",
        "tableexport": "excel/tableexport",
        "jszip": "excel/jszip",
        "xlsx": "excel/xlsx",
        "xlsxloader": 'excel/xlsx-loader',
        "CustomersData": "CustomersData"
    },
    shim: {
        'ligerui': ['jquery'],
        'ligerGrid': ['jquery', 'ligerui'],
        'ligerResizable': ['jquery', 'ligerui'],
        'ligerDrag': ['jquery', 'ligerui'],
        'tableexport': ['xlsx'],
        xlsx: {
            exports: 'XLSX',
            deps: ['xlsxloader']
        },
        table2excel: {
            exports: 'table2excel'
        }
    }
});


define(['jquery', 'ligerui', 'tableexport'], function ($, CustomersData) {


    function tableExcelByTable() {
        $("#maingrid").tableExport({
            title: "",
            fileName: "111",
            islageruigrid: false, // 是否基于ligerui
            ligerui: {
                isGetAll: false,
                gridInstance: null
            },
            css: ["/css/ligerui_grid_export_excel.css"],

            //fileType: "xls",
            frozen: {c: 1 , r:2}
        });
    }

    return {
        table2excel: tableExcelByTable
    };
});