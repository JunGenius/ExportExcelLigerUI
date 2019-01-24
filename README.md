# ExportExcelLigerUI

       $("#").tableExport({
                  title: "excel标题名称",
                  fileName: "excel导出文件名称",
                  fileType: "xlsx",  // xls格式:导出使用table2excel.js (需传入css样式,否则无样式,默认导出xlsx)
                  islageruigrid: true, // 是否基于ligerui
                  ligerui: {
                      isGetAll: true, // 是否全部导出数据
                      gridInstance: g, // ligerGrid实例
                      serialNumberWidth: "50"
                  },
                  frozen: { c: 1, r: 2 } //冻结列 (c:列 r:行)
              });
