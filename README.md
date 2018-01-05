# ToExcel
.NET 使用NPOI导入导出标准Excel
NPOI是POI项目的.NET版本
它的一些特性：
支持对标准的Excel读写
支持对流(Stream)的读写 (而Jet OLEDB和Office COM都只能针对文件)
支持大部分Office COM组件的常用功能
性能优异 (相对于前面的方法)
使用简单，易上手
但因为是开源组件，还是有一些限制，每个sheet中 最大导出数据量为65535条。
数据量大的可以分sheet,在此demo中就有分sheet的方法。
#excel文件转成流
