一、需要导入的jar

　　1.commons-collections4-4.1.jar

　　2.poi-3.17-beta1.jar

　　3.poi-ooxml-3.17-beta1.jar

　　4.poi-ooxml-schemas-3.17-beta1.jar

　　5.xmlbeans-2.6.0.jar

二、主要API

　　1.import org.apache.poi.ss.usermodel.Workbook,对应Excel文档；

　　2.import org.apache.poi.hssf.usermodel.HSSFWorkbook，对应xls格式的Excel文档；

　　3.import org.apache.poi.xssf.usermodel.XSSFWorkbook，对应xlsx格式的Excel文档；

　　4.import org.apache.poi.ss.usermodel.Sheet，对应Excel文档中的一个sheet；

　　5.import org.apache.poi.ss.usermodel.Row，对应一个sheet中的一行；

　　6.import org.apache.poi.ss.usermodel.Cell，对应一个单元格