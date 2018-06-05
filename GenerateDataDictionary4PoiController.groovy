package rit.tradso.system

import groovy.sql.Sql
import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import rit.tradso.base.BaseController

/**
 * POI生成数据库的所有表的数据字典，生成到一个Excel中的多个sheet页中，每个表占用一个sheet
 */
class GenerateDataDictionary4PoiController extends BaseController {
    def dataSource
    def sql
    /**
     * 使用POI生成数据字典
     */
    def generateDataDictionary4Poi = {

        long startTime = System.currentTimeMillis()
        sql = new Sql(dataSource)
        //获取数据库连接，执行sql，查询到数据库中的所有表
        String allTableSql = " select TABLE_NAME name  from INFORMATION_SCHEMA.TABLES where TABLE_SCHEMA='tradso-2018-5-14'"
        //得到ArrayList结果，泛型是LinkedHashMap，key是数据库名，值是每个表的名称
        def allTables = sql.rows(allTableSql)
        //遍历每个表，然后通过该表名获取到这个表的数据结构
        eachAllTables(allTables)

        long endTime = System.currentTimeMillis()
        long s=(endTime-startTime)/1000
        render "共耗时："+s+"秒"
    }
    /**
     * 遍历每个表，然后通过该表明获取到这个表的数据结构
     * @param allTables 数据库中所有表
     */
    def eachAllTables(def allTables) {
        //遍历每个数据表时开始创建Excel文件
        HSSFWorkbook workbook = new HSSFWorkbook();
        FileOutputStream fo=new FileOutputStream("d:\\海贸数据字典.xls")
        //遍历每张表名的同时，将对应的.groovy同名文件找到，并存储在Map中，key为表名，键为该文件的绝对路径
        HashMap<String, String> pathMaps = new HashMap<>();
        //应提前将所有对应表的实体类文件找到
        allTables.each{Map map ->
            //每个表的名称
            String tableName = map?.name
            String newTableName = tableName.replace("_", "");
            //查询每个表名对应的java文件
            findJavaFile(pathMaps, newTableName)
        }
        println pathMaps.size()
        allTables.each { Map map ->
            //每个表的名称
            String tableName = map?.name
            //查询每个表名对应的java文件
           // findJavaFile(pathMaps, tableName)
            //执行查询每个表的表结构，根据表名称
            eachOneTables(workbook, fo, pathMaps, tableName)//根据每个表的名称，获取表结构数据
        }
        workbook.write(fo)
        workbook.close()
        fo.close()

    }

    /**
     * 根据每个表的名称，获取表结构数据
     * @param tableName 表名称
     */
    def eachOneTables(HSSFWorkbook workbook, FileOutputStream fo, Map pathMaps, String tableName) {
        String oneTableSql = """
            SELECT
                COLUMN_NAME COLUMN_NAME,
                DATA_TYPE DATA_TYPE,
                IFNULL(CHARACTER_MAXIMUM_LENGTH,0) length,
                case when IS_NULLABLE='no' then '' when IS_NULLABLE='YES' then '√' else 0 end isNull
            FROM
                INFORMATION_SCHEMA. COLUMNS
            WHERE
                table_schema = 'tradso-2018-5-14'
            AND 
                table_name = '${tableName}'
        """
        //得到每张表的数据结构
        def oneTable = sql.rows(oneTableSql)
        String newTableName = tableName.replace("_", "");
        //查询到每个表的数据结构后，就可以构建POI生成excel了
        generateExcel4Poi(workbook, fo, pathMaps, tableName, newTableName, oneTable)
    }

    /**
     * 查询到每个表的数据结构后，就可以构建POI生成excel了
     * @param oneTable 每张表中的所有数据结构
     */
    def generateExcel4Poi(HSSFWorkbook workbook, FileOutputStream fo, Map pathMaps, String tableName, String newTableName,
                          def oneTable) {
        //判断数据库表是否有对应的实体类文件
        if(pathMaps.get(newTableName)!=null){
            //字段描述，从.groovy文件中提取
            String name=pathMaps.get(newTableName)
            FileReader groovyFile=new FileReader(name)
            //读取文件
            BufferedReader fis=new BufferedReader(groovyFile)
            ArrayList<String> list=new ArrayList<>()
            String lines
            while ((lines=fis.readLine())!=null){
                list.add(lines)
            }
            fis.close()//关闭输入流
            println "表名："+newTableName

            //开始创建每个sheet页
            HSSFSheet sheet = workbook.createSheet(tableName)
            HSSFCellStyle setBorder = workbook.createCellStyle();
            setBorder.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
            setBorder.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
            setBorder.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
            setBorder.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
            //设置单元格样式
            sheet.setColumnWidth(5, 16*2 * 256);
            HSSFRow row = sheet.createRow(0);
            row.createCell(0).setCellValue("序号");
            row.createCell(1).setCellValue("列名");
            row.createCell(2).setCellValue("字段类型");
            row.createCell(3).setCellValue("长度");
            row.createCell(4).setCellValue("是否为空");
            row.createCell(5).setCellValue("字段描述");


            oneTable.eachWithIndex{Map map,int i->
                int index=0;
                //第1行
                HSSFRow newRow = sheet.createRow(i+1);
                //第一个单元格
                newRow.createCell(index).setCellValue(i+1)
                newRow.createCell(index+1).setCellValue(map?.COLUMN_NAME)
                newRow.createCell(index+2).setCellValue(map?.DATA_TYPE)
                newRow.createCell(index+3).setCellValue(map?.length)
                newRow.createCell(index+4).setCellValue(map?.isNull)

                String colName=map?.COLUMN_NAME
                colName=colName.replace("_","")
                if(colName.endsWith("id")){
                    colName=colName.substring(0,colName.length()-2)
                }
                println colName
                for(String ll:list){
                    if(ll!=null&&ll.toLowerCase().contains(colName.toLowerCase())){
                        int num=ll.indexOf("//")
                        if(num!=-1) {
                            newRow.createCell(index + 5).setCellValue(ll.substring(num + 2, ll.length()))
                        }
                    }
                }
            }


        }
    }

    /**
     * 传入一个表名，遍历项目下的所有.groovy文件，找到同名文件，将文件里的所有内容保存到Map中
     */
    def findJavaFile(Map map, String tableName) {
        //获取项目路径：递归遍历
        String basePath = grailsApplication.getMainContext().servletContext.getRealPath("/")
        basePath = basePath.replace("\\web-app\\", "")
        File file = new File(basePath)
        getRealFile(map, tableName, file)
    }

    /**
     * 递归遍历和数据表同名.groovy文件
     */
    def getRealFile(Map map, String tableName, File path) {
        File[] files = path.listFiles()
        for (File ff : files) {
            if (ff.isFile() && ff.getName().equalsIgnoreCase(tableName + ".groovy")) {
                map.put(tableName, ff)
            } else if (ff.isDirectory()) {
                getRealFile(map, tableName, ff);
            }
        }
    }


}
