package twim.excel.util;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Component
public class ExcelUtil {
    public String getCellValue(XSSFCell cell){
        String value = "";

        if(cell == null){
            return value;
        }

        switch (cell.getCellTypeEnum()){
            case STRING:
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                //마지막에 ""을 붙이는 이유는?
                value = (int) cell.getNumericCellValue() + "";
                break;
            default:
                break;
        }
        return value;
    }

    public List<Map<String, Object>> getListData(MultipartFile file, int startRowNum, int columnLength){
        List<Map<String, Object>> excelList = new ArrayList<Map<String, Object>>();

        try{
            //OPCPackage를 사용해서 엑셀파일을 읽는 방법 .open(new File()) 이후 workbook 을 통해 읽어 드립니다.
            OPCPackage opcPackage = OPCPackage.open(file.getInputStream());

            //SuppressWarning 어노테이션을 사용하여 컴파일 단위의 서브세트와 관련된 컴파일 경고를 사용하지 않도록 설정할 수 있다.
            //resource는 닫기 가능 유형의 자원 사용에 관련된 경고 억제
            //출처: https://ktko.tistory.com/entry/Java의-SuppressWarnings-사용하기 [KTKO 개발 블로그와 여행 일기:티스토리]
            @SuppressWarnings("resource")
            XSSFWorkbook workbook = new XSSFWorkbook(opcPackage);

            //읽은 엑셀파일에서 0번째 Sheet를 가져옵니다.
            XSSFSheet sheet = workbook.getSheetAt(0);

            int rowIndex = 0;
            int columnIndex = 0;


            for(rowIndex = startRowNum; rowIndex < sheet.getLastRowNum() + 1; rowIndex++){
                // 엑셀 파일에서 한 행을 가져옵니다.
                XSSFRow row = sheet.getRow(rowIndex);

                if(row.getCell(0) != null && !row.getCell(0).toString().isBlank()){

                    Map<String, Object> map = new HashMap<String, Object>();

                    //속성의 개수를 구한 후
                    int cells = columnLength;

                    // 개수 만큼 한 행에 있는 셀의 값을 받아서 Map에 추가합니다.
                    for(columnIndex = 0; columnIndex <= cells; columnIndex++){
                        XSSFCell cell = row.getCell(columnIndex);
                        map.put(String.valueOf(columnIndex), getCellValue(cell));
                    }

                    // 그리고 행 마다 map으로 만들어서 List에 추가
                    excelList.add(map);
                }
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }

        return excelList;
    }
}
