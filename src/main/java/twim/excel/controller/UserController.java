package twim.excel.controller;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import twim.excel.entity.User;
import twim.excel.repository.UserRepository;
import twim.excel.util.ExcelUtil;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.channels.FileLock;
import java.util.*;

@RestController
@RequiredArgsConstructor
public class UserController {
    private final ExcelUtil excelUtil;
    private final UserRepository userRepository;

    @PostMapping("/add/excel")
    public Map<String, String> addExcel(@RequestParam("file") MultipartFile file){
        Map<String, String> response = new HashMap<>();

        if(file.isEmpty()){
            response.put("message", "빈 파일 입니다.");
            return response;
        }

        String contentType = file.getContentType();

        System.out.println(contentType);

        if(!contentType.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")){
            response.put("message", "파일의 확장자를 확인해주세요.");
            return response;
        }

        List<User> listUser = new ArrayList<User>();

        List<Map<String, Object>> listMap = excelUtil.getListData(file, 1, 3);

        for (Map<String, Object> userMap : listMap) {
            User user = new User(userMap.get("0").toString(), userMap.get("1").toString(), userMap.get("2").toString());

            userRepository.save(user);
        }

        response.put("message", "엑셀 파일 내 사용자 입력 성공");
        return response;
    }

    @GetMapping("/user")
    public List<User> findAllUser(){
        return userRepository.findAll();
    }

    @GetMapping("/download/excel")
    public void queryToExcelFile() throws IOException {
        Workbook workbook = null;
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;

        workbook = new SXSSFWorkbook(); //엑셀 파일 생성
        sheet = workbook.createSheet("sheet1"); //시트 생성

        List<User> list = userRepository.findAll();
        File file = new File("C:\\Users\\TWIM\\Pictures\\excel\\query.xlsx");
        FileOutputStream fos = new FileOutputStream(file);


        row = sheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("username");
        cell = row.createCell(1);
        cell.setCellValue("age");
        cell = row.createCell(2);
        cell.setCellValue("tel");

        int curRow = 1;

        for (int i = 0; i < list.size(); i++) {
            row = sheet.createRow(curRow);

            cell = row.createCell(0);
            cell.setCellValue(list.get(i).getUsername());
            cell = row.createCell(1);
            cell.setCellValue(list.get(i).getAge());
            cell = row.createCell(2);
            cell.setCellValue(list.get(i).getTel());

            curRow++;
        }

        workbook.write(fos);
        if(fos != null) {
            fos.close();
        }
    }
}
