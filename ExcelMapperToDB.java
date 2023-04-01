package com.utils;

import com.model.MailContactEntity;
import com.model.ErrorEntity;
import com.model.SMSEntity;
import com.repository.MailContactRepository;
import com.repository.ErrorRepository;
import com.repository.SMSsRepository;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

@Service
public class ExcelHelper {

    private final ErrorRepository errorRepository;
    private final SMSsRepository smssRepository;
    private final MailContactRepository mailContactRepository;

    public ExcelHelper(ErrorRepository errorRepository, smssRepository smsRepository, MailContactRepository mailContactRepository) {
        this.errorRepository = errorRepository;
        this.smssRepository = smsRepository;
        this.mailContactRepository = mailContactRepository;
    }


    public void readExcel(String fileLocation) throws IOException {
        FileInputStream file = new FileInputStream(fileLocation);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);

        List<ErrorEntity> errorEntityList = new ArrayList<>();
        List<SMSEntity> smsEntityList = new ArrayList<>();
        List<MailContactEntity> mailContactEntityList = new ArrayList<>();

        int rows = sheet.getLastRowNum();
        int columns = sheet.getRow(2).getLastCellNum() - 1;

        for(int r=0; r<=rows; r++) {
            XSSFRow row = sheet.getRow(r);
            System.out.println("Row: " + r + 1);

            if(r != 0) {
                ErrorEntity errorEntity = new ErrorEntity();
                SMSEntity smsEntity = new SMSEntity();
                List<MailContactEntity> mailContactEntities = new ArrayList<>();


                for (int c = 0; c <= columns; c++) {
                    XSSFCell cell = row.getCell(c);
                    System.out.println("Cell: " + c + 1);

                    if (cell == null) {
                        System.out.println("Skipping cell");

                    } else {
                        switch (cell.getCellType()) {
                            case STRING:
                                assignErrorsValues(cell, errorEntity, c);
                                assignSMSValues(cell, smsEntity, c);
                                assignMailValues(cell, mailContactEntities, c);
                                break;
                            case NUMERIC:
                                System.out.print("NUMERIC");
                            case ERROR:
                                System.out.println("ERROR");
                            case BLANK:
                                System.out.println("BLANK, skipping");
                            case FORMULA:
                                System.out.println("FORMULA");
                        }
                    }
                }
                System.out.println("==========");
                srrorEntityList.add(errorEntity);
                smsEntityList.add(sMSEntity);
                mailContactEntityList.addAll(mailContactEntities);
            }
        }
        smssRepository.deleteAll();
        mailContactRepository.deleteAll();
        errorRepository.deleteAll();

        ErrorEntityList =
                ErrorEntityList.stream().filter(entity -> entity.getErrorCode() != null).collect(Collectors.toList());
        SMSEntityList =
                SMSEntityList.stream().filter(entity -> entity.getErrorCode() != null).collect(Collectors.toList());
        MailContactEntityList =
                MailContactEntityList.stream().filter(entity -> entity.getErrorCode() != null).collect(Collectors.toList());
        ErrorRepository.saveAll(ErrorEntityList);
        SMSsRepository.saveAll(SMSEntityList);
        MailContactRepository.saveAll(MailContactEntityList);
    }

    private void assignErrorsValues(XSSFCell cell, ErrorEntity entity, int cellNumber) {
        if (cellNumber == 6) {
            entity.setFlow(cell.getStringCellValue());
        } else if (cellNumber == 8) {
            entity.setTestProcess(cell.getStringCellValue());
        } else if (cellNumber == 11) {
            entity.setCode(cell.getStringCellValue());
        } else if (cellNumber == 15) {
            entity.setDescription1(cell.getStringCellValue());
        } else if (cellNumber == 16) {
            entity.setDescription2(cell.getStringCellValue());
        } else if (cellNumber == 17) {
            entity.setEmailSubject(cell.getStringCellValue());
        } else if (cellNumber == 18) {
            entity.setSeverityCode(cell.getStringCellValue());
        } else if (cellNumber == 19) {
            entity.setDescription3(cell.getStringCellValue());
        } else if (cellNumber == 21) {
            entity.setResolution(cell.getStringCellValue());
        } else if (cellNumber == 22) {
            entity.setResolution2(cell.getStringCellValue());
        }
    }

    private void smsValues(XSSFCell cell, SMSEntity entity, int cellNumber) {
        if (cellNumber == 3) {
            entity.setContact(cell.getStringCellValue());
        } else if (cellNumber == 11) {
            entity.setCode(cell.getStringCellValue());
        }
    }


    private void assignEmailValues(XSSFCell cell, List<MailContactEntity> list, int cellNumber) {

        if(cellNumber == 2) {
            String cellValue = cell.getStringCellValue();
            List<String> listOfEmails = Arrays.asList(cellValue.split("; "));
            listOfEmails.forEach(email -> {
                MailContactEntity entity2 = new MailContactEntity();
                entity2.setContact(email);
                list.add(entity2);
            });
        } else if (cellNumber == 11) {
            list.forEach(emailContactEntity -> {
                emailContactEntity.setCode(cell.getStringCellValue());
            });
        }
    }
}
