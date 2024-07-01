package com.linus.test.excel;

import com.linus.excel.ColumnConfiguration;
import com.linus.excel.PojoSheetWriter;
import com.linus.excel.po.User;
import com.linus.excel.util.ColumnConfigurationParserForPojo;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class BigDropdownListTest {

  private String[] options = {"EF/OC",
          "WINIT",
          "GOODCANG",
          "4PX",
          "ARMLOGI",
          "AYBASES",
          "CATHY LOGISTIC",
          "CHINA POST",
          "Chukou1",
          "CONTINENTAL",
          "DEWELL",
          "DISCOVERY",
          "ECOF",
          "EDAEU",
          "EDAYUN",
          "EUCASHBOX",
          "GENIQUA",
          "GIGA",
          "GOODCHAINS",
          "JW",
          "LECANGS",
          "LLP",
          "MBB",
          "NEWEGG",
          "PANEXWD",
          "PRIMEROAD",
          "QXBOX",
          "RYLH",
          "SENDEX",
          "SUNWARD",
          "WEDO SCM",
          "WESTERN POST",
          "OTHERS",
          "JDL",
          "JINDA",
          "Heloo",
          "World"};

  @Test
  public void test() throws IntrospectionException, IOException {

    Workbook wb = new XSSFWorkbook();
    Sheet sheet = wb.createSheet("Detail");

    ArrayList<ColumnConfiguration> configs = ColumnConfigurationParserForPojo.getColumnConfigurations(
            User.class, Locale.SIMPLIFIED_CHINESE, null);

    PojoSheetWriter<User> writer = new PojoSheetWriter<User>(wb, configs);

    writer.createOptions(Arrays.asList(options), "warehouse");

    writer.writeSheet(wb, sheet, getUserList(), true);
    writer.createDropdown(wb, sheet, 1, "warehouse");

    File file = new File("excel/pojo/bigdropdown.xlsx");
    if (!file.exists()) {
      file.createNewFile();
    }
    FileOutputStream fos = new FileOutputStream(file);
    wb.write(fos);
  }

  public List<User> getUserList() {
    List<User> list = new ArrayList<User>(1);

    User u = new User();
    u.setFirstName("Linus");
    u.setAge(18);
    u.setBirthday(new Date());
    u.setEmail("linus.yan@hotmail.com");

    list.add(u);

    return list;
  }
}
