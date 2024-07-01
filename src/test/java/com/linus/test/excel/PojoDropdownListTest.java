package com.linus.test.excel;

import com.linus.excel.ColumnConfiguration;
import com.linus.excel.PojoSheetWriter;
import com.linus.excel.po.User;
import com.linus.excel.util.ColumnConfigurationParserForPojo;
import com.linus.excel.validation.RangeColumnConstraint;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class PojoDropdownListTest {

  private String[] genderOptions = {"Male", "Female"};

  @Test
  public void test() throws IntrospectionException, IOException {

    Workbook wb = new XSSFWorkbook();
    Sheet sheet = wb.createSheet("Detail");

    ArrayList<ColumnConfiguration> configs = ColumnConfigurationParserForPojo.getColumnConfigurations(
            User.class, Locale.SIMPLIFIED_CHINESE, null);

    for (ColumnConfiguration config : configs) {
      if ("gender".equals(config.getKey())) {
        RangeColumnConstraint rangeColumnConstraint = new RangeColumnConstraint();
        rangeColumnConstraint.setMustInRange(true);
        rangeColumnConstraint.setPickList(genderOptions);
        config.setConstraints(Arrays.asList(rangeColumnConstraint));
        break;
      }
    }

    PojoSheetWriter<User> writer = new PojoSheetWriter<User>(wb, configs);

    writer.writeSheet(wb, sheet, getUserList(), true);

    File file = new File("excel/pojo/dropdown.xlsx");
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
