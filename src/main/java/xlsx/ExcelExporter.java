package xlsx;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.BiPredicate;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelExporter {

  HashMap<String, CellStyle> styles = new HashMap<>();
  LinkedHashMap<String, Method> methods = new LinkedHashMap<>();
  ExcelMapper mapper;
  List<?> datas;
  String sheetTitle = " ";

  XSSFWorkbook workbook = null;
  XSSFSheet sheet = null;
  CreationHelper createHelper = null;

  public ExcelExporter(ExcelMapper mapper, List<?> datas) {
    this.mapper = mapper;
    this.datas = datas;
    createWorkbook();
  }

  public ExcelExporter(ExcelMapper mapper, List<?> datas, String sheetTitle) {
    this.mapper = mapper;
    this.datas = datas;
    this.sheetTitle = sheetTitle;
    createWorkbook();
  }

  /** @return the sheetTitle */
  public String getSheetTitle() {
    return sheetTitle;
  }

  /** @param sheetTitle the sheetTitle to set */
  public void setSheetTitle(String sheetTitle) {
    this.sheetTitle = sheetTitle;
  }

  private void createWorkbook() {
    workbook = new XSSFWorkbook();
    sheet = workbook.createSheet(sheetTitle);
    createHelper = workbook.getCreationHelper();
  }

  public void generate() {
    Class<?> clazz = datas.get(0).getClass();
    mapper
        .getMapper()
        .entrySet()
        .stream()
        .forEach(
            entry -> {
              String methodName = "get" + capitalize(entry.getKey());
              Method method = getMethod(clazz, methodName);
              if (method != null) {
                methods.put(entry.getKey(), method);

                String format = entry.getValue().getFormat();

                if (format != null) {
                  styles.put(format, workbook.createCellStyle());
                  styles
                      .get(format)
                      .setDataFormat(createHelper.createDataFormat().getFormat(format));
                }
              }
            });

    Integer rowNum = 0;

    for (Integer d = 0; d < datas.size(); ++d) {
      XSSFRow row = sheet.createRow(rowNum);
      Integer cellNum = 0;
      for (Map.Entry<String, Method> entry : methods.entrySet()) {

        try {
          Object obj = entry.getValue().invoke(datas.get(d));
          if (obj != null) {
            System.out.println(rowNum + " - " + cellNum.toString() + " = " + obj.toString());
            Cell cell = row.createCell(cellNum);
            CellStyle cellStyle = styles.get(mapper.getMapper().get(entry.getKey()).getFormat());
            System.out.println("Style: " + cellStyle.getDataFormatString());
            cell.setCellStyle(cellStyle);

            switch (mapper.getMapper().get(entry.getKey()).getExcelCellType()) {
              case DATE:
              case DATETIME:
                cell.setCellValue((Date) obj);
                break;
              case DECIMAL:
                cell.setCellValue((Double) obj);
                break;
              case NUMERIC:
                cell.setCellValue((Integer) obj);
                break;
              default:
                cell.setCellValue(obj.toString());
                break;
            }
          }
          cellNum++;

        } catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
          // TODO Auto-generated catch block
          e.printStackTrace();
        }
      }
      rowNum++;
    }
    ;

    try (FileOutputStream file = new FileOutputStream(new File("workbook.xlsx"))) {
      workbook.write(file);
      workbook.close();
    } catch (FileNotFoundException e) { // TODO Auto-generated catch block
      e.printStackTrace();
    } catch (IOException e) { // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }

  private Method getMethod(Class<?> aClass, String methodName) {
    BiPredicate<Method, String> areMethodsEquals =
        (m, s) -> {
          return s.equals(m.getName());
        };

    Optional<Method> method =
        Arrays.stream(aClass.getMethods())
            .filter(m -> areMethodsEquals.test(m, methodName))
            .collect(Collectors.reducing((a, b) -> null));

    return method.isPresent() ? method.get() : null;
  }

  private String capitalize(String s) {
    return Character.toUpperCase(s.charAt(0)) + s.substring(1);
  }
}
