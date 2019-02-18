package xlsx;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
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
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** @author thierry */
public class ExcelExporter implements AutoCloseable {

  private HashMap<String, CellStyle> styles = new HashMap<>();
  private LinkedHashMap<String, Method> methods = new LinkedHashMap<>();
  private List<Integer> sizes = new ArrayList<>();
  private ExcelMapper mapper;
  private List<?> datas;

  private XSSFWorkbook workbook = null;
  private XSSFSheet sheet = null;
  private String sheetTitle = " ";

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

  /** @return the workbook */
  public XSSFWorkbook getWorkbook() {
    return workbook;
  }

  /**
   * Generate Excel file
   *
   * @throws IllegalAccessException
   * @throws IllegalArgumentException
   * @throws InvocationTargetException
   */
  public void generate()
      throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {

    cacheMethodsAndStyles();

    sizes.clear();

    for (Integer i = 0; i < methods.size(); ++i) {
      sizes.add(0);
    }

    for (Integer rowNum = 0; rowNum < datas.size(); ++rowNum) {
      XSSFRow row = sheet.createRow(rowNum);
      Integer cellNum = 0;
      for (Map.Entry<String, Method> entry : methods.entrySet()) {

        Object obj = entry.getValue().invoke(datas.get(rowNum));
        if (obj != null) {

          System.out.println(rowNum + " - " + cellNum.toString() + " = " + obj.toString());
          Cell cell = row.createCell(cellNum);
          CellStyle cellStyle = styles.get(mapper.getMapper().get(entry.getKey()).getFormat());
          cell.setCellStyle(cellStyle);

          switch (mapper.getMapper().get(entry.getKey()).getExcelCellType()) {
            case DATE:
            case DATETIME:
              cell.setCellValue((Date) obj);
              break;
            case BIG_DECIMAL:
              cell.setCellValue(((BigDecimal) obj).doubleValue());
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

          Integer length = new DataFormatter().formatCellValue(cell).length();

          if (length > sizes.get(cellNum)) {
            sizes.set(cellNum, length);
          }
        }
        cellNum++;
      }
    }
    ;

    for (Integer i = 0; i < sizes.size(); ++i) {
      int width = (int) ((sizes.get(i) + 2.53) * 256);
      sheet.setColumnWidth(i, width);
    }
  }

  /** caching methods and cells styles to speed up the generation */
  private void cacheMethodsAndStyles() {

    if (methods.isEmpty() && !datas.isEmpty()) {

      Class<?> clazz = datas.get(0).getClass();
      CreationHelper createHelper = workbook.getCreationHelper();
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
    }
  }

  /** Init the workbook, the default sheet and the createHelper */
  private void createWorkbook() {
    if (workbook == null) {
      workbook = new XSSFWorkbook();
      sheet = workbook.createSheet(sheetTitle);
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

  @Override
  public void close() throws Exception {
    System.out.println("In close");
    if (workbook != null) {
      System.out.println("workbook exists, closing it");
      workbook.close();
    }
  }
}
