package xlsx;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Main {

  public Main() {}

  public static void main(String[] args) throws Exception {

    final String file = "test.xlsx";

    List<TestDto> datas = new ArrayList<>();

    for (Integer i = 0; i < 10; ++i) {
      TestDto dto = new TestDto();
      dto.setDate(new Date());
      dto.setInteger(i);
      if (i != 5) {
        dto.setLng(i * 1.659);
      }
      dto.setBigDec(BigDecimal.valueOf(i + 1).multiply(BigDecimal.valueOf(2.321)));
      dto.setString("str");
      datas.add(dto);
    }

    ExcelMapper mapper = new ExcelMapper();
    mapper.addMappingAsString("string");
    mapper.addMappingAsNumeric("integer");
    mapper.addMappingAsDateTime("date");
    mapper.addMapping("lng", new CellType(ExcelCellType.DECIMAL, "000.00"));
    mapper.addMappingAsBigDecimal("BigDec");

    if (Files.exists(Paths.get(file))) {
      Files.delete(Paths.get(file));
    }

    try (ExcelExporter exporter = new ExcelExporter(mapper, datas)) {
      exporter.generate();
      try (OutputStream outputStream = new FileOutputStream(file)) {
        exporter.getWorkbook().write(outputStream);
      }
    }
  }
}
