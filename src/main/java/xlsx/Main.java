package xlsx;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Main {

  public Main() {}

  public static void main(String[] args) {
    List<TestDto> datas = new ArrayList<>();

    for (Integer i = 0; i < 10; ++i) {
      TestDto dto = new TestDto();
      dto.setDate(new Date());
      dto.setInteger(i);
      if (i != 5) dto.setLng(i * 1.659);
      dto.setString("str");
      datas.add(dto);
    }

    ExcelMapper mapper = new ExcelMapper();
    mapper.addMappingAsString("string");
    mapper.addMappingAsNumeric("integer");
    mapper.addMappingAsDateTime("date");
    mapper.addMapping("lng", new CellType(ExcelCellType.DECIMAL, "000.00"));

    ExcelExporter exporter = new ExcelExporter(mapper, datas);
    exporter.generate();
  }
}
