package xlsx;

import java.util.LinkedHashMap;

public class ExcelMapper {

  private LinkedHashMap<String, CellType> mapper = new LinkedHashMap<>();

  public ExcelMapper() {}

  public void addMapping(String field, CellType cellType) {
    mapper.put(field, cellType);
  }

  public void addMappingAsDecimal(String field) {
    mapper.put(field, new CellType(ExcelCellType.DECIMAL, "0.0"));
  }

  public void addMappingAsBigDecimal(String field) {
    mapper.put(field, new CellType(ExcelCellType.BIG_DECIMAL, "0.0"));
  }

  public void addMappingAsDate(String field) {
    mapper.put(field, new CellType(ExcelCellType.DATE, "dd/mm/yyyy"));
  }

  public void addMappingAsDateTime(String field) {
    mapper.put(field, new CellType(ExcelCellType.DATETIME, "dd/mm/yyyy hh:mm:ss"));
  }

  public void addMappingAsNumeric(String field) {
    mapper.put(field, new CellType(ExcelCellType.NUMERIC, "0"));
  }

  public void addMappingAsString(String field) {
    mapper.put(field, new CellType(ExcelCellType.STRING, ""));
  }

  /** @return the mapper */
  public LinkedHashMap<String, CellType> getMapper() {
    return mapper;
  }
}
