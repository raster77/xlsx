package xlsx;

public class CellType {

  private ExcelCellType excelCellType;
  private String format = null;

  public CellType(ExcelCellType excelCellType, String format) {
    this.excelCellType = excelCellType;
    this.format = format;
  }

  /** @return the excelCellType */
  public ExcelCellType getExcelCellType() {
    return excelCellType;
  }

  /** @return the format */
  public String getFormat() {
    return format;
  }
}
