package xlsx;

import java.io.Serializable;
import java.util.Date;

public class TestDto implements Serializable {

  /** */
  private static final long serialVersionUID = 1L;

  private String string;
  private Integer integer;
  private Date date;
  private Double lng;

  public TestDto() {}

  /** @return the string */
  public String getString() {
    return string;
  }

  /** @param string the string to set */
  public void setString(String string) {
    this.string = string;
  }

  /** @return the integer */
  public Integer getInteger() {
    return integer;
  }

  /** @param integer the integer to set */
  public void setInteger(Integer integer) {
    this.integer = integer;
  }

  /** @return the date */
  public Date getDate() {
    return date;
  }

  /** @param date the date to set */
  public void setDate(Date date) {
    this.date = date;
  }

  /** @return the lng */
  public Double getLng() {
    return lng;
  }

  /** @param lng the lng to set */
  public void setLng(Double lng) {
    this.lng = lng;
  }
}
