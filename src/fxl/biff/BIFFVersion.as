package fxl.biff
{
  /**
   * Used to represent the different versions of BIFF files that exist in the wild
   * The following table shows which Excel version writes which file format for worksheet and workbook documents:
   * Excel version | MS Windows | Release | Apple Macintosh | Release | BIFF version | Document type
   *   Excel 2.x   | Excel 2.0  |   1987  | Excel 2.2       |   1989  |    BIFF2     |   Worksheet
   *   Excel 3.0   | Excel 3.0  |   1990  | Excel 3.0       |   1990  |    BIFF3     |   Worksheet
   *   Excel 4.0   | Excel 4.0  |   1992  | Excel 4.0       |   1992  |    BIFF4     |   Worksheet
   *   Excel 5.0   | Excel 5.0  |   1993  | Excel 5.0       |   1993  |    BIFF5     |   Workbook
   *   Excel 7.0   | Excel 95   |   1995  | -               |         |    BIFF5     |   Workbook
   *   Excel 8.0   | Excel 97   |   1997  | Excel 98        |   1998  |    BIFF8     |   Workbook
   *   Excel 9.0   | Excel 2000 |   1999  | Excel 2001      |   2000  |    BIFF8     |   Workbook
   *   Excel 10.0  | Excel XP   |   2001  | Excel v.X       |   2001  |    BIFF8     |   Workbook
   *   Excel 11.0  | Excel 2003 |   2003  | Excel 2004      |   2004  |    BIFF8     |   Workbook
   */
  public class BIFFVersion
  {
    public static const BIFF0:uint = 0;
    public static const BIFF2:uint = 2;
    public static const BIFF3:uint = 3;
    public static const BIFF4:uint = 4;
    public static const BIFF5:uint = 5;
    public static const BIFF8:uint = 8;
  }
}